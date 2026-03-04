/**
 * MCP Server — Employee Training
 *
 * Registers the training-media tool and its UI resource.
 * The tool recommends AI/ML courses with embedded MS Learn video previews.
 */
import {
  registerAppResource,
  registerAppTool,
  RESOURCE_MIME_TYPE,
} from "@modelcontextprotocol/ext-apps/server";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { CallToolResult, ReadResourceResult } from "@modelcontextprotocol/sdk/types.js";
import fs from "node:fs/promises";
import path from "node:path";

import { z } from "zod";

import { trainingMediaData } from "./mock-data/training-media.js";

// Works both from source (server.ts) and compiled (dist/server.js)
const DIST_DIR = import.meta.filename.endsWith(".ts")
  ? path.join(import.meta.dirname, "dist", "ui")
  : path.join(import.meta.dirname, "ui");

/**
 * Helper to register a tool + its corresponding resource in one call.
 */
function registerToolWithUI<T extends z.ZodRawShape>(
  server: McpServer,
  name: string,
  title: string,
  description: string,
  resourceFileName: string,
  inputSchema: T,
  handler: (args: z.infer<z.ZodObject<T>>) => Promise<CallToolResult> | CallToolResult,
  uiMeta?: {
    csp?: {
      connectDomains?: string[];
      resourceDomains?: string[];
      frameDomains?: string[];
    };
  },
) {
  const resourceUri = `ui://${name}/${resourceFileName}`;

  registerAppTool(
    server,
    name,
    {
      title,
      description,
      inputSchema,
      annotations: { readOnlyHint: true },
      _meta: { ui: { resourceUri } },
    },
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    handler as any,
  );

  registerAppResource(
    server,
    resourceUri,
    resourceUri,
    { mimeType: RESOURCE_MIME_TYPE, description: title },
    async (): Promise<ReadResourceResult> => {
      const html = await fs.readFile(
        path.join(DIST_DIR, resourceFileName),
        "utf-8",
      );
      return {
        contents: [
          {
            uri: resourceUri,
            mimeType: RESOURCE_MIME_TYPE,
            text: html,
            ...(uiMeta ? { _meta: { ui: uiMeta } } : {}),
          },
        ],
      };
    },
  );
}

/**
 * Creates a new MCP server instance with all tools and resources registered.
 */
export function createServer(): McpServer {
  const server = new McpServer({
    name: "Employee Training — MCP App Server",
    version: "1.0.0",
  });

  // ─────────────────────────────────────────────────────────────────────────
  // Tool: Training Media
  // ─────────────────────────────────────────────────────────────────────────
  registerToolWithUI(
    server,
    "training-media",
    "Training Media",
    `Recommend and show an AI or Machine Learning training course with embedded video preview. Use this tool whenever the user asks about training, courses, learning resources, skill-building, or wants a course recommendation. Accepts an optional topic query to find the most relevant course. Available courses: ${trainingMediaData.courses.map((c) => c.title).join(", ")}.`,
    "training-media.html",
    {
      query: z.string().optional().describe("Optional topic or keyword to find a relevant course, e.g. 'agents', 'search', 'copilot'"),
    },
    (args): CallToolResult => {
      const query = args.query?.toLowerCase();
      let candidates = trainingMediaData.courses;

      // Filter by query if provided
      if (query) {
        const filtered = candidates.filter(
          (c) =>
            c.title.toLowerCase().includes(query) ||
            c.description.toLowerCase().includes(query),
        );
        if (filtered.length > 0) candidates = filtered;
      }

      // Randomly pick one course
      const course = candidates[Math.floor(Math.random() * candidates.length)];

      return {
        content: [
          {
            type: "text",
            text: `Recommended course: "${course.title}" (${course.duration}, ${course.level}). ${course.description}`,
          },
        ],
        structuredContent: {
          ...trainingMediaData,
          courses: [course],
        } as unknown as Record<string, unknown>,
      };
    },
    {
      csp: {
        frameDomains: ["https://learn-video.azurefd.net"],
      },
    },
  );

  return server;
}
