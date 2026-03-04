import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type {
  CallToolResult,
  ReadResourceResult,
} from "@modelcontextprotocol/sdk/types.js";
import {
  registerAppTool,
  registerAppResource,
  RESOURCE_MIME_TYPE,
} from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  listNewAssignments,
  getAssignmentsByIds,
  getAvailableTechnicians,
  getTechniciansByIds,
} from "./dispatch.js";
import { assignments } from "./mockData.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ASSETS_DIR = path.resolve(__dirname, "..", "..", "assets");

// ── Widget HTML loader ─────────────────────────────────────────────────
async function readWidgetHtml(name: string): Promise<string> {
  const filePath = path.join(ASSETS_DIR, `${name}.html`);
  try {
    return await fs.readFile(filePath, "utf-8");
  } catch {
    return `<html><body><p>Widget "${name}" not built yet. Run: npm run build:widgets</p></body></html>`;
  }
}

// ── Resource URIs ──────────────────────────────────────────────────────
const LIST_URI = "ui://fieldops/dispatch-list.html";
const MAP_URI = "ui://fieldops/dispatch-map.html";
const PLAN_URI = "ui://fieldops/dispatch-plan.html";

// ── MCP Server Factory ─────────────────────────────────────────────────
export function createMcpServer(): McpServer {
  const server = new McpServer({
    name: "fieldops-mcp",
    version: "1.0.0",
  });

  // ══════════════════════════════════════════════════════════════════════
  //  APP RESOURCES — Widget HTML (mimeType: text/html;profile=mcp-app)
  // ══════════════════════════════════════════════════════════════════════

  registerAppResource(
    server,
    "Dispatch List",
    LIST_URI,
    {
      mimeType: RESOURCE_MIME_TYPE,
      description: "Assignment list interactive widget",
    },
    async (): Promise<ReadResourceResult> => {
      const html = await readWidgetHtml("dispatch-list");
      return {
        contents: [
          { uri: LIST_URI, mimeType: RESOURCE_MIME_TYPE, text: html },
        ],
      };
    }
  );

  registerAppResource(
    server,
    "Dispatch Map",
    MAP_URI,
    {
      mimeType: RESOURCE_MIME_TYPE,
      description: "Assignment map interactive widget",
    },
    async (): Promise<ReadResourceResult> => {
      const html = await readWidgetHtml("dispatch-map");
      return {
        contents: [
          { uri: MAP_URI, mimeType: RESOURCE_MIME_TYPE, text: html },
        ],
      };
    }
  );

  registerAppResource(
    server,
    "Dispatch Plan",
    PLAN_URI,
    {
      mimeType: RESOURCE_MIME_TYPE,
      description: "Dispatch plan review interactive widget",
    },
    async (): Promise<ReadResourceResult> => {
      const html = await readWidgetHtml("dispatch-plan");
      return {
        contents: [
          { uri: PLAN_URI, mimeType: RESOURCE_MIME_TYPE, text: html },
        ],
      };
    }
  );

  // ══════════════════════════════════════════════════════════════════════
  //  APP TOOLS — Widget tools (tool + UI resource linked via _meta)
  // ══════════════════════════════════════════════════════════════════════

  // ── List New Assignments ─────────────────────────────────────────────
  registerAppTool(
    server,
    "list-new-assignments",
    {
      title: "List New Assignments",
      description:
        "Use this first for dispatch intake. Returns unassigned, recently created work orders (default: status=New and created in last 24 hours), optionally filtered by priority/region/team. Output includes assignment IDs, site, SLA due, required skills, status, and coordinates for downstream tools. Call before map or planning when user asks for new work.",
      inputSchema: {
        priority: z
          .enum(["Low", "Medium", "High"])
          .optional()
          .describe(
            "Optional priority filter. Use High/Medium/Low to narrow returned assignments."
          ),
        region: z
          .string()
          .optional()
          .describe(
            "Optional region filter. Example: North, South, East, West."
          ),
        team: z
          .string()
          .optional()
          .describe("Optional team filter. Example: Alpha, Beta."),
        maxHoursOld: z
          .number()
          .int()
          .min(1)
          .max(168)
          .optional()
          .describe(
            "Optional recency window in hours. Defaults to 24 when omitted."
          ),
        maxResults: z
          .number()
          .int()
          .min(1)
          .max(100)
          .optional()
          .describe(
            "Optional max number of assignments to return. Defaults to 50."
          ),
      },
      annotations: { readOnlyHint: true },
      _meta: { ui: { resourceUri: LIST_URI } },
    },
    async ({
      priority,
      region,
      team,
      maxHoursOld,
      maxResults,
    }): Promise<CallToolResult> => {
      const data = listNewAssignments({
        priority,
        region,
        team,
        maxHoursOld,
      }).slice(0, maxResults ?? 50);

      return {
        content: [
          {
            type: "text",
            text: "Suggested next step: list available technicians and create a dispatch plan.",
          },
        ],
        structuredContent: {
          view: "list",
          assignments: data.map((a) => ({
            id: a.id,
            site: a.site,
            category: a.category,
            priority: a.priority,
            slaDue: a.slaDue,
            estimatedStartDateTime: a.estimatedStartDateTime,
            estimatedEndDateTime: a.estimatedEndDateTime,
            description: a.description,
            requiredSkills: a.requiredSkills,
            customerName: a.customerName,
            customerPhone: a.customerPhone,
            customerProfilePicUrl: a.customerProfilePicUrl,
            assetId: a.assetId,
            estimatedDurationMinutes: a.estimatedDurationMinutes,
            estimatedTechnicianArrivalDateTime:
              a.estimatedTechnicianArrivalDateTime,
            assignedTechnicianId: a.assignedTechnicianId,
            siteImageUrl: a.siteImageUrl,
            tags: a.tags,
            status: a.status,
            address: a.location.address,
            lat: a.location.lat,
            lng: a.location.lng,
          })),
        },
      };
    }
  );

  // ── Show Assignments on Map ──────────────────────────────────────────
  registerAppTool(
    server,
    "show-assignments-on-map",
    {
      title: "Show Assignments on Map",
      description:
        "Render assignment pins on a Mapbox map view. Use after list-new-assignments and pass selected assignmentIds (or all returned IDs) to visualize location and inspect details. Returns map-ready pin payload and list metadata for map/list toggle. This is a read-only visualization tool and should not mutate dispatch state.",
      inputSchema: {
        assignmentIds: z
          .array(z.string())
          .optional()
          .describe(
            "Optional list of assignment IDs to display. If omitted, the tool maps current default new assignments."
          ),
      },
      annotations: { readOnlyHint: true },
      _meta: { ui: { resourceUri: MAP_URI } },
    },
    async ({ assignmentIds }): Promise<CallToolResult> => {
      const selected = assignmentIds?.length
        ? getAssignmentsByIds(assignmentIds)
        : listNewAssignments({ maxHoursOld: 24 });

      return {
        content: [
          {
            type: "text",
            text: "Suggested next step: list available technicians and create a dispatch plan.",
          },
        ],
        structuredContent: {
          view: "map",
          assignments: selected.map((a) => ({
            id: a.id,
            site: a.site,
            category: a.category,
            priority: a.priority,
            slaDue: a.slaDue,
            estimatedStartDateTime: a.estimatedStartDateTime,
            estimatedEndDateTime: a.estimatedEndDateTime,
            description: a.description,
            requiredSkills: a.requiredSkills,
            customerName: a.customerName,
            customerPhone: a.customerPhone,
            customerProfilePicUrl: a.customerProfilePicUrl,
            estimatedTechnicianArrivalDateTime:
              a.estimatedTechnicianArrivalDateTime,
            assignedTechnicianId: a.assignedTechnicianId,
            siteImageUrl: a.siteImageUrl,
            tags: a.tags,
            status: a.status,
            address: a.location.address,
            lat: a.location.lat,
            lng: a.location.lng,
          })),
        },
      };
    }
  );

  // ── Get Available Technicians ────────────────────────────────────────
  registerAppTool(
    server,
    "get-available-technicians",
    {
      description:
        "Return currently available technicians and their dispatch metadata for planning. Input can include optional region filter. Output includes technician IDs, skills, current location, availability, rating, shift, and profile metadata so Copilot can choose assignment-technician pairings.",
      inputSchema: {
        region: z
          .string()
          .optional()
          .describe(
            "Optional region filter to return technicians from a specific region."
          ),
      },
      annotations: { readOnlyHint: true },
      _meta: {},
    },
    async ({ region }) => {
      const available = getAvailableTechnicians({ region });
      return {
        content: [
          {
            type: "text" as const,
            text: "Suggested next step: create a dispatch plan for assignments.",
          },
        ],
        structuredContent: {
          technicians: available.map((tech) => ({
            id: tech.id,
            name: tech.name,
            profilePicUrl: tech.profilePicUrl,
            phone: tech.phone,
            rating: tech.rating,
            yearsExperience: tech.yearsExperience,
            shift: tech.shift,
            vehicleType: tech.vehicleType,
            skills: tech.skills,
            certifications: tech.certifications,
            languages: tech.languages,
            available: tech.available,
            lat: tech.location.lat,
            lng: tech.location.lng,
            address: tech.location.address,
          })),
        },
      };
    }
  );

  // ── Create Dispatch Plan ─────────────────────────────────────────────
  registerAppTool(
    server,
    "create-dispatch-plan",
    {
      title: "Create Dispatch Plan",
      description:
        "Assigns technicians to assignments and renders a review-ready dispatch plan. Before creating the plan, perform a prerequisite data check: if assignments are not available, call list-new-assignments; if available technicians are not available, call get-available-technicians. This check must be performed before optimization and plan generation. Think hard before generating the plan. Then optimize technician-to-assignment pairing carefully, considering skills, ETA, distance, and workload continuity, and return mapping with reasons. Each planItem requires assignmentId, technicianId, and reason; etaMinutes is recommended, skillMatch and distanceKm are optional. The MCP server does not compute or alter assignment-technician mapping; it renders the mapping provided in planItems.",
      inputSchema: {
        planItems: z
          .array(
            z.object({
              assignmentId: z
                .string()
                .describe(
                  "Required assignment ID selected for dispatch."
                ),
              technicianId: z
                .string()
                .describe(
                  "Required technician ID selected by the external matching/planning logic."
                ),
              reason: z
                .string()
                .optional()
                .describe(
                  "Clear explanation of why this technician is assigned to this assignment."
                ),
              etaMinutes: z
                .number()
                .int()
                .positive()
                .describe(
                  "Recommended ETA in minutes for technician arrival."
                ),
              skillMatch: z
                .enum(["Full", "Partial"])
                .optional()
                .describe(
                  "Optional skill match label from external matching result."
                ),
              distanceKm: z
                .number()
                .nonnegative()
                .optional()
                .describe(
                  "Optional travel distance in kilometers from technician to assignment."
                ),
            })
          )
          .describe(
            "Required final assignment-technician pairings produced by external planning logic. One item per assignment to plan."
          ),
        technicianIds: z
          .array(z.string())
          .optional()
          .describe(
            "Optional technician IDs from get-available-technicians output, used to populate the override dropdown in the UI."
          ),
        maxTravelKm: z
          .number()
          .min(1)
          .max(500)
          .optional()
          .describe(
            "Optional display constraint context for plan review. Defaults to 60."
          ),
        allowPartialSkillMatch: z
          .boolean()
          .optional()
          .describe(
            "Optional display constraint context. Defaults to true."
          ),
        travelBufferMinutes: z
          .number()
          .int()
          .min(0)
          .max(180)
          .optional()
          .describe(
            "Optional scheduling buffer between consecutive assignments for the same technician. Defaults to 30."
          ),
      },
      annotations: { readOnlyHint: true },
      _meta: { ui: { resourceUri: PLAN_URI } },
    },
    async ({
      planItems: inputPlanItems,
      technicianIds,
      maxTravelKm,
      allowPartialSkillMatch,
      travelBufferMinutes,
    }): Promise<CallToolResult> => {
      // Resolve assignment and technician details from IDs in planItems
      const assignmentIds = [
        ...new Set(inputPlanItems.map((item) => item.assignmentId)),
      ];
      const planAssignments = getAssignmentsByIds(assignmentIds);

      const planTechIds = [
        ...new Set(inputPlanItems.map((item) => item.technicianId)),
      ];
      const planTechs = getTechniciansByIds(planTechIds);

      // Build technician dropdown options from explicit input or fall back
      const technicianOptions = technicianIds?.length
        ? getTechniciansByIds(technicianIds)
        : planTechs;

      const techniciansForMap = Array.from(
        new Map(
          [...planTechs, ...technicianOptions].map((t) => [t.id, t])
        ).values()
      );

      const planItems = inputPlanItems.map((item) => {
        const technician = planTechs.find((t) => t.id === item.technicianId);
        return {
          ...item,
          technicianName: technician?.name ?? item.technicianId,
        };
      });

      return {
        content: [
          {
            type: "text",
            text: `Prepared a plan for ${planItems.length} assignments.`,
          },
        ],
        structuredContent: {
          view: "plan",
          constraints: {
            maxTravelKm: maxTravelKm ?? 60,
            allowPartialSkillMatch: allowPartialSkillMatch ?? true,
            travelBufferMinutes: travelBufferMinutes ?? 30,
          },
          planItems,
          assignments: planAssignments,
          technicians: techniciansForMap,
          technicianOptions: technicianOptions.map((t) => ({
            id: t.id,
            name: t.name,
          })),
        },
      };
    }
  );

  // ── Commit Assignments ───────────────────────────────────────────────
  registerAppTool(
    server,
    "commit-assignments",
    {
      title: "Commit Assignments",
      description:
        "Commit a reviewed dispatch plan after explicit user confirmation (for example, user says Yes, assign them). Input must include final plan rows (assignmentId + technicianId + etaMinutes), including any manual overrides from the UI. Returns confirmation summary and per-assignment status for follow-up actions. Do not call this before create-dispatch-plan and user confirmation.",
      inputSchema: {
        assignments: z
          .array(
            z.object({
              assignmentId: z
                .string()
                .describe("Required assignment ID to commit."),
              technicianId: z
                .string()
                .describe("Required final technician ID to assign."),
              etaMinutes: z
                .number()
                .int()
                .positive()
                .describe(
                  "Required final ETA in minutes for commit summary."
                ),
            })
          )
          .describe(
            "Required final assignment rows to commit. Use reviewed plan rows from create-dispatch-plan output."
          ),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false,
      },
      _meta: { ui: { visibility: ["app"] } },
    },
    async ({
      assignments: commitAssignments,
    }) => {
      const rows = commitAssignments.map((item) => {
        const assignment = assignments.find(
          (a) => a.id === item.assignmentId
        );
        const technician = getTechniciansByIds([item.technicianId])[0];

        const estimatedTechnicianArrivalDateTime = new Date(
          Date.now() + item.etaMinutes * 60 * 1000
        ).toISOString();

        return {
          assignmentId: item.assignmentId,
          site: assignment?.site ?? item.assignmentId,
          technicianId: item.technicianId,
          technicianName: technician?.name ?? item.technicianId,
          etaMinutes: item.etaMinutes,
          estimatedTechnicianArrivalDateTime,
          assignedTechnicianId: item.technicianId,
          status: "Dispatched",
        };
      });

      return {
        content: [
          {
            type: "text" as const,
            text: `${rows.length} assignments have been confirmed.`,
          },
        ],
        structuredContent: {
          summary: `${rows.length} assignments confirmed`,
          count: rows.length,
          rows,
        },
      };
    }
  );

  return server;
}
