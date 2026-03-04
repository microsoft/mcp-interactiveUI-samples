/**
 * Entry point for the MCP server.
 * Streamable HTTP transport (stateless mode).
 */
import { createMcpExpressApp } from "@modelcontextprotocol/sdk/server/express.js";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import cors from "cors";
import type { Request, Response } from "express";
import { createServer } from "./server.js";

/**
 * Starts the MCP server with Streamable HTTP transport (stateless mode).
 * Each request creates a new McpServer instance — no session persistence.
 */
export async function startStreamableHTTPServer(
  createServerFn: () => McpServer,
): Promise<void> {
  const port = parseInt(process.env.PORT ?? "3001", 10);

  const app = createMcpExpressApp({ host: "0.0.0.0" });
  app.use(cors());

  app.all("/mcp", async (req: Request, res: Response) => {
    const server = createServerFn();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined, // Stateless
    });

    res.on("close", () => {
      transport.close().catch(() => {});
      server.close().catch(() => {});
    });

    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      console.error("MCP error:", error);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        });
      }
    }
  });

  const httpServer = app.listen(port, () => {
    console.log(`\n🚀 MCP App Server listening on http://localhost:${port}/mcp\n`);
    console.log("Available tools:");
    console.log("  • training-media     — AI course recommendation with video preview\n");
  });

  const shutdown = () => {
    console.log("\nShutting down...");
    httpServer.close(() => process.exit(0));
  };

  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
}

startStreamableHTTPServer(createServer).catch((e) => {
  console.error(e);
  process.exit(1);
});
