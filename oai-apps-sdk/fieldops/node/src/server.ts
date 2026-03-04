import cors from "cors";
import express from "express";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createMcpServer } from "./mcp/createServer.js";

const port = Number(process.env.PORT ?? 8787);
const app = express();

app.use(cors());
app.use(express.json({ limit: "4mb" }));

app.post("/mcp", async (req, res) => {
  const server = createMcpServer();
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
  });

  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);

    res.on("close", () => {
      void transport.close();
      void server.close();
    });
  } catch (error) {
    // TODO(template): Add structured logging + telemetry correlation here.
    console.error("MCP request failed", error);
    if (!res.headersSent) {
      res.status(500).json({ error: "MCP transport error" });
    }

    await transport.close();
    await server.close();
  }
});

app.get("/mcp", (_req, res) => {
  res.status(405).json({ error: "Method not allowed. Use POST /mcp" });
});

app.delete("/mcp", (_req, res) => {
  res.status(405).json({ error: "Method not allowed. Use POST /mcp" });
});

app.listen(port, () => {
  console.log(`Field Service Dispatch MCP server listening on http://localhost:${port}/mcp`);
});
