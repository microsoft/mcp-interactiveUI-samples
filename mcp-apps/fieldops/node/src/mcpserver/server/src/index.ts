import express from "express";
import cors from "cors";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createMcpServer } from "./mcp-server.js";

const PORT = parseInt(process.env.PORT ?? "8787", 10);

const app = express();

app.use(cors({ origin: "*" }));
app.use(express.json({ limit: "4mb" }));

// ── Health check ───────────────────────────────────────────────────────
app.get("/", (_req, res) => {
  res.json({ name: "fieldops-mcp", status: "ok" });
});

// ── MCP endpoint — Streamable HTTP (stateless) ────────────────────────
app.all("/mcp", async (req, res) => {
  try {
    console.log(`\n══ MCP Request [${req.method}] ══`);
    console.log(JSON.stringify(req.body, null, 2));

    const server = createMcpServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined, // stateless
      enableJsonResponse: true,
    });

    // Intercept response to log it
    const chunks: Buffer[] = [];
    const origWrite = res.write.bind(res);
    const origEnd = res.end.bind(res);
    res.write = (chunk: any, ...args: any[]) => {
      if (chunk) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      return origWrite(chunk, ...args);
    };
    res.end = (chunk?: any, ...args: any[]) => {
      if (chunk) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      const body = Buffer.concat(chunks).toString("utf-8");
      try {
        console.log(`\n══ MCP Response ══`);
        console.log(JSON.stringify(JSON.parse(body), null, 2));
      } catch {
        console.log(`\n══ MCP Response (raw) ══`);
        console.log(body);
      }
      return origEnd(chunk, ...args);
    };

    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err) {
    console.error("MCP error:", err);
    if (!res.headersSent) {
      res.status(500).json({ error: "Internal server error" });
    }
  }
});

// ── Start ──────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(
    `🚀 Field Service Dispatch MCP server running at http://localhost:${PORT}`
  );
  console.log(`   MCP endpoint: http://localhost:${PORT}/mcp`);
});
