// ─────────────────────────────────────────────────────────────────────────────
// Express server entry point
// POST /mcp — StreamableHTTP stateless MCP transport
// GET  /    — health check
// ─────────────────────────────────────────────────────────────────────────────

import express from "express";
import cors from "cors";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createMcpServer } from "./mcp-server.js";
import { getDb, getUsageStats } from "./db.js";
import { seedDatabase } from "./seed.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = parseInt(process.env.PORT ?? "3001", 10);

// ─────────────────────────────────────────────────────────────────────────────
// Auto-seed — if the approvals table is empty, seed with demo data.
// Runs once at startup so users get sample data immediately.
// ─────────────────────────────────────────────────────────────────────────────

async function autoSeedIfEmpty(): Promise<void> {
  try {
    const db = getDb();
    const row = db.prepare("SELECT COUNT(*) AS cnt FROM approvals").get() as { cnt: number };
    if (row.cnt === 0) {
      console.log("📦 Database empty — seeding with demo data...");
      await seedDatabase();
      console.log("✅ Auto-seed complete.");
    } else {
      console.log(`📋 Database has ${row.cnt} approvals — skipping seed.`);
    }
  } catch (err) {
    console.error("⚠️  Auto-seed failed (non-fatal):", err);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CORS — allow Copilot and localhost origins
// ─────────────────────────────────────────────────────────────────────────────

const ALLOWED_ORIGINS_EXACT = new Set([
  "http://localhost:3001",
  "http://localhost:5173",
]);

const ALLOWED_ORIGIN_SUFFIXES = [
  ".microsoft.com",
  ".devtunnels.ms",
  ".azurewebsites.net",
  ".ngrok-free.app",
  ".ngrok.io",
];

function isAllowedOrigin(origin: string | undefined): boolean {
  if (!origin) return true; // allow non-browser requests (e.g. curl, inspector)
  if (ALLOWED_ORIGINS_EXACT.has(origin)) return true;

  const additionalOrigins = process.env.ADDITIONAL_ALLOWED_ORIGINS?.split(",") ?? [];
  if (additionalOrigins.includes(origin)) return true;

  return ALLOWED_ORIGIN_SUFFIXES.some((suffix) => origin.endsWith(suffix));
}

const app = express();

app.use(
  cors({
    origin: (origin, callback) => {
      if (isAllowedOrigin(origin)) {
        callback(null, true);
      } else {
        callback(new Error(`Origin not allowed: ${origin}`));
      }
    },
    methods: ["GET", "POST", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization", "Mcp-Session-Id"],
    credentials: true,
  }),
);

app.use(express.json({ limit: "4mb" }));

// ─────────────────────────────────────────────────────────────────────────────
// Health check
// ─────────────────────────────────────────────────────────────────────────────

app.get("/", (_req, res) => {
  res.json({ name: "approvals-mcp", version: "1.0.0", status: "ok" });
});

app.get("/stats", (_req, res) => {
  res.json(getUsageStats());
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP endpoint — stateless StreamableHTTP transport
// A fresh server + transport is created per request (stateless pattern).
// ─────────────────────────────────────────────────────────────────────────────

app.all("/mcp", async (req, res) => {
  try {
    const server = createMcpServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined, // stateless
      enableJsonResponse: true,      // avoid 30s SSE timeout
    });

    res.on("close", () => {
      void transport.close();
      void server.close();
    });

    await server.connect(transport);
    await transport.handleRequest(req, res, req.body as Record<string, unknown>);
  } catch (err) {
    console.error("MCP transport error:", err);
    if (!res.headersSent) {
      res.status(500).json({ error: "Internal MCP server error" });
    }
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// Start
// ─────────────────────────────────────────────────────────────────────────────

autoSeedIfEmpty().then(() => {
  app.listen(PORT, () => {
    console.log(`\n🚀 Approvals MCP server listening on http://localhost:${PORT}/mcp`);
    console.log(`   Health check: http://localhost:${PORT}/`);
    console.log(`   Inspector:    npx @modelcontextprotocol/inspector http://localhost:${PORT}/mcp\n`);
  });
});
