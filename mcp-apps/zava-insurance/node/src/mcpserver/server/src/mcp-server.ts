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
import * as db from "./db.js";

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

// ── JSON parse helper ──────────────────────────────────────────────────
function safeJsonParse(value: unknown): unknown {
  if (typeof value !== "string") return value;
  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}

function parseClaim(entity: db.ClaimEntity) {
  return {
    id: entity.rowKey,
    claimNumber: entity.claimNumber,
    policyNumber: entity.policyNumber,
    policyHolderName: entity.policyHolderName,
    policyHolderEmail: entity.policyHolderEmail,
    property: entity.property,
    dateOfLoss: entity.dateOfLoss,
    dateReported: entity.dateReported,
    status: entity.status,
    damageTypes: safeJsonParse(entity.damageTypes),
    description: entity.description,
    estimatedLoss: entity.estimatedLoss,
    adjusterAssigned: entity.adjusterAssigned,
    notes: safeJsonParse(entity.notes),
    createdAt: entity.createdAt,
    updatedAt: entity.updatedAt,
  };
}

function parseContractor(entity: db.ContractorEntity) {
  return {
    id: entity.rowKey,
    name: entity.name,
    businessName: entity.businessName,
    email: entity.email,
    phone: entity.phone,
    address: safeJsonParse(entity.address),
    licenseNumber: entity.licenseNumber,
    insuranceCertificate: entity.insuranceCertificate,
    specialties: safeJsonParse(entity.specialties),
    rating: entity.rating,
    isPreferred: entity.isPreferred,
    isActive: entity.isActive,
  };
}

function parseInspection(entity: db.InspectionEntity) {
  return {
    id: entity.rowKey,
    claimId: entity.claimId,
    claimNumber: entity.claimNumber,
    taskType: entity.taskType,
    priority: entity.priority,
    status: entity.status,
    scheduledDate: entity.scheduledDate,
    inspectorId: entity.inspectorId,
    property: entity.property,
    instructions: entity.instructions,
    photos: safeJsonParse(entity.photos),
    findings: entity.findings,
    recommendedActions: safeJsonParse(entity.recommendedActions),
    flaggedIssues: safeJsonParse(entity.flaggedIssues),
    createdAt: entity.createdAt,
    updatedAt: entity.updatedAt,
    completedDate: entity.completedDate,
  };
}

function parseInspector(entity: db.InspectorEntity) {
  return {
    id: entity.rowKey,
    name: entity.name,
    email: entity.email,
    phone: entity.phone,
    licenseNumber: entity.licenseNumber,
    specializations: safeJsonParse(entity.specializations),
  };
}

function parsePurchaseOrder(entity: db.PurchaseOrderEntity) {
  return {
    id: entity.rowKey,
    poNumber: entity.poNumber,
    claimId: entity.claimId,
    claimNumber: entity.claimNumber,
    contractorId: entity.contractorId,
    workDescription: entity.workDescription,
    lineItems: safeJsonParse(entity.lineItems),
    subtotal: entity.subtotal,
    tax: entity.tax,
    total: entity.total,
    status: entity.status,
    createdDate: entity.createdDate,
    notes: safeJsonParse(entity.notes),
  };
}

// ── Resource URIs ──────────────────────────────────────────────────────
const DASHBOARD_URI = "ui://zava/claims-dashboard.html";
const DETAIL_URI = "ui://zava/claim-detail.html";
const CONTRACTORS_URI = "ui://zava/contractors-list.html";

// ── MCP Server Factory ─────────────────────────────────────────────────
export function createMcpServer(): McpServer {
  const server = new McpServer({
    name: "zava-insurance-mcp",
    version: "1.0.0",
  });

  // ══════════════════════════════════════════════════════════════════════
  //  APP RESOURCES — Widget HTML (mimeType: text/html;profile=mcp-app)
  // ══════════════════════════════════════════════════════════════════════
  registerAppResource(
    server,
    "Claims Dashboard",
    DASHBOARD_URI,
    {
      mimeType: RESOURCE_MIME_TYPE,
      description: "Claims dashboard interactive widget",
    },
    async (): Promise<ReadResourceResult> => {
      const html = await readWidgetHtml("claims-dashboard");
      return {
        contents: [
          { uri: DASHBOARD_URI, mimeType: RESOURCE_MIME_TYPE, text: html },
        ],
      };
    }
  );

  registerAppResource(
    server,
    "Claim Detail",
    DETAIL_URI,
    {
      mimeType: RESOURCE_MIME_TYPE,
      description: "Claim detail interactive widget",
    },
    async (): Promise<ReadResourceResult> => {
      const html = await readWidgetHtml("claim-detail");
      return {
        contents: [
          { uri: DETAIL_URI, mimeType: RESOURCE_MIME_TYPE, text: html },
        ],
      };
    }
  );

  registerAppResource(
    server,
    "Contractors List",
    CONTRACTORS_URI,
    {
      mimeType: RESOURCE_MIME_TYPE,
      description: "Contractors list interactive widget",
    },
    async (): Promise<ReadResourceResult> => {
      const html = await readWidgetHtml("contractors-list");
      return {
        contents: [
          {
            uri: CONTRACTORS_URI,
            mimeType: RESOURCE_MIME_TYPE,
            text: html,
          },
        ],
      };
    }
  );

  // ══════════════════════════════════════════════════════════════════════
  //  APP TOOLS — Widget tools (tool + UI resource linked via _meta)
  // ══════════════════════════════════════════════════════════════════════

  // ── Show Claims Dashboard ────────────────────────────────────────────
  registerAppTool(
    server,
    "show-claims-dashboard",
    {
      title: "Show Claims Dashboard",
      description:
        "Displays the Zava Insurance claims dashboard showing all claims with status overview, filters, and summary metrics. " +
        "Supports filtering by status and/or policy holder name. When the user mentions a person's name, first name, last name, " +
        "or partial name, always pass it as the policyHolderName parameter. The name filter is case-insensitive and supports partial matches.",
      inputSchema: {
        status: z
          .string()
          .optional()
          .describe(
            "Filter claims by status keyword (e.g. 'Open', 'Approved', 'Pending', 'Denied', 'Closed')"
          ),
        policyHolderName: z
          .string()
          .optional()
          .describe(
            "Filter claims by policy holder name. Supports partial, case-insensitive matching."
          ),
      },
      annotations: { readOnlyHint: true },
      _meta: { ui: { resourceUri: DASHBOARD_URI } },
    },
    async ({ status, policyHolderName }): Promise<CallToolResult> => {
      let claims = (await db.getAllClaims()).map(parseClaim);
      if (status) {
        const s = status.toLowerCase();
        claims = claims.filter((c) => c.status.toLowerCase().includes(s));
      }
      if (policyHolderName) {
        const n = policyHolderName.toLowerCase();
        claims = claims.filter((c) =>
          c.policyHolderName.toLowerCase().includes(n)
        );
      }

      const allInspections = (await db.getAllInspections()).map(parseInspection);
      const allPurchaseOrders = (await db.getAllPurchaseOrders()).map(
        parsePurchaseOrder
      );

      const contractorIds = [
        ...new Set(allPurchaseOrders.map((po) => po.contractorId)),
      ];
      const allContractors: Record<string, any> = {};
      for (const cid of contractorIds) {
        const c = await db.getContractorById(cid);
        if (c) allContractors[cid] = parseContractor(c);
      }

      const inspectorIds = [
        ...new Set(allInspections.map((i) => i.inspectorId)),
      ];
      const allInspectors: Record<string, any> = {};
      for (const iid of inspectorIds) {
        const i = await db.getInspectorById(iid);
        if (i) allInspectors[iid] = parseInspector(i);
      }

      return {
        content: [
          { type: "text", text: `Loaded ${claims.length} claims.` },
        ],
        structuredContent: {
          claims,
          inspections: allInspections,
          purchaseOrders: allPurchaseOrders,
          contractors: allContractors,
          inspectors: allInspectors,
        },
      };
    }
  );

  // ── Show Claim Detail ────────────────────────────────────────────────
  registerAppTool(
    server,
    "show-claim-detail",
    {
      title: "Show Claim Detail",
      description:
        "Displays detailed information about a specific insurance claim including related inspections, " +
        "purchase orders, and contractor assignments. Use claim ID (e.g. '1', '2') or claim number (e.g. 'CN202504990').",
      inputSchema: {
        claimId: z
          .string()
          .describe("The claim ID or claim number to look up"),
      },
      annotations: { readOnlyHint: true },
      _meta: { ui: { resourceUri: DETAIL_URI } },
    },
    async ({ claimId }): Promise<CallToolResult> => {
      let claim = await db.getClaimById(claimId);
      if (!claim) {
        const all = await db.getAllClaims();
        claim = all.find((c) => c.claimNumber === claimId) ?? null;
      }
      if (!claim) {
        return {
          content: [
            { type: "text", text: `Claim "${claimId}" not found.` },
          ],
        };
      }
      const parsed = parseClaim(claim);
      const inspections = (
        await db.getInspectionsByClaimId(parsed.id)
      ).map(parseInspection);
      const purchaseOrders = (
        await db.getPurchaseOrdersByClaimId(parsed.id)
      ).map(parsePurchaseOrder);

      const contractorIds = [
        ...new Set(purchaseOrders.map((po) => po.contractorId)),
      ];
      const contractors: Record<string, any> = {};
      for (const cid of contractorIds) {
        const c = await db.getContractorById(cid);
        if (c) contractors[cid] = parseContractor(c);
      }

      const inspectorIds = [
        ...new Set(inspections.map((i) => i.inspectorId)),
      ];
      const inspectors: Record<string, any> = {};
      for (const iid of inspectorIds) {
        const i = await db.getInspectorById(iid);
        if (i) inspectors[iid] = parseInspector(i);
      }

      return {
        content: [
          {
            type: "text",
            text: `Claim ${parsed.claimNumber} - ${parsed.policyHolderName}`,
          },
        ],
        structuredContent: {
          claim: parsed,
          inspections,
          purchaseOrders,
          contractors,
          inspectors,
        },
      };
    }
  );

  // ── Show Contractors ─────────────────────────────────────────────────
  registerAppTool(
    server,
    "show-contractors",
    {
      title: "Show Contractors",
      description:
        "Displays the list of contractors available for insurance repair work. Optionally filter by specialty or preferred status.",
      inputSchema: {
        specialty: z
          .string()
          .optional()
          .describe(
            "Filter by contractor specialty (e.g. 'Roofing', 'Water Damage', 'Fire')"
          ),
        preferredOnly: z
          .boolean()
          .optional()
          .describe("Show only preferred contractors"),
      },
      annotations: { readOnlyHint: true },
      _meta: { ui: { resourceUri: CONTRACTORS_URI } },
    },
    async ({ specialty, preferredOnly }): Promise<CallToolResult> => {
      let contractors = (await db.getAllContractors()).map(parseContractor);
      if (specialty) {
        const s = specialty.toLowerCase();
        contractors = contractors.filter((c) => {
          const specs = Array.isArray(c.specialties) ? c.specialties : [];
          return specs.some((sp: string) => sp.toLowerCase().includes(s));
        });
      }
      if (preferredOnly) {
        contractors = contractors.filter((c) => c.isPreferred);
      }
      return {
        content: [
          {
            type: "text",
            text: `Loaded ${contractors.length} contractors.`,
          },
        ],
        structuredContent: { contractors },
      };
    }
  );

  // ══════════════════════════════════════════════════════════════════════
  //  DATA TOOLS — Non-widget tools (callable by both model and app)
  // ══════════════════════════════════════════════════════════════════════

  // ── Update Claim Status ──────────────────────────────────────────────
  server.tool(
    "update-claim-status",
    "Updates the status of an insurance claim. Use claim ID (e.g. '1', '2').",
    {
      claimId: z.string().describe("The claim ID"),
      status: z
        .string()
        .describe(
          "New status (e.g. 'Approved', 'Denied', 'Closed', 'Open - Under Investigation')"
        ),
      note: z
        .string()
        .optional()
        .describe("Optional note to add to the claim"),
    },
    async ({ claimId, status, note }) => {
      const existing = await db.getClaimById(claimId);
      if (!existing) {
        return {
          content: [
            {
              type: "text" as const,
              text: `Claim "${claimId}" not found.`,
            },
          ],
        };
      }
      const updates: Record<string, unknown> = {
        status,
        updatedAt: new Date().toISOString(),
      };
      if (note) {
        const existingNotes = safeJsonParse(existing.notes);
        const notes = Array.isArray(existingNotes)
          ? [...existingNotes, note]
          : [note];
        updates.notes = notes;
      }
      await db.updateClaim(claimId, updates);
      return {
        content: [
          {
            type: "text" as const,
            text: `✅ Claim ${existing.claimNumber} status updated to "${status}".`,
          },
        ],
      };
    }
  );

  // ── Update Inspection ────────────────────────────────────────────────
  server.tool(
    "update-inspection",
    "Updates an inspection record — status, findings, recommended actions, property, or inspector assignment.",
    {
      inspectionId: z
        .string()
        .describe("The inspection ID (e.g. 'insp-001')"),
      status: z
        .string()
        .optional()
        .describe(
          "New status (e.g. 'completed', 'scheduled', 'in-progress', 'cancelled')"
        ),
      findings: z.string().optional().describe("Updated findings text"),
      recommendedActions: z
        .array(z.string())
        .optional()
        .describe("Updated recommended actions"),
      property: z.string().optional().describe("Updated property address"),
      inspectorId: z
        .string()
        .optional()
        .describe("Inspector ID to assign (e.g. 'inspector-003')"),
    },
    async ({
      inspectionId,
      status,
      findings,
      recommendedActions,
      property,
      inspectorId,
    }) => {
      const existing = await db.getInspectionById(inspectionId);
      if (!existing) {
        return {
          content: [
            {
              type: "text" as const,
              text: `Inspection "${inspectionId}" not found.`,
            },
          ],
        };
      }
      const updates: Record<string, unknown> = {
        updatedAt: new Date().toISOString(),
      };
      if (status) {
        updates.status = status;
        if (status === "completed")
          updates.completedDate = new Date().toISOString();
      }
      if (findings) updates.findings = findings;
      if (recommendedActions) updates.recommendedActions = recommendedActions;
      if (property) updates.property = property;
      if (inspectorId) updates.inspectorId = inspectorId;
      await db.updateInspection(inspectionId, updates);
      return {
        content: [
          {
            type: "text" as const,
            text: `✅ Inspection ${inspectionId} updated.`,
          },
        ],
      };
    }
  );

  // ── Update Purchase Order ────────────────────────────────────────────
  server.tool(
    "update-purchase-order",
    "Updates a purchase order status (e.g. approve, reject, complete).",
    {
      purchaseOrderId: z
        .string()
        .describe("The purchase order ID (e.g. 'po-001')"),
      status: z
        .string()
        .describe(
          "New status (e.g. 'approved', 'rejected', 'completed', 'in-progress')"
        ),
      note: z.string().optional().describe("Optional note to add"),
    },
    async ({ purchaseOrderId, status, note }) => {
      const existing = await db.getPurchaseOrderById(purchaseOrderId);
      if (!existing) {
        return {
          content: [
            {
              type: "text" as const,
              text: `Purchase order "${purchaseOrderId}" not found.`,
            },
          ],
        };
      }
      const updates: Record<string, unknown> = { status };
      if (note) {
        const existingNotes = safeJsonParse(existing.notes);
        const notes = Array.isArray(existingNotes)
          ? [...existingNotes, note]
          : [note];
        updates.notes = notes;
      }
      await db.updatePurchaseOrder(purchaseOrderId, updates);
      return {
        content: [
          {
            type: "text" as const,
            text: `✅ Purchase order ${existing.poNumber} status updated to "${status}".`,
          },
        ],
      };
    }
  );

  // ── Get Claim Summary ────────────────────────────────────────────────
  server.tool(
    "get-claim-summary",
    "Returns a text summary for a specific claim with key details. Use claim ID or claim number.",
    {
      claimId: z.string().describe("Claim ID or claim number"),
    },
    async ({ claimId }) => {
      let claim = await db.getClaimById(claimId);
      if (!claim) {
        const all = await db.getAllClaims();
        claim = all.find((c) => c.claimNumber === claimId) ?? null;
      }
      if (!claim) {
        return {
          content: [
            {
              type: "text" as const,
              text: `Claim "${claimId}" not found.`,
            },
          ],
        };
      }
      const c = parseClaim(claim);
      const inspections = await db.getInspectionsByClaimId(c.id);
      const pos = await db.getPurchaseOrdersByClaimId(c.id);
      const summary = [
        `📋 Claim: ${c.claimNumber}`,
        `👤 Policy Holder: ${c.policyHolderName} (${c.policyHolderEmail})`,
        `📍 Property: ${c.property}`,
        `📅 Date of Loss: ${new Date(c.dateOfLoss).toLocaleDateString()}`,
        `💰 Estimated Loss: $${c.estimatedLoss.toLocaleString()}`,
        `📊 Status: ${c.status}`,
        `🔧 Damage Types: ${(c.damageTypes as string[]).join(", ")}`,
        `🔍 Inspections: ${inspections.length}`,
        `📦 Purchase Orders: ${pos.length}`,
        `📝 Description: ${c.description}`,
      ].join("\n");
      return {
        content: [{ type: "text" as const, text: summary }],
      };
    }
  );

  // ── Create Inspection ────────────────────────────────────────────────
  server.tool(
    "create-inspection",
    "Creates a new inspection record. Only claimNumber is required. ID is auto-generated, status defaults to 'open'. claimId is optional.",
    {
      claimNumber: z
        .string()
        .describe("The claim number (e.g. 'CN202504990')"),
      claimId: z.string().optional().describe("Optional claim ID"),
      taskType: z
        .string()
        .optional()
        .describe(
          "Type of inspection: 'initial', 're-inspection', 'final'. Defaults to 'initial'"
        ),
      priority: z
        .string()
        .optional()
        .describe("Priority: 'low', 'medium', 'high'. Defaults to 'medium'"),
      status: z
        .string()
        .optional()
        .describe("Status. Defaults to 'open'"),
      scheduledDate: z
        .string()
        .optional()
        .describe("Scheduled date (ISO string)"),
      inspectorId: z
        .string()
        .optional()
        .describe("Inspector ID to assign"),
      property: z.string().optional().describe("Property address"),
      instructions: z
        .string()
        .optional()
        .describe("Inspection instructions"),
    },
    async ({
      claimNumber,
      claimId,
      taskType,
      priority,
      status,
      scheduledDate,
      inspectorId,
      property,
      instructions,
    }) => {
      let resolvedClaimId = claimId ?? "";
      let resolvedProperty = property ?? "";
      if (!resolvedClaimId || !resolvedProperty) {
        const allClaims = await db.getAllClaims();
        const matchedClaim = allClaims.find(
          (c) => c.claimNumber === claimNumber
        );
        if (matchedClaim) {
          if (!resolvedClaimId) resolvedClaimId = matchedClaim.rowKey;
          if (!resolvedProperty) resolvedProperty = matchedClaim.property;
        }
      }
      const newInspection = await db.createInspection({
        claimNumber,
        claimId: resolvedClaimId,
        property: resolvedProperty,
        taskType,
        priority,
        status,
        scheduledDate,
        inspectorId,
        instructions,
      });
      const parsed = parseInspection(newInspection);
      return {
        content: [
          {
            type: "text" as const,
            text: `✅ Inspection ${parsed.id} created for claim ${claimNumber}. Status: ${parsed.status}`,
          },
        ],
      };
    }
  );

  // ── List Inspectors ──────────────────────────────────────────────────
  server.tool(
    "list-inspectors",
    "Lists all available inspectors with their specializations.",
    {},
    async () => {
      const inspectors = (await db.getAllInspectors()).map(parseInspector);
      const lines = inspectors.map(
        (i) =>
          `• ${i.name} (${i.id}) — ${(i.specializations as string[]).join(", ")} — ${i.email}`
      );
      return {
        content: [
          {
            type: "text" as const,
            text: `👷 ${inspectors.length} Inspectors:\n${lines.join("\n")}`,
          },
        ],
      };
    }
  );

  return server;
}
