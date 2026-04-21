// ─────────────────────────────────────────────────────────────────────────────
// MCP server — all tools + resource templates
// Uses McpServer (high-level API) with registerTool / registerResource
// ─────────────────────────────────────────────────────────────────────────────

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ListResourceTemplatesRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import {
  getApprovals,
  getApprovalById,
  getSimilarApprovals,
  approveApproval,
  rejectApproval,
  bulkDecide,
  addComment,
  createApproval,
  getUserBySubject,
  insertUsageEvent,
} from "./db.js";
import { computeRisk } from "./risk.js";
import type { ApprovalFilters, ApprovalType, SortField } from "./types.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ASSETS_DIR = path.resolve(__dirname, "..", "assets");

// ─────────────────────────────────────────────────────────────────────────────
// Widget descriptors
// ─────────────────────────────────────────────────────────────────────────────

const WIDGET_MIME = "text/html+skybridge";

interface WidgetDef {
  id: string;
  title: string;
  templateUri: string;
  invoking: string;
  invoked: string;
}

const WIDGETS = {
  pendingApprovals: {
    id: "pending-approvals",
    title: "Pending Approvals",
    templateUri: "ui://widget/pending-approvals.html",
    invoking: "Loading your approvals…",
    invoked: "Approvals ready",
  },
  approvalDetail: {
    id: "approval-detail",
    title: "Approval Detail",
    templateUri: "ui://widget/approval-detail.html",
    invoking: "Opening approval…",
    invoked: "Approval loaded",
  },
  createApproval: {
    id: "create-approval",
    title: "Create Approval",
    templateUri: "ui://widget/create-approval.html",
    invoking: "Opening create form…",
    invoked: "Form ready",
  },
} satisfies Record<string, WidgetDef>;

const ALL_WIDGETS = Object.values(WIDGETS);

function descriptorMeta(widget: WidgetDef) {
  return {
    "openai/outputTemplate": widget.templateUri,
    "openai/toolInvocation/invoking": widget.invoking,
    "openai/toolInvocation/invoked": widget.invoked,
    "openai/widgetAccessible": true,
  } as const;
}

function invocationMeta(widget: WidgetDef) {
  return {
    "openai/toolInvocation/invoking": widget.invoking,
    "openai/toolInvocation/invoked": widget.invoked,
  } as const;
}

// ─────────────────────────────────────────────────────────────────────────────
// Widget HTML loader
// ─────────────────────────────────────────────────────────────────────────────

function loadWidgetHtml(widgetId: string): string {
  const htmlPath = path.join(ASSETS_DIR, `${widgetId}.html`);
  try {
    return fs.readFileSync(htmlPath, "utf-8");
  } catch {
    // Widget not yet built — return placeholder for Phase 1
    return `<!doctype html><html><head><meta charset="utf-8"/><title>${widgetId}</title></head>
<body style="font-family:sans-serif;padding:24px;background:#f5f5f5;">
  <h2>Widget: ${widgetId}</h2>
  <p style="color:#666;">Widget HTML not yet built. Run <code>npm run build:widgets</code>.</p>
  <pre id="data" style="background:#fff;padding:12px;border-radius:8px;font-size:12px;overflow:auto;"></pre>
  <script>
    const set = (e) => { try { document.getElementById('data').textContent = JSON.stringify(window.openai?.toolOutput, null, 2); } catch {} };
    window.addEventListener('openai:set_globals', set);
    set();
  </script>
</body></html>`;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Actor identity — extract user from request metadata
// Falls back to demo user mgr_201 until Phase 6b auth is wired
// ─────────────────────────────────────────────────────────────────────────────

interface RequestMeta {
  subject?: string;
}

function resolveActor(meta?: RequestMeta): { actorId: string; actorName: string } {
  if (meta?.subject) {
    const user = getUserBySubject(meta.subject);
    if (user) return { actorId: user.id, actorName: user.name };
  }
  // Default demo user
  return { actorId: "mgr_201", actorName: "Samantha Sandy" };
}

// ─────────────────────────────────────────────────────────────────────────────
// Server factory
// ─────────────────────────────────────────────────────────────────────────────

export function createMcpServer(_baseUrl?: string): McpServer {
  const server = new McpServer({
    name: "approvals-mcp",
    version: "1.0.0",
  });

  // ── Resources (widget HTML via text/html+skybridge) ──────────────────────

  for (const widget of ALL_WIDGETS) {
    server.registerResource(
      widget.id,
      widget.templateUri,
      {
        name: widget.title,
        description: `${widget.title} widget`,
        mimeType: WIDGET_MIME,
      },
      async () => ({
        contents: [
          {
            uri: widget.templateUri,
            mimeType: WIDGET_MIME,
            text: loadWidgetHtml(widget.id),
            _meta: descriptorMeta(widget),
          },
        ],
      }),
    );
  }

  // Resource templates (required for host widget discovery)
  server.server.setRequestHandler(ListResourceTemplatesRequestSchema, async () => ({
    resourceTemplates: ALL_WIDGETS.map((w) => ({
      uriTemplate: w.templateUri,
      name: w.title,
      description: `${w.title} widget markup`,
      mimeType: WIDGET_MIME,
      _meta: descriptorMeta(w),
    })),
  }));

  // ── READ TOOLS ───────────────────────────────────────────────────────────

  // list_pending_approvals
  server.registerTool(
    "list_pending_approvals",
    {
      title: "List Pending Approvals",
      description:
        "Returns a summary count and paginated list of approvals assigned to the current user. " +
        "Supports filtering by status, risk level, type, requester, due date, and attachment presence. " +
        "Use this as the primary entry point to the approvals workspace. " +
        "IMPORTANT: To bulk-approve or bulk-reject a filtered set (e.g. 'approve all low-risk purchase orders'), " +
        "call this tool with the appropriate filters AND set open_bulk_action to 'approve' or 'reject' — " +
        "the widget will open with matching items pre-selected and the bulk confirm dialog ready. " +
        "Do NOT call bulk_decide_approvals directly from chat.",
      inputSchema: {
        status: z
          .array(z.enum(["pending", "approved", "rejected", "cancelled"]))
          .optional()
          .describe("Filter by status. Defaults to pending only."),
        risk_level: z
          .array(z.enum(["low", "medium", "high"]))
          .optional()
          .describe("Filter by risk level."),
        type: z
          .array(
            z.enum([
              "promo_pricing",
              "purchase_order",
              "capex",
              "contract_sow",
              "travel_exception",
              "marketing_budget",
              "vendor_onboarding",
            ]),
          )
          .optional()
          .describe("Filter by approval type."),
        requester: z
          .string()
          .optional()
          .describe("Filter by requester name (partial match)."),
        has_attachments: z
          .boolean()
          .optional()
          .describe("Filter to only approvals with or without attachments."),
        due_before: z
          .string()
          .optional()
          .describe("Filter approvals due before this ISO date (e.g. 2026-03-21)."),
        query: z
          .string()
          .optional()
          .describe("Free-text search across title, summary, requester, and fields."),
        sort: z
          .enum(["due_date", "amount", "submitted_date", "risk"])
          .optional()
          .describe("Sort field. Defaults to due_date ascending."),
        limit: z
          .number()
          .int()
          .min(1)
          .max(100)
          .optional()
          .describe("Page size. Defaults to 20."),
        cursor: z.string().optional().describe("Pagination cursor from previous response."),
        open_bulk_action: z
          .enum(["approve", "reject"])
          .optional()
          .describe("When set, the widget opens with all filtered items pre-selected and the bulk action dialog open. Use for bulk approve/reject intents."),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
      _meta: descriptorMeta(WIDGETS.pendingApprovals),
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(
        extra?.meta as RequestMeta | undefined,
      );
      insertUsageEvent(actorId, actorName, "list_pending_approvals", "pendingApprovals");

      const filters: ApprovalFilters = {
        status: (input.status as ApprovalFilters["status"]) ?? ["pending"],
        risk_level: input.risk_level as ApprovalFilters["risk_level"],
        type: input.type as ApprovalFilters["type"],
        requester: input.requester,
        has_attachments: input.has_attachments,
        due_before: input.due_before,
        query: input.query,
        sort: (input.sort as SortField) ?? "due_date",
        limit: input.limit ?? 20,
        cursor: input.cursor,
      };

      const result = getApprovals(filters, actorId);

      const summaryText =
        `${result.summary.total_pending} pending approvals — ` +
        `${result.summary.urgent} urgent, ${result.summary.overdue} overdue, ` +
        `${result.summary.high_risk} high risk. ` +
        `Total value pending: USD ${result.summary.total_amount_pending.toLocaleString()}.`;

      return {
        content: [{ type: "text", text: summaryText }],
        structuredContent: result,
        _meta: invocationMeta(WIDGETS.pendingApprovals),
      };
    },
  );

  // search_approvals
  server.registerTool(
    "search_approvals",
    {
      title: "Search Approvals",
      description:
        "Search and browse approvals matching criteria — returns a list. " +
        "Use for discovery queries like 'Which approvals are due this week?' or 'Show me high-risk approvals'. " +
        "Do NOT use when the user asks to view, open, or see details on a specific named approval — use fetch_approval_details with 'query' instead.",
      inputSchema: {
        query: z.string().describe("Search query — natural language or keywords."),
        status: z
          .array(z.enum(["pending", "approved", "rejected", "cancelled"]))
          .optional(),
        limit: z.number().int().min(1).max(50).optional(),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
      _meta: descriptorMeta(WIDGETS.pendingApprovals),
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      insertUsageEvent(actorId, actorName, "search_approvals", "searchResults");

      const result = getApprovals(
        {
          query: input.query,
          status: input.status as ApprovalFilters["status"],
          limit: input.limit ?? 20,
        },
        actorId,
      );

      return {
        content: [
          {
            type: "text",
            text: `Found ${result.items.length} approval(s) matching "${input.query}".`,
          },
        ],
        structuredContent: result,
        _meta: invocationMeta(WIDGETS.pendingApprovals),
      };
    },
  );

  // get_approval_details
  server.registerTool(
    "get_approval_details",
    {
      title: "Get Approval Details",
      description:
        "Returns full details for a single approval including fields, risk assessment, " +
        "approver chain, comments, audit history, and attachments. Opens the detail view widget.",
      inputSchema: {
        approval_id: z.string().describe("The approval ID (e.g. apr_1001)."),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
      _meta: descriptorMeta(WIDGETS.approvalDetail),
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      insertUsageEvent(actorId, actorName, "get_approval_details", "approvalDetail");
      const detail = getApprovalById(input.approval_id, actorId);

      if (!detail) {
        return {
          content: [
            {
              type: "text",
              text: `Approval ${input.approval_id} not found.`,
            },
          ],
          isError: true,
        };
      }

      return {
        content: [
          {
            type: "text",
            text: `Approval: ${detail.title} — ${detail.status} — ${detail.risk_level} risk.`,
          },
        ],
        structuredContent: detail,
        _meta: invocationMeta(WIDGETS.approvalDetail),
      };
    },
  );

  // fetch_approval_details — drill-down tool for the pending-approvals widget.
  // outputTemplate points back to pending-approvals.html so the host updates
  // toolOutput in the same widget (with ApprovalDetail payload), allowing the
  // widget to route to DetailPanel without switching widgets.
  server.registerTool(
    "fetch_approval_details",
    {
      title: "Fetch Approval Details (inline)",
      description:
        "Opens a specific approval in a full detail view. Pass the approval name or ID as 'query' — " +
        "the server resolves it in one call, no prior search needed. " +
        "IMPORTANT: When user says 'approve X' or 'reject X' from chat, " +
        "call this with open_dialog set to 'approve' or 'reject' — NEVER call approve_approval or reject_approval directly.",
      inputSchema: {
        query: z.string().describe("The approval name, title, or ID (e.g. 'Tokyo Sales Conference' or 'apr_1001'). The server resolves it automatically."),
        open_dialog: z
          .enum(["approve", "reject"])
          .optional()
          .describe("When set, opens the approve or reject confirmation dialog immediately on load, pre-filled with a suggested reason."),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
      _meta: {
        "openai/outputTemplate": WIDGETS.pendingApprovals.templateUri,
        "openai/toolInvocation/invoking": "Loading approval details…",
        "openai/toolInvocation/invoked": "Approval loaded",
        "openai/widgetAccessible": true,
      },
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      insertUsageEvent(actorId, actorName, "fetch_approval_details", "approvalDetail");
      // If query looks like an ID, use it directly; otherwise search by name
      const isId = /^apr_\d+$/i.test(input.query.trim());
      let resolvedId: string | undefined = isId ? input.query.trim() : undefined;
      if (!resolvedId) {
        const results = getApprovals({ query: input.query, limit: 1 }, actorId);
        resolvedId = results.items[0]?.id;
      }
      if (!resolvedId) {
        return {
          content: [{ type: "text", text: `No approval found matching "${input.query}".` }],
          isError: true,
        };
      }
      const detail = getApprovalById(resolvedId, actorId);
      if (!detail) {
        return {
          content: [{ type: "text", text: `Approval ${resolvedId} not found.` }],
          isError: true,
        };
      }
      return {
        content: [{ type: "text", text: `Approval: ${detail.title}` }],
        structuredContent: detail,
        _meta: { "openai/toolInvocation/invoked": "Approval loaded" },
      };
    },
  );

  // get_risk_assessment (data tool — no widget)
  server.registerTool(
    "get_risk_assessment",
    {
      title: "Get Risk Assessment",
      description:
        "Returns the backend risk assessment for an approval, including risk level and specific risk reasons. " +
        "Use before making decisions or when the user asks why something is risky.",
      inputSchema: {
        approval_id: z.string().describe("The approval ID."),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      insertUsageEvent(actorId, actorName, "get_risk_assessment");
      const detail = getApprovalById(input.approval_id);
      if (!detail) {
        return {
          content: [{ type: "text", text: `Approval ${input.approval_id} not found.` }],
          isError: true,
        };
      }

      const risk = detail.risk_assessment;
      const reasonsText = risk.risk_reasons.join("; ");

      return {
        content: [
          {
            type: "text",
            text: `Risk: ${risk.risk_level.toUpperCase()} — ${reasonsText}`,
          },
        ],
        structuredContent: risk,
      };
    },
  );

  // get_similar_approvals (data tool — no widget)
  server.registerTool(
    "get_similar_approvals",
    {
      title: "Get Similar Approvals",
      description:
        "Returns up to 5 previously submitted approvals of the same type for AI comparison context. " +
        "Use when the user asks to compare or check precedent for an approval.",
      inputSchema: {
        approval_id: z.string().describe("The approval ID to find similar items for."),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      insertUsageEvent(actorId, actorName, "get_similar_approvals");
      const similar = getSimilarApprovals(input.approval_id);

      return {
        content: [
          {
            type: "text",
            text: `Found ${similar.length} similar approval(s) of the same type.`,
          },
        ],
        structuredContent: { items: similar },
      };
    },
  );

  // show_create_form (launcher tool)
  server.registerTool(
    "show_create_form",
    {
      title: "Show Create Approval Form",
      description:
        "Opens the approval creation form widget. Use when the user wants to submit a new approval request. " +
        "ALWAYS extract as much context as possible from the user's message to prefill the form — " +
        "type, title, summary, amount, vendor, department, dates, etc. " +
        "The more you prefill, the less the user has to type.",
      inputSchema: {
        type_hint: z
          .enum([
            "promo_pricing",
            "purchase_order",
            "capex",
            "contract_sow",
            "travel_exception",
            "marketing_budget",
            "vendor_onboarding",
          ])
          .optional()
          .describe("Approval type inferred from context (capex, purchase_order, travel_exception, etc.)."),
        title: z
          .string()
          .optional()
          .describe("Pre-filled title derived from the user's description, e.g. 'New Data Center Equipment'."),
        summary: z
          .string()
          .optional()
          .describe("Pre-filled business justification or summary extracted from the conversation."),
        prefill_fields: z
          .union([z.record(z.unknown()), z.string()])
          .optional()
          .describe(
            "Key-value map of type-specific field values to pre-populate (object or JSON-encoded string). " +
            "Common keys by type — " +
            "capex: amount, project_name, department, justification, expected_roi; " +
            "purchase_order: amount, currency, vendor, department, justification; " +
            "travel_exception: traveler_name, destination, travel_dates, estimated_cost, policy_exception_reason; " +
            "marketing_budget: campaign_name, channel, amount, target_audience, start_date, end_date; " +
            "contract_sow: vendor_agency, contract_value, currency, start_date, end_date, deliverables; " +
            "vendor_onboarding: vendor_name, service_category, annual_value, compliance_status; " +
            "promo_pricing: publisher, retailer, country, discount_pct, start_date, end_date. " +
            "Only include keys where values can be reasonably inferred from context."
          ),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
      _meta: descriptorMeta(WIDGETS.createApproval),
    },
    async (input) => {
      const typeName = input.type_hint ? input.type_hint.replace(/_/g, " ") : "approval";
      let prefill: Record<string, unknown> | null = null;
      if (typeof input.prefill_fields === "string") {
        try { prefill = JSON.parse(input.prefill_fields) as Record<string, unknown>; } catch { prefill = null; }
      } else if (input.prefill_fields) {
        prefill = input.prefill_fields as Record<string, unknown>;
      }
      return {
        content: [
          {
            type: "text",
            text: `Opening ${typeName} creation form${input.title ? ` for "${input.title}"` : ""}.`,
          },
        ],
        structuredContent: {
          type_hint: input.type_hint ?? null,
          title: input.title ?? null,
          summary: input.summary ?? null,
          prefill_fields: prefill,
        },
        _meta: invocationMeta(WIDGETS.createApproval),
      };
    },
  );

  // ── WRITE TOOLS ──────────────────────────────────────────────────────────

  // approve_approval
  server.registerTool(
    "approve_approval",
    {
      title: "Approve Approval",
      description:
        "Approves a single pending approval. Only call this from within the widget confirm dialog — " +
        "NEVER call directly in response to a user chat message. " +
        "For chat-initiated approvals, use fetch_approval_details with open_dialog='approve' instead.",
      inputSchema: {
        approval_id: z.string().describe("The approval ID to approve."),
        comment: z
          .string()
          .optional()
          .describe("Optional comment to attach to the approval decision."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      insertUsageEvent(actorId, actorName, "approve_approval");
      const result = approveApproval(input.approval_id, actorId, actorName, input.comment);

      if (!result.success) {
        return {
          content: [{ type: "text", text: `Failed to approve: ${result.error}` }],
          isError: true,
          structuredContent: { success: false, error: result.error },
        };
      }

      return {
        content: [{ type: "text", text: `Approval ${input.approval_id} approved successfully.` }],
        structuredContent: { success: true, approval_id: input.approval_id },
      };
    },
  );

  // reject_approval
  server.registerTool(
    "reject_approval",
    {
      title: "Reject Approval",
      description:
        "Rejects a single pending approval. A rejection reason is required. Only call this from " +
        "within the widget reject dialog — NEVER call directly in response to a user chat message. " +
        "For chat-initiated rejections, use fetch_approval_details with open_dialog='reject' instead.",
      inputSchema: {
        approval_id: z.string().describe("The approval ID to reject."),
        reason: z
          .string()
          .min(1)
          .describe("Rejection reason — required, will be stored in audit history."),
        comment: z
          .string()
          .optional()
          .describe("Optional additional comment."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      insertUsageEvent(actorId, actorName, "reject_approval");
      const result = rejectApproval(
        input.approval_id,
        actorId,
        actorName,
        input.reason,
        input.comment,
      );

      if (!result.success) {
        return {
          content: [{ type: "text", text: `Failed to reject: ${result.error}` }],
          isError: true,
          structuredContent: { success: false, error: result.error },
        };
      }

      return {
        content: [
          { type: "text", text: `Approval ${input.approval_id} rejected. Reason: ${input.reason}` },
        ],
        structuredContent: { success: true, approval_id: input.approval_id },
      };
    },
  );

  // bulk_decide_approvals
  server.registerTool(
    "bulk_decide_approvals",
    {
      title: "Bulk Decide Approvals",
      description:
        "Approve or reject multiple approvals at once. The user must have reviewed and confirmed " +
        "the batch in the widget before this is called. Returns per-item results.",
      inputSchema: {
        approval_ids: z
          .array(z.string())
          .min(1)
          .describe("List of approval IDs to act on."),
        decision: z
          .enum(["approve", "reject"])
          .describe("The decision to apply to all selected approvals."),
        comment: z
          .string()
          .optional()
          .describe("Shared comment applied to all items."),
        reason: z
          .string()
          .optional()
          .describe("Rejection reason — required when decision is 'reject'."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
    },
    async (input, extra) => {
      if (input.decision === "reject" && !input.reason) {
        return {
          content: [{ type: "text", text: "Rejection reason is required for bulk reject." }],
          isError: true,
          structuredContent: { success: false, error: "reason required" },
        };
      }

      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      insertUsageEvent(actorId, actorName, "bulk_decide_approvals");
      const results = bulkDecide(
        input.approval_ids,
        input.decision,
        actorId,
        actorName,
        input.comment,
        input.reason,
      );

      const succeeded = results.filter((r) => r.success).length;
      const failed = results.length - succeeded;

      return {
        content: [
          {
            type: "text",
            text: `Bulk ${input.decision}: ${succeeded} succeeded, ${failed} failed.`,
          },
        ],
        structuredContent: { results, succeeded, failed },
      };
    },
  );

  // add_comment_to_approval
  server.registerTool(
    "add_comment_to_approval",
    {
      title: "Add Comment to Approval",
      description:
        "Adds a comment to an approval without changing its status. " +
        "Use for notes, questions to the requester, or follow-up context.",
      inputSchema: {
        approval_id: z.string().describe("The approval ID."),
        body: z.string().min(1).describe("Comment text."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      const comment = addComment(input.approval_id, actorId, actorName, input.body);

      return {
        content: [
          { type: "text", text: `Comment added to approval ${input.approval_id}.` },
        ],
        structuredContent: { success: true, comment },
      };
    },
  );

  // create_approval_request
  server.registerTool(
    "create_approval_request",
    {
      title: "Create Approval Request",
      description:
        "Creates a new approval request. Risk level is computed by the backend. " +
        "Returns the new approval ID and computed risk. " +
        "Typically called after the user completes the create form widget.",
      inputSchema: {
        type: z
          .enum([
            "promo_pricing",
            "purchase_order",
            "capex",
            "contract_sow",
            "travel_exception",
            "marketing_budget",
            "vendor_onboarding",
          ])
          .describe("Approval type."),
        title: z.string().min(1).describe("Short descriptive title."),
        summary: z.string().min(1).describe("Business justification or request summary."),
        fields: z
          .array(z.object({ label: z.string(), value: z.string() }))
          .describe("Dynamic fields for this approval type."),
        attachment_ids: z
          .array(z.string())
          .optional()
          .describe("File IDs uploaded via window.openai.uploadFile()."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
    },
    async (input, extra) => {
      const { actorId, actorName } = resolveActor(extra?.meta as RequestMeta | undefined);
      insertUsageEvent(actorId, actorName, "create_approval_request", "createApproval");
      const defaultAssignees = ["mgr_201", "mgr_202"]; // TODO: route by type in Phase 6

      const result = createApproval(
        {
          type: input.type as ApprovalType,
          title: input.title,
          summary: input.summary,
          requester_id: actorId,
          fields: input.fields,
          attachment_ids: input.attachment_ids,
        },
        defaultAssignees,
      );

      return {
        content: [
          {
            type: "text",
            text: `New approval created: ${result.id} — ${result.risk_level} risk. Assigned to ${result.assignee_ids.join(", ")}.`,
          },
        ],
        structuredContent: result,
      };
    },
  );

  return server;
}
