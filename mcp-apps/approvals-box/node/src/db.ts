// ─────────────────────────────────────────────────────────────────────────────
// SQLite data layer — node:sqlite (built-in, Node 22+, requires --experimental-sqlite)
// ─────────────────────────────────────────────────────────────────────────────

// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore — node:sqlite is experimental; types land in @types/node ≥22.x
import { DatabaseSync } from "node:sqlite";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type {
  Approval,
  ApprovalListItem,
  Attachment,
  AuditEvent,
  Comment,
  User,
  ApprovalFilters,
  ApprovalStatus,
  RiskLevel,
  ApprovalType,
  ApprovalSummary,
  ListApprovalsOutput,
  ApprovalDetail,
  AttachmentPreviewOutput,
} from "./types.js";
import { computeRisk, computeAllowedActions } from "./risk.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DB_PATH = process.env.DB_PATH ?? path.resolve(__dirname, "..", "approvals.db");

// ─────────────────────────────────────────────────────────────────────────────
// Database initialisation
// ─────────────────────────────────────────────────────────────────────────────

// node:sqlite DatabaseSync type (experimental — no official @types yet)
type NodeSqliteDb = InstanceType<typeof DatabaseSync>;

let _db: NodeSqliteDb | null = null;

export function getDb(): NodeSqliteDb {
  if (!_db) {
    _db = new DatabaseSync(DB_PATH);
    (_db as NodeSqliteDb).exec("PRAGMA journal_mode = WAL");
    (_db as NodeSqliteDb).exec("PRAGMA foreign_keys = ON");
    initSchema(_db);
  }
  return _db;
}

function initSchema(db: NodeSqliteDb): void {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id           TEXT PRIMARY KEY,
      name         TEXT NOT NULL,
      email        TEXT NOT NULL,
      role         TEXT NOT NULL DEFAULT 'approver',
      subject_id   TEXT
    );

    CREATE TABLE IF NOT EXISTS approvals (
      id              TEXT PRIMARY KEY,
      title           TEXT NOT NULL,
      type            TEXT NOT NULL,
      display_type    TEXT NOT NULL,
      status          TEXT NOT NULL DEFAULT 'pending',
      risk_level      TEXT NOT NULL DEFAULT 'low',
      requester_id    TEXT NOT NULL,
      requester_name  TEXT NOT NULL,
      assignee_ids    TEXT NOT NULL DEFAULT '[]',   -- JSON array
      submitted_at    TEXT NOT NULL,
      due_at          TEXT NOT NULL,
      amount          REAL NOT NULL DEFAULT 0,
      currency        TEXT NOT NULL DEFAULT 'USD',
      country         TEXT NOT NULL DEFAULT 'USA',
      summary         TEXT NOT NULL DEFAULT '',
      fields          TEXT NOT NULL DEFAULT '[]',   -- JSON array
      attachment_ids  TEXT NOT NULL DEFAULT '[]',   -- JSON array
      comment_ids     TEXT NOT NULL DEFAULT '[]',   -- JSON array
      history_ids     TEXT NOT NULL DEFAULT '[]',   -- JSON array
      allowed_actions TEXT NOT NULL DEFAULT '[]'    -- JSON array
    );

    CREATE TABLE IF NOT EXISTS risk_assessments (
      approval_id             TEXT PRIMARY KEY,
      risk_level              TEXT NOT NULL,
      risk_reasons            TEXT NOT NULL DEFAULT '[]', -- JSON array
      recommended_review_mode TEXT NOT NULL DEFAULT 'fullscreen'
    );

    CREATE TABLE IF NOT EXISTS attachments (
      id               TEXT PRIMARY KEY,
      approval_id      TEXT NOT NULL,
      name             TEXT NOT NULL,
      mime_type        TEXT NOT NULL,
      size_bytes       INTEGER NOT NULL DEFAULT 0,
      preview_supported INTEGER NOT NULL DEFAULT 0,
      storage_key      TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS comments (
      id          TEXT PRIMARY KEY,
      approval_id TEXT NOT NULL,
      author_id   TEXT NOT NULL,
      author_name TEXT NOT NULL,
      body        TEXT NOT NULL,
      created_at  TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS audit_events (
      id          TEXT PRIMARY KEY,
      approval_id TEXT NOT NULL,
      actor_id    TEXT NOT NULL,
      actor_name  TEXT NOT NULL,
      action      TEXT NOT NULL,
      reason      TEXT,
      comment     TEXT,
      timestamp   TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS usage_log (
      id         INTEGER PRIMARY KEY AUTOINCREMENT,
      ts         TEXT NOT NULL,
      actor_id   TEXT NOT NULL,
      actor_name TEXT NOT NULL,
      tool       TEXT NOT NULL,
      widget     TEXT
    );
  `);
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function parseJSON<T>(str: string, fallback: T): T {
  try {
    return JSON.parse(str) as T;
  } catch {
    return fallback;
  }
}

function rowToApproval(row: Record<string, unknown>): Approval {
  return {
    id: row.id as string,
    title: row.title as string,
    type: row.type as ApprovalType,
    display_type: row.display_type as string,
    status: row.status as ApprovalStatus,
    risk_level: row.risk_level as RiskLevel,
    requester_id: row.requester_id as string,
    requester_name: row.requester_name as string,
    assignee_ids: parseJSON<string[]>(row.assignee_ids as string, []),
    submitted_at: row.submitted_at as string,
    due_at: row.due_at as string,
    amount: row.amount as number,
    currency: row.currency as string,
    country: row.country as string,
    summary: row.summary as string,
    fields: parseJSON(row.fields as string, []),
    attachment_ids: parseJSON<string[]>(row.attachment_ids as string, []),
    comment_ids: parseJSON<string[]>(row.comment_ids as string, []),
    history_ids: parseJSON<string[]>(row.history_ids as string, []),
    allowed_actions: parseJSON<string[]>(row.allowed_actions as string, []),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Users
// ─────────────────────────────────────────────────────────────────────────────

export function upsertUser(user: User): void {
  const db = getDb();
  db.prepare(`
    INSERT INTO users (id, name, email, role, subject_id)
    VALUES (@id, @name, @email, @role, @subject_id)
    ON CONFLICT(id) DO UPDATE SET
      name = excluded.name,
      email = excluded.email,
      role = excluded.role,
      subject_id = excluded.subject_id
  `).run(user);
}

export function getUserById(id: string): User | undefined {
  const db = getDb();
  return db.prepare("SELECT * FROM users WHERE id = ?").get(id) as User | undefined;
}

export function getUserBySubject(subjectId: string): User | undefined {
  const db = getDb();
  return db.prepare("SELECT * FROM users WHERE subject_id = ?").get(subjectId) as User | undefined;
}

// ─────────────────────────────────────────────────────────────────────────────
// Approvals — Read
// ─────────────────────────────────────────────────────────────────────────────

export function getApprovals(filters: ApprovalFilters = {}, actorId = "mgr_201"): ListApprovalsOutput {
  const db = getDb();

  let sql = "SELECT * FROM approvals WHERE 1=1";
  const params: unknown[] = [];

  // Status filter
  if (filters.status && filters.status.length > 0) {
    sql += ` AND status IN (${filters.status.map(() => "?").join(",")})`;
    params.push(...filters.status);
  }

  // Risk level filter
  if (filters.risk_level && filters.risk_level.length > 0) {
    sql += ` AND risk_level IN (${filters.risk_level.map(() => "?").join(",")})`;
    params.push(...filters.risk_level);
  }

  // Type filter
  if (filters.type && filters.type.length > 0) {
    sql += ` AND type IN (${filters.type.map(() => "?").join(",")})`;
    params.push(...filters.type);
  }

  // Requester filter
  if (filters.requester) {
    sql += " AND LOWER(requester_name) LIKE ?";
    params.push(`%${filters.requester.toLowerCase()}%`);
  }

  // Has attachments filter
  if (filters.has_attachments === true) {
    sql += " AND attachment_ids != '[]'";
  } else if (filters.has_attachments === false) {
    sql += " AND attachment_ids = '[]'";
  }

  // Due before filter
  if (filters.due_before) {
    sql += " AND due_at <= ?";
    params.push(filters.due_before);
  }

  // Text search across title, summary, requester name
  if (filters.query) {
    const q = `%${filters.query.toLowerCase()}%`;
    sql += " AND (LOWER(title) LIKE ? OR LOWER(summary) LIKE ? OR LOWER(requester_name) LIKE ? OR LOWER(fields) LIKE ?)";
    params.push(q, q, q, q);
  }

  // Sort
  switch (filters.sort) {
    case "due_date":
      sql += " ORDER BY due_at ASC";
      break;
    case "amount":
      sql += " ORDER BY amount DESC";
      break;
    case "submitted_date":
      sql += " ORDER BY submitted_at DESC";
      break;
    case "risk":
      sql += " ORDER BY CASE risk_level WHEN 'high' THEN 0 WHEN 'medium' THEN 1 ELSE 2 END ASC";
      break;
    default:
      sql += " ORDER BY due_at ASC";
  }

  // Cursor-based pagination (cursor is the last ID seen)
  const limit = Math.min(filters.limit ?? 20, 100);

  if (filters.cursor) {
    // Simple offset cursor: cursor = base64(offset)
    const offset = parseInt(Buffer.from(filters.cursor, "base64").toString("utf8"), 10);
    sql += ` LIMIT ${limit} OFFSET ${offset}`;
  } else {
    sql += ` LIMIT ${limit}`;
  }

  const rows = db.prepare(sql).all(...params) as Record<string, unknown>[];
  const approvals = rows.map(rowToApproval);

  // Build base SQL without pagination for summary counts
  const baseSql = sql
    .replace(/ ORDER BY .+/, "")
    .replace(/ LIMIT \d+( OFFSET \d+)?/, "");

  const totalRow = db.prepare(baseSql.replace(/SELECT \*/, "SELECT COUNT(*) as cnt")).get(...params) as { cnt: number } | undefined;
  const total = totalRow?.cnt ?? approvals.length;

  const now = Date.now();
  const in24h = now + 24 * 60 * 60 * 1000;
  const nowIso = new Date(now).toISOString();
  const in24hIso = new Date(in24h).toISOString();

  // Summary counts from all matching rows (not just current page)
  const urgentRow = db.prepare(baseSql.replace(/SELECT \*/, "SELECT COUNT(*) as cnt") + ` AND status = 'pending' AND due_at > ? AND due_at <= ?`).get(...params, nowIso, in24hIso) as { cnt: number } | undefined;
  const overdueRow = db.prepare(baseSql.replace(/SELECT \*/, "SELECT COUNT(*) as cnt") + ` AND status = 'pending' AND due_at < ?`).get(...params, nowIso) as { cnt: number } | undefined;
  const highRiskRow = db.prepare(baseSql.replace(/SELECT \*/, "SELECT COUNT(*) as cnt") + ` AND risk_level = 'high'`).get(...params) as { cnt: number } | undefined;
  const amountRow = db.prepare(baseSql.replace(/SELECT \*/, "SELECT COALESCE(SUM(amount),0) as total") + ` AND status = 'pending'`).get(...params) as { total: number } | undefined;

  const summary: ApprovalSummary = {
    total_pending: total,
    urgent: urgentRow?.cnt ?? 0,
    overdue: overdueRow?.cnt ?? 0,
    high_risk: highRiskRow?.cnt ?? 0,
    total_amount_pending: amountRow?.total ?? 0,
  };

  const offset = filters.cursor
    ? parseInt(Buffer.from(filters.cursor, "base64").toString("utf8"), 10)
    : 0;
  const nextOffset = offset + limit;
  const next_cursor =
    nextOffset < total
      ? Buffer.from(String(nextOffset)).toString("base64")
      : null;

  const items: ApprovalListItem[] = approvals.map((a) => ({
    id: a.id,
    title: a.title,
    type: a.type,
    display_type: a.display_type,
    requester_name: a.requester_name,
    submitted_at: a.submitted_at,
    due_at: a.due_at,
    amount: a.amount,
    currency: a.currency,
    risk_level: a.risk_level,
    has_attachments: a.attachment_ids.length > 0,
    status: a.status,
    allowed_actions: computeAllowedActions(a.status, actorId, a.assignee_ids),
  }));

  return { summary, items, next_cursor };
}

export function getApprovalById(id: string, actorId = "mgr_201"): ApprovalDetail | null {
  const db = getDb();
  const row = db.prepare("SELECT * FROM approvals WHERE id = ?").get(id) as
    | Record<string, unknown>
    | undefined;
  if (!row) return null;

  const approval = rowToApproval(row);
  approval.allowed_actions = computeAllowedActions(approval.status, actorId, approval.assignee_ids);

  // Risk
  const riskRow = db
    .prepare("SELECT * FROM risk_assessments WHERE approval_id = ?")
    .get(id) as Record<string, unknown> | undefined;
  const risk_assessment = riskRow
    ? {
        approval_id: riskRow.approval_id as string,
        risk_level: riskRow.risk_level as RiskLevel,
        risk_reasons: parseJSON<string[]>(riskRow.risk_reasons as string, []),
        recommended_review_mode: riskRow.recommended_review_mode as "inline" | "fullscreen",
      }
    : computeRisk({
        id,
        type: approval.type,
        amount: approval.amount,
        currency: approval.currency,
        country: approval.country,
        fields: approval.fields,
        attachment_ids: approval.attachment_ids,
        due_at: approval.due_at,
        summary: approval.summary,
      });

  // Attachments
  const attachments = db
    .prepare("SELECT * FROM attachments WHERE approval_id = ?")
    .all(id) as Attachment[];

  // Comments
  const comments = db
    .prepare("SELECT * FROM comments WHERE approval_id = ? ORDER BY created_at ASC")
    .all(id) as Comment[];

  // Audit history
  const history = db
    .prepare("SELECT * FROM audit_events WHERE approval_id = ? ORDER BY timestamp ASC")
    .all(id) as AuditEvent[];

  // Approver chain
  const approver_chain = approval.assignee_ids.map((uid) => {
    const user = getUserById(uid);
    return {
      user_id: uid,
      name: user?.name ?? uid,
      status: approval.status === "pending" ? "pending" : approval.status,
    };
  });

  return {
    ...approval,
    risk_assessment,
    approver_chain,
    comments,
    history,
    attachments,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Attachments
// ─────────────────────────────────────────────────────────────────────────────

export function getAttachmentById(id: string): Attachment | undefined {
  const db = getDb();
  return db.prepare("SELECT * FROM attachments WHERE id = ?").get(id) as
    | Attachment
    | undefined;
}

export function getAttachmentPreview(id: string, baseUrl: string): AttachmentPreviewOutput | null { // sync
  const att = getAttachmentById(id);
  if (!att) return null;

  const download_url = `${baseUrl}/attachments/${att.storage_key}`;

  // For .txt files, attempt to read preview text from local storage
  let preview_text: string | undefined;
  if (att.mime_type === "text/plain" && att.preview_supported) {
    try {
      const storagePath = path.resolve(__dirname, "..", "storage", att.storage_key);
      const content = fs.readFileSync(storagePath, "utf-8");
      preview_text = content.slice(0, 2000); // first 2000 chars
    } catch {
      // File not found in storage — just return metadata
    }
  }

  return {
    id: att.id,
    approval_id: att.approval_id,
    name: att.name,
    mime_type: att.mime_type,
    size_bytes: att.size_bytes,
    preview_supported: att.preview_supported,
    preview_text,
    download_url,
  };
}

export function insertAttachment(att: Attachment): void {
  const db = getDb();
  db.prepare(`
    INSERT INTO attachments (id, approval_id, name, mime_type, size_bytes, preview_supported, storage_key)
    VALUES (@id, @approval_id, @name, @mime_type, @size_bytes, @preview_supported, @storage_key)
    ON CONFLICT(id) DO NOTHING
  `).run({ ...att, preview_supported: att.preview_supported ? 1 : 0 });
}

// ─────────────────────────────────────────────────────────────────────────────
// Write actions
// ─────────────────────────────────────────────────────────────────────────────

export function approveApproval(
  approvalId: string,
  actorId: string,
  actorName: string,
  comment?: string,
): { success: boolean; error?: string } {
  const db = getDb();
  const row = db.prepare("SELECT * FROM approvals WHERE id = ?").get(approvalId) as
    | Record<string, unknown>
    | undefined;

  if (!row) return { success: false, error: `Approval ${approvalId} not found` };

  const approval = rowToApproval(row);
  const allowed = computeAllowedActions(approval.status, actorId, approval.assignee_ids);

  if (!allowed.includes("approve")) {
    return {
      success: false,
      error: approval.status !== "pending"
        ? `Cannot approve: approval is already ${approval.status}`
        : "You are not authorized to approve this request",
    };
  }

  db.prepare("UPDATE approvals SET status = 'approved' WHERE id = ?").run(approvalId);
  writeAuditEvent(db, {
    id: `evt_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    approval_id: approvalId,
    actor_id: actorId,
    actor_name: actorName,
    action: "approved",
    comment,
    timestamp: new Date().toISOString(),
  });

  return { success: true };
}

export function rejectApproval(
  approvalId: string,
  actorId: string,
  actorName: string,
  reason: string,
  comment?: string,
): { success: boolean; error?: string } {
  const db = getDb();
  const row = db.prepare("SELECT * FROM approvals WHERE id = ?").get(approvalId) as
    | Record<string, unknown>
    | undefined;

  if (!row) return { success: false, error: `Approval ${approvalId} not found` };

  const approval = rowToApproval(row);
  const allowed = computeAllowedActions(approval.status, actorId, approval.assignee_ids);

  if (!allowed.includes("reject")) {
    return {
      success: false,
      error: approval.status !== "pending"
        ? `Cannot reject: approval is already ${approval.status}`
        : "You are not authorized to reject this request",
    };
  }

  if (!reason || reason.trim().length === 0) {
    return { success: false, error: "Rejection reason is required" };
  }

  db.prepare("UPDATE approvals SET status = 'rejected' WHERE id = ?").run(approvalId);
  writeAuditEvent(db, {
    id: `evt_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    approval_id: approvalId,
    actor_id: actorId,
    actor_name: actorName,
    action: "rejected",
    reason: reason.trim(),
    comment,
    timestamp: new Date().toISOString(),
  });

  return { success: true };
}

export interface BulkDecideResult {
  approval_id: string;
  success: boolean;
  error?: string;
}

export function bulkDecide(
  approvalIds: string[],
  decision: "approve" | "reject",
  actorId: string,
  actorName: string,
  comment?: string,
  reason?: string,
): BulkDecideResult[] {
  return approvalIds.map((id) => {
    if (decision === "approve") {
      const result = approveApproval(id, actorId, actorName, comment);
      return { approval_id: id, ...result };
    } else {
      const result = rejectApproval(id, actorId, actorName, reason ?? comment ?? "", comment);
      return { approval_id: id, ...result };
    }
  });
}

export function addComment(
  approvalId: string,
  authorId: string,
  authorName: string,
  body: string,
): Comment {
  const db = getDb();
  const comment: Comment = {
    id: `com_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    approval_id: approvalId,
    author_id: authorId,
    author_name: authorName,
    body: body.trim(),
    created_at: new Date().toISOString(),
  };

  db.prepare(`
    INSERT INTO comments (id, approval_id, author_id, author_name, body, created_at)
    VALUES (@id, @approval_id, @author_id, @author_name, @body, @created_at)
  `).run(comment);

  // Append to approval's comment_ids
  const row = db.prepare("SELECT comment_ids FROM approvals WHERE id = ?").get(approvalId) as
    | { comment_ids: string }
    | undefined;
  if (row) {
    const ids = parseJSON<string[]>(row.comment_ids, []);
    ids.push(comment.id);
    db.prepare("UPDATE approvals SET comment_ids = ? WHERE id = ?").run(
      JSON.stringify(ids),
      approvalId,
    );
  }

  writeAuditEvent(db, {
    id: `evt_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    approval_id: approvalId,
    actor_id: authorId,
    actor_name: authorName,
    action: "commented",
    comment: body.trim(),
    timestamp: new Date().toISOString(),
  });

  return comment;
}

// ─────────────────────────────────────────────────────────────────────────────
// Create approval
// ─────────────────────────────────────────────────────────────────────────────

export function createApproval(
  input: import("./types.js").CreateApprovalInput,
  defaultAssigneeIds: string[],
): import("./types.js").CreateApprovalOutput {
  const db = getDb();
  const id = `apr_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
  const now = new Date().toISOString();

  // Due date: 3 business days from now
  const due = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString();

  const risk = computeRisk({
    id,
    type: input.type,
    amount: Number(input.fields.find((f) => f.label.toLowerCase().includes("amount"))?.value ?? 0),
    currency: input.fields.find((f) => f.label.toLowerCase().includes("currency"))?.value ?? "USD",
    country: input.fields.find((f) => f.label.toLowerCase().includes("country"))?.value ?? "USA",
    fields: input.fields,
    attachment_ids: input.attachment_ids ?? [],
    due_at: due,
    summary: input.summary,
  });

  const approval: Omit<Approval, "comment_ids" | "history_ids" | "allowed_actions"> & {
    comment_ids: string[];
    history_ids: string[];
    allowed_actions: string[];
  } = {
    id,
    title: input.title,
    type: input.type,
    display_type: input.type
      .split("_")
      .map((w) => w[0].toUpperCase() + w.slice(1))
      .join(" "),
    status: "pending",
    risk_level: risk.risk_level,
    requester_id: input.requester_id,
    requester_name:
      getUserById(input.requester_id)?.name ?? input.requester_id,
    assignee_ids: defaultAssigneeIds,
    submitted_at: now,
    due_at: due,
    amount: Number(input.fields.find((f) => f.label.toLowerCase().includes("amount"))?.value ?? 0),
    currency: input.fields.find((f) => f.label.toLowerCase().includes("currency"))?.value ?? "USD",
    country: input.fields.find((f) => f.label.toLowerCase().includes("country"))?.value ?? "USA",
    summary: input.summary,
    fields: input.fields,
    attachment_ids: input.attachment_ids ?? [],
    comment_ids: [],
    history_ids: [],
    allowed_actions: [],
  };

  db.prepare(`
    INSERT INTO approvals (
      id, title, type, display_type, status, risk_level,
      requester_id, requester_name, assignee_ids,
      submitted_at, due_at, amount, currency, country,
      summary, fields, attachment_ids, comment_ids, history_ids, allowed_actions
    ) VALUES (
      @id, @title, @type, @display_type, @status, @risk_level,
      @requester_id, @requester_name, @assignee_ids,
      @submitted_at, @due_at, @amount, @currency, @country,
      @summary, @fields, @attachment_ids, @comment_ids, @history_ids, @allowed_actions
    )
  `).run({
    ...approval,
    assignee_ids: JSON.stringify(approval.assignee_ids),
    fields: JSON.stringify(approval.fields),
    attachment_ids: JSON.stringify(approval.attachment_ids),
    comment_ids: JSON.stringify(approval.comment_ids),
    history_ids: JSON.stringify(approval.history_ids),
    allowed_actions: JSON.stringify(approval.allowed_actions),
  });

  // Save risk assessment
  db.prepare(`
    INSERT INTO risk_assessments (approval_id, risk_level, risk_reasons, recommended_review_mode)
    VALUES (@approval_id, @risk_level, @risk_reasons, @recommended_review_mode)
    ON CONFLICT(approval_id) DO UPDATE SET
      risk_level = excluded.risk_level,
      risk_reasons = excluded.risk_reasons,
      recommended_review_mode = excluded.recommended_review_mode
  `).run({
    approval_id: id,
    risk_level: risk.risk_level,
    risk_reasons: JSON.stringify(risk.risk_reasons),
    recommended_review_mode: risk.recommended_review_mode,
  });

  // Write submitted audit event
  writeAuditEvent(db, {
    id: `evt_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    approval_id: id,
    actor_id: input.requester_id,
    actor_name: approval.requester_name,
    action: "submitted",
    timestamp: now,
  });

  return {
    id,
    status: "pending",
    risk_level: risk.risk_level,
    risk_reasons: risk.risk_reasons,
    assignee_ids: defaultAssigneeIds,
    submitted_at: now,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Similar approvals (for AI context)
// ─────────────────────────────────────────────────────────────────────────────

export function getSimilarApprovals(approvalId: string, limit = 5): ApprovalListItem[] {
  const db = getDb();
  const row = db.prepare("SELECT type, risk_level FROM approvals WHERE id = ?").get(approvalId) as
    | { type: string; risk_level: string }
    | undefined;
  if (!row) return [];

  const rows = db
    .prepare(
      `SELECT * FROM approvals
       WHERE type = ? AND id != ?
       ORDER BY submitted_at DESC
       LIMIT ?`,
    )
    .all(row.type, approvalId, limit) as Record<string, unknown>[];

  return rows.map(rowToApproval).map((a) => ({
    id: a.id,
    title: a.title,
    display_type: a.display_type,
    requester_name: a.requester_name,
    submitted_at: a.submitted_at,
    due_at: a.due_at,
    amount: a.amount,
    currency: a.currency,
    risk_level: a.risk_level,
    has_attachments: a.attachment_ids.length > 0,
    status: a.status,
    allowed_actions: [],
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Usage logging
// ─────────────────────────────────────────────────────────────────────────────

export function insertUsageEvent(
  actorId: string,
  actorName: string,
  tool: string,
  widget?: string,
): void {
  getDb()
    .prepare(
      "INSERT INTO usage_log (ts, actor_id, actor_name, tool, widget) VALUES (?, ?, ?, ?, ?)",
    )
    .run(new Date().toISOString(), actorId, actorName, tool, widget ?? null);
}

export function getUsageStats(): {
  total_calls: number;
  unique_users: number;
  by_user: { actor_id: string; actor_name: string; calls: number; last_seen: string }[];
  by_tool: { tool: string; calls: number }[];
  by_widget: { widget: string; shown: number }[];
  recent: { ts: string; actor_name: string; tool: string; widget: string | null }[];
} {
  const db = getDb();
  const total_calls = (db.prepare("SELECT COUNT(*) AS n FROM usage_log").get() as { n: number }).n;
  const unique_users = (db.prepare("SELECT COUNT(DISTINCT actor_id) AS n FROM usage_log").get() as { n: number }).n;
  const by_user = db
    .prepare("SELECT actor_id, actor_name, COUNT(*) AS calls, MAX(ts) AS last_seen FROM usage_log GROUP BY actor_id ORDER BY calls DESC")
    .all() as { actor_id: string; actor_name: string; calls: number; last_seen: string }[];
  const by_tool = db
    .prepare("SELECT tool, COUNT(*) AS calls FROM usage_log GROUP BY tool ORDER BY calls DESC")
    .all() as { tool: string; calls: number }[];
  const by_widget = db
    .prepare("SELECT widget, COUNT(*) AS shown FROM usage_log WHERE widget IS NOT NULL GROUP BY widget ORDER BY shown DESC")
    .all() as { widget: string; shown: number }[];
  const recent = db
    .prepare("SELECT ts, actor_name, tool, widget FROM usage_log ORDER BY id DESC LIMIT 50")
    .all() as { ts: string; actor_name: string; tool: string; widget: string | null }[];
  return { total_calls, unique_users, by_user, by_tool, by_widget, recent };
}

// ─────────────────────────────────────────────────────────────────────────────
// Internal helpers
// ─────────────────────────────────────────────────────────────────────────────

function writeAuditEvent(
  db: NodeSqliteDb,
  event: Omit<AuditEvent, "reason" | "comment"> & { reason?: string; comment?: string },
): void {
  db.prepare(`
    INSERT INTO audit_events (id, approval_id, actor_id, actor_name, action, reason, comment, timestamp)
    VALUES (@id, @approval_id, @actor_id, @actor_name, @action, @reason, @comment, @timestamp)
  `).run({
    ...event,
    reason: event.reason ?? null,
    comment: event.comment ?? null,
  });

  // Append to approval's history_ids
  const row = db
    .prepare("SELECT history_ids FROM approvals WHERE id = ?")
    .get(event.approval_id) as { history_ids: string } | undefined;
  if (row) {
    const ids = parseJSON<string[]>(row.history_ids, []);
    ids.push(event.id);
    db.prepare("UPDATE approvals SET history_ids = ? WHERE id = ?").run(
      JSON.stringify(ids),
      event.approval_id,
    );
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Seed helpers (used by seed.ts)
// ─────────────────────────────────────────────────────────────────────────────

export function insertApproval(approval: Approval): void {
  const db = getDb();
  db.prepare(`
    INSERT INTO approvals (
      id, title, type, display_type, status, risk_level,
      requester_id, requester_name, assignee_ids,
      submitted_at, due_at, amount, currency, country,
      summary, fields, attachment_ids, comment_ids, history_ids, allowed_actions
    ) VALUES (
      @id, @title, @type, @display_type, @status, @risk_level,
      @requester_id, @requester_name, @assignee_ids,
      @submitted_at, @due_at, @amount, @currency, @country,
      @summary, @fields, @attachment_ids, @comment_ids, @history_ids, @allowed_actions
    )
    ON CONFLICT(id) DO NOTHING
  `).run({
    ...approval,
    assignee_ids: JSON.stringify(approval.assignee_ids),
    fields: JSON.stringify(approval.fields),
    attachment_ids: JSON.stringify(approval.attachment_ids),
    comment_ids: JSON.stringify(approval.comment_ids),
    history_ids: JSON.stringify(approval.history_ids),
    allowed_actions: JSON.stringify(approval.allowed_actions),
  });
}

export function insertRiskAssessment(r: import("./types.js").RiskAssessment): void {
  const db = getDb();
  db.prepare(`
    INSERT INTO risk_assessments (approval_id, risk_level, risk_reasons, recommended_review_mode)
    VALUES (@approval_id, @risk_level, @risk_reasons, @recommended_review_mode)
    ON CONFLICT(approval_id) DO NOTHING
  `).run({ ...r, risk_reasons: JSON.stringify(r.risk_reasons) });
}

export function insertComment(c: Comment): void {
  const db = getDb();
  db.prepare(`
    INSERT INTO comments (id, approval_id, author_id, author_name, body, created_at)
    VALUES (@id, @approval_id, @author_id, @author_name, @body, @created_at)
    ON CONFLICT(id) DO NOTHING
  `).run(c);
}

export function insertAuditEvent(e: AuditEvent): void {
  const db = getDb();
  db.prepare(`
    INSERT INTO audit_events (id, approval_id, actor_id, actor_name, action, reason, comment, timestamp)
    VALUES (@id, @approval_id, @actor_id, @actor_name, @action, @reason, @comment, @timestamp)
    ON CONFLICT(id) DO NOTHING
  `).run(e);
}

export function clearAll(): void {
  const db = getDb();
  db.exec(`
    DELETE FROM audit_events;
    DELETE FROM comments;
    DELETE FROM attachments;
    DELETE FROM risk_assessments;
    DELETE FROM approvals;
    DELETE FROM users;
  `);
}
