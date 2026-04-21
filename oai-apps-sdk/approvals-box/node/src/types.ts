// ─────────────────────────────────────────────────────────────────────────────
// Core domain types — shared by db, risk, mcp-server
// ─────────────────────────────────────────────────────────────────────────────

export type RiskLevel = "low" | "medium" | "high";
export type ApprovalStatus = "pending" | "approved" | "rejected" | "cancelled";
export type ApprovalType =
  | "promo_pricing"
  | "purchase_order"
  | "capex"
  | "contract_sow"
  | "travel_exception"
  | "marketing_budget"
  | "vendor_onboarding";

export type AuditAction =
  | "submitted"
  | "approved"
  | "rejected"
  | "bulk_approved"
  | "bulk_rejected"
  | "commented"
  | "reopened";

export type ReviewMode = "inline" | "fullscreen";

// ─────────────────────────────────────────────────────────────────────────────
// Approval
// ─────────────────────────────────────────────────────────────────────────────

export interface ApprovalField {
  label: string;
  value: string;
}

export interface Approval {
  id: string;
  title: string;
  type: ApprovalType;
  display_type: string;
  status: ApprovalStatus;
  risk_level: RiskLevel;
  requester_id: string;
  requester_name: string;
  /** User IDs who must act on this approval */
  assignee_ids: string[];
  submitted_at: string; // ISO 8601
  due_at: string;       // ISO 8601
  amount: number;
  currency: string;
  country: string;
  summary: string;
  fields: ApprovalField[];
  attachment_ids: string[];
  comment_ids: string[];
  history_ids: string[];
  /** Backend-computed — never set by model or widget */
  allowed_actions: string[];
}

// ─────────────────────────────────────────────────────────────────────────────
// Risk
// ─────────────────────────────────────────────────────────────────────────────

export interface RiskAssessment {
  approval_id: string;
  risk_level: RiskLevel;
  risk_reasons: string[];
  recommended_review_mode: ReviewMode;
}

// ─────────────────────────────────────────────────────────────────────────────
// Attachment
// ─────────────────────────────────────────────────────────────────────────────

export interface Attachment {
  id: string;
  approval_id: string;
  name: string;
  mime_type: string;
  size_bytes: number;
  preview_supported: boolean;
  storage_key: string; // local relative path or Azure Blob key
}

// ─────────────────────────────────────────────────────────────────────────────
// Comment
// ─────────────────────────────────────────────────────────────────────────────

export interface Comment {
  id: string;
  approval_id: string;
  author_id: string;
  author_name: string;
  body: string;
  created_at: string; // ISO 8601
}

// ─────────────────────────────────────────────────────────────────────────────
// Audit Event
// ─────────────────────────────────────────────────────────────────────────────

export interface AuditEvent {
  id: string;
  approval_id: string;
  actor_id: string;
  actor_name: string;
  action: AuditAction;
  reason?: string;
  comment?: string;
  timestamp: string; // ISO 8601
}

// ─────────────────────────────────────────────────────────────────────────────
// User (lightweight — expanded with Azure AD in Phase 6b)
// ─────────────────────────────────────────────────────────────────────────────

export interface User {
  id: string;
  name: string;
  email: string;
  role: "approver" | "requester" | "admin";
  /** openai/subject value from ChatGPT request — maps to this user */
  subject_id?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Tool I/O shapes (used by mcp-server and widgets via toolOutput)
// ─────────────────────────────────────────────────────────────────────────────

export interface ApprovalSummary {
  total_pending: number;
  urgent: number;        // due within 24h
  overdue: number;       // past due_at and still pending
  high_risk: number;
  total_amount_pending: number; // sum of amounts for all pending items (USD equivalent)
}

export interface ApprovalListItem {
  id: string;
  title: string;
  type: ApprovalType;
  display_type: string;
  requester_name: string;
  submitted_at: string;
  due_at: string;
  amount: number;
  currency: string;
  risk_level: RiskLevel;
  has_attachments: boolean;
  status: ApprovalStatus;
  allowed_actions: string[];
}

export interface ListApprovalsOutput {
  summary: ApprovalSummary;
  items: ApprovalListItem[];
  next_cursor: string | null;
}

export interface ApprovalDetail extends Approval {
  risk_assessment: RiskAssessment;
  approver_chain: { user_id: string; name: string; status: string }[];
  comments: Comment[];
  history: AuditEvent[];
  attachments: Attachment[];
}

export interface AttachmentPreviewOutput {
  id: string;
  approval_id: string;
  name: string;
  mime_type: string;
  size_bytes: number;
  preview_supported: boolean;
  preview_text?: string;     // populated for .txt files
  download_url: string;
}

export interface CreateApprovalInput {
  type: ApprovalType;
  title: string;
  summary: string;
  requester_id: string;
  fields: ApprovalField[];
  attachment_ids?: string[];
}

export interface CreateApprovalOutput {
  id: string;
  status: ApprovalStatus;
  risk_level: RiskLevel;
  risk_reasons: string[];
  assignee_ids: string[];
  submitted_at: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Filters
// ─────────────────────────────────────────────────────────────────────────────

export type SortField = "due_date" | "amount" | "submitted_date" | "risk";

export interface ApprovalFilters {
  status?: ApprovalStatus[];
  risk_level?: RiskLevel[];
  type?: ApprovalType[];
  requester?: string;
  assignee_id?: string;
  has_attachments?: boolean;
  due_before?: string; // ISO date
  query?: string;      // text search
  sort?: SortField;
  limit?: number;
  cursor?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Dynamic field schemas per approval type (used by create form)
// ─────────────────────────────────────────────────────────────────────────────

export interface FieldSchema {
  label: string;
  key: string;
  type: "text" | "number" | "date" | "select";
  required: boolean;
  options?: string[]; // for select type
  placeholder?: string;
}

export const APPROVAL_TYPE_DISPLAY: Record<ApprovalType, string> = {
  promo_pricing: "C2C Promo Pricing",
  purchase_order: "Purchase Order",
  capex: "Capital Expenditure",
  contract_sow: "Contract / SOW",
  travel_exception: "Travel Exception",
  marketing_budget: "Marketing Budget",
  vendor_onboarding: "Vendor Onboarding",
};

export const APPROVAL_FIELD_SCHEMAS: Record<ApprovalType, FieldSchema[]> = {
  promo_pricing: [
    { label: "Publisher Name", key: "publisher_name", type: "text", required: true },
    { label: "Retailer Name", key: "retailer_name", type: "text", required: true },
    { label: "Country", key: "country", type: "text", required: true, placeholder: "e.g. USA, MEX" },
    { label: "Campaign Start Date", key: "start_date", type: "date", required: true },
    { label: "Campaign End Date", key: "end_date", type: "date", required: true },
    { label: "Discount %", key: "discount_pct", type: "number", required: false },
  ],
  purchase_order: [
    { label: "Vendor Name", key: "vendor_name", type: "text", required: true },
    { label: "Amount", key: "amount", type: "number", required: true },
    { label: "Currency", key: "currency", type: "select", required: true, options: ["USD", "EUR", "GBP", "CAD", "MXN"] },
    { label: "Business Justification", key: "justification", type: "text", required: true },
    { label: "Delivery Date", key: "delivery_date", type: "date", required: false },
  ],
  capex: [
    { label: "Project Name", key: "project_name", type: "text", required: true },
    { label: "Amount", key: "amount", type: "number", required: true },
    { label: "Currency", key: "currency", type: "select", required: true, options: ["USD", "EUR", "GBP"] },
    { label: "Department", key: "department", type: "text", required: true },
    { label: "Expected ROI", key: "expected_roi", type: "text", required: false },
    { label: "Business Justification", key: "justification", type: "text", required: true },
  ],
  contract_sow: [
    { label: "Vendor / Agency Name", key: "vendor_name", type: "text", required: true },
    { label: "Contract Value", key: "amount", type: "number", required: true },
    { label: "Currency", key: "currency", type: "select", required: true, options: ["USD", "EUR", "GBP"] },
    { label: "Start Date", key: "start_date", type: "date", required: true },
    { label: "End Date", key: "end_date", type: "date", required: true },
    { label: "Scope Summary", key: "scope_summary", type: "text", required: true },
  ],
  travel_exception: [
    { label: "Traveler Name", key: "traveler_name", type: "text", required: true },
    { label: "Destination", key: "destination", type: "text", required: true },
    { label: "Travel Dates", key: "travel_dates", type: "text", required: true },
    { label: "Estimated Cost", key: "amount", type: "number", required: true },
    { label: "Currency", key: "currency", type: "select", required: true, options: ["USD", "EUR", "GBP"] },
    { label: "Business Justification", key: "justification", type: "text", required: true },
    { label: "Policy Exception Reason", key: "exception_reason", type: "text", required: true },
  ],
  marketing_budget: [
    { label: "Campaign Name", key: "campaign_name", type: "text", required: true },
    { label: "Budget Amount", key: "amount", type: "number", required: true },
    { label: "Currency", key: "currency", type: "select", required: true, options: ["USD", "EUR", "GBP"] },
    { label: "Channel", key: "channel", type: "text", required: true, placeholder: "e.g. Paid Social, Events" },
    { label: "Quarter", key: "quarter", type: "select", required: true, options: ["Q1", "Q2", "Q3", "Q4"] },
    { label: "Business Justification", key: "justification", type: "text", required: true },
  ],
  vendor_onboarding: [
    { label: "Vendor Name", key: "vendor_name", type: "text", required: true },
    { label: "Vendor Type", key: "vendor_type", type: "select", required: true, options: ["Supplier", "Agency", "Contractor", "Consultant", "Technology"] },
    { label: "Annual Spend Estimate", key: "amount", type: "number", required: false },
    { label: "Currency", key: "currency", type: "select", required: false, options: ["USD", "EUR", "GBP"] },
    { label: "Business Contact", key: "business_contact", type: "text", required: true },
    { label: "Business Justification", key: "justification", type: "text", required: true },
  ],
};
