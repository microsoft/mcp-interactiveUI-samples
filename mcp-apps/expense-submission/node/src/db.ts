import { TableClient } from "@azure/data-tables";

const AZURITE_FALLBACK =
  "DefaultEndpointsProtocol=http;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;TableEndpoint=http://127.0.0.1:10002/devstoreaccount1;";

function getConnectionString() {
  return process.env.AZURE_STORAGE_CONNECTION_STRING ?? AZURITE_FALLBACK;
}

function getOpts() {
  const cs = getConnectionString();
  return cs.includes("127.0.0.1") || cs.includes("devstoreaccount1")
    ? { allowInsecureConnection: true }
    : {};
}

// ── Table Clients (lazy – evaluated after dotenv.config()) ─────────────
let _expensesTable: TableClient | null = null;
let _draftsTable: TableClient | null = null;

export function getExpensesTable(): TableClient {
  return _expensesTable ??= TableClient.fromConnectionString(getConnectionString(), "Expenses", getOpts());
}

export function getDraftsTable(): TableClient {
  return _draftsTable ??= TableClient.fromConnectionString(getConnectionString(), "Drafts", getOpts());
}

export async function ensureTables() {
  for (const table of [getExpensesTable(), getDraftsTable()]) {
    try { await table.createTable(); } catch { /* already exists */ }
  }
}

// ── Entity Interfaces ──────────────────────────────────────────────────
export interface ExpenseEntity {
  partitionKey: string;
  rowKey: string;
  description: string;
  category: string;
  merchant: string;
  date: string;
  amount: number;
  currency: string;
  cardLastFour: string;
  status: string;
  businessPurpose: string;
  projectCode: string;
  attendees: string;
  notes: string;
}

export interface DraftEntity {
  partitionKey: string;
  rowKey: string;
  draftId: string;
  expenses: string;      // JSON-serialised Expense[]
  status: string;        // "draft" | "submitted"
  createdAt: string;
  updatedAt: string;
}

// ── Expense (application-level shape returned to tools / widgets) ──────
export interface Expense {
  expense_id: string;
  description: string;
  category: string;
  merchant: string;
  date: string;
  amount: number;
  currency: string;
  card_last_four: string;
  status: string;
  receipt_attachment?: { file_name: string; file_url: string; mime_type: string };
  receipt_source?: 'email' | 'odsp';
  receipt_match?: 'matched';
  business_purpose?: string;
  project_code?: string;
  attendees?: string;
  notes?: string;
}

// ── Mapping helpers ────────────────────────────────────────────────────
function entityToExpense(e: ExpenseEntity): Expense {
  const expense: Expense = {
    expense_id: e.rowKey,
    description: e.description,
    category: e.category,
    merchant: e.merchant,
    date: e.date,
    amount: e.amount,
    currency: e.currency,
    card_last_four: e.cardLastFour,
    status: e.status,
  };
  if (e.businessPurpose) expense.business_purpose = e.businessPurpose;
  if (e.projectCode) expense.project_code = e.projectCode;
  if (e.attendees) expense.attendees = e.attendees;
  if (e.notes) expense.notes = e.notes;
  return expense;
}

// ── Expense CRUD (read-only) ───────────────────────────────────────────
export async function getAllExpenses(): Promise<Expense[]> {
  const results: Expense[] = [];
  for await (const entity of getExpensesTable().listEntities<ExpenseEntity>()) {
    results.push(entityToExpense(entity));
  }
  return results;
}

export async function getExpensesByIds(ids: string[]): Promise<Expense[]> {
  const results: Expense[] = [];
  for (const id of ids) {
    try {
      const entity = await getExpensesTable().getEntity<ExpenseEntity>("expenses", id);
      results.push(entityToExpense(entity));
    } catch { /* not found — skip */ }
  }
  return results;
}

// ── Draft CRUD ─────────────────────────────────────────────────────────
const DRAFT_PK = "drafts";
const DRAFT_RK = "current";

export interface Draft {
  draft_id: string;
  expenses: Expense[];
  status: string;
}

export async function getCurrentDraft(): Promise<Draft | null> {
  try {
    const entity = await getDraftsTable().getEntity<DraftEntity>(DRAFT_PK, DRAFT_RK);
    if (entity.status === "submitted") return null;
    return {
      draft_id: entity.draftId,
      expenses: JSON.parse(entity.expenses),
      status: entity.status,
    };
  } catch { return null; }
}

export async function upsertDraft(draftId: string, expenses: Expense[]): Promise<Draft> {
  const now = new Date().toISOString();
  const existing = await getCurrentDraft();
  await getDraftsTable().upsertEntity(
    {
      partitionKey: DRAFT_PK,
      rowKey: DRAFT_RK,
      draftId,
      expenses: JSON.stringify(expenses),
      status: "draft",
      createdAt: existing ? (existing as any).createdAt ?? now : now,
      updatedAt: now,
    },
    "Replace"
  );
  return { draft_id: draftId, expenses, status: "draft" };
}

export async function deleteDraft(): Promise<void> {
  try {
    await getDraftsTable().deleteEntity(DRAFT_PK, DRAFT_RK);
  } catch { /* nothing to delete */ }
}
