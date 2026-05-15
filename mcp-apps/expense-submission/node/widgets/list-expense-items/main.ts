import {
  App,
  applyDocumentTheme,
  applyHostStyleVariables,
  applyHostFonts,
} from "@modelcontextprotocol/ext-apps";

// ── Types ──────────────────────────────────────────────────────────────────
interface ReceiptAttachment {
  file_name: string;
  file_url: string;
  mime_type: string;
}
interface Expense {
  expense_id: string;
  description: string;
  category: string;
  merchant: string;
  date: string;
  amount: number;
  currency: string;
  card_last_four: string;
  status: string;
  receipt_attachment?: ReceiptAttachment;
  receipt_match?: 'matched';
  notes?: string;
}
interface ExpenseResult {
  success?: boolean;
  total_count?: number;
  expenses?: Expense[];
  start_date?: string;
  end_date?: string;
}

// ── State ──────────────────────────────────────────────────────────────────
const selectedIds = new Set<string>();
let allData: ExpenseResult = {};   // full unfiltered data from server
let lastData: ExpenseResult = {};  // filtered view passed to render
let filterStart = '';
let filterEnd = '';

// ── Helpers ────────────────────────────────────────────────────────────────
function parseData(result: unknown): ExpenseResult {
  if (!result || typeof result !== "object") return {};
  const r = result as Record<string, unknown>;
  const sc = r.structuredContent as Record<string, unknown> | undefined;
  if (sc && Array.isArray(sc.expenses)) return sc as ExpenseResult;
  if (Array.isArray(r.expenses)) return r as ExpenseResult;
  for (const item of (r.content as Array<{ text?: string }> | undefined) ?? []) {
    try {
      const p = JSON.parse(item.text ?? "") as Record<string, unknown>;
      if (Array.isArray(p.expenses)) return p as ExpenseResult;
    } catch { /**/ }
  }
  return {};
}

function applyStyles(body: HTMLElement, insets?: { top?: number; right?: number; bottom?: number; left?: number }) {
  if (insets) {
    const { top = 0, right = 0, bottom = 0, left = 0 } = insets;
    body.style.padding = `${top}px ${right}px ${bottom}px ${left}px`;
  }
}

function fmtAmt(n: number, currency = "USD") {
  return new Intl.NumberFormat("en-US", { style: "currency", currency, minimumFractionDigits: 2 }).format(n);
}

function fmtDate(iso: string) {
  return new Date(iso).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

function esc(s: string) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ── Render ──────────────────────────────────────────────────────────────────
function applyClientFilter(data: ExpenseResult): ExpenseResult {
  let expenses = data.expenses ?? [];
  if (filterStart) {
    const sd = new Date(filterStart).getTime();
    expenses = expenses.filter(e => new Date(e.date).getTime() >= sd);
  }
  if (filterEnd) {
    const ed = new Date(filterEnd).getTime();
    expenses = expenses.filter(e => new Date(e.date).getTime() <= ed);
  }
  return { ...data, expenses, total_count: expenses.length };
}

function buildListHtml(expenses: Expense[]): string {
  const selectedTotal = expenses.filter(e => selectedIds.has(e.expense_id)).reduce((s, e) => s + (e.amount || 0), 0);
  const allSelected = expenses.length > 0 && selectedIds.size === expenses.length;

  if (expenses.length === 0) {
    return `<p style="padding:40px 24px;text-align:center;font-size:14px;color:var(--color-text-secondary,#aaa)">No expenses found for the selected period.</p>`;
  }

  const rows = expenses.map(e => {
    const checked = selectedIds.has(e.expense_id);
    const ra = e.receipt_attachment;
    const receiptEl = ra
      ? `<span class="chip chip-attached">
           <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48"/></svg>
           ${esc(ra.file_name)}
         </span>`
      : `<span class="chip chip-missing">No receipt</span>`;

    const matchBadge = e.receipt_match === 'matched'
      ? `<span class="match-badge match-matched" title="Receipt matched to expense">
           <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
           Matched
         </span>`
      : '';

    return `
      <div class="row" data-id="${esc(e.expense_id)}">
        <div class="row-main" onclick="window.__toggleSelect('${esc(e.expense_id)}')">
          <input type="checkbox" class="cb" ${checked ? "checked" : ""} onclick="event.stopPropagation(); window.__toggleSelect('${esc(e.expense_id)}')" />
          <div class="row-body">
            <div class="exp-id">${esc(e.expense_id)}</div>
            <div class="merchant">${esc(e.merchant)}</div>
            <div class="meta">${esc(e.category)}<span class="bull">&middot;</span>${fmtDate(e.date)}</div>
          </div>
          <div class="row-right">
            <div class="amount">${fmtAmt(e.amount, e.currency)}</div>
            <div class="receipt-row">${receiptEl}${matchBadge}</div>
          </div>
        </div>
      </div>`;
  }).join(`<div class="divider"></div>`);

  const createBtn = selectedIds.size > 0
    ? `<button class="btn-create" onclick="window.__createReport()">Create Report (${selectedIds.size})</button>`
    : `<button class="btn-create btn-disabled" disabled>Select expenses to create report</button>`;

  return `
    <div class="select-all-bar">
      <input type="checkbox" class="cb" id="selectAll" ${allSelected ? "checked" : ""} onclick="window.__toggleAll()" />
      <label for="selectAll">Select all</label>
    </div>
    <div class="list">${rows}</div>
    <div class="footer">
      <div class="selected-info">${selectedIds.size > 0 ? `${selectedIds.size} selected &middot; ${fmtAmt(selectedTotal)}` : "None selected"}</div>
      ${createBtn}
    </div>`;
}

function updateListInline() {
  const filtered = applyClientFilter(allData);
  lastData = filtered;
  const expenses = filtered.expenses ?? [];
  const total = expenses.reduce((s, e) => s + (e.amount || 0), 0);

  const listContainer = document.getElementById('listContainer');
  const headerSub = document.getElementById('headerSub');
  if (listContainer) listContainer.innerHTML = buildListHtml(expenses);
  if (headerSub) headerSub.textContent = `${expenses.length} items${expenses.length > 0 ? ` · ${fmtAmt(total)}` : ''}`;
}

function render(root: HTMLElement, data: ExpenseResult) {
  allData = data;

  // Sync filter inputs from tool response if not yet set by user
  if (!filterStart && data.start_date) filterStart = data.start_date.slice(0, 10);
  if (!filterEnd && data.end_date) filterEnd = data.end_date.slice(0, 10);

  const filtered = applyClientFilter(data);
  lastData = filtered;
  const expenses = filtered.expenses ?? [];
  const total = expenses.reduce((s, e) => s + (e.amount || 0), 0);

  root.innerHTML = `
    <style>
      *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
      body {
        font-family: var(--font-sans, -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif);
        background: var(--color-background-primary, #fff);
        color: var(--color-text-primary, #0f0f0f);
        font-size: 16px; line-height: 1.5;
      }
      .header { padding: 16px 16px 12px; border-bottom: 1px solid var(--color-border-default, #ebebeb); }
      .header-top { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; }
      .header-title { font-size: 20px; font-weight: 700; line-height: 1.2; }
      .header-sub { font-size: 12px; color: var(--color-text-secondary, #999); margin-top: 2px; }

      .filter-bar { display: flex; align-items: center; gap: 8px; padding: 10px 16px; border-bottom: 1px solid var(--color-border-default, #f0f0f0); flex-wrap: wrap; }
      .filter-bar label { font-size: 12px; font-weight: 600; color: var(--color-text-secondary, #999); }
      .filter-bar input[type="date"] {
        font-family: inherit; font-size: 13px; padding: 4px 8px; border: 1px solid var(--color-border-default, #ddd);
        border-radius: 6px; background: var(--color-background-secondary, #fafafa); color: var(--color-text-primary, #111);
        outline: none;
      }
      .filter-bar input[type="date"]:focus { border-color: var(--color-accent, #0078d4); }
      .btn-filter {
        padding: 5px 14px; border: none; border-radius: 6px; cursor: pointer;
        font-family: inherit; font-size: 12px; font-weight: 700;
        background: var(--color-accent, #0078d4); color: #fff;
      }
      .btn-filter:hover { opacity: .9; }

      .select-all-bar { display: flex; align-items: center; gap: 8px; padding: 6px 16px; border-bottom: 1px solid var(--color-border-default, #f0f0f0); }
      .select-all-bar label { font-size: 12px; font-weight: 600; cursor: pointer; color: var(--color-text-secondary, #999); }

      .list { padding: 2px 0 0; }
      .divider { height: 1px; background: var(--color-border-default, #f0f0f0); margin: 0 16px; }
      .row { position: relative; }
      .row-main {
        display: flex; align-items: center; gap: 8px;
        padding: 10px 14px 10px 16px; cursor: pointer; user-select: none;
        transition: background .12s;
      }
      .row-main:hover { background: var(--color-background-secondary, #fafafa); }
      .cb { width: 16px; height: 16px; accent-color: var(--color-accent, #0078d4); cursor: pointer; flex-shrink: 0; }
      .row-body { flex: 1; min-width: 0; }
      .exp-id { font-size: 10px; font-weight: 600; letter-spacing: .4px; text-transform: uppercase; color: var(--color-text-secondary, #bbb); margin-bottom: 1px; }
      .merchant { font-size: 14px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .meta { font-size: 12px; color: var(--color-text-secondary, #999); margin-top: 1px; display: flex; align-items: center; flex-wrap: wrap; gap: 4px; }
      .bull { color: var(--color-border-default, #d4d4d4); }
      .row-right { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; flex-shrink: 0; padding-left: 8px; }
      .amount { font-size: 15px; font-weight: 700; white-space: nowrap; }
      .chip { display: inline-flex; align-items: center; gap: 4px; padding: 2px 7px; border-radius: 99px; font-size: 11px; font-weight: 500; white-space: nowrap; }
      .chip-attached { background: #f0fdf4; border: 1px solid #d1fae5; color: #15803d; }
      .chip-missing { color: var(--color-text-secondary, #c0c0c0); font-size: 11px; }

      .receipt-row { display: flex; align-items: center; gap: 4px; flex-wrap: wrap; justify-content: flex-end; }
      .match-badge {
        display: inline-flex; align-items: center; gap: 3px;
        padding: 1px 6px; border-radius: 99px; font-size: 10px; font-weight: 600;
        white-space: nowrap; opacity: .85;
      }
      .match-matched { background: #ecfdf5; border: 1px solid #d1fae5; color: #059669; }

      .footer { padding: 10px 16px; border-top: 1px solid var(--color-border-default, #ebebeb); display: flex; align-items: center; justify-content: space-between; }
      .selected-info { font-size: 13px; font-weight: 600; color: var(--color-text-secondary, #999); }
      .btn-create {
        padding: 7px 16px; border: none; border-radius: 8px; cursor: pointer;
        font-family: inherit; font-size: 13px; font-weight: 700; white-space: nowrap;
        background: var(--color-accent, #0078d4); color: #fff;
        transition: opacity .15s, transform .1s;
      }
      .btn-create:hover:not(.btn-disabled) { opacity: .9; }
      .btn-disabled { background: var(--color-background-secondary, #e8e8e8) !important; color: var(--color-text-secondary, #bbb) !important; cursor: not-allowed; }
    </style>

    <div class="header">
      <div class="header-top">
        <div>
          <div class="header-title">Expense Items</div>
          <div class="header-sub" id="headerSub">${expenses.length} items${expenses.length > 0 ? ` &middot; ${fmtAmt(total)}` : ''}</div>
        </div>
      </div>
    </div>

    <div class="filter-bar">
      <label>From</label>
      <input type="date" id="filterStart" value="${esc(filterStart)}" />
      <label>To</label>
      <input type="date" id="filterEnd" value="${esc(filterEnd)}" />
      <button class="btn-filter" onclick="window.__applyFilter()">Apply</button>
    </div>

    <div id="listContainer">${buildListHtml(expenses)}</div>
  `;
}

// ── App ──────────────────────────────────────────────────────────────────────
const root = document.getElementById("root") as HTMLElement;
root.innerHTML = `<p style="padding:40px 24px;text-align:center;font-size:14px;color:var(--color-text-secondary,#aaa)">Loading expense items…</p>`;

const app = new App({ name: "Expense Items", version: "1.0.0" });

function getContextText(): string {
  const selected = (lastData.expenses ?? []).filter(e => selectedIds.has(e.expense_id));
  if (selected.length === 0) return 'No expense items are currently selected.';
  return `User has selected ${selected.length} expense items:\n\n` +
    selected.map(e => `- ${e.expense_id}: ${e.merchant} | ${e.category} | ${fmtDate(e.date)} | ${fmtAmt(e.amount, e.currency)}`).join('\n');
}

(window as unknown as Record<string, unknown>).__toggleSelect = async (id: string) => {
  if (selectedIds.has(id)) selectedIds.delete(id);
  else selectedIds.add(id);
  updateListInline();
  await app.updateModelContext({ content: [{ type: 'text', text: getContextText() }] });
};

(window as unknown as Record<string, unknown>).__toggleAll = async () => {
  const expenses = lastData.expenses ?? [];
  if (selectedIds.size === expenses.length) selectedIds.clear();
  else expenses.forEach(e => selectedIds.add(e.expense_id));
  updateListInline();
  await app.updateModelContext({ content: [{ type: 'text', text: getContextText() }] });
};

(window as unknown as Record<string, unknown>).__applyFilter = async () => {
  const startEl = document.getElementById('filterStart') as HTMLInputElement | null;
  const endEl = document.getElementById('filterEnd') as HTMLInputElement | null;
  filterStart = startEl?.value ?? '';
  filterEnd = endEl?.value ?? '';
  selectedIds.clear();
  updateListInline();
  await app.updateModelContext({ content: [{ type: 'text', text: getContextText() }] });
};

(window as unknown as Record<string, unknown>).__createReport = async () => {
  if (selectedIds.size === 0) return;
  await app.updateModelContext({ content: [{ type: 'text', text: getContextText() }] });
  try {
    await app.sendMessage({
      role: 'user',
      content: [{ type: 'text', text: 'Think hard and create a draft expense report from the selected expense items.' }],
    });
  } catch (err) {
    console.error('Create report request failed:', err instanceof Error ? err.message : String(err));
  }
};

app.ontoolinput = () => {
  root.innerHTML = `<p style="padding:40px 24px;text-align:center;font-size:14px;color:var(--color-text-secondary,#aaa)">Fetching expense items…</p>`;
};

app.ontoolresult = (result: unknown) => {
  const data = parseData(result);
  render(root, data);
};

app.onhostcontextchanged = (ctx: {
  theme?: unknown;
  styles?: { variables?: unknown; css?: { fonts?: unknown } };
  safeAreaInsets?: { top?: number; right?: number; bottom?: number; left?: number };
}) => {
  if (ctx.theme) applyDocumentTheme(ctx.theme as Parameters<typeof applyDocumentTheme>[0]);
  if (ctx.styles?.variables) applyHostStyleVariables(ctx.styles.variables as Parameters<typeof applyHostStyleVariables>[0]);
  if (ctx.styles?.css?.fonts) applyHostFonts(ctx.styles.css.fonts as Parameters<typeof applyHostFonts>[0]);
  applyStyles(document.body, ctx.safeAreaInsets);
};
app.onteardown = async () => ({});
await app.connect();
console.log('[list-expense-items] Host capabilities:', JSON.stringify(app.getHostCapabilities()));
