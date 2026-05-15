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
  receipt_source?: 'email' | 'odsp';
  receipt_match?: 'matched';
  notes?: string;
}
interface DraftResult {
  success?: boolean;
  draft_id?: string;
  total_count?: number;
  expenses?: Expense[];
}

// ── State ──────────────────────────────────────────────────────────────────
const expandedIds = new Set<string>();
let lastData: DraftResult = {};
let _submitting = false;
let _submitted = false;

// ── Helpers ────────────────────────────────────────────────────────────────
function parseData(result: unknown): DraftResult {
  if (!result || typeof result !== "object") return {};
  const r = result as Record<string, unknown>;
  const sc = r.structuredContent as Record<string, unknown> | undefined;
  if (sc && Array.isArray(sc.expenses)) return sc as DraftResult;
  if (Array.isArray(r.expenses)) return r as DraftResult;
  for (const item of (r.content as Array<{ text?: string }> | undefined) ?? []) {
    try {
      const p = JSON.parse(item.text ?? "") as Record<string, unknown>;
      if (Array.isArray(p.expenses)) return p as DraftResult;
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
function isRec(e: Expense) {
  return e.status === "reconciled" || e.status === "receipt_submitted";
}

// ── Render ──────────────────────────────────────────────────────────────────
function render(root: HTMLElement, data: DraftResult) {
  lastData = data;
  const expenses = data.expenses ?? [];
  const draftId = data.draft_id ?? "";

  if (expenses.length === 0) {
    root.innerHTML = `<p style="padding:40px 24px;text-align:center;font-size:14px;color:var(--color-text-secondary,#aaa)">No expenses in this draft.</p>`;
    return;
  }

  const reconciledCount = expenses.filter(isRec).length;
  const total = expenses.reduce((s, e) => s + (e.amount || 0), 0);
  const allDone = reconciledCount === expenses.length;
  const pct = Math.round((reconciledCount / expenses.length) * 100);

  const statusBadge = _submitted
    ? `<span class="status-badge status-pending">Pending Approval</span>`
    : `<span class="status-badge status-draft">Draft</span>`;

  const reportId = draftId ? (_submitted ? draftId.replace(/^DRAFT-/i, '') : draftId) : '';

  const submitBtn = _submitted
    ? `<button class="btn-submit btn-submitted" disabled>Submitted</button>`
    : allDone
      ? `<button class="btn-submit" onclick="window.__submitForApproval()">Submit for Approval</button>`
      : `<button class="btn-submit btn-disabled" disabled title="Attach receipts to all ${expenses.length} expenses to enable submission">Submit for Approval</button>`;

  const receiptHint = !_submitted && !allDone
    ? `<div class="receipt-hint">Attach receipts to all expenses to submit for approval</div>`
    : '';

  const rows = expenses.map(e => {
    const rec = isRec(e);
    const open = expandedIds.has(e.expense_id);
    const ra = e.receipt_attachment;

    const sourceTag = ra
      ? (e.receipt_source === 'email'
        ? `<span class="source-tag source-email" title="Extracted from email">
             <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="2" y="4" width="20" height="16" rx="2"/><polyline points="22 4 12 13 2 4"/></svg>
             Email
           </span>`
        : `<span class="source-tag source-file" title="Uploaded file">
             <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
             Uploaded File
           </span>`)
      : '';

    const receiptEl = ra
      ? `<span class="chip chip-attached" onclick="event.stopPropagation(); window.__openLink('${esc(ra.file_url)}')" style="cursor:pointer">
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
      <div class="row ${rec ? "row-rec" : "row-nr"}" data-id="${esc(e.expense_id)}">
        <div class="row-main" onclick="window.__toggle('${esc(e.expense_id)}')">
          <div class="row-left">
            <div class="dot ${rec ? "dot-rec" : "dot-nr"}"></div>
            <div class="row-body">
              <div class="exp-id">${esc(e.expense_id)}</div>
              <div class="merchant">${esc(e.merchant)}</div>
              <div class="meta">${esc(e.category)}<span class="bull">&middot;</span>${fmtDate(e.date)}</div>
            </div>
          </div>
          <div class="row-right">
            <div class="amount">${fmtAmt(e.amount, e.currency)}</div>
            <div class="receipt-row">${receiptEl}${sourceTag}${matchBadge}</div>
          </div>
          <svg class="chev ${open ? "chev-open" : ""}" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
        </div>
        <div class="detail ${open ? "detail-open" : ""}">
          <label class="note-lbl">Notes</label>
          <textarea class="note-ta" rows="3" placeholder="Add any notes…">${esc(e.notes || "")}</textarea>
        </div>
      </div>`;
  }).join(`<div class="divider"></div>`);

  root.innerHTML = `
    <style>
      *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
      body {
        font-family: var(--font-sans, -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif);
        background: var(--color-background-primary, #fff);
        color: var(--color-text-primary, #0f0f0f);
        font-size: 16px; line-height: 1.5;
      }
      .header { padding: 16px 16px 14px; border-bottom: 1px solid var(--color-border-default, #ebebeb); }
      .header-top { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; margin-bottom: 14px; }
      .header-title { font-size: 20px; font-weight: 700; line-height: 1.2; display: flex; align-items: center; gap: 8px; }
      .header-sub { font-size: 12px; color: var(--color-text-secondary, #999); margin-top: 4px; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
      .status-badge { display: inline-flex; font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 99px; letter-spacing: .4px; text-transform: uppercase; white-space: nowrap; }
      .status-draft { background: #fef3c7; color: #92400e; }
      .status-pending { background: #dbeafe; color: #1e40af; }
      .receipt-hint { font-size: 12px; color: #b45309; margin-top: 6px; display: flex; align-items: center; gap: 5px; }
      .receipt-hint::before { content: '\u26A0'; font-size: 13px; }
      .footer-bar { padding: 12px 16px; border-top: 1px solid var(--color-border-default, #ebebeb); display: flex; align-items: center; justify-content: space-between; }
      .footer-total { font-size: 18px; font-weight: 800; white-space: nowrap; }
      .footer-total-label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .4px; color: var(--color-text-secondary, #999); }
      .progress-wrap { display: flex; align-items: center; gap: 10px; }
      .progress-track { flex: 1; height: 4px; border-radius: 99px; background: var(--color-border-default, #e8e8e8); overflow: hidden; }
      .progress-fill { height: 100%; border-radius: 99px; background: #16a34a; width: ${pct}%; transition: width .4s ease; }
      .progress-label { font-size: 12px; font-weight: 600; color: var(--color-text-secondary, #999); white-space: nowrap; }
      .list { padding: 2px 0 0; }
      .divider { height: 1px; background: var(--color-border-default, #f0f0f0); margin: 0 16px; }
      .row { position: relative; }
      .row-rec { border-left: 2px solid #16a34a; }
      .row-nr  { border-left: 2px solid #f59e0b; }
      .row-main { display: flex; align-items: center; gap: 8px; padding: 10px 14px 10px 14px; cursor: pointer; user-select: none; transition: background .12s; }
      .row-main:hover { background: var(--color-background-secondary, #fafafa); }
      .row-left { display: flex; align-items: flex-start; gap: 10px; flex: 1; min-width: 0; }
      .dot { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; margin-top: 6px; }
      .dot-rec { background: #16a34a; }
      .dot-nr  { background: #f59e0b; }
      .row-body { flex: 1; min-width: 0; }
      .exp-id { font-size: 10px; font-weight: 600; letter-spacing: .4px; text-transform: uppercase; color: var(--color-text-secondary, #bbb); margin-bottom: 1px; }
      .merchant { font-size: 14px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .meta { font-size: 12px; color: var(--color-text-secondary, #999); margin-top: 1px; display: flex; align-items: center; flex-wrap: wrap; gap: 4px; }
      .bull { color: var(--color-border-default, #d4d4d4); }
      .row-right { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; flex-shrink: 0; padding-left: 8px; }
      .amount { font-size: 15px; font-weight: 700; white-space: nowrap; }
      .receipt-row { display: flex; align-items: center; gap: 4px; flex-wrap: wrap; justify-content: flex-end; }
      .chip { display: inline-flex; align-items: center; gap: 4px; padding: 2px 7px; border-radius: 99px; font-size: 11px; font-weight: 500; white-space: nowrap; text-decoration: none; }
      .chip-attached { background: var(--color-background-secondary, #f5f5f5); border: 1px solid var(--color-border-default, #e0e0e0); color: var(--color-text-primary, #333); }
      .chip-attached:hover { background: var(--color-background-tertiary, #ebebeb); }
      .source-tag { display: inline-flex; align-items: center; gap: 3px; padding: 1px 6px; border-radius: 99px; font-size: 10px; font-weight: 600; white-space: nowrap; }
      .source-email { background: #eff6ff; border: 1px solid #bfdbfe; color: #2563eb; }
      .source-file { background: #f5f3ff; border: 1px solid #ddd6fe; color: #7c3aed; }
      .chip-missing { color: var(--color-text-secondary, #c0c0c0); font-size: 11px; }
      .match-badge { display: inline-flex; align-items: center; gap: 3px; padding: 1px 6px; border-radius: 99px; font-size: 10px; font-weight: 600; white-space: nowrap; opacity: .85; }
      .match-matched { background: #ecfdf5; border: 1px solid #d1fae5; color: #059669; }
      .chev { flex-shrink: 0; color: var(--color-text-secondary, #c0c0c0); transition: transform .2s ease; }
      .chev-open { transform: rotate(180deg); }
      .detail { max-height: 0; overflow: hidden; transition: max-height .22s ease; }
      .detail-open { max-height: 180px; }
      .note-lbl { display: block; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .5px; color: var(--color-text-secondary, #bbb); padding: 0 16px; margin-bottom: 4px; margin-top: 2px; }
      .note-ta { display: block; width: calc(100% - 32px); margin: 0 16px 12px; font-family: inherit; font-size: 14px; padding: 8px 10px; border: 1px solid var(--color-border-default, #e0e0e0); border-radius: 8px; background: var(--color-background-secondary, #fafafa); color: var(--color-text-primary, #111); resize: none; outline: none; line-height: 1.5; }
      .note-ta:focus { border-color: var(--color-accent, #0078d4); background: var(--color-background-primary, #fff); }
      .btn-submit { padding: 10px 20px; border: none; border-radius: 8px; cursor: pointer; font-family: inherit; font-size: 14px; font-weight: 700; white-space: nowrap; background: var(--color-accent, #0078d4); color: #fff; transition: opacity .15s; flex-shrink: 0; }
      .btn-submit:hover:not(.btn-disabled):not(.btn-submitted) { opacity: .9; }
      .btn-disabled { background: var(--color-background-secondary, #e8e8e8) !important; color: var(--color-text-secondary, #bbb) !important; cursor: not-allowed; }
      .btn-submitted { background: var(--color-background-secondary, #e8e8e8) !important; color: var(--color-text-secondary, #bbb) !important; cursor: default; }
      .btn-portal { padding: 10px 16px; border: 1px solid var(--color-border-default, #d4d4d4); border-radius: 8px; cursor: pointer; font-family: inherit; font-size: 13px; font-weight: 600; white-space: nowrap; background: var(--color-background-primary, #fff); color: var(--color-text-primary, #333); transition: background .15s; flex-shrink: 0; display: inline-flex; align-items: center; gap: 5px; }
      .btn-portal:hover { background: var(--color-background-secondary, #f5f5f5); }
    </style>

    <div class="header">
      <div class="header-top">
        <div>
          <div class="header-title">Expense Report ${statusBadge}</div>
          <div class="header-sub">${reportId ? esc(reportId) : ''} &middot; ${expenses.length} items</div>
        </div>
        <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">
          ${submitBtn}
          <button class="btn-portal" onclick="window.__openPortal()">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 01-2 2H5a2 2 0 01-2-2V8a2 2 0 012-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
            Open in Expense Portal
          </button>
        </div>
      </div>
      <div class="progress-wrap">
        <div class="progress-track"><div class="progress-fill"></div></div>
        <span class="progress-label">${reconciledCount} / ${expenses.length} receipts</span>
      </div>
      ${receiptHint}
    </div>

    <div class="list">${rows}</div>

    <div class="footer-bar">
      <div>
        <div class="footer-total-label">Total</div>
        <div class="footer-total">${fmtAmt(total)}</div>
      </div>
    </div>
  `;

  (window as unknown as Record<string, unknown>).__toggle = (id: string) => {
    if (expandedIds.has(id)) expandedIds.delete(id);
    else expandedIds.add(id);
    render(root, lastData);
  };
}

// ── App ──────────────────────────────────────────────────────────────────────
const root = document.getElementById("root") as HTMLElement;
root.innerHTML = `<p style="padding:40px 24px;text-align:center;font-size:14px;color:var(--color-text-secondary,#aaa)">Creating expense report draft…</p>`;

const app = new App({ name: "Expense Report Draft", version: "1.0.0" });

(window as unknown as Record<string, unknown>).__openLink = (url: string) => {
  app.openLink({ url }).catch(err => console.warn('openLink failed:', err));
};

(window as unknown as Record<string, unknown>).__openPortal = () => {
  app.openLink({ url: 'https://expenseassist.contoso.com' }).catch(err => console.warn('openPortal failed:', err));
};

(window as unknown as Record<string, unknown>).__submitForApproval = async () => {
  if (_submitting) return;
  const btn = document.querySelector<HTMLButtonElement>('.btn-submit');
  if (!btn || btn.disabled) return;

  if (!app.getHostCapabilities()?.serverTools) {
    btn.textContent = 'Not supported by host';
    setTimeout(() => render(root, lastData), 3000);
    return;
  }

  _submitting = true;
  btn.textContent = 'Submitting…';
  btn.disabled = true;
  try {
    const expenseIds = (lastData.expenses ?? []).map(e => e.expense_id);
    const result = await app.callServerTool(
      { name: 'submit_expense_report', arguments: { expense_ids: expenseIds, draft_id: lastData.draft_id } },
      { timeout: 15000 },
    );
    const res = result as { structuredContent?: { message?: string }; content?: Array<{ text?: string }> };
    const msg = res.structuredContent?.message ?? res.content?.[0]?.text ?? 'Report submitted!';
    _submitted = true;
    console.log('submit_expense_report result:', msg);
    render(root, lastData);
  } catch (err) {
    const errMsg = err instanceof Error ? err.message : String(err);
    console.error('submit_expense_report failed:', errMsg);
    btn.textContent = `Error: ${errMsg.slice(0, 40)}`;
    setTimeout(() => { _submitting = false; render(root, lastData); }, 4000);
  }
};

app.ontoolinput = () => {
  root.innerHTML = `<p style="padding:40px 24px;text-align:center;font-size:14px;color:var(--color-text-secondary,#aaa)">Loading report draft…</p>`;
};

app.ontoolresult = (result: unknown) => {
  const data = parseData(result);
  render(root, data);
  const expenses = data.expenses ?? [];
  if (expenses.length > 0) {
    const withReceipt = expenses.filter(e => e.receipt_attachment);
    const withoutReceipt = expenses.filter(e => !e.receipt_attachment);

    let text = `Draft report ${data.draft_id ?? ''} with ${expenses.length} expenses (${withReceipt.length} with receipts, ${withoutReceipt.length} without):\n\n` +
      expenses.map(e => `- ${e.expense_id}: ${e.merchant} | ${e.category} | ${fmtDate(e.date)} | ${fmtAmt(e.amount, e.currency)} | receipt: ${e.receipt_attachment?.file_name ?? 'none'}`).join('\n');

    if (withoutReceipt.length > 0) {
      text += `\n\nExpense IDs still missing receipts: ${withoutReceipt.map(e => e.expense_id).join(', ')}`;
    }

    app.updateModelContext({ content: [{ type: 'text', text }] });
  }
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
