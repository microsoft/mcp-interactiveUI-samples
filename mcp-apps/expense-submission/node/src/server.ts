import express, { Request, Response } from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { registerAppResource, registerAppTool, RESOURCE_MIME_TYPE } from '@modelcontextprotocol/ext-apps/server';
import { createRemoteJWKSet, jwtVerify } from 'jose';
import { z } from 'zod';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { getAllExpenses, getExpensesByIds, getCurrentDraft, upsertDraft, deleteDraft, type Expense } from './db.js';

dotenv.config();

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ASSETS_DIR = path.join(__dirname, '..', 'assets');
const RECEIPTS_DIR = path.join(__dirname, '..', 'sample-receipts');

async function readWidgetHtml(name: string): Promise<string> {
  const filePath = path.join(ASSETS_DIR, `${name}.html`);
  try {
    return await fs.promises.readFile(filePath, 'utf8');
  } catch {
    return `<html><body><p>Widget "${name}" not built yet. Run: npm run build:widgets</p></body></html>`;
  }
}

// =============================================================================
// Configuration
// =============================================================================
const OBO_CLIENT_ID = process.env.ENTRA_CLIENT_ID ?? '';
const OBO_CLIENT_SECRET = process.env.ENTRA_CLIENT_SECRET ?? '';
const ENTRA_TENANT_ID = process.env.ENTRA_TENANT_ID ?? 'common';
// TODO [PRODUCTION]: Remove DEV_API_KEY. API key auth is for local development only.
// In production all requests must be authenticated via Entra JWT tokens.
const DEV_API_KEY = 'mock_mcp_api_key';
const PORT = parseInt(process.env.PORT ?? '3001', 10);

// =============================================================================
// JWKS for Microsoft Entra (cached automatically by jose)
// =============================================================================
const microsoftJwks = createRemoteJWKSet(
  new URL('https://login.microsoftonline.com/common/discovery/keys')
);

// =============================================================================
// OBO Token Cache
// TODO [PRODUCTION]: Replace this in-memory Map with a distributed cache (e.g.
// Redis) so tokens are shared across multiple server instances and survive restarts.
// =============================================================================
interface CachedOboToken {
  accessToken: string;
  expiresAt: number;
}
const oboTokenCache = new Map<string, CachedOboToken>();

// Thrown when OBO token exchange fails due to missing user consent
class ConsentRequiredError extends Error {
  constructor(scope: string, detail?: string) {
    super(`User consent required for scope: ${scope}${detail ? ` (${detail})` : ''}`);
    this.name = 'ConsentRequiredError';
  }
}


/**
 * Save receipt content to the draft receipts directory and return the server-relative URL.
 * Files are stored under sample-receipts/drafts/ so they can be cleaned up
 * when a new draft is created or the current draft is submitted.
 */
async function saveReceiptLocally(
  opts: { fileName: string; contentBytes?: string; downloadUrl?: string; authToken?: string }
): Promise<string> {
  const draftsDir = path.join(RECEIPTS_DIR, 'drafts');
  await fs.promises.mkdir(draftsDir, { recursive: true });

  // Generate a unique filename to avoid collisions
  const ext = path.extname(opts.fileName) || '.bin';
  const localName = `${crypto.randomUUID()}${ext}`;
  const localPath = path.join(draftsDir, localName);

  if (opts.contentBytes) {
    // Email attachment: base64 content already available
    await fs.promises.writeFile(localPath, Buffer.from(opts.contentBytes, 'base64'));
    console.log(`  Saved email attachment to drafts/${localName}`);
  } else if (opts.downloadUrl) {
    // ODSP / SAS URL: download the file content
    const headers: Record<string, string> = {};
    if (opts.authToken) {
      headers['Authorization'] = `Bearer ${opts.authToken}`;
    }
    const resp = await fetch(opts.downloadUrl, { headers });
    if (!resp.ok) {
      throw new Error(`Failed to download receipt (${resp.status}): ${resp.statusText}`);
    }
    const buffer = Buffer.from(await resp.arrayBuffer());
    await fs.promises.writeFile(localPath, buffer);
    console.log(`  Downloaded and saved ODSP receipt to drafts/${localName}`);
  } else {
    throw new Error('saveReceiptLocally requires either contentBytes or downloadUrl');
  }

  return `/receipts/drafts/${localName}`;
}

/**
 * Remove all downloaded receipt files from the drafts folder.
 * Called when a new draft is created or after submission.
 */
async function cleanupDraftReceipts(): Promise<void> {
  const draftsDir = path.join(RECEIPTS_DIR, 'drafts');
  try {
    await fs.promises.rm(draftsDir, { recursive: true, force: true });
    console.log('Cleaned up draft receipt files.');
  } catch {
    // Directory may not exist — that's fine
  }
}

/**
 * Convert a server-relative receipt path to a full URL using the given base.
 */
function toFullReceiptUrl(relativePath: string, baseUrl: string): string {
  return `${baseUrl}${relativePath}`;
}

function getTidFromToken(token: string): string {
  try {
    const payload = JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString('utf8'));
    return (payload.tid as string) || 'common';
  } catch {
    return 'common';
  }
}

/**
 * Fetch a file from SharePoint/OneDrive using a Graph OBO token.
 * Supports sharing links (encoded via base64url) and direct drive item URLs.
 * Returns the driveItem metadata (name, size, mimeType, webUrl) or throws.
 */
/**
 * Fetch a file from SharePoint/OneDrive using a Graph OBO token.
 * Supports sharing links (encoded via base64url) and direct drive item URLs.
 * Returns the driveItem metadata including the pre-authenticated SAS download URL,
 * or throws on error.
 */
async function fetchReceiptViaGraph(
  fileUrl: string,
  graphToken: string
): Promise<{ name: string; size: number; mimeType: string; webUrl: string; downloadUrl: string }> {
  const headers = { Authorization: `Bearer ${graphToken}` };

  // Encode sharing URL as base64url for the /shares endpoint
  const encoded = Buffer.from(fileUrl).toString('base64')
    .replace(/=/g, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');
  // @microsoft.graph.downloadUrl is an OData annotation — it's returned by default but NOT via $select
  const sharesUrl = `https://graph.microsoft.com/v1.0/shares/u!${encoded}/driveItem`;

  const resp = await fetch(sharesUrl, { headers });
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({})) as Record<string, unknown>;
    const msg = (err?.error as Record<string, unknown>)?.message ?? resp.statusText;
    throw new Error(`Graph /shares fetch failed (${resp.status}): ${msg}`);
  }
  const item = await resp.json() as Record<string, unknown>;
  const file = item.file as Record<string, unknown> | undefined;
  return {
    name: (item.name as string) ?? '',
    size: (item.size as number) ?? 0,
    mimeType: (file?.mimeType as string) ?? 'application/octet-stream',
    webUrl: (item.webUrl as string) ?? fileUrl,
    downloadUrl: (item['@microsoft.graph.downloadUrl'] as string) ?? '',
  };
}

/**
 * Fetch file attachments from an Outlook email given its OWA URL.
 * 1. Extracts the ItemID query-parameter and URL-decodes it.
 * 2. Calls /me/translateExchangeIds to convert the EWS ID to a Graph REST ID.
 * 3. Calls /me/messages/{id}/attachments to list all attachments.
 * Returns metadata for every #fileAttachment found.
 */
async function fetchAttachmentsFromEmail(
  emailUrl: string,
  graphToken: string
): Promise<Array<{ name: string; size: number; mimeType: string; contentBytes: string; downloadUrl: string }>> {
  const headers = { Authorization: `Bearer ${graphToken}`, 'Content-Type': 'application/json' };

  // 1. Extract & decode ItemID from the OWA URL (case-insensitive lookup)
  const parsed = new URL(emailUrl);
  let rawItemId: string | null = null;
  for (const [key, value] of parsed.searchParams.entries()) {
    if (key.toLowerCase() === 'itemid') {
      rawItemId = value;
      break;
    }
  }
  if (!rawItemId) {
    throw new Error(`Could not extract ItemID query parameter from the email URL. Params: ${[...parsed.searchParams.keys()].join(', ')}`);
  }
  const itemId = decodeURIComponent(rawItemId);

  // 2. Translate EWS ID → Graph REST ID
  const translateResp = await fetch('https://graph.microsoft.com/v1.0/me/translateExchangeIds', {
    method: 'POST',
    headers,
    body: JSON.stringify({
      inputIds: [itemId],
      sourceIdType: 'ewsId',
      targetIdType: 'restId',
    }),
  });
  if (!translateResp.ok) {
    const err = await translateResp.json().catch(() => ({})) as Record<string, unknown>;
    const msg = (err?.error as Record<string, unknown>)?.message ?? translateResp.statusText;
    throw new Error(`Graph /translateExchangeIds failed (${translateResp.status}): ${msg}`);
  }
  const translateData = await translateResp.json() as Record<string, unknown>;
  const values = translateData.value as Array<Record<string, unknown>> | undefined;
  const messageId = values?.[0]?.targetId as string | undefined;
  if (!messageId) {
    throw new Error('Could not translate Exchange ID to a Graph-compatible message ID');
  }

  // 3. List attachments on the message
  const attachResp = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments`,
    { headers: { Authorization: `Bearer ${graphToken}` } }
  );
  if (!attachResp.ok) {
    const err = await attachResp.json().catch(() => ({})) as Record<string, unknown>;
    const msg = (err?.error as Record<string, unknown>)?.message ?? attachResp.statusText;
    throw new Error(`Graph /messages/.../attachments failed (${attachResp.status}): ${msg}`);
  }
  const attachData = await attachResp.json() as Record<string, unknown>;
  const attachments = (attachData.value ?? []) as Array<Record<string, unknown>>;

  // Return only file attachments (not item or reference attachments)
  return attachments
    .filter(a => a['@odata.type'] === '#microsoft.graph.fileAttachment')
    .map(a => ({
      name: (a.name as string) ?? 'attachment',
      size: (a.size as number) ?? 0,
      mimeType: (a.contentType as string) ?? 'application/octet-stream',
      contentBytes: (a.contentBytes as string) ?? '',
      downloadUrl: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(a.id as string)}/$value`,
    }));
}

async function exchangeTokenForGraph(
  userToken: string,
  // ── Graph API scope for OBO token exchange ──
  // Option 1: Files.Read.All — reads any file the user can access across
  //   OneDrive and SharePoint. Requires tenant admin consent in most orgs.
  // Option 2 (preferred): Files.SelectedOperations.Selected — reads only the
  //   files the user explicitly attached in the current interaction. No admin
  //   consent required. Requires Sydney/TuringBot to grant file-level access
  //   to the agent's Entra app before the tool call.
  // scope = 'https://graph.microsoft.com/Files.Read.All'
  scope = 'https://graph.microsoft.com/Files.SelectedOperations.Selected'
): Promise<string | null> {
  if (!OBO_CLIENT_SECRET) {
    console.error('ENTRA_CLIENT_SECRET not set. Cannot perform OBO token exchange.');
    console.error('Please set ENTRA_CLIENT_SECRET in your .env file.');
    return null;
  }

  const cacheKey = `${userToken.slice(0, 20)}:${scope}`;
  const cached = oboTokenCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now() / 1000) {
    console.log('Using cached Graph token');
    return cached.accessToken;
  }

  // For multi-tenant apps OBO requires a tenant-specific endpoint, not /common/
  const tid = getTidFromToken(userToken);
  // Log token claims to help diagnose OBO issues
  try {
    const claims = JSON.parse(Buffer.from(userToken.split('.')[1], 'base64').toString('utf8'));
    console.log(`OBO assertion token — aud: ${claims.aud}, scp: ${claims.scp}, appid: ${claims.appid}`);
  } catch { /* ignore */ }
  const tokenUrl = `https://login.microsoftonline.com/${tid}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    client_id: OBO_CLIENT_ID,
    client_secret: OBO_CLIENT_SECRET,
    assertion: userToken,
    scope,
    requested_token_use: 'on_behalf_of',
  });

  try {
    console.log(`Exchanging token for Graph API (scope: ${scope})...`);
    const resp = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
    const data = (await resp.json()) as Record<string, unknown>;

    if (typeof data.access_token === 'string') {
      console.log('OBO token exchange successful!');
      oboTokenCache.set(cacheKey, {
        accessToken: data.access_token,
        expiresAt:
          Date.now() / 1000 +
          (typeof data.expires_in === 'number' ? data.expires_in : 3600) -
          60,
      });
      return data.access_token;
    }

    console.error(
      'OBO token exchange failed:',
      data.error_description ?? JSON.stringify(data)
    );
    // Detect consent-required errors and throw so callers can return 401
    const errCode = (data.error as string) ?? '';
    const errDesc = (data.error_description as string) ?? '';
    if (
      errCode === 'interaction_required' ||
      errDesc.includes('AADSTS65001') ||
      errDesc.includes('AADSTS50013') ||
      errDesc.includes('consent')
    ) {
      throw new ConsentRequiredError(scope, errDesc);
    }
    return null;
  } catch (err) {
    console.error('OBO token exchange error:', err);
    return null;
  }
}

// =============================================================================
// Authentication — MS SSO JWT | API Key
// =============================================================================
async function authenticateRequest(req: Request): Promise<string | null> {
  const auth = req.headers.authorization;
  if (!auth?.startsWith('Bearer ')) {
    console.log('[AUTH] No Bearer token in request');
    return null;
  }
  const token = auth.slice(7);

  // Decode and log JWT claims (without verifying) for diagnostics
  try {
    const claims = JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString('utf8'));
    console.log('[AUTH] Incoming token claims:', JSON.stringify({
      aud: claims.aud,
      iss: claims.iss,
      appid: claims.appid,
      azp: claims.azp,
      scp: claims.scp,
      roles: claims.roles,
      tid: claims.tid,
      sub: claims.sub,
      exp: claims.exp,
    }, null, 2));
  } catch { console.log('[AUTH] Could not decode token claims'); }

  // Validate audience: only accept Copilot auth gateway format
  // api://auth-{uuid}/{CLIENT_ID}
  const COPILOT_AUTH_GATEWAY_RE = new RegExp(
    `^api://auth-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/${OBO_CLIENT_ID.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`
  );
  const isAudienceValid = (aud: string): boolean => COPILOT_AUTH_GATEWAY_RE.test(aud);

  // 1. Verify JWT signature against Microsoft JWKS, then check audience
  try {
    const { payload } = await jwtVerify(token, microsoftJwks);
    const aud = typeof payload.aud === 'string' ? payload.aud : '';
    if (!isAudienceValid(aud)) {
      console.log(`[AUTH] JWT signature valid but audience rejected: ${aud}`);
    } else {
      console.log('[AUTH] Token verified by MS SSO verifier');
      return token;
    }
  } catch (err) {
    console.log('[AUTH] MS SSO verification failed:', (err as Error).message);
  }

  // 2. Try API key (development/testing)
  if (token === DEV_API_KEY) {
    console.log('[AUTH] Token verified by API key verifier');
    return token;
  }

  console.log('[AUTH] All verification methods failed — returning null');
  return null;
}

// =============================================================================
// Resource URIs
// =============================================================================
const LIST_EXPENSE_ITEMS_RESOURCE_URI = 'ui://list-expense-items/app.html';
const EXPENSE_REPORT_DRAFT_RESOURCE_URI = 'ui://expense-report-draft/app.html';

// =============================================================================
// Zod Schemas
// =============================================================================
const ReceiptItemSchema = z.object({
  expense_id: z
    .string()
    .describe(
      "The selected expense ID to attach the receipt to (e.g. 'EXP-2001'). " +
        "Receipts that do not match any existing expense will be skipped."
    ),
  file_name: z
    .string()
    .optional()
    .describe(
      'The display name of the file or email subject. ' +
        'For odsp: the file name (e.g. "receipt.pdf"). ' +
        'For email: the email subject (e.g. "Dinner receipt from vendor").'
    ),
  file_url: z
    .string()
    .describe(
      'The URL of the file or email. ' +
        'For odsp: the SharePoint/OneDrive file URL (e.g. "https://contoso.sharepoint.com/sites/docs/receipt.pdf"). ' +
        'For email: the Outlook Web (OWA) URL (e.g. "https://outlook.office365.com/owa/?ItemID=AAMk...&exvsurl=1&viewmodel=ReadMessageItem").'
    ),
  type: z
    .enum(['odsp', 'email'])
    .default('odsp')
    .describe(
      'The type of receipt source. ' +
        '"odsp" (default) for OneDrive/SharePoint file links. ' +
        '"email" for Outlook Web email URLs whose file attachments should be extracted as receipts.'
    ),
});

// =============================================================================
// MCP Server Factory — creates a fresh server per request (stateless HTTP)
// =============================================================================
function createMcpServer(userToken: string | null, baseUrl: string): McpServer {
  const server = new McpServer({
    name: 'Expense Submission Service',
    version: '1.0.0',
  });

  // ---------------------------------------------------------------------------
  // Resources
  // ---------------------------------------------------------------------------
  registerAppResource(
    server,
    'List Expense Items UI',
    LIST_EXPENSE_ITEMS_RESOURCE_URI,
    {
      description: 'Interactive table view for expense line items with selection',
      mimeType: RESOURCE_MIME_TYPE,
    },
    async () => ({
      contents: [
        {
          uri: LIST_EXPENSE_ITEMS_RESOURCE_URI,
          mimeType: RESOURCE_MIME_TYPE,
          text: await readWidgetHtml('list-expense-items'),
        },
      ],
    })
  );

  registerAppResource(
    server,
    'Expense Report Draft UI',
    EXPENSE_REPORT_DRAFT_RESOURCE_URI,
    {
      description: 'Draft expense report view with submit for approval',
      mimeType: RESOURCE_MIME_TYPE,
    },
    async () => ({
      contents: [
        {
          uri: EXPENSE_REPORT_DRAFT_RESOURCE_URI,
          mimeType: RESOURCE_MIME_TYPE,
          text: await readWidgetHtml('expense-report-draft'),
        },
      ],
    })
  );

  // ---------------------------------------------------------------------------
  // Tool: list_expense_items
  // ---------------------------------------------------------------------------
  registerAppTool(
    server,
    'list_expense_items',
    {
      description:
        'Fetch expense line items from credit card transactions. ' +
        'Optionally filter by date range using start_date and end_date (ISO 8601). ' +
        'Returns expense details including expense id, category, merchant, date/time, and amount. ' +
        'Call this first so the user can select expenses before creating a draft report.',
      annotations: { readOnlyHint: true },
      _meta: { ui: { resourceUri: LIST_EXPENSE_ITEMS_RESOURCE_URI } },
      inputSchema: {
        start_date: z.string().optional().describe(
          'Optional start date (ISO 8601) to filter expenses. Only expenses on or after this date are returned.'
        ),
        end_date: z.string().optional().describe(
          'Optional end date (ISO 8601) to filter expenses. Only expenses on or before this date are returned.'
        ),
      },
    },
    async (args) => {
      const { start_date, end_date } = args as { start_date?: string; end_date?: string };
      console.log('Fetching expense line items from CRM...');

      if (!userToken) {
        const errorResult = {
          error: 'No user token found. Authentication required.',
        };
        return {
          content: [
            {
              type: 'text' as const,
              text: errorResult.error,
            },
          ],
          structuredContent: errorResult,
          isError: true,
        };
      }

      // Always return all expenses; pass date range to widget for client-side filtering
      const all = await getAllExpenses();

      console.log(`Found ${all.length} expense records`);
      const result = {
        success: true,
        total_count: all.length,
        expenses: all,
        ...(start_date ? { start_date } : {}),
        ...(end_date ? { end_date } : {}),
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: `Fetched ${result.total_count} expense record(s).`,
          },
        ],
        structuredContent: result,
        _meta: { ui: { resourceUri: LIST_EXPENSE_ITEMS_RESOURCE_URI } },
      };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: add_expense_receipts
  // ---------------------------------------------------------------------------
  registerAppTool(
    server,
    'add_expense_receipts',
    {
      description:
        'Attach receipts to an expense report draft. ' +
        'If the draft expense report details (draft_id and expense list) are already available in your conversation context, ' +
        'use them directly — do NOT call fetch_draft_expense_report again. ' +
        'Only call fetch_draft_expense_report first if the draft details are missing from context. ' +
        'Match each receipt file against the expense items in the draft. ' +
        'Include the expense_id, file_name, file_url, and type for each receipt. ' +
        'Receipts that do not match any draft expense are skipped.',
      annotations: { readOnlyHint: false },
      _meta: { ui: { resourceUri: EXPENSE_REPORT_DRAFT_RESOURCE_URI } },
      inputSchema: { receipts: z.array(ReceiptItemSchema) },
    },
    async (args) => {
      const { receipts } = args as { receipts: Array<{ expense_id: string; file_name?: string; file_url: string; type?: 'odsp' | 'email' }> };
      console.log(`Submitting ${receipts.length} receipt(s)...`);

      if (!userToken) {
        const errorResult = {
          error: 'No user token found. Authentication required.',
        };
        return {
          content: [
            {
              type: 'text' as const,
              text: errorResult.error,
            },
          ],
          structuredContent: errorResult,
          isError: true,
        };
      }

      // Determine which Graph scopes are needed based on receipt types
      const hasOdsp = receipts.some(r => (r.type ?? 'odsp') === 'odsp');
      const hasEmail = receipts.some(r => r.type === 'email');

      let graphTokenOdsp: string | null = null;
      let graphTokenMail: string | null = null;

      try {
        if (hasOdsp) {
          // Use Files.SelectedOperations.Selected (preferred, least-privilege) or
          // Files.Read.All (broader access, requires admin consent).
          // graphTokenOdsp = await exchangeTokenForGraph(userToken, 'https://graph.microsoft.com/Files.Read.All');
          graphTokenOdsp = await exchangeTokenForGraph(userToken, 'https://graph.microsoft.com/Files.SelectedOperations.Selected');
        }
        if (hasEmail) {
          graphTokenMail = await exchangeTokenForGraph(userToken, 'https://graph.microsoft.com/Mail.Read');
        }
      } catch (err) {
        if (err instanceof ConsentRequiredError) {
          console.error('Consent required:', err.message);
          return {
            content: [{ type: 'text' as const, text: 'Authorization required. Please sign in and grant the required permissions to continue.' }],
            structuredContent: { error: 'consent_required', message: err.message },
            isError: true,
          };
        }
        throw err;
      }

      // Load the current draft — receipts are attached to draft expenses, not the static table.
      const draft = await getCurrentDraft();
      if (!draft) {
        return {
          content: [{ type: 'text' as const, text: 'No draft expense report exists. Create one first.' }],
          structuredContent: { error: 'No draft expense report exists. Create one first.' },
          isError: true,
        };
      }

      let submitted = 0;
      let skipped = 0;
      let failed = 0;

      for (const receipt of receipts) {
        const expenseId = receipt.expense_id;
        const receiptType = receipt.type ?? 'odsp';
        console.log(`  Processing receipt — expense_id: ${expenseId}, type: ${receiptType}, file_name: ${receipt.file_name ?? '(none)'}, file_url: ${receipt.file_url}`);

        let receiptAttachment: { file_name: string; file_url: string; mime_type: string };

        if (receiptType === 'email') {
          // ── Email receipt: extract file attachments from the OWA email URL ──
          const emailUrl = receipt.file_url;

          if (!graphTokenMail) {
            console.warn(`  No Mail.Read Graph token — cannot process email receipt for ${expenseId}`);
            failed++;
            continue;
          }
          try {
            const attachments = await fetchAttachmentsFromEmail(emailUrl, graphTokenMail);
            console.log(`  Email attachments returned: ${attachments.length}`);
            if (attachments.length === 0) {
              console.warn(`  Email has no file attachments — skipping receipt for ${expenseId}`);
              failed++;
              continue;
            }
            // Use the first file attachment as the receipt
            const att = attachments[0];
            console.log(`  Email attachment: ${att.name} (${att.size} bytes, ${att.mimeType})`);
            console.log(`  contentBytes present: ${!!att.contentBytes}, length: ${att.contentBytes?.length ?? 0}`);
            console.log(`  downloadUrl: ${att.downloadUrl}`);
            // Save attachment content locally so the widget can access it without a Graph token
            const localUrl = await saveReceiptLocally({ fileName: att.name, contentBytes: att.contentBytes });
            console.log(`  Saved locally: ${localUrl}`);
            receiptAttachment = {
              file_name: att.name,
              file_url: toFullReceiptUrl(localUrl, baseUrl),
              mime_type: att.mimeType,
            };
            console.log(`  Final receipt URL: ${receiptAttachment.file_url}`);
            if (attachments.length > 1) {
              console.log(`  Note: email contained ${attachments.length} file attachments; using the first one.`);
            }
          } catch (err) {
            console.error(`  Email attachment extraction failed for ${expenseId}:`, err);
            failed++;
            continue;
          }
        } else {
          // ── ODSP receipt: fetch/verify via SharePoint/OneDrive sharing link ──
          const file_name = receipt.file_name ?? receipt.file_url;
          const file_url = receipt.file_url;

          let resolvedName = file_name;
          let resolvedMime = 'application/octet-stream';

          if (graphTokenOdsp) {
            try {
              const item = await fetchReceiptViaGraph(file_url, graphTokenOdsp);
              resolvedName = item.name || file_name;
              resolvedMime = item.mimeType;
              console.log(`  Graph file verified: ${resolvedName} (${item.size} bytes, ${resolvedMime})`);
              console.log(`  Graph downloadUrl present: ${!!item.downloadUrl}`);

              if (!item.downloadUrl) {
                console.error(`  No download URL returned by Graph for ${file_url} — using SharePoint link as fallback`);
                receiptAttachment = {
                  file_name: resolvedName,
                  file_url: file_url,
                  mime_type: resolvedMime,
                };
              } else {
                // Download the file server-side so the widget can access it without a Graph token
                const localUrl = await saveReceiptLocally({ fileName: resolvedName, downloadUrl: item.downloadUrl });
                receiptAttachment = {
                  file_name: resolvedName,
                  file_url: toFullReceiptUrl(localUrl, baseUrl),
                  mime_type: resolvedMime,
                };
              }
            } catch (err) {
              console.error(`  Graph file fetch/download failed for ${file_url}:`, err);
              failed++;
              continue;
            }
          } else {
            console.warn('  No Graph token — skipping file verification (consent may be required).');
            receiptAttachment = {
              file_name: file_name,
              file_url: file_url,
              mime_type: resolvedMime,
            };
          }
        }

        // Match against expenses in the current draft
        const draftIdx = expenseId
          ? draft.expenses.findIndex(e => e.expense_id === expenseId)
          : -1;

        if (draftIdx >= 0) {
          console.log(`  Receipt matched to expense ${expenseId}: ${receiptAttachment.file_name}`);
          draft.expenses[draftIdx] = {
            ...draft.expenses[draftIdx],
            status: 'receipt_submitted',
            receipt_attachment: receiptAttachment,
            receipt_source: receiptType as 'email' | 'odsp',
            receipt_match: 'matched',
          };
          submitted++;
        } else {
          console.log(`  No draft expense matched '${expenseId}' — skipping receipt`);
          skipped++;
        }
      }

      // Persist the updated draft back to the Drafts table
      await upsertDraft(draft.draft_id, draft.expenses);

      const processedExpenses = draft.expenses.filter(e => e.receipt_match);
      const stillMissing = draft.expenses.filter(e => !e.receipt_attachment);
      let summaryText = `Processed ${receipts.length} receipt(s): ` +
        `${submitted} matched to existing expenses, ` +
        `${skipped} skipped (no match), ${failed} failed.`;
      if (stillMissing.length > 0) {
        summaryText += `\n\nExpense IDs still missing receipts: ${stillMissing.map(e => e.expense_id).join(', ')}`;
      }

      const toolResult = {
        content: [
          {
            type: 'text' as const,
            text: summaryText,
          },
        ],
        structuredContent: {
          success: failed === 0,
          draft_id: draft.draft_id,
          total_count: draft.expenses.length,
          expenses: draft.expenses,
        },
        _meta: { ui: { resourceUri: EXPENSE_REPORT_DRAFT_RESOURCE_URI } },
      };

      // Small delay so the widget has time to finish loading and register
      // its ontoolresult handler before the result is delivered.
      await new Promise(resolve => setTimeout(resolve, 4000));

      return toolResult;
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: create_expense_report_draft
  // ---------------------------------------------------------------------------
  registerAppTool(
    server,
    'create_expense_report_draft',
    {
      description: 'Create draft expense for selected expense IDs. Make sure you have selected expense ids in your context before you call this tool. Otherwise ignore and call list_expense_items for user to first select the expenses, DO not assume any default expense selection and do not create a draft report without explicit expense IDs in context',
      annotations: { readOnlyHint: false },
      _meta: { ui: { resourceUri: EXPENSE_REPORT_DRAFT_RESOURCE_URI } },
      inputSchema: {
        expense_ids: z.array(z.string()).describe(
          'List of selected expense IDs to include in the draft report.'
        ),
      },
    },
    async (args) => {
      const { expense_ids } = args as { expense_ids: string[] };
      console.log(`create_expense_report_draft: creating draft with ${expense_ids.length} expense(s).`);

      if (!userToken) {
        return {
          content: [{ type: 'text' as const, text: 'No user token found. Authentication required.' }],
          structuredContent: { error: 'No user token found. Authentication required.' },
          isError: true,
        };
      }

      // Always replace any existing draft with the new selection
      const draftId = `DRAFT-${Date.now()}`;
      // Clean up receipt files from the previous draft
      await cleanupDraftReceipts();
      const selected = await getExpensesByIds(expense_ids);
      let draft = await upsertDraft(draftId, selected);
      console.log(`create_expense_report_draft: created draft ${draft.draft_id} with ${selected.length} expense(s).`);

      const result = {
        success: true,
        draft_id: draft.draft_id,
        total_count: draft.expenses.length,
        expenses: draft.expenses,
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: `Draft report ${draft.draft_id} with ${draft.expenses.length} expense(s).`,
          },
        ],
        structuredContent: result,
        _meta: { ui: { resourceUri: EXPENSE_REPORT_DRAFT_RESOURCE_URI } },
      };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: fetch_draft_expense_report
  // ---------------------------------------------------------------------------
  registerAppTool(
    server,
    'fetch_draft_expense_report',
    {
      description:
        'Fetch the current draft expense report. Returns expense details of the ' +
        'previously created draft. If no draft exists, returns a message indicating ' +
        'no draft has been created yet.',
      annotations: { readOnlyHint: true },
      _meta: {},
      inputSchema: {},
    },
    async () => {
      const draft = await getCurrentDraft();
      if (!draft) {
        return {
          content: [{ type: 'text' as const, text: 'No draft expense report has been created yet.' }],
          structuredContent: { success: false, message: 'No draft expense report has been created yet.' },
        };
      }

      const result = {
        success: true,
        draft_id: draft.draft_id,
        total_count: draft.expenses.length,
        expenses: draft.expenses,
      };

      const withoutReceipt = draft.expenses.filter(e => !e.receipt_attachment);
      let text = `Draft report ${draft.draft_id} contains ${draft.expenses.length} expense(s).`;
      if (withoutReceipt.length > 0) {
        text += `\n\nExpense IDs still missing receipts: ${withoutReceipt.map(e => e.expense_id).join(', ')}`;
      }

      return {
        content: [
          {
            type: 'text' as const,
            text,
          },
        ],
        structuredContent: result,
      };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: submit_expense_report  (app-only — initiated by widget Submit button)
  // ---------------------------------------------------------------------------
  registerAppTool(
    server,
    'submit_expense_report',
    {
      description:
        'Submits the finalised expense report for manager approval. ' +
        'Call this tool only after all receipts have been attached to their ' +
        'corresponding expense line items (i.e. every expense has status ' +
        '"receipt_submitted"). Accepts the full list of expense IDs to include ' +
        'in the report and an optional submitter note. Returns a confirmation ' +
        'with the generated report reference number.',
      annotations: { readOnlyHint: false },
      _meta: {
        ui: {
          visibility: ['app'],
        },
      },
      inputSchema: {
        expense_ids: z.array(z.string()).describe(
          'List of expense IDs to include in the submitted report.'
        ),
        draft_id: z.string().optional().describe(
          'Optional draft report ID if submitting from an existing draft.'
        ),
        note: z.string().optional().describe(
          'Optional submitter note to attach to the expense report.'
        ),
      },
    },
    async (args) => {
      const { expense_ids, draft_id, note } = args as { expense_ids: string[]; draft_id?: string; note?: string };
      console.log(`submit_expense_report: submitting ${expense_ids.length} expense(s).`);

      // Generate a report reference and record submission
      const reportRef = `RPT-${Date.now()}`;
      const result = {
        success: true,
        report_reference: reportRef,
        draft_id: draft_id ?? null,
        submitted_count: expense_ids.length,
        message: `Expense report ${reportRef} submitted successfully with ${expense_ids.length} expense(s).${draft_id ? ` Draft: ${draft_id}.` : ''}${note ? ` Note: ${note}` : ''}`,
      };

      console.log(`submit_expense_report: report ${reportRef} created.`);

      // Delete the draft from the Drafts table after successful submission
      await deleteDraft();
      // Clean up downloaded receipt files
      await cleanupDraftReceipts();

      return {
        content: [{ type: 'text' as const, text: result.message }],
        structuredContent: result,
        _meta: {},
      };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: sign_out
  // ---------------------------------------------------------------------------
  server.tool(
    'sign_out',
    'Clears the cached OBO (On-Behalf-Of) Graph token for the current user and ' +
    'triggers a fresh sign-in and consent flow. Call this tool when the user wants ' +
    'to sign out, switch accounts, re-grant consent, or reset their authentication ' +
    'state. After calling this tool the user must re-authenticate — ask them to ' +
    'refresh or re-open the agent conversation to sign in again.',
    {},
    async () => {
      if (userToken) {
        const prefix = userToken.slice(0, 20);
        let evicted = 0;
        for (const key of oboTokenCache.keys()) {
          if (key.startsWith(prefix)) {
            oboTokenCache.delete(key);
            evicted++;
          }
        }
        console.log(`sign_out: cleared ${evicted} cached OBO token(s) for the current user.`);
      } else {
        console.log('sign_out: no user token present, nothing to clear.');
      }

      const result = {
        success: true,
        message:
          'Your cached credentials have been cleared. Please refresh or re-open this ' +
          'conversation to sign in again and grant any required consent.',
      };
      return {
        content: [
          {
            type: 'text' as const,
            text: result.message,
          },
        ],
        structuredContent: result,
      };
    }
  );

  return server;
}

// =============================================================================
// Express App
// =============================================================================
const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve receipt files (sample PDFs + downloaded attachments)
app.use('/receipts', express.static(RECEIPTS_DIR));

// Health check
app.get('/health', (_req, res) => {
  res.json({ status: 'ok', service: 'Expense Submission MCP Server' });
});

// =============================================================================
// OAuth proxy — /authorize and /token
// These endpoints proxy the OAuth 2.0 Authorization Code flow to Microsoft
// Entra ID so that clients like VS Code can complete Entra SSO against this
// server's URL without requiring a separate identity endpoint.
// =============================================================================

// GET /authorize — redirect the browser to Entra's authorize endpoint
app.get('/authorize', (req: Request, res: Response) => {
  const entraAuthorizeUrl = new URL(
    `https://login.microsoftonline.com/${ENTRA_TENANT_ID}/oauth2/v2.0/authorize`
  );

  // Forward all query parameters from the client (client_id, response_type,
  // redirect_uri, state, code_challenge, code_challenge_method, scope, etc.)
  for (const [key, value] of Object.entries(req.query)) {
    if (typeof value === 'string') {
      entraAuthorizeUrl.searchParams.set(key, value);
    }
  }

  // Ensure client_id is set (use ours if the caller didn't provide one)
  if (!entraAuthorizeUrl.searchParams.has('client_id')) {
    entraAuthorizeUrl.searchParams.set('client_id', OBO_CLIENT_ID);
  }

  // Default scope if none provided
  if (!entraAuthorizeUrl.searchParams.has('scope')) {
    entraAuthorizeUrl.searchParams.set(
      'scope',
      `api://${OBO_CLIENT_ID}/access_as_user openid profile offline_access`
    );
  }

  console.log(`OAuth /authorize → redirecting to Entra: ${entraAuthorizeUrl.toString()}`);
  res.redirect(entraAuthorizeUrl.toString());
});

// POST /token — proxy the token exchange to Entra's token endpoint
app.post('/token', async (req: Request, res: Response) => {
  const entraTokenUrl = `https://login.microsoftonline.com/${ENTRA_TENANT_ID}/oauth2/v2.0/token`;

  // Build form body — forward everything the client sent
  const body = new URLSearchParams();
  const source = req.is('application/x-www-form-urlencoded') ? req.body : req.query;
  for (const [key, value] of Object.entries(source as Record<string, string>)) {
    if (typeof value === 'string') {
      body.set(key, value);
    }
  }

  // Ensure client_id and client_secret are included
  if (!body.has('client_id')) {
    body.set('client_id', OBO_CLIENT_ID);
  }
  if (!body.has('client_secret') && OBO_CLIENT_SECRET) {
    body.set('client_secret', OBO_CLIENT_SECRET);
  }

  // Default scope if none provided
  if (!body.has('scope')) {
    body.set(
      'scope',
      `api://${OBO_CLIENT_ID}/access_as_user openid profile offline_access`
    );
  }

  console.log(`OAuth /token → proxying to Entra (grant_type: ${body.get('grant_type')})`);

  try {
    const tokenResp = await fetch(entraTokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
    const data = await tokenResp.json();
    res.status(tokenResp.status).json(data);
  } catch (err) {
    console.error('OAuth /token proxy error:', err);
    res.status(502).json({ error: 'token_proxy_error', error_description: 'Failed to exchange token with Entra ID' });
  }
});

// MCP endpoint — stateless HTTP (new server + transport per request)
app.all('/mcp', async (req: Request, res: Response) => {
  console.log(`\n[MCP] ── ${req.method} /mcp ──`);
  console.log(`[MCP] JSON-RPC method: ${(req.body as { method?: string })?.method ?? '(none)'}`);
  console.log(`[MCP] Has Authorization header: ${!!req.headers.authorization}`);
  const rpcMethod =
    typeof (req.body as { method?: unknown })?.method === 'string'
      ? ((req.body as { method: string }).method)
      : '';
  const resourceUri = (req.body as { params?: { uri?: unknown } })?.params?.uri;
  const isUiResourceRead =
    rpcMethod === 'resources/read' &&
    typeof resourceUri === 'string' &&
    resourceUri.startsWith('ui://');
  const isResourceList = rpcMethod === 'resources/list';
  const canBypassAuth = isUiResourceRead || isResourceList;

  // Some hosts fetch UI resources without forwarding bearer auth.
  // Keep tools/auth-sensitive operations protected while allowing UI HTML retrieval.
  const userToken = canBypassAuth ? null : await authenticateRequest(req);
  if (!canBypassAuth && !userToken) {
    res.status(401).json({
      jsonrpc: '2.0',
      error: {
        code: -32001,
        message: 'Unauthorized: valid Bearer token required',
      },
      id: null,
    });
    return;
  }

  // Create a per-request server (stateless) and transport
  const proto = req.headers['x-forwarded-proto'] || req.protocol || 'https';
  const host = req.headers['x-forwarded-host'] || req.headers.host || `localhost:${PORT}`;
  const baseUrl = `${proto}://${host}`;
  const server = createMcpServer(userToken, baseUrl);
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined, // stateless mode — no sessions
    enableJsonResponse: true,
  });

  res.on('finish', () => {
    server.close().catch(() => {});
  });

  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err) {
    console.error('MCP request error:', err);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: '2.0',
        error: { code: -32603, message: 'Internal server error' },
        id: null,
      });
    }
  }
});

// =============================================================================
// Start Server
// =============================================================================
app.listen(PORT, '0.0.0.0', () => {
  console.log('Starting Expense Submission MCP Server (TypeScript)...');
  console.log('Supported authentication methods:');
  console.log(
    '  - MS SSO JWT tokens (audiences: CLIENT_ID, api://CLIENT_ID, Application ID URI)'
  );
  console.log('  - API Key (development): mock_mcp_api_key');
  console.log(`Server running on http://0.0.0.0:${PORT}/mcp`);
});
