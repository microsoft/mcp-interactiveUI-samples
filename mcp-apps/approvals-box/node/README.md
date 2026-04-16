# Approvals Box (MCP App Server)

## 1. What this sample is about

This directory contains a **Model Context Protocol (MCP) App Server** for AI-powered approval queue management. Built with Node.js, Express, and the MCP Apps SDK, it exposes a set of tools so you can manage your approval queue through natural conversation in Microsoft 365 Copilot — with visually rich inline widgets rendered directly in the chat.

The sample includes three interactive widgets:

- `pending-approvals` — Paginated list of approvals with status badges, risk indicators, amount, requester, and due date. Supports bulk approve/reject.
- `approval-detail` — Full detail view for a single approval: fields, approver chain, risk assessment, comments, and audit history. Includes inline approve/reject dialogs.
- `create-approval` — Guided form to submit a new approval request with pre-filled fields, type selection, and summary.

### Features

- **Risk triage** — Every approval is automatically scored (low/medium/high) with specific risk reasons surfaced in the detail widget
- **Bulk decisions** — Approve or reject multiple items at once directly from the list widget
- **Smart summaries** — Ask Copilot to summarize overdue items, flag high-risk approvals, or compare similar requests
- **Pre-filled forms** — Copilot populates the create form based on context from the conversation
- **AI-drafted rejections** — Ask Copilot to draft a rejection reason before confirming
- **Auto-seeded demo data** — A realistic dataset of ~50 approvals across 7 types is seeded automatically on first run

## 2. Sample prompts

| Prompt | What it does |
|---|---|
| Show me my pending approvals. | Opens the list widget with all pending approvals, sorted by due date. |
| Which approvals are high risk? | Filters the list to high-risk items and opens the widget. |
| Show me the details for the Activision InComm approval. | Resolves the approval by name and opens the detail widget. |
| Approve the pending purchase order from Kevin Walsh. | Opens the detail widget with the approve dialog pre-opened. |
| Bulk reject all low-priority travel exceptions. | Opens the list filtered to travel exceptions with the bulk-reject dialog. |
| Create a new vendor onboarding approval for Contoso. | Opens the create form with type and vendor fields pre-filled. |
| Which approvals are overdue and high risk? | Lists overdue high-risk approvals and summarizes the most urgent ones. |
| Draft a rejection reason for the capex request. | AI drafts a rejection note; confirm to reject. |

## 3. Pre-requisites

### Required (to run locally)

- Node.js 22+ (required for the built-in `node:sqlite` module)
- npm 10+

## 4. Development (run locally + dev tunnel)

### Step 1: Clone this repo and navigate to the approvals-box folder

```bash
git clone https://github.com/microsoft/mcp-interactiveUI-samples.git
cd mcp-interactiveUI-samples/mcp-apps/approvals-box/node
```

### Step 2: Install dependencies

```bash
npm install
```

### Step 3: Build and run

```bash
npm run build
npm start
```

Server endpoint: `http://localhost:3001/mcp`

> **Note:** The server uses Node.js's built-in `node:sqlite` module (Node 22+). Demo data is seeded automatically on first start — no manual setup needed.

### Expose via dev tunnel (example with ngrok)

```bash
ngrok http 3001
```

Use the public HTTPS URL from ngrok and append `/mcp` as your MCP spec URL.

### Run in watch mode (for development)

```bash
npm run dev
```

## 5. How to test in Copilot

1. Open [appPackage/ai-plugin.json](appPackage/ai-plugin.json).
2. Replace the spec URL with your tunnel MCP URL:

   ```json
   "runtimes": [
     {
       "type": "RemoteMCPServer",
       "spec": {
         "url": "<your-tunnel-url>/mcp",
         "mcp_tool_description": {
           "...": "..."
         }
       }
     }
   ]
   ```

3. Zip the [appPackage](appPackage) folder.
4. Sideload the package into Teams via the Teams Admin Center or Developer Portal.
5. Open Microsoft 365 Copilot and try the sample prompts above.

If you are customizing tool definitions, you can use M365 Agents Toolkit to generate a declarative agent from your MCP server URL. See [M365 Agents Toolkit Instructions](../../../M365-Agents-Toolkit-Instructions.md) for details.

## 6. Next steps

- Customize the seed data in `src/seed.ts` to match your own approval types and workflows
- Add new approval types to the `ApprovalType` enum in `src/types.ts`
- Adjust risk scoring logic in `src/risk.ts`
- Add new tools and widgets to extend the agent's capabilities
