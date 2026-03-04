# HR Consultant MCP Server

MCP server with rich Fluent UI React widgets for managing HR consultants, projects, and assignments.

<a href="https://www.youtube.com/watch?v=kNXT7Syf9fQ" target="_blank"><img src="../../demos/fake-play-thumbnail.png" alt="Watch the demo"></a>

> **<a href="https://www.youtube.com/watch?v=kNXT7Syf9fQ" target="_blank">Watch the demo on YouTube</a>** | [Demo video file](../../demos/trey-demo.mp4)

## Prerequisites

| Requirement | Version |
|---|---|
| [Node.js](https://nodejs.org/) | 18, 20, or 22 |
| npm | ≥ 9 |

Azurite is included as a dev dependency — no separate install needed.

## Getting Started

> All commands below should be run from the `src/mcpserver/` directory.

### 1. Install dependencies

```bash
npm run install:all
```

This installs packages for the root, server, and widgets workspaces.

### 2. Start Azurite (local Azure Table Storage)

Open a **separate terminal** and run:

```bash
npm run start:azurite
```

This starts the local Azure Table Storage emulator on port 10002. **Keep this terminal running** throughout development.

### 3. Seed the database

In a **new terminal** (with Azurite still running):

```bash
npm run seed
```

Populates the local database with sample consultants, projects, and assignments.

### 4. Create a dev tunnel

Use [Dev Tunnels](https://learn.microsoft.com/azure/developer/dev-tunnels/) to expose your local MCP server publicly:

```bash
devtunnel host -p 3001 --allow-anonymous
```

Copy the forwarded URL (e.g. `https://<tunnel-id>-3001.aue.devtunnels.ms`). You'll need this in the next step.

### 5. Create your `.env` file

Copy the sample and update it with your dev tunnel URL:

```bash
cp .env.sample .env
```

Then edit `.env` and replace the placeholder values:

```dotenv
SERVER_BASE_URL=https://<your-tunnel-id>-3001.aue.devtunnels.ms/
ADDITIONAL_ALLOWED_ORIGINS=https://<your-tunnel-id>-3001.aue.devtunnels.ms,http://localhost:6274
```

> `SERVER_BASE_URL` is injected into the widgets so they can call back to the MCP server. Without it, widgets will default to `http://localhost:3001`.

### 6. Build the widgets

```bash
npm run build:widgets
```

Compiles the React + Fluent UI widgets into single-file HTML assets in the `assets/` folder.

### 7. Start the MCP server

```bash
npm run start:server
```

The server starts on **http://localhost:3001** with the MCP endpoint at **http://localhost:3001/mcp**.

### 8. Test the server (MCP Inspector)

With the server running, open a **new terminal** and launch the [MCP Inspector](https://modelcontextprotocol.io/docs/tools/inspector):

```bash
npm run inspector
```

This opens a browser-based UI where you can:

- Browse all registered tools and their schemas
- Call tools with custom inputs and inspect the JSON responses
- Verify that widget HTML is returned correctly

Use the Inspector to confirm the server is working before connecting it to Copilot.

## Connect to Microsoft 365 Copilot

To connect this MCP server to a Microsoft 365 Copilot Declarative Agent, see the [Trey Research Declarative Agent README](../../README.md) for full instructions on provisioning, dev tunnel configuration in the agent manifest, and testing in Copilot.

## MCP Tools

### Widget Tools

| Tool | Widget | Description |
|---|---|---|
| `show-hr-dashboard` | Dashboard | KPIs, consultant cards, project list. Optional filters: `consultantName`, `projectName`, `skill`, `role`, `billable`. |
| `show-consultant-profile` | Profile | Detailed profile card with contact info, skills, certifications, roles, and assignments. Requires `consultantId`. |
| `show-project-details` | Dashboard | Project detail with assigned consultants and forecasted hours. Requires `projectId`. |
| `search-consultants` | Bulk Editor | Search consultants by `skill` or `name`, results shown in the bulk editor for viewing and editing. |
| `show-bulk-editor` | Bulk Editor | View and edit consultant records. Optional filters: `skill`, `name`. |

### Data Tools

| Tool | Description |
|---|---|
| `update-consultant` | Update a single consultant's name, email, phone, skills, or roles. |
| `bulk-update-consultants` | Batch-update multiple consultant records at once. |
| `assign-consultant-to-project` | Assign a consultant to a project with a role, optional rate. |
| `bulk-assign-consultants` | Assign multiple consultants to a project at once. |
| `remove-assignment` | Remove a consultant's assignment from a project. |

## Sample Prompts

| Prompt | What it does |
|---|---|
| *Show the HR dashboard* | Opens the HR consultant dashboard widget |
| *I need a React developer for the Copilot project at Consolidated Messenger. Find someone with React skills, show me their profile, and assign them as a Developer.* | Searches consultants by skill, displays a profile card, and assigns the consultant to a project — all by name, no IDs needed |
| *Show me the HR dashboard filtered to only billable assignments. Which consultants have the most forecasted hours, and are any of them over-allocated?* | Opens the interactive dashboard with a billable filter applied, then the AI analyzes forecast data across consultants to surface workload insights |
| *We need to staff the Disaster Recovery project at Relecloud. Show me the project details, then find all consultants who have Python or Java skills and bulk-assign them as Developers at $120/hr.* | Chains project lookup, skill-based consultant search, and bulk assignment in a single conversation — replacing multiple clicks across an HR system |
| *Compare Avery Howard and Sanjay Puranik — show me both their profiles side by side. Who has more certifications, and which projects are they currently assigned to?* | Fetches two consultant profiles by name and synthesizes a comparison of certifications, skills, and active assignments |

## Development

```bash
npm run dev:server         # Server with hot-reload (tsx --watch)
npm run build:widgets      # Rebuild widgets after changes
npm run inspector          # Launch MCP Inspector for testing
```
