/**
 * HR Consultant MCP Server factory.
 *
 * Uses the MCP Apps standard (@modelcontextprotocol/ext-apps) for
 * widget resources and tool registration (replaces the old OpenAI
 * Apps SDK / text/html+skybridge approach).
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerAppTool, registerAppResource, RESOURCE_MIME_TYPE, } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import * as db from "./db.js";
import { getPublicServerUrl } from "./index.js";
// ─── Widget HTML loader ────────────────────────────────────────────
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ASSETS_DIR = path.resolve(__dirname, "..", "..", "assets");
/**
 * Read a widget's built HTML and inject the public server URL so the
 * widget can call back to this server even when it is loaded through a
 * tunnel or proxy (avoids private-network / mixed-content blocks).
 *
 * Injects a small `<script>` right after `<head>` that sets
 * `window.__SERVER_BASE_URL__`.
 */
function readWidgetHtml(componentName) {
    if (!fs.existsSync(ASSETS_DIR)) {
        throw new Error(`Widget assets not found at ${ASSETS_DIR}. Run "npm run build:widgets" first.`);
    }
    let html;
    const directPath = path.join(ASSETS_DIR, `${componentName}.html`);
    if (fs.existsSync(directPath)) {
        html = fs.readFileSync(directPath, "utf8");
    }
    else {
        // Try hashed fallback
        const candidates = fs
            .readdirSync(ASSETS_DIR)
            .filter((f) => f.startsWith(`${componentName}-`) && f.endsWith(".html"))
            .sort();
        const fallback = candidates[candidates.length - 1];
        if (fallback) {
            html = fs.readFileSync(path.join(ASSETS_DIR, fallback), "utf8");
        }
    }
    if (!html) {
        throw new Error(`Widget HTML for "${componentName}" not found in ${ASSETS_DIR}.`);
    }
    // Inject public server URL into the widget
    const serverUrl = getPublicServerUrl();
    const injection = `<script>window.__SERVER_BASE_URL__=${JSON.stringify(serverUrl)};</script>`;
    html = html.replace("<head>", `<head>${injection}`);
    return html;
}
// ─── Widget URI definitions ────────────────────────────────────────
const DASHBOARD_URI = "ui://trey-hr/hr-dashboard.html";
const PROFILE_URI = "ui://trey-hr/consultant-profile.html";
const BULK_EDITOR_URI = "ui://trey-hr/bulk-editor.html";
// ─── Entity → plain object helpers ─────────────────────────────────
function parseConsultant(c) {
    return {
        id: c.rowKey,
        name: c.name,
        email: c.email,
        phone: c.phone,
        photoUrl: c.consultantPhotoUrl,
        location: JSON.parse(c.location || "{}"),
        skills: JSON.parse(c.skills || "[]"),
        certifications: JSON.parse(c.certifications || "[]"),
        roles: JSON.parse(c.roles || "[]"),
    };
}
function parseProject(p) {
    return {
        id: p.rowKey,
        name: p.name,
        description: p.description,
        clientName: p.clientName,
        clientContact: p.clientContact,
        clientEmail: p.clientEmail,
        location: JSON.parse(p.location || "{}"),
    };
}
function parseAssignment(a) {
    return {
        id: a.rowKey,
        projectId: a.projectId,
        consultantId: a.consultantId,
        role: a.role,
        billable: a.billable,
        rate: a.rate,
        forecast: JSON.parse(a.forecast || "[]"),
        delivered: JSON.parse(a.delivered || "[]"),
    };
}
// ─── Server factory ────────────────────────────────────────────────
export function createHRServer() {
    const server = new McpServer({ name: "trey-hr-consultant", version: "1.0.0" });
    // ─── Widget Resources ──────────────────────────────────────────
    registerAppResource(server, "HR Dashboard", DASHBOARD_URI, {
        mimeType: RESOURCE_MIME_TYPE,
        description: "HR Dashboard widget markup",
    }, async () => {
        const html = readWidgetHtml("hr-dashboard");
        return { contents: [{ uri: DASHBOARD_URI, mimeType: RESOURCE_MIME_TYPE, text: html }] };
    });
    registerAppResource(server, "Consultant Profile", PROFILE_URI, {
        mimeType: RESOURCE_MIME_TYPE,
        description: "Consultant Profile widget markup",
    }, async () => {
        const html = readWidgetHtml("consultant-profile");
        return { contents: [{ uri: PROFILE_URI, mimeType: RESOURCE_MIME_TYPE, text: html }] };
    });
    registerAppResource(server, "Bulk Editor", BULK_EDITOR_URI, {
        mimeType: RESOURCE_MIME_TYPE,
        description: "Bulk Editor widget markup",
    }, async () => {
        const html = readWidgetHtml("bulk-editor");
        return { contents: [{ uri: BULK_EDITOR_URI, mimeType: RESOURCE_MIME_TYPE, text: html }] };
    });
    // ─── Widget Tools (render UI) ──────────────────────────────────
    // show-hr-dashboard
    registerAppTool(server, "show-hr-dashboard", {
        title: "Show HR Dashboard",
        description: "Display the HR consultant dashboard with KPIs. Accepts optional filters: consultantName, projectName, skill, role, billable — the dashboard auto-applies them so users see a focused view.",
        inputSchema: {
            consultantName: z.string().optional().describe("Optional consultant name to pre-filter the dashboard (partial match, case-insensitive)."),
            projectName: z.string().optional().describe("Optional project name to pre-filter the dashboard (partial match, case-insensitive)."),
            skill: z.string().optional().describe("Optional skill to pre-filter the dashboard — shows only consultants with this skill and their assignments."),
            role: z.string().optional().describe("Optional role to pre-filter assignments (e.g. 'Developer', 'Architect')."),
            billable: z.boolean().optional().describe("Optional — set true to show only billable assignments, false for non-billable."),
        },
        annotations: { readOnlyHint: true },
        _meta: { ui: { resourceUri: DASHBOARD_URI } },
    }, async ({ consultantName, projectName, skill, role, billable }) => {
        const [consultants, projects, assignments] = await Promise.all([
            db.getAllConsultants(),
            db.getAllProjects(),
            db.getAllAssignments(),
        ]);
        const totalBillableHours = assignments.reduce((sum, a) => {
            if (!a.billable)
                return sum;
            const forecast = JSON.parse(a.forecast || "[]");
            return sum + forecast.reduce((s, f) => s + f.hours, 0);
        }, 0);
        // Build active filter hints to pass to the widget
        const activeFilters = {};
        const filterDescParts = [];
        if (consultantName) {
            activeFilters.consultantName = consultantName;
            filterDescParts.push(`consultant: "${consultantName}"`);
        }
        if (projectName) {
            const q = projectName.toLowerCase();
            const matchedIds = projects.filter((p) => p.name.toLowerCase().includes(q)).map((p) => p.rowKey);
            activeFilters.projectIds = matchedIds;
            activeFilters.projectName = projectName;
            filterDescParts.push(`project: "${projectName}"`);
        }
        if (skill) {
            activeFilters.skill = skill;
            filterDescParts.push(`skill: "${skill}"`);
        }
        if (role) {
            activeFilters.role = role;
            filterDescParts.push(`role: "${role}"`);
        }
        if (billable !== undefined) {
            activeFilters.billable = billable;
            filterDescParts.push(billable ? "billable only" : "non-billable only");
        }
        const dashboardData = {
            consultants: consultants.map(parseConsultant),
            projects: projects.map(parseProject),
            assignments: assignments.map((a) => {
                const parsed = parseAssignment(a);
                const proj = projects.find((p) => p.rowKey === a.projectId);
                const cons = consultants.find((c) => c.rowKey === a.consultantId);
                return {
                    ...parsed,
                    projectName: proj?.name ?? "Unknown",
                    clientName: proj?.clientName ?? "Unknown",
                    consultantName: cons?.name ?? "Unknown",
                };
            }),
            summary: {
                totalConsultants: consultants.length,
                totalProjects: projects.length,
                totalAssignments: assignments.length,
                totalBillableHours,
            },
            ...(Object.keys(activeFilters).length > 0 ? { filters: activeFilters } : {}),
        };
        const filterDesc = filterDescParts.length > 0
            ? ` (filtered by ${filterDescParts.join(", ")})`
            : "";
        return {
            content: [
                {
                    type: "text",
                    text: `HR Dashboard: ${consultants.length} consultants, ${projects.length} projects, ${totalBillableHours} billable hours forecasted.${filterDesc}`,
                },
            ],
            structuredContent: dashboardData,
        };
    });
    // show-consultant-profile
    registerAppTool(server, "show-consultant-profile", {
        title: "Show Consultant Profile",
        description: "Display a detailed profile card for a specific consultant, including contact info, skills, certifications, roles, and current assignments.",
        inputSchema: {
            consultantId: z.string().describe("The ID of the consultant to view."),
        },
        annotations: { readOnlyHint: true },
        _meta: { ui: { resourceUri: PROFILE_URI } },
    }, async ({ consultantId }) => {
        const consultant = await db.getConsultantById(consultantId);
        if (!consultant) {
            return {
                content: [{ type: "text", text: `Consultant ${consultantId} not found.` }],
                isError: true,
            };
        }
        const assignments = await db.getAssignmentsByConsultant(consultantId);
        const allProjects = await db.getAllProjects();
        const projectMap = new Map(allProjects.map((p) => [p.rowKey, parseProject(p)]));
        const enrichedAssignments = assignments.map((a) => ({
            ...parseAssignment(a),
            projectName: projectMap.get(a.projectId)?.name ?? "Unknown",
            clientName: projectMap.get(a.projectId)?.clientName ?? "Unknown",
        }));
        const profileData = {
            consultant: parseConsultant(consultant),
            assignments: enrichedAssignments,
        };
        return {
            content: [
                {
                    type: "text",
                    text: `Profile for ${consultant.name}: ${JSON.parse(consultant.skills || "[]").join(", ")} | ${enrichedAssignments.length} active assignment(s).`,
                },
            ],
            structuredContent: profileData,
        };
    });
    // search-consultants
    registerAppTool(server, "search-consultants", {
        title: "Search Consultants",
        description: "Search consultants by skill or name. Returns matching consultants in the bulk editor widget for easy viewing and editing.",
        inputSchema: {
            skill: z.string().optional().describe("Skill to search for (partial match)."),
            name: z.string().optional().describe("Name to search for (partial match)."),
        },
        annotations: { readOnlyHint: true },
        _meta: { ui: { resourceUri: BULK_EDITOR_URI } },
    }, async ({ skill, name: nameFilter }) => {
        let results = await db.getAllConsultants();
        if (skill) {
            results = results.filter((c) => {
                const skills = JSON.parse(c.skills || "[]");
                return skills.some((s) => s.toLowerCase().includes(skill.toLowerCase()));
            });
        }
        if (nameFilter) {
            results = results.filter((c) => c.name.toLowerCase().includes(nameFilter.toLowerCase()));
        }
        return {
            content: [
                {
                    type: "text",
                    text: `Found ${results.length} consultant(s) matching criteria.`,
                },
            ],
            structuredContent: {
                consultants: results.map(parseConsultant),
            },
        };
    });
    // show-bulk-editor
    registerAppTool(server, "show-bulk-editor", {
        title: "Show Bulk Editor",
        description: "Open the bulk editor widget to view and edit consultant records. Accepts optional filters: skill, name — to show only matching consultants.",
        inputSchema: {
            skill: z.string().optional().describe("Optional skill to filter consultants (partial match, case-insensitive)."),
            name: z.string().optional().describe("Optional name to filter consultants (partial match, case-insensitive)."),
        },
        annotations: { readOnlyHint: false },
        _meta: { ui: { resourceUri: BULK_EDITOR_URI } },
    }, async ({ skill, name: nameFilter }) => {
        let consultants = await db.getAllConsultants();
        if (skill) {
            consultants = consultants.filter((c) => {
                const skills = JSON.parse(c.skills || "[]");
                return skills.some((s) => s.toLowerCase().includes(skill.toLowerCase()));
            });
        }
        if (nameFilter) {
            consultants = consultants.filter((c) => c.name.toLowerCase().includes(nameFilter.toLowerCase()));
        }
        const filterDesc = [skill && `skill: "${skill}"`, nameFilter && `name: "${nameFilter}"`].filter(Boolean).join(", ");
        return {
            content: [
                {
                    type: "text",
                    text: `Bulk editor loaded with ${consultants.length} consultant record(s)${filterDesc ? ` (filtered by ${filterDesc})` : ""}.`,
                },
            ],
            structuredContent: {
                consultants: consultants.map(parseConsultant),
            },
        };
    });
    // show-project-details
    registerAppTool(server, "show-project-details", {
        title: "Show Project Details",
        description: "Display detailed information about a specific project including its assigned consultants and forecasted hours.",
        inputSchema: {
            projectId: z.string().describe("The project ID."),
        },
        annotations: { readOnlyHint: true },
        _meta: { ui: { resourceUri: DASHBOARD_URI } },
    }, async ({ projectId }) => {
        const project = await db.getProjectById(projectId);
        if (!project) {
            return {
                content: [{ type: "text", text: `Project ${projectId} not found.` }],
                isError: true,
            };
        }
        const assignments = await db.getAssignmentsByProject(projectId);
        const allConsultants = await db.getAllConsultants();
        const consultantMap = new Map(allConsultants.map((c) => [c.rowKey, parseConsultant(c)]));
        const enrichedAssignments = assignments.map((a) => ({
            ...parseAssignment(a),
            consultantName: consultantMap.get(a.consultantId)?.name ?? "Unknown",
        }));
        const totalBillableHours = enrichedAssignments.reduce((sum, a) => {
            return sum + a.forecast.reduce((s, f) => s + f.hours, 0);
        }, 0);
        return {
            content: [
                {
                    type: "text",
                    text: `Project "${project.name}" for ${project.clientName}: ${enrichedAssignments.length} assignment(s).`,
                },
            ],
            structuredContent: {
                projects: [parseProject(project)],
                assignments: enrichedAssignments,
                consultants: allConsultants.map((c) => parseConsultant(c)),
                summary: {
                    totalConsultants: allConsultants.length,
                    totalProjects: 1,
                    totalAssignments: enrichedAssignments.length,
                    totalBillableHours,
                },
            },
        };
    });
    // ─── Data-only tools (no UI) ───────────────────────────────────
    // update-consultant
    server.tool("update-consultant", "Update a single consultant's information (name, email, phone, skills, roles).", {
        consultantId: z.string().describe("The ID of the consultant to update."),
        name: z.string().optional().describe("Updated name."),
        email: z.string().optional().describe("Updated email."),
        phone: z.string().optional().describe("Updated phone."),
        skills: z.array(z.string()).optional().describe("Updated skills list."),
        roles: z.array(z.string()).optional().describe("Updated roles list."),
    }, async ({ consultantId, ...updates }) => {
        const updated = await db.updateConsultant(consultantId, updates);
        if (!updated) {
            return {
                content: [{ type: "text", text: `Consultant ${consultantId} not found.` }],
                isError: true,
            };
        }
        return {
            content: [
                {
                    type: "text",
                    text: `Updated consultant ${updated.name} (ID: ${consultantId}).`,
                },
            ],
        };
    });
    // bulk-update-consultants
    server.tool("bulk-update-consultants", "Batch-update multiple consultant records at once.", {
        consultantIds: z.array(z.string()).describe("Array of consultant IDs to update."),
        name: z.string().optional().describe("New name for all."),
        email: z.string().optional().describe("New email for all."),
        phone: z.string().optional().describe("New phone for all."),
        skills: z.array(z.string()).optional().describe("New skills list for all."),
        roles: z.array(z.string()).optional().describe("New roles list for all."),
    }, async ({ consultantIds, ...changes }) => {
        const results = [];
        for (const consultantId of consultantIds) {
            const updated = await db.updateConsultant(consultantId, changes);
            results.push(updated
                ? `✓ Updated ${updated.name}`
                : `✗ Consultant ${consultantId} not found`);
        }
        return {
            content: [
                {
                    type: "text",
                    text: `Bulk update complete:\n${results.join("\n")}`,
                },
            ],
        };
    });
    // assign-consultant-to-project
    server.tool("assign-consultant-to-project", "Assign a single consultant to a project with a specified role, optional billing rate, and optional forecast hours.", {
        projectId: z.string().describe("The project ID to assign the consultant to."),
        consultantId: z.string().describe("The consultant ID to assign."),
        role: z.string().describe("The role the consultant will play on the project (e.g. Architect, Developer, Designer, Project lead)."),
        billable: z.boolean().optional().describe("Whether the assignment is billable. Defaults to true."),
        rate: z.number().optional().describe("Hourly rate for the consultant on this project. Defaults to 0."),
    }, async ({ projectId, consultantId, role, billable, rate }) => {
        const project = await db.getProjectById(projectId);
        if (!project) {
            return {
                content: [{ type: "text", text: `Project ${projectId} not found.` }],
                isError: true,
            };
        }
        const consultant = await db.getConsultantById(consultantId);
        if (!consultant) {
            return {
                content: [{ type: "text", text: `Consultant ${consultantId} not found.` }],
                isError: true,
            };
        }
        await db.createAssignment({ projectId, consultantId, role, billable, rate });
        return {
            content: [
                {
                    type: "text",
                    text: `Assigned ${consultant.name} to project "${project.name}" as ${role}${billable === false ? " (non-billable)" : ""}${rate ? ` at $${rate}/hr` : ""}.`,
                },
            ],
        };
    });
    // bulk-assign-consultants
    server.tool("bulk-assign-consultants", "Assign multiple consultants to a project at once. Each assignment includes a role, optional billing rate, and optional forecast.", {
        projectId: z.string().describe("The project ID to assign consultants to."),
        consultantIds: z.array(z.string()).describe("Array of consultant IDs to assign."),
        role: z.string().describe("The role for all assigned consultants."),
        billable: z.boolean().optional().describe("Whether the assignments are billable. Defaults to true."),
        rate: z.number().optional().describe("Hourly rate for all assigned consultants. Defaults to 0."),
    }, async ({ projectId, consultantIds, role, billable, rate }) => {
        const project = await db.getProjectById(projectId);
        if (!project) {
            return {
                content: [{ type: "text", text: `Project ${projectId} not found.` }],
                isError: true,
            };
        }
        const results = [];
        for (const consultantId of consultantIds) {
            const consultant = await db.getConsultantById(consultantId);
            if (!consultant) {
                results.push(`✗ Consultant ${consultantId} not found`);
                continue;
            }
            await db.createAssignment({ projectId, consultantId, role, billable, rate });
            results.push(`✓ Assigned ${consultant.name} as ${role}`);
        }
        return {
            content: [
                {
                    type: "text",
                    text: `Bulk assignment to "${project.name}" complete:\n${results.join("\n")}`,
                },
            ],
        };
    });
    // remove-assignment
    server.tool("remove-assignment", "Remove a consultant's assignment from a project.", {
        projectId: z.string().describe("The project ID."),
        consultantId: z.string().describe("The consultant ID to remove from the project."),
    }, async ({ projectId, consultantId }) => {
        const removed = await db.deleteAssignment(projectId, consultantId);
        if (!removed) {
            return {
                content: [{ type: "text", text: `Assignment for consultant ${consultantId} on project ${projectId} not found.` }],
                isError: true,
            };
        }
        return {
            content: [
                {
                    type: "text",
                    text: `Removed assignment: consultant ${consultantId} from project ${projectId}.`,
                },
            ],
        };
    });
    return server;
}
//# sourceMappingURL=mcp-server.js.map