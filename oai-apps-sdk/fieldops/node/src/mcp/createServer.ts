import { readFileSync } from "node:fs";
import path from "node:path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  getAssignmentsByIds,
  getAvailableTechnicians,
  getTechniciansByIds,
  listNewAssignments
} from "../logic/dispatch.js";
import { ListResourceTemplatesRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { WIDGET_MIME, TEMPLATE_URIS } from "./resources.js";
import { assignments } from "../data/mockData.js";

function widgetPath(fileName: string): string {
  return path.resolve(process.cwd(), "widgets", fileName);
}

function loadWidget(fileName: string): string {
  return readFileSync(widgetPath(fileName), "utf8");
}

type WidgetMeta = { templateUri: string; invoking: string; invoked: string };

const WIDGET_META: Record<string, WidgetMeta> = {
  list: { templateUri: TEMPLATE_URIS.list, invoking: "Finding new assignments\u2026", invoked: "New assignments ready." },
  map: { templateUri: TEMPLATE_URIS.map, invoking: "Loading map\u2026", invoked: "Map is ready." },
  plan: { templateUri: TEMPLATE_URIS.plan, invoking: "Building dispatch plan\u2026", invoked: "Dispatch plan ready." },
};

function widgetDescriptorMeta(meta: WidgetMeta) {
  return {
    "openai/outputTemplate": meta.templateUri,
    "openai/toolInvocation/invoking": meta.invoking,
    "openai/toolInvocation/invoked": meta.invoked,
    "openai/widgetAccessible": true,
  };
}

function widgetInvocationMeta(meta: WidgetMeta) {
  return {
    "openai/toolInvocation/invoking": meta.invoking,
    "openai/toolInvocation/invoked": meta.invoked,
  };
}

export function createMcpServer(): McpServer {
  const server = new McpServer({ name: "field-service-dispatch-agent", version: "0.1.0" });

  server.registerResource("dispatch-list", TEMPLATE_URIS.list, {}, async () => ({
    contents: [
      {
        uri: TEMPLATE_URIS.list,
        mimeType: WIDGET_MIME,
        text: loadWidget("list.html"),
        _meta: widgetDescriptorMeta(WIDGET_META.list),
      }
    ]
  }));

  server.registerResource("dispatch-map", TEMPLATE_URIS.map, {}, async () => ({
    contents: [
      {
        uri: TEMPLATE_URIS.map,
        mimeType: WIDGET_MIME,
        text: loadWidget("map.html"),
        _meta: widgetDescriptorMeta(WIDGET_META.map),
      }
    ]
  }));

  server.registerResource("dispatch-plan", TEMPLATE_URIS.plan, {}, async () => ({
    contents: [
      {
        uri: TEMPLATE_URIS.plan,
        mimeType: WIDGET_MIME,
        text: loadWidget("dispatch-plan.html"),
        _meta: widgetDescriptorMeta(WIDGET_META.plan),
      }
    ]
  }));

  server.registerTool(
    "list_new_assignments",
    {
      title: "List New Assignments",
      description:
        "Use this first for dispatch intake. Returns unassigned, recently created work orders (default: status=New and created in last 24 hours), optionally filtered by priority/region/team. Output includes assignment IDs, site, SLA due, required skills, status, and coordinates for downstream tools. Call before map or planning when user asks for new work.",
      inputSchema: {
        priority: z
          .enum(["Low", "Medium", "High"])
          .optional()
          .describe("Optional priority filter. Use High/Medium/Low to narrow returned assignments."),
        region: z
          .string()
          .optional()
          .describe("Optional region filter. Example: North, South, East, West."),
        team: z
          .string()
          .optional()
          .describe("Optional team filter. Example: Alpha, Beta."),
        maxHoursOld: z
          .number()
          .int()
          .min(1)
          .max(168)
          .optional()
          .describe("Optional recency window in hours. Defaults to 24 when omitted."),
        maxResults: z
          .number()
          .int()
          .min(1)
          .max(100)
          .optional()
          .describe("Optional max number of assignments to return. Defaults to 50.")
      },
      annotations: { readOnlyHint: true },
      _meta: widgetDescriptorMeta(WIDGET_META.list),
    },
    async (input) => {
      const data = listNewAssignments(input).slice(0, input.maxResults ?? 50);
      return {
        structuredContent: {
          view: "list",
          assignments: data.map((assignment) => ({
            id: assignment.id,
            site: assignment.site,
            category: assignment.category,
            priority: assignment.priority,
            slaDue: assignment.slaDue,
            estimatedStartDateTime: assignment.estimatedStartDateTime,
            estimatedEndDateTime: assignment.estimatedEndDateTime,
            description: assignment.description,
            requiredSkills: assignment.requiredSkills,
            customerName: assignment.customerName,
            customerPhone: assignment.customerPhone,
            customerProfilePicUrl: assignment.customerProfilePicUrl,
            assetId: assignment.assetId,
            estimatedDurationMinutes: assignment.estimatedDurationMinutes,
            estimatedTechnicianArrivalDateTime: assignment.estimatedTechnicianArrivalDateTime,
            assignedTechnicianId: assignment.assignedTechnicianId,
            siteImageUrl: assignment.siteImageUrl,
            tags: assignment.tags,
            status: assignment.status,
            address: assignment.location.address,
            lat: assignment.location.lat,
            lng: assignment.location.lng
          }))
        },
        content: [{
          type: "text",
          text: "Suggested next step: list available technicians and create a dispatch plan."
        }],
        _meta: widgetInvocationMeta(WIDGET_META.list),
      };
    }
  );

  server.registerTool(
    "show_assignments_on_map",
    {
      title: "Show Assignments on Map",
      description:
        "Render assignment pins on a Mapbox map view. Use after list_new_assignments and pass selected assignmentIds (or all returned IDs) to visualize location and inspect details. Returns map-ready pin payload and list metadata for map/list toggle. This is a read-only visualization tool and should not mutate dispatch state.",
      inputSchema: {
        assignmentIds: z
          .array(z.string())
          .optional()
          .describe(
            "Optional list of assignment IDs to display. If omitted, the tool maps current default new assignments."
          )
      },
      annotations: { readOnlyHint: true },
      _meta: widgetDescriptorMeta(WIDGET_META.map),
    },
    async (input) => {
      const selected = input.assignmentIds?.length
        ? getAssignmentsByIds(input.assignmentIds)
        : listNewAssignments({ maxHoursOld: 24 });

      return {
        structuredContent: {
          view: "map",
          assignments: selected.map((assignment) => ({
            id: assignment.id,
            site: assignment.site,
            category: assignment.category,
            priority: assignment.priority,
            slaDue: assignment.slaDue,
            estimatedStartDateTime: assignment.estimatedStartDateTime,
            estimatedEndDateTime: assignment.estimatedEndDateTime,
            description: assignment.description,
            requiredSkills: assignment.requiredSkills,
            customerName: assignment.customerName,
            customerPhone: assignment.customerPhone,
            customerProfilePicUrl: assignment.customerProfilePicUrl,
            estimatedTechnicianArrivalDateTime: assignment.estimatedTechnicianArrivalDateTime,
            assignedTechnicianId: assignment.assignedTechnicianId,
            siteImageUrl: assignment.siteImageUrl,
            tags: assignment.tags,
            status: assignment.status,
            address: assignment.location.address,
            lat: assignment.location.lat,
            lng: assignment.location.lng
          }))
        },
        content: [{
          type: "text",
          text: "Suggested next step: list available technicians and create a dispatch plan."
        }],
        _meta: widgetInvocationMeta(WIDGET_META.map),
      };
    }
  );

  server.registerTool(
    "get_available_technicians",
    {
      title: "Get Available Technicians",
      description:
        "Return currently available technicians and their dispatch metadata for planning. Input can include optional region filter. Output includes technician IDs, skills, current location, availability, rating, shift, and profile metadata so Copilot can choose assignment-technician pairings.",
      inputSchema: {
        region: z
          .string()
          .optional()
          .describe("Optional region filter to return technicians from a specific region.")
      },
      annotations: { readOnlyHint: true },
      _meta: {
        "openai/toolInvocation/invoking": "Checking technician availability…",
        "openai/toolInvocation/invoked": "Technician list ready."
      }
    },
    async (input) => {
      const available = getAvailableTechnicians(input);
      return {
        structuredContent: {
          technicians: available.map((tech) => ({
            id: tech.id,
            name: tech.name,
            profilePicUrl: tech.profilePicUrl,
            phone: tech.phone,
            rating: tech.rating,
            yearsExperience: tech.yearsExperience,
            shift: tech.shift,
            vehicleType: tech.vehicleType,
            skills: tech.skills,
            certifications: tech.certifications,
            languages: tech.languages,
            available: tech.available,
            lat: tech.location.lat,
            lng: tech.location.lng,
            address: tech.location.address
          }))
        },
        content: [{
          type: "text",
          text: "Suggested next step: create a dispatch plan for assignments."
        }],
        _meta: {
          "openai/toolInvocation/invoking": "Checking technician availability\u2026",
          "openai/toolInvocation/invoked": "Technician list ready.",
        }
      };
    }
  );

  server.registerTool(
    "create_dispatch_plan",
    {
      title: "Create Dispatch Plan",
      description:
        "Assigns technicians to assignments and renders a review-ready dispatch plan. Before creating the plan, perform a prerequisite data check: if assignments are not available, call list_new_assignments; if available technicians are not available, call get_available_technicians. This check must be performed before optimization and plan generation. Think hard before generating the plan. Then optimize technician-to-assignment pairing carefully, considering skills, ETA, distance, and workload continuity, and return mapping with reasons. Each planItem requires assignmentId, technicianId, and reason; etaMinutes is recommended, skillMatch and distanceKm are optional. The MCP server does not compute or alter assignment-technician mapping; it renders the mapping provided in planItems.",
      inputSchema: {
        planItems: z.array(
          z.object({
            assignmentId: z
              .string()
              .describe("Required assignment ID selected for dispatch."),
            technicianId: z
              .string()
              .describe("Required technician ID selected by the external matching/planning logic."),
            reason: z
              .string()
              .optional()
              .describe("Clear explanation of why this technician is assigned to this assignment."),
            etaMinutes: z
              .number()
              .int()
              .positive()
              .describe("Recommended ETA in minutes for technician arrival."),
            skillMatch: z
              .enum(["Full", "Partial"])
              .optional()
              .describe("Optional skill match label from external matching result."),
            distanceKm: z
              .number()
              .nonnegative()
              .optional()
              .describe("Optional travel distance in kilometers from technician to assignment.")
          })
        ).describe(
          "Required final assignment-technician pairings produced by external planning logic. One item per assignment to plan."
        ),
        technicianIds: z
          .array(z.string())
          .optional()
          .describe(
            "Optional technician IDs from get_available_technicians output, used to populate the override dropdown in the UI."
          ),
        maxTravelKm: z
          .number()
          .min(1)
          .max(500)
          .optional()
          .describe("Optional display constraint context for plan review. Defaults to 60."),
        allowPartialSkillMatch: z
          .boolean()
          .optional()
          .describe("Optional display constraint context. Defaults to true."),
        travelBufferMinutes: z
          .number()
          .int()
          .min(0)
          .max(180)
          .optional()
          .describe(
            "Optional scheduling buffer between consecutive assignments for the same technician. Defaults to 30."
          )
      },
      annotations: { readOnlyHint: true },
      _meta: widgetDescriptorMeta(WIDGET_META.plan),
    },
    async (input) => {
      // Resolve assignment and technician details from IDs in planItems (for display enrichment only)
      const assignmentIds = [...new Set(input.planItems.map((item) => item.assignmentId))];
      const planAssignments = getAssignmentsByIds(assignmentIds);

      const planTechIds = [...new Set(input.planItems.map((item) => item.technicianId))];
      const planTechs = getTechniciansByIds(planTechIds);

      // Build technician dropdown options from explicit input, or fall back to plan technicians
      const technicianOptions = input.technicianIds?.length
        ? getTechniciansByIds(input.technicianIds)
        : planTechs;

      const techniciansForMap = Array.from(
        new Map([...planTechs, ...technicianOptions].map((tech) => [tech.id, tech])).values()
      );

      const planItems = input.planItems.map((item) => {
        const technician = planTechs.find((tech) => tech.id === item.technicianId);
        return {
          ...item,
          technicianName: technician?.name ?? item.technicianId
        };
      });

      return {
        structuredContent: {
          view: "plan",
          constraints: {
            maxTravelKm: input.maxTravelKm ?? 60,
            allowPartialSkillMatch: input.allowPartialSkillMatch ?? true,
            travelBufferMinutes: input.travelBufferMinutes ?? 30
          },
          planItems,
          assignments: planAssignments,
          technicians: techniciansForMap,
          technicianOptions: technicianOptions.map((tech) => ({ id: tech.id, name: tech.name }))
        },
        content: [{ type: "text", text: `Prepared a plan for ${planItems.length} assignments.` }],
        _meta: widgetInvocationMeta(WIDGET_META.plan),
      };
    }
  );

  server.registerTool(
    "commit_assignments",
    {
      title: "Commit Assignments",
      description:
        "Commit a reviewed dispatch plan after explicit user confirmation (for example, user says Yes, assign them). Input must include final plan rows (assignmentId + technicianId + etaMinutes), including any manual overrides from the UI. Returns confirmation summary and per-assignment status for follow-up actions. Do not call this before create_dispatch_plan and user confirmation.",
      inputSchema: {
        assignments: z.array(
          z.object({
            assignmentId: z
              .string()
              .describe("Required assignment ID to commit."),
            technicianId: z
              .string()
              .describe("Required final technician ID to assign."),
            etaMinutes: z
              .number()
              .int()
              .positive()
              .describe("Required final ETA in minutes for commit summary.")
          })
        ).describe(
          "Required final assignment rows to commit. Use reviewed plan rows from create_dispatch_plan output."
        )
      },
      annotations: {
        readOnlyHint: false,
        openWorldHint: false,
        destructiveHint: false
      },
      _meta: {
        "openai/toolInvocation/invoking": "Committing assignments\u2026",
        "openai/toolInvocation/invoked": "Assignments committed.",
        "openai/visibility": "private"
      },
    },
    async (input) => {
      // TODO(template): Replace this demo-only response with writes to your system of record (dispatch backend/ERP/CRM).
      // Demo behavior: do not mutate server-side mock data; return final assignment state in tool output for widget rendering.
      const rows = input.assignments.map((item) => {
        const assignment = assignments.find((a) => a.id === item.assignmentId);
        const technician = getTechniciansByIds([item.technicianId])[0];

        const estimatedTechnicianArrivalDateTime = new Date(
          Date.now() + item.etaMinutes * 60 * 1000
        ).toISOString();

        return {
          assignmentId: item.assignmentId,
          site: assignment?.site ?? item.assignmentId,
          technicianId: item.technicianId,
          technicianName: technician?.name ?? item.technicianId,
          etaMinutes: item.etaMinutes,
          estimatedTechnicianArrivalDateTime,
          assignedTechnicianId: item.technicianId,
          status: "Dispatched"
        };
      });

      return {
        structuredContent: {
          summary: `${rows.length} assignments confirmed`,
          count: rows.length,
          rows
        },
        content: [{ type: "text", text: `${rows.length} assignments have been confirmed.` }],
        _meta: {
          "openai/toolInvocation/invoking": "Committing assignments\u2026",
          "openai/toolInvocation/invoked": "Assignments committed.",
        },
      };
    }
  );

  // Register resource templates (required for host widget discovery, per reference pattern)
  server.server.setRequestHandler(
    ListResourceTemplatesRequestSchema,
    async () => ({
      resourceTemplates: Object.values(WIDGET_META).map((meta) => ({
        uriTemplate: meta.templateUri,
        name: meta.templateUri,
        description: "Widget markup",
        mimeType: WIDGET_MIME,
        _meta: widgetDescriptorMeta(meta),
      })),
    })
  );

  return server;
}
