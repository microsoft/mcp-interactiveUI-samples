export interface ToolDescriptorMetadata {
  name: string;
  title: string;
  description: string;
}

export const toolMetadata: ToolDescriptorMetadata[] = [
  {
    name: "list_new_assignments",
    title: "List New Assignments",
    description:
      "Use this first for dispatch intake. Returns unassigned, recently created work orders (default: status=New and created in last 24 hours), optionally filtered by priority/region/team. Output includes assignment IDs, site, SLA due, required skills, status, and coordinates for downstream tools. Call before map or planning when user asks for new work."
  },
  {
    name: "show_assignments_on_map",
    title: "Show Assignments on Map",
    description:
      "Render assignment pins on a Mapbox map view. Use after list_new_assignments and pass selected assignmentIds (or all returned IDs) to visualize location and inspect details. Returns map-ready pin payload and list metadata for map/list toggle. This is a read-only visualization tool and should not mutate dispatch state."
  },
  {
    name: "get_available_technicians",
    title: "Get Available Technicians",
    description:
      "Return currently available technicians and their dispatch metadata for planning. Input can include optional region filter. Output includes technician IDs, skills, current location, availability, rating, shift, and profile metadata so Copilot can choose assignment-technician pairings."
  },
  {
    name: "create_dispatch_plan",
    title: "Create Dispatch Plan",
    description:
      "Assigns technicians to assignments and renders a review-ready optimized dispatch plan. Before creating the plan, perform a prerequisite data check: if assignments are not available, call list_new_assignments; if available technicians are not available, call get_available_technicians. This check must be performed before optimization and plan generation. Think hard before generating the plan. Then optimize technician-to-assignment pairing carefully, considering skills, ETA, distance, and workload continuity. Return planItems mapping assignments to technicians with a clear reason for each mapping. Each planItem requires assignmentId, technicianId, and reason; etaMinutes is recommended, skillMatch and distanceKm are optional. The MCP server does not compute or alter assignment-technician mapping; it renders the mapping provided in planItems."
  },
  {
    name: "commit_assignments",
    title: "Commit Assignments",
    description:
      "Commit a reviewed dispatch plan after explicit user confirmation (for example, user says Yes, assign them). Input must include final plan rows (assignmentId + technicianId + etaMinutes), including any manual overrides from the UI. Returns confirmation summary and per-assignment status for follow-up actions. Do not call this before create_dispatch_plan and user confirmation."
  }
];
