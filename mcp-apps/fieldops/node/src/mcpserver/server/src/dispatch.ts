import { assignments, technicians } from "./mockData.js";
import type { Assignment, Technician } from "./types.js";

export function listNewAssignments(filters: {
  priority?: Assignment["priority"];
  region?: string;
  team?: string;
  maxHoursOld?: number;
}): Assignment[] {
  const maxHoursOld = filters.maxHoursOld ?? 24;
  const newestAllowed = Date.now() - maxHoursOld * 60 * 60 * 1000;

  // TODO(template): Replace this in-memory filter with a backend query and pagination.
  const filteredByRecency = assignments.filter((assignment) => {
    const isNew = assignment.status === "New";
    const isRecent = new Date(assignment.createdAt).getTime() >= newestAllowed;
    const priorityMatch = !filters.priority || assignment.priority === filters.priority;
    const regionMatch = !filters.region || assignment.region === filters.region;
    const teamMatch = !filters.team || assignment.team === filters.team;

    return isNew && isRecent && priorityMatch && regionMatch && teamMatch;
  });

  if (filters.maxHoursOld !== undefined || filteredByRecency.length > 0) {
    return filteredByRecency;
  }

  return assignments.filter((assignment) => {
    const isNew = assignment.status === "New";
    const priorityMatch = !filters.priority || assignment.priority === filters.priority;
    const regionMatch = !filters.region || assignment.region === filters.region;
    const teamMatch = !filters.team || assignment.team === filters.team;

    return isNew && priorityMatch && regionMatch && teamMatch;
  });
}

export function getAssignmentsByIds(ids: string[]): Assignment[] {
  const set = new Set(ids);
  return assignments.filter((assignment) => set.has(assignment.id));
}

export function getAvailableTechnicians(filters: { region?: string }): Technician[] {
  // TODO(template): Replace this static availability with live workforce management data.
  return technicians.filter((tech) => {
    const availableMatch = tech.available;
    const regionMatch = !filters.region || tech.region === filters.region;

    return availableMatch && regionMatch;
  });
}

export function getTechniciansByIds(ids: string[]): Technician[] {
  const set = new Set(ids);
  return technicians.filter((tech) => set.has(tech.id));
}
