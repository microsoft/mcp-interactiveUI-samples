import { TableClient } from "@azure/data-tables";
// Azurite default connection string
const CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING ??
    "DefaultEndpointsProtocol=http;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;TableEndpoint=http://127.0.0.1:10002/devstoreaccount1;";
const tableOptions = { allowInsecureConnection: true };
export const consultantsTable = TableClient.fromConnectionString(CONNECTION_STRING, "Consultants", tableOptions);
export const projectsTable = TableClient.fromConnectionString(CONNECTION_STRING, "Projects", tableOptions);
export const assignmentsTable = TableClient.fromConnectionString(CONNECTION_STRING, "Assignments", tableOptions);
export async function ensureTables() {
    try {
        await consultantsTable.createTable();
    }
    catch { /* table may already exist */ }
    try {
        await projectsTable.createTable();
    }
    catch { /* table may already exist */ }
    try {
        await assignmentsTable.createTable();
    }
    catch { /* table may already exist */ }
}
export async function getAllConsultants() {
    const results = [];
    for await (const entity of consultantsTable.listEntities()) {
        results.push(entity);
    }
    return results;
}
export async function getConsultantById(id) {
    try {
        return await consultantsTable.getEntity("consultant", id);
    }
    catch {
        return null;
    }
}
export async function updateConsultant(id, updates) {
    const existing = await getConsultantById(id);
    if (!existing)
        return null;
    const merged = { ...existing };
    for (const [key, value] of Object.entries(updates)) {
        if (key === "skills" || key === "certifications" || key === "roles") {
            merged[key] = JSON.stringify(value);
        }
        else if (key === "location") {
            merged[key] = JSON.stringify(value);
        }
        else {
            merged[key] = value;
        }
    }
    await consultantsTable.updateEntity({ partitionKey: "consultant", rowKey: id, ...merged }, "Replace");
    return getConsultantById(id);
}
export async function searchConsultantsBySkill(skill) {
    const all = await getAllConsultants();
    return all.filter((c) => {
        const skills = JSON.parse(c.skills || "[]");
        return skills.some((s) => s.toLowerCase().includes(skill.toLowerCase()));
    });
}
export async function getAllProjects() {
    const results = [];
    for await (const entity of projectsTable.listEntities()) {
        results.push(entity);
    }
    return results;
}
export async function getProjectById(id) {
    try {
        return await projectsTable.getEntity("project", id);
    }
    catch {
        return null;
    }
}
export async function getAllAssignments() {
    const results = [];
    for await (const entity of assignmentsTable.listEntities()) {
        results.push(entity);
    }
    return results;
}
export async function getAssignmentsByProject(projectId) {
    const all = await getAllAssignments();
    return all.filter((a) => a.projectId === projectId);
}
export async function getAssignmentsByConsultant(consultantId) {
    const all = await getAllAssignments();
    return all.filter((a) => a.consultantId === consultantId);
}
export async function createAssignment(data) {
    const rowKey = `${data.projectId},${data.consultantId}`;
    const entity = {
        partitionKey: "assignment",
        rowKey,
        projectId: data.projectId,
        consultantId: data.consultantId,
        role: data.role,
        billable: data.billable ?? true,
        rate: data.rate ?? 0,
        forecast: JSON.stringify(data.forecast ?? []),
        delivered: JSON.stringify([]),
    };
    await assignmentsTable.upsertEntity(entity, "Replace");
    return assignmentsTable.getEntity("assignment", rowKey);
}
export async function deleteAssignment(projectId, consultantId) {
    const rowKey = `${projectId},${consultantId}`;
    try {
        await assignmentsTable.deleteEntity("assignment", rowKey);
        return true;
    }
    catch {
        return false;
    }
}
//# sourceMappingURL=db.js.map