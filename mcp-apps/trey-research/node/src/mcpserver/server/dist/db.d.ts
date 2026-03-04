import { TableClient } from "@azure/data-tables";
export declare const consultantsTable: TableClient;
export declare const projectsTable: TableClient;
export declare const assignmentsTable: TableClient;
export declare function ensureTables(): Promise<void>;
export interface ConsultantEntity {
    partitionKey: string;
    rowKey: string;
    name: string;
    email: string;
    phone: string;
    consultantPhotoUrl: string;
    location: string;
    skills: string;
    certifications: string;
    roles: string;
}
export declare function getAllConsultants(): Promise<ConsultantEntity[]>;
export declare function getConsultantById(id: string): Promise<ConsultantEntity | null>;
export declare function updateConsultant(id: string, updates: Record<string, unknown>): Promise<ConsultantEntity | null>;
export declare function searchConsultantsBySkill(skill: string): Promise<ConsultantEntity[]>;
export interface ProjectEntity {
    partitionKey: string;
    rowKey: string;
    name: string;
    description: string;
    clientName: string;
    clientContact: string;
    clientEmail: string;
    location: string;
}
export declare function getAllProjects(): Promise<ProjectEntity[]>;
export declare function getProjectById(id: string): Promise<ProjectEntity | null>;
export interface AssignmentEntity {
    partitionKey: string;
    rowKey: string;
    projectId: string;
    consultantId: string;
    role: string;
    billable: boolean;
    rate: number;
    forecast: string;
    delivered: string;
}
export declare function getAllAssignments(): Promise<AssignmentEntity[]>;
export declare function getAssignmentsByProject(projectId: string): Promise<AssignmentEntity[]>;
export declare function getAssignmentsByConsultant(consultantId: string): Promise<AssignmentEntity[]>;
export declare function createAssignment(data: {
    projectId: string;
    consultantId: string;
    role: string;
    billable?: boolean;
    rate?: number;
    forecast?: Array<{
        month: number;
        year: number;
        hours: number;
    }>;
}): Promise<AssignmentEntity>;
export declare function deleteAssignment(projectId: string, consultantId: string): Promise<boolean>;
//# sourceMappingURL=db.d.ts.map