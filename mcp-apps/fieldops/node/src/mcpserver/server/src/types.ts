export type Priority = "Low" | "Medium" | "High";

export interface GeoPoint {
  lat: number;
  lng: number;
  address: string;
}

export interface Assignment {
  id: string;
  site: string;
  category: string;
  priority: Priority;
  slaDue: string;
  estimatedStartDateTime: string;
  estimatedEndDateTime: string;
  description: string;
  requiredSkills: string[];
  customerName: string;
  customerPhone: string;
  customerProfilePicUrl: string;
  assetId: string;
  estimatedDurationMinutes: number;
  estimatedTechnicianArrivalDateTime: string | null;
  assignedTechnicianId: string | null;
  siteImageUrl: string;
  tags: string[];
  location: GeoPoint;
  status: "New" | "Dispatched" | "EnRoute";
  createdAt: string;
  region: string;
  team: string;
}

export interface Technician {
  id: string;
  name: string;
  profilePicUrl: string;
  phone: string;
  rating: number;
  yearsExperience: number;
  shift: "Morning" | "Evening" | "Night";
  vehicleType: "Bike" | "Van" | "Car";
  skills: string[];
  certifications: string[];
  languages: string[];
  location: GeoPoint;
  available: boolean;
  region: string;
}

export interface DispatchPlanItem {
  assignmentId: string;
  technicianId: string;
  technicianName: string;
  reason: string;
  etaMinutes: number;
  skillMatch: "Full" | "Partial";
  distanceKm: number;
}

export interface DispatchPlan {
  items: DispatchPlanItem[];
  constraints: {
    maxTravelKm: number;
    allowPartialSkillMatch: boolean;
  };
}
