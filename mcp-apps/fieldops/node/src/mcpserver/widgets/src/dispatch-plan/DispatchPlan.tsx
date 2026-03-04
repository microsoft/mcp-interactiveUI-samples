import React, { useEffect, useRef, useState, useCallback } from "react";
import { useMcpToolData, useMcpApp } from "../hooks/useMcpApp";

/* ── Types ──────────────────────────────────────────────────────────── */
interface PlanItem {
  assignmentId: string;
  technicianId: string;
  technicianName?: string;
  reason?: string;
  etaMinutes: number;
  skillMatch?: "Full" | "Partial";
  distanceKm?: number;
}

interface PlanAssignment {
  id: string;
  site: string;
  priority: "High" | "Medium" | "Low";
  location: { lat: number; lng: number; address: string };
  [k: string]: unknown;
}

interface PlanTechnician {
  id: string;
  name: string;
  profilePicUrl?: string;
  phone?: string;
  skills?: string[];
  rating?: number;
  shift?: string;
  vehicleType?: string;
  location: { lat: number; lng: number; address?: string };
}

interface TechDropdownOption {
  id: string;
  name: string;
}

interface PlanData {
  planItems: PlanItem[];
  assignments: PlanAssignment[];
  technicians: PlanTechnician[];
  technicianOptions?: TechDropdownOption[];
  constraints?: { maxTravelKm?: number; allowPartialSkillMatch?: boolean; travelBufferMinutes?: number };
}

declare const mapboxgl: any;

/* ── Helpers ────────────────────────────────────────────────────────── */
function badgeClass(p: string): React.CSSProperties {
  if (p === "High") return { background: "#fde7e9", color: "#d13438" };
  if (p === "Medium") return { background: "#fff4ce", color: "#8a6100" };
  return { background: "#dff6dd", color: "#107c10" };
}

const ROUTE_PALETTE = ["#0f6cbd", "#107c10", "#6b4eff", "#0078d4", "#2d7d9a", "#8e562e", "#5c2d91", "#018574"];

/* ── Component ──────────────────────────────────────────────────────── */
export function DispatchPlan() {
  const data = useMcpToolData<PlanData>();
  const { app, hostContext } = useMcpApp();

  const mapContainerRef = useRef<HTMLDivElement>(null);
  const mapRef = useRef<any>(null);
  const markersRef = useRef<any[]>([]);
  const [mapReady, setMapReady] = useState(false);
  const [scriptLoaded, setScriptLoaded] = useState(false);

  const [planItems, setPlanItems] = useState<PlanItem[]>([]);
  const [activeAssignmentId, setActiveAssignmentId] = useState("");
  const [activeTechnicianId, setActiveTechnicianId] = useState("");
  const [statusMsg, setStatusMsg] = useState<{ type: "success" | "error"; text: string } | null>(null);
  const [confirming, setConfirming] = useState(false);
  const [boundsSignature, setBoundsSignature] = useState("");

  const assignments = data?.assignments ?? [];
  const technicians = data?.technicians ?? [];
  const technicianOptions = data?.technicianOptions ?? technicians.map((t) => ({ id: t.id, name: t.name }));
  const isFullscreen = hostContext?.displayMode === "fullscreen";

  // Sync plan items from tool data
  useEffect(() => {
    if (data?.planItems) {
      setPlanItems(data.planItems);
      if (data.planItems.length > 0 && !activeAssignmentId) {
        setActiveAssignmentId(data.planItems[0].assignmentId);
        setActiveTechnicianId(data.planItems[0].technicianId);
      }
    }
  }, [data]);

  // Load Mapbox script
  useEffect(() => {
    if ((window as any).mapboxgl) { setScriptLoaded(true); return; }
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://api.mapbox.com/mapbox-gl-js/v3.12.0/mapbox-gl.css";
    document.head.appendChild(link);
    const script = document.createElement("script");
    script.src = "https://api.mapbox.com/mapbox-gl-js/v3.12.0/mapbox-gl.js";
    script.onload = () => setScriptLoaded(true);
    document.head.appendChild(script);
  }, []);

  // Init map
  useEffect(() => {
    if (!scriptLoaded || !mapContainerRef.current || mapRef.current) return;
    const mb = (window as any).mapboxgl;
    mb.accessToken = "YOUR_MAPBOX_ACCESS_TOKEN";
    const map = new mb.Map({
      container: mapContainerRef.current,
      style: "mapbox://styles/mapbox/streets-v12",
      center: [78.3915, 17.4483],
      zoom: 10,
      attributionControl: false,
    });
    map.addControl(new mb.NavigationControl(), "top-left");
    map.on("load", () => {
      // Add route line source + layers
      map.addSource("dispatch-links", { type: "geojson", data: { type: "FeatureCollection", features: [] } });
      map.addLayer({
        id: "dispatch-links-casing",
        type: "line",
        source: "dispatch-links",
        paint: { "line-color": "rgba(255,255,255,0.95)", "line-width": ["case", ["get", "selected"], 11, 8], "line-opacity": 1 },
      });
      map.addLayer({
        id: "dispatch-links-layer",
        type: "line",
        source: "dispatch-links",
        paint: { "line-color": ["case", ["get", "selected"], "#d83b01", ["get", "routeColor"]], "line-width": ["case", ["get", "selected"], 7, 5], "line-opacity": 1 },
      });
      setMapReady(true);
    });
    mapRef.current = map;
    return () => { map.remove(); mapRef.current = null; };
  }, [scriptLoaded]);

  // Render markers + route lines  
  useEffect(() => {
    const map = mapRef.current;
    if (!map || !mapReady) return;
    const mb = (window as any).mapboxgl;

    // Clear previous
    markersRef.current.forEach((m) => m.remove());
    markersRef.current = [];

    if (assignments.length === 0 && technicians.length === 0) return;

    const techById = new Map(technicians.map((t) => [t.id, t]));
    const plannedTechIds = new Set(planItems.map((p) => p.technicianId));
    const bounds = new mb.LngLatBounds();

    // Assignment markers
    for (const a of assignments) {
      const el = document.createElement("div");
      el.className = "assignment-marker";
      const popup = new mb.Popup({ offset: 20 }).setHTML(`<strong>${a.site}</strong><br/>${a.id} · ${a.priority}`);
      const marker = new mb.Marker({ element: el, anchor: "center", pitchAlignment: "map", rotationAlignment: "map" }).setLngLat([a.location.lng, a.location.lat]).setPopup(popup).addTo(map);
      marker.getElement().addEventListener("mouseenter", () => popup.addTo(map));
      marker.getElement().addEventListener("mouseleave", () => popup.remove());
      marker.getElement().addEventListener("click", () => {
        setActiveAssignmentId(a.id);
        const mapped = planItems.find((p) => p.assignmentId === a.id);
        if (mapped) setActiveTechnicianId(mapped.technicianId);
        map.flyTo({ center: [a.location.lng, a.location.lat], zoom: 14 });
      });
      markersRef.current.push(marker);
      bounds.extend([a.location.lng, a.location.lat]);
    }

    // Technician markers (only for planned ones)
    for (const t of technicians) {
      if (!plannedTechIds.has(t.id)) continue;
      const el = document.createElement("div");
      el.style.cssText = "width:34px;height:34px;border-radius:50%;overflow:hidden;border:2px solid #fff;box-shadow:0 2px 10px rgba(0,0,0,0.32);background:#0f6cbd;cursor:pointer;";
      const img = document.createElement("img");
      img.src = t.profilePicUrl ?? "https://i.pravatar.cc/80?img=1";
      img.alt = t.name;
      img.style.cssText = "width:100%;height:100%;object-fit:cover;display:block;";
      el.appendChild(img);
      const skills = (t.skills ?? []).slice(0, 3).join(", ");
      const popupHtml = `<div style="min-width:200px"><strong>${t.name}</strong><br/>${t.id}${t.rating ? ` · ⭐ ${t.rating.toFixed(1)}` : ""}${skills ? `<br/>🛠️ ${skills}` : ""}</div>`;
      const popup = new mb.Popup({ offset: 20 }).setHTML(popupHtml);
      const marker = new mb.Marker({ element: el, anchor: "center", pitchAlignment: "map", rotationAlignment: "map" }).setLngLat([t.location.lng, t.location.lat]).setPopup(popup).addTo(map);
      marker.getElement().addEventListener("mouseenter", () => popup.addTo(map));
      marker.getElement().addEventListener("mouseleave", () => popup.remove());
      marker.getElement().addEventListener("click", () => {
        setActiveTechnicianId(t.id);
        map.flyTo({ center: [t.location.lng, t.location.lat], zoom: 14 });
      });
      markersRef.current.push(marker);
      bounds.extend([t.location.lng, t.location.lat]);
    }

    // Route lines
    const assignmentsById = new Map(assignments.map((a) => [a.id, a]));
    const lineFeatures = planItems.map((row, idx) => {
      const a = assignmentsById.get(row.assignmentId);
      const t = techById.get(row.technicianId);
      if (!a || !t) return null;
      return {
        type: "Feature" as const,
        properties: {
          routeColor: ROUTE_PALETTE[idx % ROUTE_PALETTE.length],
          selected: row.assignmentId === activeAssignmentId && row.technicianId === activeTechnicianId,
        },
        geometry: {
          type: "LineString" as const,
          coordinates: [
            [a.location.lng, a.location.lat],
            [t.location.lng, t.location.lat],
          ],
        },
      };
    }).filter(Boolean);

    const src = map.getSource("dispatch-links");
    if (src) src.setData({ type: "FeatureCollection", features: lineFeatures });

    // Fit bounds once per unique signature
    const sig = assignments.map((a) => a.id).sort().join("|") + "::" + Array.from(plannedTechIds).sort().join("|");
    if (!bounds.isEmpty() && sig !== boundsSignature) {
      map.fitBounds(bounds, { padding: 40, animate: true });
      setBoundsSignature(sig);
    }
  }, [mapReady, assignments, technicians, planItems, activeAssignmentId, activeTechnicianId]);

  // Resize on fullscreen
  useEffect(() => {
    if (mapRef.current) setTimeout(() => mapRef.current.resize(), 100);
  }, [isFullscreen]);

  // Handle technician override
  const handleTechOverride = useCallback((assignmentId: string, newTechId: string) => {
    setPlanItems((prev) =>
      prev.map((p) => {
        if (p.assignmentId !== assignmentId) return p;
        const tech = technicianOptions.find((t) => t.id === newTechId);
        return { ...p, technicianId: newTechId, technicianName: tech?.name ?? p.technicianName };
      })
    );
    setActiveAssignmentId(assignmentId);
    setActiveTechnicianId(newTechId);
  }, [technicianOptions]);

  // Confirm handler
  const handleConfirm = useCallback(async () => {
    if (!app) { setStatusMsg({ type: "error", text: "App not connected." }); return; }
    setStatusMsg(null);
    setConfirming(true);
    try {
      await app.callServerTool({
        name: "commit-assignments",
        arguments: {
          assignments: planItems.map((p) => ({
            assignmentId: p.assignmentId,
            technicianId: p.technicianId,
            etaMinutes: p.etaMinutes,
          })),
        },
      });
      setStatusMsg({ type: "success", text: `Successfully committed ${planItems.length} assignment${planItems.length === 1 ? "" : "s"}.` });
    } catch (err: any) {
      setStatusMsg({ type: "error", text: err?.message ?? "Failed to commit assignments." });
    } finally {
      setConfirming(false);
    }
  }, [app, planItems]);

  const focusPair = useCallback((assignmentId: string, technicianId: string) => {
    setActiveAssignmentId(assignmentId);
    setActiveTechnicianId(technicianId);
    const a = assignments.find((x) => x.id === assignmentId);
    if (a && mapRef.current) mapRef.current.flyTo({ center: [a.location.lng, a.location.lat], zoom: 13.5 });
  }, [assignments]);

  return (
    <>
      <style>{`
        .assignment-marker {
          width: 22px; height: 22px; border-radius: 50% 50% 50% 0;
          transform: rotate(-45deg); background: #d13438; border: 2px solid #fff;
          box-shadow: 0 2px 8px rgba(0,0,0,0.28); cursor: pointer;
        }
        .assignment-marker::after {
          content: ""; position: absolute; width: 8px; height: 8px;
          background: #fff; border-radius: 50%; top: 6px; left: 6px;
        }
        .mapboxgl-popup { z-index: 8; }
      `}</style>
      <div style={{
        fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif",
        fontSize: 14, lineHeight: 1.45, color: "#1f1f1f",
        height: isFullscreen ? "100vh" : "auto", overflow: isFullscreen ? "hidden" : undefined,
        display: "flex", flexDirection: "column",
      }}>
      <div style={{ display: "block", position: "relative", borderRadius: isFullscreen ? 0 : 12, overflow: "hidden", height: isFullscreen ? "100%" : "auto" }}>
        <div
          ref={mapContainerRef}
          style={{ width: "100%", height: isFullscreen ? "100%" : undefined, aspectRatio: isFullscreen ? undefined : "4 / 3", borderRadius: isFullscreen ? 0 : 12, overflow: "hidden" }}
        />

        {/* Fullscreen button */}
        {!isFullscreen && app && (
          <button
            onClick={() => app.requestDisplayMode({ mode: "fullscreen" })}
            style={{ position: "absolute", top: 10, right: 10, zIndex: 7, width: 32, height: 32, padding: 0, borderRadius: 999, display: "inline-flex", alignItems: "center", justifyContent: "center", background: "rgba(255,255,255,0.92)", border: "1px solid rgba(0,0,0,0.12)", cursor: "pointer", fontSize: 16 }}
            title="Expand map"
          >⛶</button>
        )}

        {/* Plan panel */}
        <div style={{
          position: "absolute", top: isFullscreen ? 10 : 54, right: 10, bottom: 10,
          width: "min(40%, 420px)", border: "1px solid rgba(0,0,0,0.14)", borderRadius: 12,
          padding: 10, overflow: "hidden",
          background: "rgba(255,255,255,0.92)", backdropFilter: "blur(6px)", WebkitBackdropFilter: "blur(6px)",
          zIndex: 5, display: "flex", flexDirection: "column",
        }}>
          {/* Header */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 8 }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: "rgba(0,0,0,0.62)" }}>Assignment → Technician Mapping</div>
            <button
              onClick={handleConfirm}
              disabled={confirming || planItems.length === 0}
              style={{ fontFamily: "inherit", fontSize: 14, fontWeight: 500, padding: "6px 14px", borderRadius: 8, border: "none", cursor: confirming ? "default" : "pointer", background: "#1f1f1f", color: "#fff", whiteSpace: "nowrap", opacity: confirming ? 0.6 : 1 }}
            >
              {confirming ? "Confirming\u2026" : "Confirm Assignments"}
            </button>
          </div>

          {/* Status message */}
          {statusMsg && (
            <div style={{
              margin: "0 0 8px", fontSize: 12, lineHeight: 1.3, padding: "6px 8px", borderRadius: 8,
              background: statusMsg.type === "success" ? "#dff6dd" : "#fde7e9",
              color: statusMsg.type === "success" ? "#107c10" : "#a4262c",
            }}>
              {statusMsg.text}
            </div>
          )}

          {/* Plan rows */}
          <div style={{ flex: 1, minHeight: 0, overflowY: "auto", overscrollBehavior: "contain" }}>
            {planItems.map((item) => {
              const a = assignments.find((x) => x.id === item.assignmentId);
              const tech = technicians.find((x) => x.id === item.technicianId);
              const isActive = item.assignmentId === activeAssignmentId && item.technicianId === activeTechnicianId;
              const reasons = item.reason || `${item.skillMatch === "Full" ? "Full" : "Partial"} skill match · ${item.etaMinutes} min ETA${item.distanceKm != null ? ` · ${item.distanceKm} km` : ""}`;

              return (
                <div
                  key={item.assignmentId}
                  onClick={() => focusPair(item.assignmentId, item.technicianId)}
                  style={{
                    padding: "10px 8px", borderRadius: 10, cursor: "pointer",
                    background: isActive ? "rgba(0,0,0,0.04)" : "transparent",
                    borderTop: "1px solid rgba(0,0,0,0.05)",
                  }}
                >
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center" }}>
                    <strong>{a?.site ?? item.assignmentId}</strong>
                    <span style={{ fontSize: 12, borderRadius: 999, padding: "2px 8px", fontWeight: 600, ...badgeClass(a?.priority ?? "Low") }}>{a?.priority ?? "Low"}</span>
                  </div>
                  <div style={{ color: "rgba(0,0,0,0.4)", fontSize: 13 }}>
                    Assigned to <strong>{tech?.name ?? item.technicianName ?? item.technicianId}</strong>
                  </div>
                  <div style={{ color: "rgba(0,0,0,0.4)", fontSize: 13 }}>Why: {reasons}</div>
                  <div style={{ color: "rgba(0,0,0,0.4)", fontSize: 13, marginTop: 2 }}>Change technician</div>
                  <select
                    value={item.technicianId}
                    onClick={(e) => e.stopPropagation()}
                    onChange={(e) => handleTechOverride(item.assignmentId, e.target.value)}
                    style={{ fontFamily: "inherit", fontSize: 13, padding: "4px 8px", border: "1px solid rgba(0,0,0,0.15)", borderRadius: 6, background: "#fff", color: "#1f1f1f", width: "100%", marginTop: 4, cursor: "pointer" }}
                  >
                    {technicianOptions.map((t) => (
                      <option key={t.id} value={t.id}>{t.name}</option>
                    ))}
                  </select>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
    </>
  );
}
