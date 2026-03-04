import React, { useState, useMemo } from "react";
import { useMcpToolData, useMcpApp, useMcpTheme } from "../hooks/useMcpApp";

interface AssignmentItem {
  id: string;
  site: string;
  category: string;
  priority: "High" | "Medium" | "Low";
  slaDue: string;
  description: string;
  address: string;
}

interface ListData {
  assignments: AssignmentItem[];
}

function fmtSla(iso: string) {
  const d = new Date(iso);
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const day = d.getDate();
  const mon = months[d.getMonth()];
  let h = d.getHours();
  const m = String(d.getMinutes()).padStart(2, "0");
  const ap = h >= 12 ? "PM" : "AM";
  h = h % 12 || 12;
  return `${day} ${mon} ${h}:${m} ${ap}`;
}

function pillClass(p: string) {
  return p === "High" ? "pill pill-high" : p === "Medium" ? "pill pill-medium" : "pill pill-low";
}

export function DispatchList() {
  const data = useMcpToolData<ListData>();
  const { app, hostContext } = useMcpApp();
  const theme = useMcpTheme();
  const [filter, setFilter] = useState("");

  const assignments = data?.assignments ?? [];
  const filtered = useMemo(
    () => (filter ? assignments.filter((a) => a.priority === filter) : assignments),
    [assignments, filter]
  );

  const isDark = theme === "dark";
  const bg = isDark ? "#1e1e1e" : "#fff";
  const text = isDark ? "#e0e0e0" : "#242424";
  const subtext = isDark ? "#a0a0a0" : "#707070";
  const border = isDark ? "#3a3a3a" : "#e0e0e0";
  const hoverBg = isDark ? "#2a2a2a" : "#f5f5f5";
  const headerBg = isDark ? "#2a2a2a" : "#fafafa";

  const isFullscreen = hostContext?.displayMode === "fullscreen";

  const subtitle =
    filtered.length === assignments.length
      ? `${assignments.length} assignment${assignments.length !== 1 ? "s" : ""}`
      : `${filtered.length} of ${assignments.length} assignments`;

  return (
    <div style={{ fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, Helvetica, Arial, sans-serif", fontSize: "15px", lineHeight: 1.45, color: text, background: "transparent" }}>
      <div style={{ border: `1px solid ${border}`, borderRadius: "12px", overflow: "hidden", background: bg, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px", borderBottom: `1px solid ${border}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <div style={{ width: 36, height: 36, display: "flex", alignItems: "center", justifyContent: "center", borderRadius: 8, background: "#0f6cbd", flexShrink: 0, color: "#fff" }}>
              <svg width="18" height="18" viewBox="0 0 20 20" fill="currentColor"><path d="M6 2a2 2 0 00-2 2v12a2 2 0 002 2h8a2 2 0 002-2V4a2 2 0 00-2-2H6zm1 3h6a.5.5 0 010 1H7a.5.5 0 010-1zm0 3h6a.5.5 0 010 1H7a.5.5 0 010-1zm0 3h4a.5.5 0 010 1H7a.5.5 0 010-1z"/></svg>
            </div>
            <div>
              <div style={{ fontSize: 17, fontWeight: 600, color: text }}>{assignments.length === 0 ? "New Assignments" : "New Assignments"}</div>
              <div style={{ fontSize: 13, color: subtext, marginTop: 1 }}>{assignments.length === 0 ? "Loading\u2026" : subtitle}</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <select
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
              style={{ fontFamily: "inherit", fontSize: 14, padding: "6px 10px", border: `1px solid ${isDark ? "#555" : "#d1d1d1"}`, borderRadius: 6, background: bg, color: text, outline: "none", cursor: "pointer", minHeight: 32 }}
            >
              <option value="">All priorities</option>
              <option value="High">High</option>
              <option value="Medium">Medium</option>
              <option value="Low">Low</option>
            </select>
            {app && (
              <button
                onClick={async () => {
                  try {
                    await app.sendMessage({
                      role: "user",
                      content: [{ type: "text", text: "Show assignments on the map" }],
                    });
                  } catch (err) {
                    console.error("sendMessage failed", err);
                  }
                }}
                style={{ fontFamily: "inherit", fontSize: 14, fontWeight: 600, padding: "6px 14px", borderRadius: 6, border: "none", cursor: "pointer", background: "#0f6cbd", color: "#fff", whiteSpace: "nowrap", display: "inline-flex", alignItems: "center", gap: 6, minHeight: 32 }}
              >
                <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polygon points="1 6 1 15 5.5 12 10.5 15 15 12 15 3 10.5 6 5.5 3"/></svg>
                View on map
              </button>
            )}
            {!isFullscreen && app && (
              <button
                onClick={() => app.requestDisplayMode({ mode: "fullscreen" })}
                style={{ fontFamily: "inherit", fontSize: 14, fontWeight: 600, padding: "6px 8px", borderRadius: 6, border: "none", cursor: "pointer", background: "transparent", color: subtext }}
                title="Expand to fullscreen"
              >
                <svg width="18" height="18" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="7 2 2 2 2 7"/><polyline points="13 2 18 2 18 7"/><polyline points="7 18 2 18 2 13"/><polyline points="13 18 18 18 18 13"/></svg>
              </button>
            )}
          </div>
        </div>

        {/* Column headers */}
        <div style={{ display: "grid", gridTemplateColumns: "1.6fr 82px 110px 2.5fr 1.8fr", gap: 8, alignItems: "center", padding: "10px 16px", background: headerBg, borderBottom: `1px solid ${border}`, fontSize: 13, fontWeight: 600, color: subtext }}>
          <div>Site</div>
          <div>Priority</div>
          <div>SLA Due</div>
          <div>Description</div>
          <div>Location</div>
        </div>

        {/* Rows */}
        <div style={{ display: "flex", flexDirection: "column" }}>
          {filtered.length === 0 ? (
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "48px 16px", gap: 8 }}>
              <svg width="40" height="40" viewBox="0 0 20 20" fill={subtext}><path d="M10 2a8 8 0 100 16 8 8 0 000-16zm-.75 4.5a.75.75 0 011.5 0v4a.75.75 0 01-1.5 0v-4zm.75 7.5a.75.75 0 110-1.5.75.75 0 010 1.5z"/></svg>
              <p style={{ color: subtext, fontSize: 15, margin: 0 }}>No assignments match the current filter</p>
            </div>
          ) : (
            filtered.map((item) => (
              <div
                key={item.id}
                style={{ display: "grid", gridTemplateColumns: "1.6fr 82px 110px 2.5fr 1.8fr", gap: 8, alignItems: "start", padding: "12px 16px", minHeight: 48, borderBottom: `1px solid ${isDark ? "#2a2a2a" : "#f0f0f0"}`, cursor: "default" }}
                onMouseEnter={(e) => (e.currentTarget.style.background = hoverBg)}
                onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
              >
                <div>
                  <div style={{ fontWeight: 600, color: text, fontSize: 14 }}>{item.site}</div>
                  <div style={{ fontSize: 13, color: subtext, marginTop: 2 }}>{item.id}</div>
                  <span style={{ display: "inline-block", fontSize: 10, fontWeight: 600, color: "#0f6cbd", background: isDark ? "#0d2240" : "#eff6fc", borderRadius: 3, padding: "1px 5px", letterSpacing: "0.03em", textTransform: "uppercase", whiteSpace: "nowrap", marginTop: 4 }}>{item.category}</span>
                </div>
                <div>
                  <span className={pillClass(item.priority)} style={{
                    display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 600, padding: "3px 10px", borderRadius: 999, whiteSpace: "nowrap",
                    background: item.priority === "High" ? (isDark ? "#3d1014" : "#fdf3f4") : item.priority === "Medium" ? (isDark ? "#2d2200" : "#fff9f0") : (isDark ? "#0b2d13" : "#f1faf1"),
                    color: item.priority === "High" ? (isDark ? "#f85149" : "#c4314b") : item.priority === "Medium" ? (isDark ? "#e3b341" : "#c37d04") : (isDark ? "#56d364" : "#0e7a0d"),
                  }}>{item.priority}</span>
                </div>
                <div style={{ fontSize: 14, color: text, fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap" }}>{fmtSla(item.slaDue)}</div>
                <div>
                  <div style={{ color: isDark ? "#b0b0b0" : "#424242", fontSize: 14, lineHeight: 1.4 }}>{item.description || "\u2014"}</div>
                </div>
                <div style={{ color: subtext, fontSize: 14, lineHeight: 1.4 }}>{item.address}</div>
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
}
