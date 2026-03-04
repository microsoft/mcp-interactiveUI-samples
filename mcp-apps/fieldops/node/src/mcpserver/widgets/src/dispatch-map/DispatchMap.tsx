import React, { useEffect, useRef, useState, useCallback } from "react";
import { useMcpToolData, useMcpApp, useMcpTheme } from "../hooks/useMcpApp";

interface MapAssignment {
  id: string;
  site: string;
  category: string;
  priority: "High" | "Medium" | "Low";
  slaDue: string;
  description: string;
  address: string;
  lat: number;
  lng: number;
  requiredSkills: string[];
  siteImageUrl?: string;
}

interface MapData {
  assignments: MapAssignment[];
}

declare const mapboxgl: any;

function fmtDue(iso: string) {
  const d = new Date(iso);
  return d.toLocaleString(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
}

function pinClass(p: string) {
  return p === "High" ? "pin-high" : p === "Medium" ? "pin-medium" : "pin-low";
}

/* ── SVG arrow icons for carousel nav ───────────────────────────────── */
const ChevronLeft = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.8" strokeLinecap="round" strokeLinejoin="round" style={{ display: "block", pointerEvents: "none" }}>
    <polyline points="15 5 8 12 15 19" />
  </svg>
);
const ChevronRight = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.8" strokeLinecap="round" strokeLinejoin="round" style={{ display: "block", pointerEvents: "none" }}>
    <polyline points="9 5 16 12 9 19" />
  </svg>
);

export function DispatchMap() {
  const data = useMcpToolData<MapData>();
  const { app, hostContext } = useMcpApp();
  const theme = useMcpTheme();
  const mapContainerRef = useRef<HTMLDivElement>(null);
  const mapRef = useRef<any>(null);
  const markersRef = useRef<any[]>([]);
  const navControlRef = useRef<any>(null);
  const navPositionRef = useRef<string>("top-left");
  const carouselRef = useRef<HTMLDivElement>(null);
  const [activeId, setActiveId] = useState("");
  const [mapReady, setMapReady] = useState(false);
  const [scriptLoaded, setScriptLoaded] = useState(false);
  const [prevDisabled, setPrevDisabled] = useState(true);
  const [nextDisabled, setNextDisabled] = useState(true);
  const [showFadeLeft, setShowFadeLeft] = useState(false);
  const [showFadeRight, setShowFadeRight] = useState(false);

  const assignments = data?.assignments ?? [];
  const isFullscreen = hostContext?.displayMode === "fullscreen";

  /* ── Carousel navigation helpers ──────────────────────────────────── */
  const updateCarouselNav = useCallback(() => {
    const el = carouselRef.current;
    if (!el) return;
    const maxScroll = el.scrollWidth - el.clientWidth;
    setPrevDisabled(el.scrollLeft <= 2);
    setNextDisabled(el.scrollLeft >= maxScroll - 2);
    setShowFadeLeft(el.scrollLeft > 2);
    setShowFadeRight(el.scrollLeft < maxScroll - 2);
  }, []);

  const scrollCarousel = useCallback((direction: number) => {
    const el = carouselRef.current;
    if (!el) return;
    const card = el.querySelector<HTMLElement>("[data-carousel-card]");
    const cardWidth = card ? card.getBoundingClientRect().width + 12 : 120;
    el.scrollBy({ left: direction * cardWidth, behavior: "smooth" });
  }, []);

  // Isolate carousel scroll events from map
  useEffect(() => {
    const el = carouselRef.current;
    if (!el) return;
    const stop = (e: Event) => e.stopPropagation();
    const events = ["wheel", "touchstart", "touchmove", "pointerdown", "mousedown"] as const;
    events.forEach((evt) => el.addEventListener(evt, stop, { passive: true }));
    return () => { events.forEach((evt) => el.removeEventListener(evt, stop)); };
  }, []);

  // Update nav button state on scroll and resize
  useEffect(() => {
    const el = carouselRef.current;
    if (!el) return;
    el.addEventListener("scroll", updateCarouselNav, { passive: true });
    window.addEventListener("resize", updateCarouselNav, { passive: true });
    updateCarouselNav(); // initial
    return () => {
      el.removeEventListener("scroll", updateCarouselNav);
      window.removeEventListener("resize", updateCarouselNav);
    };
  }, [updateCarouselNav, assignments]);

  // Load Mapbox script
  useEffect(() => {
    if (typeof window !== "undefined" && (window as any).mapboxgl) {
      setScriptLoaded(true);
      return;
    }
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
    const nav = new mb.NavigationControl();
    map.addControl(nav, "top-left");
    navControlRef.current = nav;
    navPositionRef.current = "top-left";
    map.on("load", () => setMapReady(true));
    mapRef.current = map;
    return () => { map.remove(); mapRef.current = null; };
  }, [scriptLoaded]);

  // Render markers
  useEffect(() => {
    const map = mapRef.current;
    if (!map || !mapReady || assignments.length === 0) return;
    const mb = (window as any).mapboxgl;

    // clear old markers
    markersRef.current.forEach((m) => m.remove());
    markersRef.current = [];

    const bounds = new mb.LngLatBounds();
    for (const item of assignments) {
      const el = document.createElement("div");
      el.className = `assignment-marker ${pinClass(item.priority)}`;
      const popup = new mb.Popup({ offset: 22 }).setHTML(
        `<strong>${item.site}</strong><br/>${item.priority} · SLA ${new Date(item.slaDue).toLocaleTimeString()}<br/>${item.requiredSkills.join(", ")}`
      );
      const marker = new mb.Marker({ element: el }).setLngLat([item.lng, item.lat]).setPopup(popup).addTo(map);
      marker.getElement().addEventListener("mouseenter", () => popup.addTo(map));
      marker.getElement().addEventListener("mouseleave", () => popup.remove());
      marker.getElement().addEventListener("click", () => {
        setActiveId(item.id);
        map.flyTo({ center: [item.lng, item.lat], zoom: 14 });
      });
      markersRef.current.push(marker);
      bounds.extend([item.lng, item.lat]);
    }
    map.fitBounds(bounds, { padding: 40, animate: true });
    if (!activeId && assignments.length > 0) setActiveId(assignments[0].id);
  }, [mapReady, assignments]);

  // Resize map + reposition nav control on fullscreen toggle
  useEffect(() => {
    const map = mapRef.current;
    const nav = navControlRef.current;
    if (!map) return;
    setTimeout(() => map.resize(), 100);

    if (nav) {
      const desired = isFullscreen ? "bottom-right" : "top-left";
      if (desired !== navPositionRef.current) {
        map.removeControl(nav);
        map.addControl(nav, desired);
        navPositionRef.current = desired;
      }
    }
  }, [isFullscreen]);

  const focusAssignment = useCallback((id: string) => {
    const item = assignments.find((a) => a.id === id);
    if (!item || !mapRef.current) return;
    setActiveId(id);
    mapRef.current.flyTo({ center: [item.lng, item.lat], zoom: 14 });
  }, [assignments]);

  /* ── Nav button shared style ──────────────────────────────────────── */
  const navBtnStyle = (side: "left" | "right"): React.CSSProperties => ({
    position: "absolute",
    top: "50%",
    transform: "translateY(-50%)",
    [side]: 6,
    zIndex: 3,
    width: 36,
    height: 36,
    borderRadius: 999,
    border: "1px solid #d1d1d1",
    background: "rgba(255,255,255,0.95)",
    color: "#242424",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    boxShadow: "0 1px 4px rgba(0,0,0,0.12)",
    padding: 0,
  });

  return (
    <>
      <style>{`
        .assignment-marker {
          width: 22px; height: 22px; border-radius: 50% 50% 50% 0;
          transform: rotate(-45deg); border: 2px solid #fff;
          box-shadow: 0 2px 8px rgba(0,0,0,0.28); cursor: pointer;
        }
        .assignment-marker::after {
          content: ""; position: absolute; width: 8px; height: 8px;
          background: #fff; border-radius: 50%; top: 6px; left: 6px;
        }
        .pin-high { background: #d13438; }
        .pin-medium { background: #ff8c00; }
        .pin-low { background: #107c10; }
        .mapboxgl-popup { z-index: 8; }
        @keyframes shimmer {
          0% { background-position: 200% 0; }
          100% { background-position: -200% 0; }
        }
        .dispatch-carousel::-webkit-scrollbar { display: none; }
      `}</style>
      <div style={{
        fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, Helvetica, Arial, sans-serif",
        height: isFullscreen ? "100vh" : "auto",
        overflow: isFullscreen ? "hidden" : undefined,
        display: "flex",
        flexDirection: "column",
      }}>
        <div style={{ display: "block", position: "relative", borderRadius: isFullscreen ? 0 : 12, overflow: "hidden", height: isFullscreen ? "100%" : "auto" }}>
          <div
            ref={mapContainerRef}
            style={{
              width: "100%",
              height: isFullscreen ? "100%" : undefined,
              aspectRatio: isFullscreen ? undefined : "4 / 3",
              borderRadius: isFullscreen ? 0 : 12,
              overflow: "hidden",
            }}
          />

          {/* Fullscreen button */}
          {!isFullscreen && app && (
            <button
              onClick={() => app.requestDisplayMode({ mode: "fullscreen" })}
              style={{ position: "absolute", top: 10, right: 10, zIndex: 6, width: 32, height: 32, padding: 0, borderRadius: 999, display: "inline-flex", alignItems: "center", justifyContent: "center", background: "rgba(255,255,255,0.92)", border: "1px solid rgba(0,0,0,0.12)", cursor: "pointer", fontSize: 16 }}
              title="Expand map"
            >⛶</button>
          )}

          {/* Carousel wrap */}
          <div style={{
            position: "absolute",
            left: 0, right: 0,
            [isFullscreen ? "top" : "bottom"]: 0,
            padding: "8px 8px 10px",
            background: "rgba(255,255,255,0.45)",
            backdropFilter: "blur(5px)",
            WebkitBackdropFilter: "blur(5px)",
            zIndex: 5,
          }}>
            {/* Carousel shell (contains nav buttons, viewport, fades) */}
            <div style={{ position: "relative" }}>
              {/* Prev nav button */}
              <button
                onClick={() => scrollCarousel(-1)}
                disabled={prevDisabled}
                style={{ ...navBtnStyle("left"), opacity: prevDisabled ? 0.35 : 1, cursor: prevDisabled ? "default" : "pointer" }}
                aria-label="Previous assignments"
              >
                <ChevronLeft />
              </button>

              {/* Carousel viewport */}
              <div style={{ overflow: "hidden" }}>
                <div
                  ref={carouselRef}
                  className="dispatch-carousel"
                  style={{
                    display: "flex",
                    gap: 12,
                    overflowX: "auto",
                    padding: "2px 2px 6px",
                    scrollSnapType: "x proximity",
                    scrollbarWidth: "none" as const,
                    overscrollBehaviorX: "contain",
                    overscrollBehaviorY: "contain",
                    touchAction: "pan-x",
                  }}
                >
                  {assignments.length === 0
                    ? Array.from({ length: 7 }).map((_, i) => (
                        <div key={i} data-carousel-card style={{ minWidth: 100, maxWidth: 110, flexShrink: 0, cursor: "default" }}>
                          <div style={{ width: "100%", aspectRatio: "1/1", borderRadius: 10, background: "linear-gradient(90deg, rgba(255,255,255,0.15) 25%, rgba(255,255,255,0.35) 40%, rgba(255,255,255,0.15) 60%)", backgroundSize: "200% 100%", animation: "shimmer 1.2s ease-in-out infinite" }} />
                          <div style={{ height: 11, marginTop: 8, borderRadius: 8, background: "rgba(0,0,0,0.08)" }} />
                        </div>
                      ))
                    : assignments.map((item) => (
                        <div
                          key={item.id}
                          data-carousel-card
                          onClick={() => focusAssignment(item.id)}
                          style={{
                            scrollSnapAlign: "start",
                            minWidth: 100, maxWidth: 110,
                            cursor: "pointer",
                            display: "flex", flexDirection: "column",
                            userSelect: "none",
                            opacity: activeId === item.id ? 1 : 0.85,
                            flexShrink: 0,
                            transition: "transform 0.1s ease, opacity 0.1s ease",
                          }}
                        >
                          <img
                            src={item.siteImageUrl ?? ""}
                            alt={item.site}
                            style={{ width: "100%", aspectRatio: "1/1", objectFit: "cover", borderRadius: 10, border: "1px solid rgba(0,0,0,0.06)", boxShadow: "0 2px 6px rgba(0,0,0,0.06)", background: "#f5f5f5" }}
                          />
                          <div style={{ marginTop: 8 }}>
                            <div style={{ fontSize: 14, fontWeight: 600, color: "#242424", lineHeight: 1.3, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{item.site}</div>
                            <div style={{ marginTop: 4, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                              <span style={{ display: "inline-block", fontSize: 10, fontWeight: 600, color: "#0f6cbd", background: "#eff6fc", borderRadius: 3, padding: "1px 5px", textTransform: "uppercase", letterSpacing: "0.03em" }}>{item.category ?? "General"}</span>
                              <span style={{ fontSize: 13, color: "#616161", lineHeight: 1.3 }}>{fmtDue(item.slaDue)}</span>
                            </div>
                          </div>
                        </div>
                      ))}
                </div>
              </div>

              {/* Next nav button */}
              <button
                onClick={() => scrollCarousel(1)}
                disabled={nextDisabled}
                style={{ ...navBtnStyle("right"), opacity: nextDisabled ? 0.35 : 1, cursor: nextDisabled ? "default" : "pointer" }}
                aria-label="Next assignments"
              >
                <ChevronRight />
              </button>

              {/* Fade indicators */}
              {showFadeLeft && (
                <div aria-hidden="true" style={{ position: "absolute", top: 0, bottom: 0, left: 0, width: 10, pointerEvents: "none", zIndex: 2, borderLeft: "1px solid rgba(0,0,0,0.12)", background: "linear-gradient(to right, rgba(0,0,0,0.1), rgba(0,0,0,0))" }} />
              )}
              {showFadeRight && (
                <div aria-hidden="true" style={{ position: "absolute", top: 0, bottom: 0, right: 0, width: 10, pointerEvents: "none", zIndex: 2, borderRight: "1px solid rgba(0,0,0,0.12)", background: "linear-gradient(to left, rgba(0,0,0,0.1), rgba(0,0,0,0))" }} />
              )}
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
