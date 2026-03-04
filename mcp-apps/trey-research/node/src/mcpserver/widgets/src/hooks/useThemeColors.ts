/**
 * Returns a palette of semantic colours that adapts to the current MCP Apps
 * host theme ("light" | "dark") read from the MCP Apps context.
 *
 * Falls back to a `prefers-color-scheme` media-query for non-MCP hosts.
 */
import { useState, useEffect, useMemo } from "react";

type ThemeMode = "light" | "dark";

/**
 * Detect theme from OS preference (fallback when not inside an MCP host).
 * The primary theme source is useMcpTheme() from useMcpApp — this hook
 * is kept for use by useThemeColors which needs a standalone detection.
 */
function detectTheme(): ThemeMode {
  if (typeof window.matchMedia === "function") {
    if (window.matchMedia("(prefers-color-scheme: dark)").matches) return "dark";
  }
  return "light";
}

export function useThemeMode(): ThemeMode {
  const [mode, setMode] = useState<ThemeMode>(detectTheme);

  useEffect(() => {
    if (typeof window.matchMedia !== "function") return;
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    const handler = (e: MediaQueryListEvent) => setMode(e.matches ? "dark" : "light");
    mq.addEventListener("change", handler);
    return () => mq.removeEventListener("change", handler);
  }, []);

  return mode;
}

export interface ThemeColors {
  /* surfaces */
  surface: string;
  cardBg: string;
  /* text */
  textPrimary: string;
  textSecondary: string;
  textTertiary: string;
  /* brand */
  brand: string;
  brandLight: string;
  brandDark: string;
  /* semantic */
  green: string;
  greenBg: string;
  amber: string;
  amberBg: string;
  purple: string;
  purpleBg: string;
  /* borders / misc */
  divider: string;
  /* gradients */
  bannerGradient: string;
  bannerText: string;
}

const LIGHT: ThemeColors = {
  surface: "#f3f6f8",
  cardBg: "#ffffff",
  textPrimary: "#191919",
  textSecondary: "#666666",
  textTertiary: "#999999",
  brand: "#0a66c2",
  brandLight: "#e8f1fb",
  brandDark: "#004182",
  green: "#057642",
  greenBg: "#e6f4ea",
  amber: "#b24020",
  amberBg: "#fff3e0",
  purple: "#6d28d9",
  purpleBg: "#f3e8ff",
  divider: "#e8e8e8",
  bannerGradient: "linear-gradient(135deg, #0a66c2 0%, #004182 100%)",
  bannerText: "#ffffff",
};

const DARK: ThemeColors = {
  surface: "#1b1b1b",
  cardBg: "#292929",
  textPrimary: "#e8e8e8",
  textSecondary: "#a0a0a0",
  textTertiary: "#737373",
  brand: "#70b5f9",
  brandLight: "#1a3a5c",
  brandDark: "#a8d4ff",
  green: "#6ec89b",
  greenBg: "#1a3328",
  amber: "#f5a060",
  amberBg: "#3d2a1a",
  purple: "#b794f4",
  purpleBg: "#2d1f4e",
  divider: "#3a3a3a",
  bannerGradient: "linear-gradient(135deg, #1a3a5c 0%, #0d2240 100%)",
  bannerText: "#e8e8e8",
};

export function useThemeColors(): ThemeColors {
  const mode = useThemeMode();
  return useMemo(() => (mode === "dark" ? DARK : LIGHT), [mode]);
}
