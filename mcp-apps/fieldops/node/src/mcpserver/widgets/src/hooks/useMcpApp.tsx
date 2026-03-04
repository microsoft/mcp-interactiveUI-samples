/**
 * MCP Apps SDK React integration — replaces the OpenAI Apps SDK
 * (useOpenAiGlobal / window.openai) with @modelcontextprotocol/ext-apps.
 *
 * Provides a React context that manages the App lifecycle, tool results,
 * host context (theme, display mode, safe area), and exposes convenience
 * hooks for child components.
 */
import React, { createContext, useContext, useState, useEffect } from "react";
import type { App, McpUiHostContext } from "@modelcontextprotocol/ext-apps";
import { useApp } from "@modelcontextprotocol/ext-apps/react";

// ── Context value shape ────────────────────────────────────────────────
interface McpAppContextValue {
  /** The connected App instance (null while connecting) */
  app: App | null;
  /** Structured content from the most recent tool result */
  toolData: unknown;
  /** Current host theme */
  theme: "light" | "dark";
  /** Full host context */
  hostContext: McpUiHostContext | undefined;
}

const McpAppContext = createContext<McpAppContextValue>({
  app: null,
  toolData: null,
  theme: "light",
  hostContext: undefined,
});

// ── Provider ───────────────────────────────────────────────────────────
export function McpAppProvider({
  name,
  children,
}: {
  name: string;
  children: React.ReactNode;
}) {
  const [toolData, setToolData] = useState<unknown>(null);
  const [hostContext, setHostContext] = useState<McpUiHostContext | undefined>();

  const { app, error } = useApp({
    appInfo: { name, version: "1.0.0" },
    capabilities: {},
    onAppCreated: (app: App) => {
      // Register handlers BEFORE connect fires events
      app.ontoolresult = async (result) => {
        setToolData(result.structuredContent ?? null);
      };

      app.onhostcontextchanged = (ctx) => {
        setHostContext((prev) => ({ ...prev, ...ctx }));
      };

      app.onteardown = async () => {
        return {};
      };

      app.onerror = console.error;
    },
  });

  // Grab initial host context once connected
  useEffect(() => {
    if (app) {
      setHostContext(app.getHostContext());
    }
  }, [app]);

  const theme: "light" | "dark" =
    hostContext?.theme === "dark" ? "dark" : "light";

  if (error) {
    return (
      <div style={{ padding: 16, color: "#c93c37" }}>
        <strong>MCP App Error:</strong> {error.message}
      </div>
    );
  }

  if (!app) {
    return (
      <div style={{ padding: 16, opacity: 0.6 }}>Connecting…</div>
    );
  }

  return (
    <McpAppContext.Provider value={{ app, toolData, theme, hostContext }}>
      {children}
    </McpAppContext.Provider>
  );
}

// ── Hooks ──────────────────────────────────────────────────────────────

/** Full context (app, toolData, theme, hostContext) */
export function useMcpApp() {
  return useContext(McpAppContext);
}

/** Returns the structured tool data cast to T */
export function useMcpToolData<T>(): T | null {
  const { toolData } = useContext(McpAppContext);
  return (toolData as T) ?? null;
}

/** Returns the current theme ("light" | "dark") */
export function useMcpTheme(): "light" | "dark" {
  const { theme } = useContext(McpAppContext);
  return theme;
}
