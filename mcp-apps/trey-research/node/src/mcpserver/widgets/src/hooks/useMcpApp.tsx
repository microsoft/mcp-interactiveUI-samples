/**
 * MCP Apps React context – replaces the old OpenAI Apps SDK hooks
 * (useOpenAiGlobal, window.openai.*) with the MCP Apps standard.
 *
 * Provides:
 *  - useMcpToolData<T>()  → structured content delivered by the MCP tool
 *  - useMcpTheme()        → current "light" | "dark" theme
 *  - useMcpApp()          → full context (app, toolData, theme, hostContext)
 */
import React, { createContext, useContext, useEffect, useRef, useState } from "react";
import { useApp, type McpUiHostContext } from "@modelcontextprotocol/ext-apps/react";
import type { App } from "@modelcontextprotocol/ext-apps";

interface McpAppContextValue {
  /** The connected App instance (null while connecting) */
  app: App | null;
  /** Whether the app has connected to the host */
  isConnected: boolean;
  /** Structured content from the latest tool result */
  toolData: unknown;
  /** Current theme: "light" or "dark" */
  theme: "light" | "dark";
  /** Host context (theme, displayMode, etc.) */
  hostContext: McpUiHostContext | undefined;
}

const McpAppContext = createContext<McpAppContextValue | null>(null);

export function McpAppProvider({ name, children }: { name: string; children: React.ReactNode }) {
  const [toolData, setToolData] = useState<unknown>(null);
  const [theme, setTheme] = useState<"light" | "dark">("light");
  const [hostContext, setHostContext] = useState<McpUiHostContext | undefined>(undefined);

  // Refs so the onAppCreated callback can update React state
  const setToolDataRef = useRef(setToolData);
  const setThemeRef = useRef(setTheme);
  const setHostContextRef = useRef(setHostContext);
  setToolDataRef.current = setToolData;
  setThemeRef.current = setTheme;
  setHostContextRef.current = setHostContext;

  const { app, isConnected, error } = useApp({
    appInfo: { name, version: "1.0.0" },
    capabilities: {},
    onAppCreated: (app) => {
      app.ontoolresult = (result) => {
        if (result?.structuredContent) {
          setToolDataRef.current(result.structuredContent);
        }
      };
      app.onhostcontextchanged = (ctx) => {
        setHostContextRef.current((prev) => ({ ...prev, ...ctx }));
        if (ctx?.theme === "dark" || ctx?.theme === "light") {
          setThemeRef.current(ctx.theme);
        }
      };
    },
  });

  // Set initial host context after connection
  useEffect(() => {
    if (app) {
      const initial = app.getHostContext();
      if (initial) {
        setHostContext(initial);
        if (initial.theme === "dark" || initial.theme === "light") {
          setTheme(initial.theme);
        }
      }
    }
  }, [app]);

  return (
    <McpAppContext.Provider value={{ app, isConnected, toolData, theme, hostContext }}>
      {children}
    </McpAppContext.Provider>
  );
}

export function useMcpApp() {
  const ctx = useContext(McpAppContext);
  if (!ctx) throw new Error("useMcpApp must be used within McpAppProvider");
  return ctx;
}

export function useMcpToolData<T = unknown>(): T | null {
  const { toolData } = useMcpApp();
  return toolData as T | null;
}

export function useMcpTheme(): "light" | "dark" {
  const { theme } = useMcpApp();
  return theme;
}
