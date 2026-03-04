/**
 * useToolData — Custom hook wrapping the MCP App lifecycle for data-driven tools.
 *
 * - Creates an App instance via useApp()
 * - Registers all handlers BEFORE connect() via onAppCreated
 * - Stores structuredContent from ontoolresult in typed state
 * - Tracks host context changes for theming
 *
 * @returns { app, data, hostContext, isConnected, error }
 */
import type { App, McpUiHostContext } from "@modelcontextprotocol/ext-apps";
import { useApp } from "@modelcontextprotocol/ext-apps/react";
import { useEffect, useState } from "react";

export interface ToolDataState<T> {
  app: App | null;
  data: T | null;
  hostContext: McpUiHostContext | undefined;
  isConnected: boolean;
  error: Error | null;
}

export function useToolData<T>(appName: string): ToolDataState<T> {
  const [data, setData] = useState<T | null>(null);
  const [hostContext, setHostContext] = useState<McpUiHostContext | undefined>();

  const { app, isConnected, error } = useApp({
    appInfo: { name: appName, version: "1.0.0" },
    capabilities: {},
    onAppCreated: (app: App) => {
      // All handlers registered BEFORE connect() — per MCP Apps SDK best practice
      app.ontoolresult = (result) => {
        if (result.structuredContent) {
          setData(result.structuredContent as T);
        } else {
          // Fallback: try parsing text content
          const text = result.content?.find((c: { type: string }) => c.type === "text") as
            | { type: string; text: string }
            | undefined;
          if (text?.text) {
            try {
              setData(JSON.parse(text.text) as T);
            } catch {
              // not JSON; ignore
            }
          }
        }
      };

      app.ontoolinput = (input) => {
        console.info(`[${appName}] Tool input:`, input);
      };

      app.ontoolcancelled = (params) => {
        console.info(`[${appName}] Tool cancelled:`, params.reason);
      };

      app.onerror = (err) => {
        console.error(`[${appName}] Error:`, err);
      };

      app.onhostcontextchanged = (ctx) => {
        setHostContext((prev) => ({ ...prev, ...ctx }));
      };

      app.onteardown = async () => {
        console.info(`[${appName}] Teardown`);
        return {};
      };
    },
  });

  useEffect(() => {
    if (app) {
      setHostContext(app.getHostContext());
    }
  }, [app]);

  return { app, data, hostContext, isConnected, error };
}
