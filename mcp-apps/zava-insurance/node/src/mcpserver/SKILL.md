---
name: mcpserver
description: "Migrates an MCP server with interactive widgets from the OpenAI Apps SDK (window.openai, text/html+skybridge) to the MCP Apps standard (@modelcontextprotocol/ext-apps), covering server-side and client-side changes."
---

# Skill: Migrate OpenAI Apps SDK → MCP Apps

Migrate an MCP server with interactive widgets from the **OpenAI Apps SDK** (`window.openai`, `text/html+skybridge`, flat `_meta["openai/..."]` keys) to the **MCP Apps** standard (`@modelcontextprotocol/ext-apps`).

## When to Use

Use this skill when:
- An MCP server uses `text/html+skybridge` MIME type for widget resources
- Widget code references `window.openai` globals (e.g. `window.openai.callTool`, `window.openai.toolOutput`, `window.openai.theme`)
- Server code uses flat `_meta["openai/outputTemplate"]` or `_meta["openai/widgetAccessible"]` keys
- The goal is to make the server compatible with MCP Apps hosts (Claude, ChatGPT, Microsoft 365 Copilot, etc.)

## References

- MCP Apps repo: https://github.com/modelcontextprotocol/ext-apps
- API docs: https://modelcontextprotocol.github.io/ext-apps/api/
- Migration guide: https://modelcontextprotocol.github.io/ext-apps/docs/migration/openai-apps
- Patterns: https://modelcontextprotocol.github.io/ext-apps/docs/patterns

## Packages

| Package | Where | Purpose |
|---------|-------|---------|
| `@modelcontextprotocol/ext-apps` | Server + Widgets | Core MCP Apps SDK |
| `@modelcontextprotocol/sdk` | Server | MCP protocol SDK (keep existing) |
| `zod` | Server | Schema definitions for `McpServer.tool()` |

## Migration Mapping

### MIME Type

| Before | After |
|--------|-------|
| `text/html+skybridge` | `text/html;profile=mcp-app` (use `RESOURCE_MIME_TYPE` constant) |

### Server: `_meta` Keys

| OpenAI flat key | MCP Apps nested key |
|-----------------|---------------------|
| `_meta["openai/outputTemplate"]` (URI string) | `_meta.ui.resourceUri` (URI string) |
| `_meta["openai/widgetAccessible"]: true` | `_meta.ui.visibility: ["widget"]` |

### Server: Class & Helpers

| Before | After |
|--------|-------|
| `import { Server } from "@modelcontextprotocol/sdk/server/index.js"` | `import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"` |
| `new Server({ name, version }, { capabilities })` | `new McpServer({ name, version })` |
| `server.setRequestHandler(ListToolsRequestSchema, ...)` | `server.tool(name, desc, schema, handler)` or `registerAppTool(...)` |
| `server.setRequestHandler(ReadResourceRequestSchema, ...)` | `registerAppResource(...)` |
| Manual tool/resource list handlers | Automatic via `McpServer` + helpers |

### Server: Tool Registration

**Widget tools** (tools that render UI) use `registerAppTool`:

```typescript
import { registerAppTool, registerAppResource, RESOURCE_MIME_TYPE } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";

const WIDGET_URI = "ui://myapp/widget.html";

registerAppResource(server, "Widget Name", WIDGET_URI, {
  mimeType: RESOURCE_MIME_TYPE,
  description: "Description of the widget",
}, async (): Promise<ReadResourceResult> => {
  const html = await fs.readFile(widgetPath, "utf-8");
  return { contents: [{ uri: WIDGET_URI, mimeType: RESOURCE_MIME_TYPE, text: html }] };
});

registerAppTool(server, "show-widget", {
  title: "Show Widget",
  description: "Displays the widget",
  inputSchema: {
    filter: z.string().optional().describe("Optional filter"),
  },
  annotations: { readOnlyHint: true },
  _meta: { ui: { resourceUri: WIDGET_URI } },
}, async ({ filter }): Promise<CallToolResult> => {
  const data = await fetchData(filter);
  return {
    content: [{ type: "text", text: `Loaded ${data.length} items.` }],
    structuredContent: { items: data },
  };
});
```

**Data-only tools** (no UI) use `server.tool()` directly:

```typescript
server.tool("update-item", "Updates an item.", {
  id: z.string().describe("Item ID"),
  status: z.string().describe("New status"),
}, async ({ id, status }) => {
  await db.update(id, { status });
  return { content: [{ type: "text" as const, text: `Updated ${id}.` }] };
});
```

### Client (Widget): Global API

| OpenAI (`window.openai`) | MCP Apps (`App` from `@modelcontextprotocol/ext-apps`) |
|---------------------------|--------------------------------------------------------|
| `window.openai.toolOutput` | `app.ontoolresult = (result) => result.structuredContent` |
| `window.openai.callTool(name, args)` | `app.callServerTool({ name, arguments: args })` |
| `window.openai.theme` (`"light"` / `"dark"`) | `app.getHostContext()?.theme` (`"light"` / `"dark"`) |
| `window.openai.displayMode` | `app.getHostContext()?.displayMode` |
| `window.openai.requestDisplayMode(mode)` | `app.requestDisplayMode(mode)` |
| `window.addEventListener("openai:set_globals", ...)` | `app.onhostcontextchanged = (ctx) => { ... }` |
| N/A | `app.onteardown = () => { ... }` |

### Client (Widget): React Hook

| Before | After |
|--------|-------|
| `useOpenAiGlobal("toolOutput")` | `useMcpToolData<T>()` (custom hook wrapping `useApp`) |
| `useOpenAiGlobal("theme")` | `useMcpTheme()` (custom hook returning `"light"` / `"dark"`) |

## Step-by-Step Migration Process

### 1. Update Dependencies

**Server `package.json`** — add:
```json
"@modelcontextprotocol/ext-apps": "^1.0.0",
"zod": "^3.25.0"
```

**Widgets `package.json`** — add:
```json
"@modelcontextprotocol/ext-apps": "^1.0.0"
```

### 2. Create MCP Apps React Context (Widgets)

Create a shared hook file (e.g. `hooks/useMcpApp.tsx`) that wraps the `App` class:

```tsx
import React, { createContext, useContext, useEffect, useState, useRef } from "react";
import { useApp } from "@modelcontextprotocol/ext-apps/react";

interface McpAppContextValue {
  app: ReturnType<typeof useApp>;
  toolData: unknown;
  theme: "light" | "dark";
  hostContext: { theme?: string; displayMode?: string } | null;
}

const McpAppContext = createContext<McpAppContextValue | null>(null);

export function McpAppProvider({ name, children }: { name: string; children: React.ReactNode }) {
  const app = useApp({ name });
  const [toolData, setToolData] = useState<unknown>(null);
  const [theme, setTheme] = useState<"light" | "dark">("light");
  const [hostContext, setHostContext] = useState<{ theme?: string; displayMode?: string } | null>(null);

  useEffect(() => {
    app.ontoolresult = (result: any) => {
      if (result?.structuredContent) setToolData(result.structuredContent);
    };
    app.onhostcontextchanged = (ctx: any) => {
      setHostContext(ctx);
      if (ctx?.theme === "dark" || ctx?.theme === "light") setTheme(ctx.theme);
    };
    const initial = app.getHostContext?.();
    if (initial) {
      setHostContext(initial);
      if (initial.theme === "dark" || initial.theme === "light") setTheme(initial.theme);
    }
  }, [app]);

  return (
    <McpAppContext.Provider value={{ app, toolData, theme, hostContext }}>
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
  const { toolData, theme } = useMcpApp();
  return theme;
}
```

### 3. Update Widget Entry Points (`main.tsx`)

Wrap the app in `<McpAppProvider>` instead of reading `window.openai`:

```tsx
import { McpAppProvider, useMcpTheme } from "../hooks/useMcpApp";

function ThemedApp() {
  const theme = useMcpTheme();
  return (
    <FluentProvider theme={theme === "dark" ? webDarkTheme : webLightTheme}>
      <MyWidget />
    </FluentProvider>
  );
}

createRoot(document.getElementById("root")!).render(
  <McpAppProvider name="My Widget">
    <ThemedApp />
  </McpAppProvider>
);
```

### 4. Update Widget Components

Replace all `window.openai` references:

```typescript
// BEFORE
const toolOutput = useOpenAiGlobal("toolOutput");
window.openai.callTool("update-item", { id: "1", status: "done" });
window.openai.requestDisplayMode(
  window.openai.displayMode === "expanded" ? "default" : "expanded"
);

// AFTER
const toolData = useMcpToolData<MyDataType>();
const { app, hostContext } = useMcpApp();
app.callServerTool({ name: "update-item", arguments: { id: "1", status: "done" } });
app.requestDisplayMode(
  hostContext?.displayMode === "expanded" ? "default" : "expanded"
);
```

### 5. Rewrite Server (`mcp-server.ts`)

1. Replace `Server` with `McpServer`
2. Replace manual `setRequestHandler` with `registerAppTool` / `registerAppResource` / `server.tool()`
3. Use `RESOURCE_MIME_TYPE` instead of `"text/html+skybridge"`
4. Use `zod` schemas for tool input definitions
5. Return `structuredContent` (object) alongside `content` (text array) from widget tools

### 6. Update Server Entry Point (`index.ts`)

Switch from `server.connect(transport)` with a low-level `Server` to `McpServer`:

```typescript
import { createMcpServer } from "./mcp-server.js";

app.all("/mcp", async (req, res) => {
  const server = createMcpServer(); // returns McpServer
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
    enableJsonResponse: true,
  });
  await server.connect(transport);
  await transport.handleRequest(req, res, req.body);
});
```

### 7. Build & Test

```bash
npm run install:all
npm run build:widgets
npm run dev:server
```

Verify with the MCP Inspector (`npx @modelcontextprotocol/inspector`) or connect from a host like Claude.

## Common Pitfalls

| Issue | Fix |
|-------|-----|
| `window.openai is undefined` | You missed replacing a `window.openai` reference in a widget component |
| Widget shows but no data | Ensure `structuredContent` is returned from the tool handler (not just `content`) |
| Theme not updating | Wire up `app.onhostcontextchanged` and call `setTheme()` |
| `registerAppTool` type errors | Import from `@modelcontextprotocol/ext-apps/server`, use `zod` for `inputSchema` |
| SSE gateway errors | Set `enableJsonResponse: true` on `StreamableHTTPServerTransport` |
| Resource not found by host | Ensure the `resourceUri` in `_meta.ui` exactly matches the URI in `registerAppResource` |

## Files Typically Changed

| File | Change |
|------|--------|
| `server/package.json` | Add `@modelcontextprotocol/ext-apps`, `zod` |
| `widgets/package.json` | Add `@modelcontextprotocol/ext-apps` |
| `server/src/mcp-server.ts` | Full rewrite: `McpServer` + `registerAppTool` + `registerAppResource` |
| `server/src/index.ts` | Update imports, `createMcpServer()` now returns `McpServer` |
| `widgets/src/hooks/useMcpApp.tsx` | New file: MCP Apps React context |
| `widgets/src/hooks/useThemeColors.ts` | Update import to use `useMcpTheme` |
| `widgets/src/**/main.tsx` | Wrap in `McpAppProvider`, use `useMcpTheme` |
| `widgets/src/**/*.tsx` | Replace all `window.openai.*` calls |
| `widgets/src/hooks/useOpenAiGlobal.ts` | Can be deleted after migration |
