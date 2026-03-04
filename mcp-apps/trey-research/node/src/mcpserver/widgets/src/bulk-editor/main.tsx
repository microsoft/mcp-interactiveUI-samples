import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme, webDarkTheme } from "@fluentui/react-components";
import { BulkEditor } from "./BulkEditor";
import { McpAppProvider, useMcpTheme } from "../hooks/useMcpApp";

function ThemedApp() {
  const theme = useMcpTheme();
  return (
    <FluentProvider theme={theme === "dark" ? webDarkTheme : webLightTheme}>
      <BulkEditor />
    </FluentProvider>
  );
}

createRoot(document.getElementById("root")!).render(
  <McpAppProvider name="Bulk Editor">
    <ThemedApp />
  </McpAppProvider>
);
