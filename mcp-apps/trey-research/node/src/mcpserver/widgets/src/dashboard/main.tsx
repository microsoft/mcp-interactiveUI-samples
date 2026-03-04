import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme, webDarkTheme } from "@fluentui/react-components";
import { Dashboard } from "./Dashboard";
import { McpAppProvider, useMcpTheme } from "../hooks/useMcpApp";

function ThemedApp() {
  const theme = useMcpTheme();
  return (
    <FluentProvider theme={theme === "dark" ? webDarkTheme : webLightTheme}>
      <Dashboard />
    </FluentProvider>
  );
}

createRoot(document.getElementById("root")!).render(
  <McpAppProvider name="HR Dashboard">
    <ThemedApp />
  </McpAppProvider>
);
