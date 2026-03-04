import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme, webDarkTheme } from "@fluentui/react-components";
import { McpAppProvider, useMcpTheme } from "../hooks/useMcpApp";
import { ClaimsDashboard } from "./ClaimsDashboard";

function ThemedApp() {
  const theme = useMcpTheme();
  return (
    <FluentProvider theme={theme === "dark" ? webDarkTheme : webLightTheme}>
      <ClaimsDashboard />
    </FluentProvider>
  );
}

function App() {
  return (
    <McpAppProvider name="Zava Claims Dashboard">
      <ThemedApp />
    </McpAppProvider>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
