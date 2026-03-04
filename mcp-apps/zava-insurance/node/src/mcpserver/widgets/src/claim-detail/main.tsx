import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme, webDarkTheme } from "@fluentui/react-components";
import { McpAppProvider, useMcpTheme } from "../hooks/useMcpApp";
import { ClaimDetail } from "./ClaimDetail";

function ThemedApp() {
  const theme = useMcpTheme();
  return (
    <FluentProvider theme={theme === "dark" ? webDarkTheme : webLightTheme}>
      <ClaimDetail />
    </FluentProvider>
  );
}

function App() {
  return (
    <McpAppProvider name="Zava Claim Detail">
      <ThemedApp />
    </McpAppProvider>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
