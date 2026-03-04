import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme, webDarkTheme } from "@fluentui/react-components";
import { McpAppProvider, useMcpTheme } from "../hooks/useMcpApp";
import { ContractorsList } from "./ContractorsList";

function ThemedApp() {
  const theme = useMcpTheme();
  return (
    <FluentProvider theme={theme === "dark" ? webDarkTheme : webLightTheme}>
      <ContractorsList />
    </FluentProvider>
  );
}

function App() {
  return (
    <McpAppProvider name="Zava Contractors List">
      <ThemedApp />
    </McpAppProvider>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
