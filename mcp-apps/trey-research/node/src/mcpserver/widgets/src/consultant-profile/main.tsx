import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme, webDarkTheme } from "@fluentui/react-components";
import { ConsultantProfile } from "./ConsultantProfile";
import { McpAppProvider, useMcpTheme } from "../hooks/useMcpApp";

function ThemedApp() {
  const theme = useMcpTheme();
  return (
    <FluentProvider theme={theme === "dark" ? webDarkTheme : webLightTheme}>
      <ConsultantProfile />
    </FluentProvider>
  );
}

createRoot(document.getElementById("root")!).render(
  <McpAppProvider name="Consultant Profile">
    <ThemedApp />
  </McpAppProvider>
);
