import React from "react";
import { createRoot } from "react-dom/client";
import { McpAppProvider } from "../hooks/useMcpApp";
import { DispatchMap } from "./DispatchMap";

createRoot(document.getElementById("root")!).render(
  <McpAppProvider name="Field Ops Dispatch Map">
    <DispatchMap />
  </McpAppProvider>
);
