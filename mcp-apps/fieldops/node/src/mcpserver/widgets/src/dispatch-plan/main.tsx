import React from "react";
import { createRoot } from "react-dom/client";
import { McpAppProvider } from "../hooks/useMcpApp";
import { DispatchPlan } from "./DispatchPlan";

createRoot(document.getElementById("root")!).render(
  <McpAppProvider name="Field Ops Dispatch Plan">
    <DispatchPlan />
  </McpAppProvider>
);
