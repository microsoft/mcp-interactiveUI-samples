import React from "react";
import { createRoot } from "react-dom/client";
import { McpAppProvider } from "../hooks/useMcpApp";
import { DispatchList } from "./DispatchList";

function App() {
  return (
    <McpAppProvider name="Field Ops Dispatch List">
      <DispatchList />
    </McpAppProvider>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
