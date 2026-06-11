import { StrictMode, useState } from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";
import FileGate from "./FileGate.jsx";
import "./index.css";

function Root() {
  const [ready, setReady] = useState(false);
  // App.jsx kicks off file reading on mount, so only render it once the gate
  // has populated the store and installed the window.fs shim.
  return ready ? <App /> : <FileGate onReady={() => setReady(true)} />;
}

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <Root />
  </StrictMode>
);
