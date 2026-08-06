import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "./lib/apiStorage";
import "./index.css";
import DrumTracker from "./DrumTracker.jsx";

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <DrumTracker />
  </StrictMode>
);
