import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

// Aplica o tema salvo antes de renderizar (evita "flash" de tela clara).
try {
  const saved = localStorage.getItem("fc-theme");
  if (saved === "dark" || saved === "light") document.documentElement.setAttribute("data-theme", saved);
} catch {}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
