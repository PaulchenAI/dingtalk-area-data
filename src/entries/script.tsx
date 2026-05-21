import React, { useEffect } from "react";
import ReactDOM from "react-dom/client";
import { initScript } from "dingtalk-docs-cool-app";

declare const process: any;

function App() {
  useEffect(() => {
    const publicUrl = process?.env?.PUBLIC_URL || "";

    initScript({
      scriptUrl: new URL(`${publicUrl}/static/js/script.code.js`, window.location.href),
      onError: (error) => {
        console.error("init script failed:", error);
      },
    });
  }, []);

  return null;
}

const root = ReactDOM.createRoot(document.getElementById("root_script")!);

root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
