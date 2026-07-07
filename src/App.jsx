import { useState, Suspense, lazy } from "react";

const OrionImportBuilder = lazy(() => import("./tools/OrionImportBuilder.jsx"));
const ModelAuditTool = lazy(() => import("./tools/ModelAuditTool.jsx"));

const TOOLS = [
  { key: "import", label: "Orion Import Builder", component: OrionImportBuilder },
  { key: "audit", label: "Model Audit Tool", component: ModelAuditTool },
];

export default function App() {
  const [active, setActive] = useState("import");
  const ActiveTool = TOOLS.find(t => t.key === active)?.component;

  return (
    <div>
      <nav
        style={{
          display: "flex",
          gap: 4,
          padding: "10px 16px",
          background: "#0f172a",
          borderBottom: "1px solid #1e293b",
          position: "sticky",
          top: 0,
          zIndex: 100,
        }}
      >
        {TOOLS.map(t => (
          <button
            key={t.key}
            onClick={() => setActive(t.key)}
            style={{
              padding: "6px 14px",
              borderRadius: 6,
              border: "none",
              fontSize: 13,
              fontWeight: 600,
              cursor: "pointer",
              background: active === t.key ? "#3b82f6" : "transparent",
              color: active === t.key ? "#fff" : "#94a3b8",
              transition: "all 0.15s",
            }}
          >
            {t.label}
          </button>
        ))}
      </nav>
      <Suspense fallback={<div style={{ padding: 40, textAlign: "center", color: "#94a3b8" }}>Loading…</div>}>
        <ActiveTool />
      </Suspense>
    </div>
  );
}
