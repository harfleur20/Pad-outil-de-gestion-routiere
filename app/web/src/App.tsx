import { useMemo, useState } from "react";
import RoadsPage from "./pages/RoadsPage";
import RoadDetailPage from "./pages/RoadDetailPage";
import EvaluationPage from "./pages/EvaluationPage";

type Screen = "roads" | "detail" | "evaluation";

export default function App() {
  const [screen, setScreen] = useState<Screen>("roads");

  const ScreenContent = useMemo(() => {
    if (screen === "roads") return <RoadsPage />;
    if (screen === "detail") return <RoadDetailPage />;
    return <EvaluationPage />;
  }, [screen]);

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero__brand">
          <span className="brand-mark" aria-hidden>
            PAD
          </span>
          <div>
            <h1>PAD Maintenance Routiere</h1>
            <p>Outil d'aide a la decision pour la maintenance des voies</p>
          </div>
        </div>
        <nav className="hero__nav">
          <button className={screen === "roads" ? "active" : ""} onClick={() => setScreen("roads")}>Catalogue voies</button>
          <button className={screen === "detail" ? "active" : ""} onClick={() => setScreen("detail")}>Fiche voie</button>
          <button className={screen === "evaluation" ? "active" : ""} onClick={() => setScreen("evaluation")}>Aide decision</button>
        </nav>
      </header>

      <main className="content">{ScreenContent}</main>
    </div>
  );
}
