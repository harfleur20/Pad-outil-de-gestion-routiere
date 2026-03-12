export default function EvaluationPage() {
  return (
    <section className="panel">
      <h2>Aide a la decision</h2>
      <div className="grid grid--two">
        <article className="card">
          <h3>Entree</h3>
          <label>Voie</label>
          <input value="Rue du Port de Peche" readOnly />
          <label>Degradation</label>
          <select defaultValue="FISSURES TRANSVERSALE">
            <option>FISSURES TRANSVERSALE</option>
            <option>NIDS DE POULE</option>
            <option>ORNIERAGES</option>
          </select>
          <label>Deflexion D</label>
          <input type="number" placeholder="Ex: 80" />
        </article>

        <article className="card">
          <h3>Sortie</h3>
          <p><strong>Cause probable:</strong> Fatigue de la chaussee / perte de portance.</p>
          <p><strong>Action:</strong> Renforcement lourd.</p>
          <p><strong>Assainissement:</strong> Curage et nettoyage des caniveaux.</p>
        </article>
      </div>
    </section>
  );
}
