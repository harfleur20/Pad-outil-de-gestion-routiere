export default function RoadDetailPage() {
  return (
    <section className="panel">
      <h2>Fiche voie</h2>
      <div className="grid">
        <article className="card">
          <h3>Identification</h3>
          <p><strong>Voie:</strong> Rue du Port de Peche</p>
          <p><strong>SAP:</strong> SAP1</p>
          <p><strong>PK:</strong> 0+000 a 0+730</p>
        </article>
        <article className="card">
          <h3>Etat chaussee</h3>
          <p>Etat moyen avec degradations localisees.</p>
          <p>Revetement: BB</p>
        </article>
        <article className="card">
          <h3>Assainissement</h3>
          <p>Caniveaux obstrues a certains endroits.</p>
          <p>Action suggeree: curage general.</p>
        </article>
      </div>
    </section>
  );
}
