const roadRows = [
  { code: "Rue.07", name: "Rue du Port de Peche", sap: "SAP1", start: "SCDP", end: "DAP" },
  { code: "Bvd.03", name: "Boulevard des Portiques", sap: "SAP2", start: "Gate B2", end: "Gate B1" },
  { code: "Av.01", name: "Avenue SAWA BEACH", sap: "SAP4", start: "Contournement", end: "Monument" }
];

export default function RoadsPage() {
  return (
    <section className="panel">
      <h2>Catalogue des voies</h2>
      <p className="muted">Filtrage SAP et recherche seront relies a SQLite dans l'etape suivante.</p>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Code</th>
              <th>Voie</th>
              <th>SAP</th>
              <th>Debut</th>
              <th>Fin</th>
            </tr>
          </thead>
          <tbody>
            {roadRows.map((row) => (
              <tr key={row.code}>
                <td>{row.code}</td>
                <td>{row.name}</td>
                <td>{row.sap}</td>
                <td>{row.start}</td>
                <td>{row.end}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}
