import { useCallback, useEffect, useMemo, useState } from "react";
import { padApi } from "./lib/pad-api";
import type {
  DataIntegrityReport,
  DataStatus,
  DecisionHistoryItem,
  DecisionResult,
  DrainageRule,
  DegradationItem,
  MaintenanceSolutionTemplate,
  RoadCatalogItem,
  SapSector,
  SheetColumnKey,
  SheetDefinition,
  SheetRow,
  SheetRowPayload
} from "./types/pad-api";

const DEFAULT_IMPORT_PATH = "C:\\Users\\harfl\\OneDrive\\Desktop\\programme ayissi.xlsx";
const MAX_ROWS = 1000;
const MAX_HISTORY = 500;
const DRAINAGE_OPERATORS: Array<DrainageRule["matchOperator"]> = ["CONTAINS", "EQUALS", "REGEX", "ALWAYS"];

type Feuil4Snapshot = {
  domain: string;
  sapSector: string;
  roadLabel: string;
  roadMatch: string;
  pkStart: string;
  pkEnd: string;
  observation: string;
  causeMatch: string;
  deflectionValue: string;
  severity: string;
  recommendation: string;
};

function createEmptyCells(columns: SheetColumnKey[]): Partial<Record<SheetColumnKey, string>> {
  return columns.reduce((acc, column) => {
    acc[column] = "";
    return acc;
  }, {} as Partial<Record<SheetColumnKey, string>>);
}

function getColumnLabel(sheet: SheetDefinition | null, column: SheetColumnKey) {
  return sheet?.columnLabels?.[column] || column;
}

function toPayload(
  columns: SheetColumnKey[],
  cells: Partial<Record<SheetColumnKey, string>>
): SheetRowPayload {
  const payload: SheetRowPayload = {};
  for (const column of columns) {
    payload[column] = String(cells[column] ?? "").trim();
  }
  return payload;
}

function toErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message) {
    return error.message;
  }
  return "Operation impossible.";
}

function fmtDate(value: string) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return date.toLocaleString();
}

function csvEscape(value: unknown) {
  const text = String(value ?? "");
  return `"${text.replace(/"/g, '""')}"`;
}

function normalizeLabel(value: unknown) {
  return String(value ?? "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();
}

function toDisplay(value: unknown) {
  const text = String(value ?? "").trim();
  return text || "-";
}

function toFrenchBoolean(value: unknown) {
  const normalized = normalizeLabel(value);
  if (["TRUE", "VRAI", "1", "YES", "OUI"].includes(normalized)) {
    return "VRAI";
  }
  if (["FALSE", "FAUX", "0", "NO", "NON"].includes(normalized)) {
    return "FAUX";
  }
  return toDisplay(value);
}

function toSolutionSourceLabel(source: DegradationItem["solutionSource"]) {
  if (source === "OVERRIDE") {
    return "Personnalisee";
  }
  if (source === "TEMPLATE") {
    return "Modele";
  }
  return "A parametrer";
}

function isToDetermineIntervention(value: unknown) {
  const normalized = normalizeLabel(value);
  if (!normalized) {
    return true;
  }
  return /A\s*DETERMINER|\(A\s*D\)/.test(normalized);
}

function getIntegrityIssueTotal(report: DataIntegrityReport | null) {
  if (!report) {
    return 0;
  }
  return report.issues.reduce((sum, issue) => sum + Number(issue.count || 0), 0);
}

function buildFeuil4Snapshot(rows: SheetRow[]): Feuil4Snapshot | null {
  if (rows.length === 0) {
    return null;
  }

  const rowByA = (prefix: string) =>
    rows.find((row) => normalizeLabel(row.A).startsWith(normalizeLabel(prefix))) ?? null;

  const domainRow = rowByA("Domaine");
  const sapSectorRow = rowByA("Secteur d'Activites Portuaires");
  const roadRow = rowByA("Blvd ou Rue");
  const pkHeaderRow = rowByA("Pk debut");
  const observationRow = rowByA("Entree observation faite sur la chaussee");
  const causeRow = rowByA("Entree la cause");
  const deflectionRow =
    rows.find((row) => normalizeLabel(row.B) === normalizeLabel("Valeur de la deflexion")) ?? null;
  const severityRow =
    rows.find((row) => ["FAIBLE", "MOYEN", "FORT", "TRES FORT"].includes(normalizeLabel(row.C))) ?? null;
  const recommendationRow =
    rows.find((row) => /(RENFORCEMENT|REHABILITATION|ENTRETIEN)/.test(normalizeLabel(row.D))) ?? null;

  const pkValuesRow = pkHeaderRow ? rows.find((row) => row.rowNo === pkHeaderRow.rowNo + 1) ?? null : null;

  return {
    domain: toDisplay(domainRow?.B),
    sapSector: toDisplay(sapSectorRow?.B),
    roadLabel: toDisplay(roadRow?.B),
    roadMatch: toFrenchBoolean(roadRow?.F),
    pkStart: toDisplay(pkValuesRow?.A),
    pkEnd: toDisplay(pkValuesRow?.B),
    observation: toDisplay(observationRow?.F),
    causeMatch: toFrenchBoolean(causeRow?.F),
    deflectionValue: toDisplay(deflectionRow?.D),
    severity: toDisplay(severityRow?.C),
    recommendation: toDisplay(recommendationRow?.D)
  };
}

export default function App() {
  const hasElectronBridge = Boolean(window.padApp);
  const appName = window.padApp?.appName || "PAD Maintenance Routiere";
  const appVersion = window.padApp?.appVersion || "0.0.0";
  const [status, setStatus] = useState<DataStatus | null>(null);
  const [integrityReport, setIntegrityReport] = useState<DataIntegrityReport | null>(null);
  const [definitions, setDefinitions] = useState<SheetDefinition[]>([]);
  const [activeView, setActiveView] = useState<string>("decision");

  const [rows, setRows] = useState<SheetRow[]>([]);
  const [search, setSearch] = useState("");
  const [importPath, setImportPath] = useState(DEFAULT_IMPORT_PATH);
  const [draftCells, setDraftCells] = useState<Partial<Record<SheetColumnKey, string>>>({});
  const [editingRowId, setEditingRowId] = useState<number | null>(null);

  const [sapSectors, setSapSectors] = useState<SapSector[]>([]);
  const [roads, setRoads] = useState<RoadCatalogItem[]>([]);
  const [degradations, setDegradations] = useState<DegradationItem[]>([]);
  const [selectedSap, setSelectedSap] = useState("");
  const [roadSearch, setRoadSearch] = useState("");
  const [selectedRoadId, setSelectedRoadId] = useState<number | "">("");
  const [selectedDegradationId, setSelectedDegradationId] = useState<number | "">("");
  const [deflectionValue, setDeflectionValue] = useState("");
  const [askDrainage, setAskDrainage] = useState(true);
  const [decisionResult, setDecisionResult] = useState<DecisionResult | null>(null);
  const [degradationSearch, setDegradationSearch] = useState("");
  const [historyRows, setHistoryRows] = useState<DecisionHistoryItem[]>([]);
  const [historySap, setHistorySap] = useState("");
  const [historySearch, setHistorySearch] = useState("");
  const [feuil4Snapshot, setFeuil4Snapshot] = useState<Feuil4Snapshot | null>(null);
  const [solutionTemplates, setSolutionTemplates] = useState<MaintenanceSolutionTemplate[]>([]);
  const [selectedTemplateKey, setSelectedTemplateKey] = useState("");
  const [solutionDraft, setSolutionDraft] = useState("");
  const [drainageRules, setDrainageRules] = useState<DrainageRule[]>([]);
  const [editingDrainageRuleId, setEditingDrainageRuleId] = useState<number | null>(null);
  const [drainageRuleOrder, setDrainageRuleOrder] = useState("");
  const [drainageRuleOperator, setDrainageRuleOperator] = useState<DrainageRule["matchOperator"]>("CONTAINS");
  const [drainageRulePattern, setDrainageRulePattern] = useState("");
  const [drainageRuleAskRequired, setDrainageRuleAskRequired] = useState(false);
  const [drainageRuleNeedsAttention, setDrainageRuleNeedsAttention] = useState(false);
  const [drainageRuleRecommendation, setDrainageRuleRecommendation] = useState("");
  const [drainageRuleActive, setDrainageRuleActive] = useState(true);

  const [isBusy, setIsBusy] = useState(false);
  const [isLoadingRows, setIsLoadingRows] = useState(false);
  const [isDecisionBusy, setIsDecisionBusy] = useState(false);
  const [isHistoryLoading, setIsHistoryLoading] = useState(false);
  const [isSolutionBusy, setIsSolutionBusy] = useState(false);
  const [isDrainageRuleBusy, setIsDrainageRuleBusy] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");

  const activeSheetName = activeView.startsWith("sheet:") ? activeView.replace("sheet:", "") : "";
  const activeSheet = useMemo(
    () => definitions.find((sheet) => sheet.name === activeSheetName) ?? null,
    [definitions, activeSheetName]
  );
  const activeColumns = activeSheet?.columns ?? [];
  const selectedRoad = useMemo(
    () => roads.find((road) => road.id === selectedRoadId) ?? null,
    [roads, selectedRoadId]
  );
  const selectedDegradation = useMemo(
    () => degradations.find((item) => item.id === selectedDegradationId) ?? null,
    [degradations, selectedDegradationId]
  );
  const filteredDegradations = useMemo(() => {
    const searchTerm = degradationSearch.trim().toLowerCase();
    if (!searchTerm) {
      return degradations;
    }
    return degradations.filter((item) => {
      return item.name.toLowerCase().includes(searchTerm) || item.causes.join(" ").toLowerCase().includes(searchTerm);
    });
  }, [degradationSearch, degradations]);
  const integrityIssueTotal = useMemo(() => getIntegrityIssueTotal(integrityReport), [integrityReport]);

  const refreshStatus = useCallback(async () => {
    const nextStatus = await padApi.getDataStatus();
    setStatus(nextStatus);
    if (nextStatus.lastImportPath) {
      setImportPath(nextStatus.lastImportPath);
    }
  }, []);

  const refreshDecisionCatalogs = useCallback(async () => {
    const [sectors, degradationCatalog] = await Promise.all([padApi.listSapSectors(), padApi.listDegradations()]);
    setSapSectors(sectors);
    setDegradations(degradationCatalog);
  }, []);

  const loadRoads = useCallback(async () => {
    const roadList = await padApi.listRoads({
      sapCode: selectedSap || undefined,
      search: roadSearch.trim() || undefined
    });
    setRoads(roadList);
  }, [selectedSap, roadSearch]);

  const loadRows = useCallback(async () => {
    if (!activeSheetName || activeView === "decision" || activeView === "catalogue" || activeView === "degradations" || activeView === "history") {
      setRows([]);
      return;
    }

    setIsLoadingRows(true);
    try {
      const nextRows = await padApi.listSheetRows(activeSheetName, {
        search: search.trim() || undefined,
        limit: MAX_ROWS
      });
      setRows(nextRows);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsLoadingRows(false);
    }
  }, [activeSheetName, activeView, search]);

  const loadHistory = useCallback(async () => {
    setIsHistoryLoading(true);
    try {
      const items = await padApi.listDecisionHistory({
        sapCode: historySap || undefined,
        search: historySearch.trim() || undefined,
        limit: MAX_HISTORY
      });
      setHistoryRows(items);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsHistoryLoading(false);
    }
  }, [historySap, historySearch]);

  const loadSolutionTemplates = useCallback(async () => {
    const templates = await padApi.listSolutionTemplates();
    setSolutionTemplates(templates);
  }, []);

  const loadDrainageRules = useCallback(async () => {
    const items = await padApi.listDrainageRules();
    setDrainageRules(items);
  }, []);

  const loadIntegrityReport = useCallback(async () => {
    const report = await padApi.getDataIntegrityReport();
    setIntegrityReport(report);
  }, []);

  const loadFeuil4Snapshot = useCallback(async () => {
    try {
      const feuil4Rows = await padApi.listSheetRows("Feuil4", { limit: 300 });
      setFeuil4Snapshot(buildFeuil4Snapshot(feuil4Rows));
    } catch {
      setFeuil4Snapshot(null);
    }
  }, []);

  const resetDraft = useCallback(() => {
    setEditingRowId(null);
    setDraftCells(createEmptyCells(activeColumns));
  }, [activeColumns]);

  useEffect(() => {
    if (!hasElectronBridge) {
      return;
    }

    let cancelled = false;

    async function bootstrap() {
      setIsBusy(true);
      try {
        const [sheetDefinitions] = await Promise.all([
          padApi.listSheetDefinitions(),
          refreshStatus(),
          refreshDecisionCatalogs(),
          loadRoads(),
          loadHistory(),
          loadFeuil4Snapshot(),
          loadSolutionTemplates(),
          loadDrainageRules(),
          loadIntegrityReport()
        ]);

        if (cancelled) {
          return;
        }

        setDefinitions(sheetDefinitions);
        setError("");
      } catch (err) {
        if (!cancelled) {
          setError(toErrorMessage(err));
        }
      } finally {
        if (!cancelled) {
          setIsBusy(false);
        }
      }
    }

    bootstrap();

    return () => {
      cancelled = true;
    };
  }, [
    hasElectronBridge,
    loadDrainageRules,
    loadFeuil4Snapshot,
    loadHistory,
    loadIntegrityReport,
    loadRoads,
    loadSolutionTemplates,
    refreshDecisionCatalogs,
    refreshStatus
  ]);

  useEffect(() => {
    if (!hasElectronBridge) {
      return;
    }
    loadRoads().catch((err) => setError(toErrorMessage(err)));
  }, [hasElectronBridge, loadRoads]);

  useEffect(() => {
    if (!hasElectronBridge) {
      return;
    }
    loadHistory().catch((err) => setError(toErrorMessage(err)));
  }, [hasElectronBridge, loadHistory]);

  useEffect(() => {
    if (selectedRoadId === "") {
      return;
    }
    const exists = roads.some((road) => road.id === selectedRoadId);
    if (!exists) {
      setSelectedRoadId("");
      setDecisionResult(null);
    }
  }, [roads, selectedRoadId]);

  useEffect(() => {
    if (!selectedDegradation) {
      setSelectedTemplateKey("");
      setSolutionDraft("");
      return;
    }
    setSelectedTemplateKey(selectedDegradation.templateKey ?? "");
    setSolutionDraft(selectedDegradation.solution || "");
  }, [selectedDegradation]);

  useEffect(() => {
    if (activeView === "decision" || activeView === "catalogue" || activeView === "degradations" || activeView === "history") {
      return;
    }
    resetDraft();
  }, [activeSheetName, activeView, resetDraft]);

  useEffect(() => {
    if (!hasElectronBridge || activeView === "decision" || activeView === "catalogue" || activeView === "degradations" || activeView === "history") {
      return;
    }
    loadRows();
  }, [hasElectronBridge, activeView, loadRows]);

  async function handleImport() {
    setIsBusy(true);
    try {
      const nextStatus = await padApi.importFromExcel(importPath.trim() || undefined);
      setStatus(nextStatus);
      await Promise.all([
        refreshDecisionCatalogs(),
        loadRoads(),
        loadHistory(),
        loadFeuil4Snapshot(),
        loadSolutionTemplates(),
        loadDrainageRules(),
        loadIntegrityReport()
      ]);
      if (activeView.startsWith("sheet:")) {
        await loadRows();
      }
      setNotice("Import Excel termine.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  async function handlePickImportPath() {
    try {
      const selected = await padApi.pickExcelFile();
      if (selected) {
        setImportPath(selected);
      }
    } catch (err) {
      setError(toErrorMessage(err));
    }
  }

  async function handleRefresh() {
    setIsBusy(true);
    try {
      await Promise.all([
        refreshStatus(),
        refreshDecisionCatalogs(),
        loadRoads(),
        loadHistory(),
        loadFeuil4Snapshot(),
        loadSolutionTemplates(),
        loadDrainageRules(),
        loadIntegrityReport()
      ]);
      if (activeView.startsWith("sheet:")) {
        await loadRows();
      }
      setError("");
      setNotice("Donnees actualisees.");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleEvaluateDecision() {
    if (!selectedRoadId) {
      setError("Selectionne une voie.");
      return;
    }
    if (!selectedDegradationId) {
      setError("Selectionne une degradation.");
      return;
    }

    const parsedD = Number(deflectionValue);
    const payload = {
      roadId: Number(selectedRoadId),
      degradationId: Number(selectedDegradationId),
      deflectionValue: Number.isFinite(parsedD) ? parsedD : undefined,
      askDrainage
    };

    setIsDecisionBusy(true);
    try {
      const result = await padApi.evaluateDecision(payload);
      setDecisionResult(result);
      await Promise.all([refreshStatus(), loadHistory()]);
      setError("");
      setNotice("Decision calculee et enregistree dans l'historique.");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsDecisionBusy(false);
    }
  }

  function handleEdit(row: SheetRow) {
    if (!activeSheet) {
      return;
    }

    const nextCells = createEmptyCells(activeSheet.columns);
    for (const column of activeSheet.columns) {
      nextCells[column] = String(row[column] ?? "");
    }

    setEditingRowId(row.id);
    setDraftCells(nextCells);
    setNotice(`Edition de la ligne #${row.id}`);
    setError("");
  }

  async function handleSaveRow() {
    if (!activeSheet) {
      return;
    }

    setIsBusy(true);
    try {
      const payload = toPayload(activeSheet.columns, draftCells);
      if (editingRowId) {
        await padApi.updateSheetRow(activeSheet.name, editingRowId, payload);
        setNotice(`Ligne #${editingRowId} mise a jour.`);
      } else {
        await padApi.createSheetRow(activeSheet.name, payload);
        setNotice("Nouvelle ligne ajoutee.");
      }

      await Promise.all([refreshStatus(), loadRows(), loadIntegrityReport()]);
      if (activeSheet.name === "Feuil4") {
        await loadFeuil4Snapshot();
      }
      resetDraft();
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleDeleteRow(rowId?: number) {
    if (!activeSheet) {
      return;
    }

    const id = rowId ?? editingRowId;
    if (!id) {
      return;
    }

    if (!window.confirm("Supprimer cette ligne ?")) {
      return;
    }

    setIsBusy(true);
    try {
      await padApi.deleteSheetRow(activeSheet.name, id);
      await Promise.all([refreshStatus(), loadRows(), loadIntegrityReport()]);
      if (activeSheet.name === "Feuil4") {
        await loadFeuil4Snapshot();
      }
      if (editingRowId === id) {
        resetDraft();
      }
      setNotice(`Ligne #${id} supprimee.`);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleClearHistory() {
    if (!window.confirm("Supprimer tout l'historique des decisions ?")) {
      return;
    }

    setIsBusy(true);
    try {
      await padApi.clearDecisionHistory();
      await Promise.all([refreshStatus(), loadHistory()]);
      setNotice("Historique vide.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleAssignSolutionTemplate() {
    if (!selectedDegradation) {
      setError("Selectionne une degradation.");
      return;
    }
    if (!selectedTemplateKey) {
      setError("Selectionne un modele de solution.");
      return;
    }

    setIsSolutionBusy(true);
    try {
      await padApi.assignTemplateToDegradation(selectedDegradation.code, selectedTemplateKey);
      await Promise.all([refreshDecisionCatalogs(), loadSolutionTemplates()]);
      setNotice("Modele de solution applique a la degradation.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsSolutionBusy(false);
    }
  }

  async function handleSaveSolutionOverride() {
    if (!selectedDegradation) {
      setError("Selectionne une degradation.");
      return;
    }
    if (!solutionDraft.trim()) {
      setError("Saisis une solution personnalisee.");
      return;
    }

    setIsSolutionBusy(true);
    try {
      await padApi.setDegradationSolutionOverride(selectedDegradation.code, solutionDraft.trim());
      await refreshDecisionCatalogs();
      setNotice("Solution personnalisee enregistree.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsSolutionBusy(false);
    }
  }

  async function handleClearSolutionOverride() {
    if (!selectedDegradation) {
      setError("Selectionne une degradation.");
      return;
    }

    setIsSolutionBusy(true);
    try {
      await padApi.clearDegradationSolutionOverride(selectedDegradation.code);
      await refreshDecisionCatalogs();
      setNotice("Solution personnalisee retiree (retour au modele).");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsSolutionBusy(false);
    }
  }

  function resetDrainageRuleEditor() {
    setEditingDrainageRuleId(null);
    setDrainageRuleOrder("");
    setDrainageRuleOperator("CONTAINS");
    setDrainageRulePattern("");
    setDrainageRuleAskRequired(false);
    setDrainageRuleNeedsAttention(false);
    setDrainageRuleRecommendation("");
    setDrainageRuleActive(true);
  }

  function handleEditDrainageRule(rule: DrainageRule) {
    setEditingDrainageRuleId(rule.id);
    setDrainageRuleOrder(String(rule.ruleOrder));
    setDrainageRuleOperator(rule.matchOperator);
    setDrainageRulePattern(rule.pattern || "");
    setDrainageRuleAskRequired(rule.askRequired);
    setDrainageRuleNeedsAttention(rule.needsAttention);
    setDrainageRuleRecommendation(rule.recommendation || "");
    setDrainageRuleActive(rule.isActive);
    setError("");
    setNotice(`Edition regle assainissement #${rule.id}`);
  }

  async function handleSaveDrainageRule() {
    const parsedOrder = Number(drainageRuleOrder);
    if (!Number.isFinite(parsedOrder) || parsedOrder <= 0) {
      setError("Rule order invalide (nombre > 0 attendu).");
      return;
    }
    if (!drainageRuleRecommendation.trim()) {
      setError("La recommandation assainissement est obligatoire.");
      return;
    }
    if (drainageRuleOperator !== "ALWAYS" && !drainageRulePattern.trim()) {
      setError("Le pattern est obligatoire sauf pour l'operateur ALWAYS.");
      return;
    }

    setIsDrainageRuleBusy(true);
    try {
      await padApi.upsertDrainageRule({
        id: editingDrainageRuleId ?? undefined,
        ruleOrder: Math.trunc(parsedOrder),
        matchOperator: drainageRuleOperator,
        pattern: drainageRuleOperator === "ALWAYS" ? "" : drainageRulePattern.trim(),
        askRequired: drainageRuleAskRequired,
        needsAttention: drainageRuleNeedsAttention,
        recommendation: drainageRuleRecommendation.trim(),
        isActive: drainageRuleActive
      });
      await loadDrainageRules();
      resetDrainageRuleEditor();
      setError("");
      setNotice("Regle assainissement enregistree.");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsDrainageRuleBusy(false);
    }
  }

  async function handleDeleteDrainageRule(ruleId?: number) {
    const id = Number(ruleId ?? editingDrainageRuleId);
    if (!Number.isFinite(id) || id <= 0) {
      setError("Selectionne une regle assainissement a supprimer.");
      return;
    }

    if (!window.confirm("Supprimer cette regle assainissement ?")) {
      return;
    }

    setIsDrainageRuleBusy(true);
    try {
      await padApi.deleteDrainageRule(id);
      await loadDrainageRules();
      if (editingDrainageRuleId === id) {
        resetDrainageRuleEditor();
      }
      setError("");
      setNotice(`Regle assainissement #${id} supprimee.`);
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsDrainageRuleBusy(false);
    }
  }

  function exportHistoryCsv() {
    if (historyRows.length === 0) {
      setError("Aucune ligne d'historique a exporter.");
      return;
    }

    const headers = [
      "Date",
      "SAP",
      "Code voie",
      "Designation voie",
      "Debut",
      "Fin",
      "Degradation",
      "Cause probable",
      "Solution maintenance",
      "Intervention contextuelle",
      "Deflexion D",
      "Severite",
      "Recommendation deflexion",
      "Alerte assainissement",
      "Recommendation assainissement"
    ];

    const lines = [
      headers.map(csvEscape).join(";"),
      ...historyRows.map((row) =>
        [
          row.createdAt,
          row.sapCode,
          row.roadCode,
          row.roadDesignation,
          row.startLabel,
          row.endLabel,
          row.degradationName,
          row.probableCause,
          row.maintenanceSolution,
          row.contextualIntervention,
          row.deflectionValue,
          row.deflectionSeverity,
          row.deflectionRecommendation,
          row.drainageNeedsAttention ? "oui" : "non",
          row.drainageRecommendation
        ]
          .map(csvEscape)
          .join(";")
      )
    ];

    const csv = lines.join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `pad-historique-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
    setNotice("Export CSV genere.");
  }

  function renderDecisionView() {
    return (
      <main className="workspace workspace--decision">
        <section className="panel decision-form">
          <h2>Aide a la decision maintenance</h2>
          <p className="muted">Selectionne une voie, une degradation et la valeur de deflexion (optionnelle).</p>

          <label htmlFor="sap">Secteur SAP</label>
          <select id="sap" value={selectedSap} onChange={(event) => setSelectedSap(event.target.value)}>
            <option value="">Tous les secteurs</option>
            {sapSectors.map((sector) => (
              <option key={sector.code} value={sector.code}>
                {sector.code}
              </option>
            ))}
          </select>

          <label htmlFor="road-search">Recherche voie</label>
          <input
            id="road-search"
            value={roadSearch}
            onChange={(event) => setRoadSearch(event.target.value)}
            placeholder="Code ou designation"
          />

          <label htmlFor="road">Voie</label>
          <select
            id="road"
            value={selectedRoadId}
            onChange={(event) => setSelectedRoadId(event.target.value ? Number(event.target.value) : "")}
          >
            <option value="">Selectionner une voie</option>
            {roads.map((road) => (
              <option key={road.id} value={road.id}>
                {road.sapCode || "SAP?"} | {road.roadCode} | {road.designation}
              </option>
            ))}
          </select>

          <label htmlFor="degradation">Degradation</label>
          <select
            id="degradation"
            value={selectedDegradationId}
            onChange={(event) => setSelectedDegradationId(event.target.value ? Number(event.target.value) : "")}
          >
            <option value="">Selectionner une degradation</option>
            {degradations.map((item) => (
              <option key={item.id} value={item.id}>
                {item.name}
              </option>
            ))}
          </select>

          <label htmlFor="deflection">Valeur de deflexion D</label>
          <input
            id="deflection"
            type="number"
            value={deflectionValue}
            onChange={(event) => setDeflectionValue(event.target.value)}
            placeholder="Ex: 80"
          />

          <label className="checkbox-row" htmlFor="ask-drainage">
            <input
              id="ask-drainage"
              type="checkbox"
              checked={askDrainage}
              onChange={(event) => setAskDrainage(event.target.checked)}
            />
            Interroger explicitement l'assainissement
          </label>

          <button className="primary" type="button" disabled={isDecisionBusy || isBusy} onClick={handleEvaluateDecision}>
            {isDecisionBusy ? "Analyse..." : "Analyser"}
          </button>

          {selectedRoad ? (
            <div className="card">
              <h3>Resume voie</h3>
              <p>
                <strong>SAP:</strong> {selectedRoad.sapCode || "-"}
              </p>
              <p>
                <strong>PK/borne debut:</strong> {selectedRoad.startLabel || "-"}
              </p>
              <p>
                <strong>PK/borne fin:</strong> {selectedRoad.endLabel || "-"}
              </p>
              <p>
                <strong>Longueur:</strong> {selectedRoad.lengthM ?? "-"} m
              </p>
              <p>
                <strong>Etat chaussee:</strong> {selectedRoad.pavementState || "-"}
              </p>
              <p>
                <strong>Assainissement (type caniveaux / description):</strong> {selectedRoad.drainageType || "-"} /{" "}
                {selectedRoad.drainageState || "-"}
              </p>
            </div>
          ) : null}

          {feuil4Snapshot ? (
            <div className="card">
              <h3>Repere Feuil4 (valeurs importees)</h3>
              <p>
                <strong>Domaine:</strong> {feuil4Snapshot.domain}
              </p>
              <p>
                <strong>SAP:</strong> {feuil4Snapshot.sapSector}
              </p>
              <p>
                <strong>Blvd/Rue:</strong> {feuil4Snapshot.roadLabel} ({feuil4Snapshot.roadMatch})
              </p>
              <p>
                <strong>Pk debut/fin:</strong> {feuil4Snapshot.pkStart} / {feuil4Snapshot.pkEnd}
              </p>
              <p>
                <strong>Observation:</strong> {feuil4Snapshot.observation}
              </p>
              <p>
                <strong>Validation cause:</strong> {feuil4Snapshot.causeMatch}
              </p>
              <p>
                <strong>D:</strong> {feuil4Snapshot.deflectionValue} | <strong>Etat:</strong> {feuil4Snapshot.severity}
              </p>
              <p>
                <strong>Intervention:</strong> {feuil4Snapshot.recommendation}
              </p>
            </div>
          ) : null}
        </section>

        <section className="panel decision-output">
          <h2>Resultat automatique</h2>
          {!decisionResult ? <p className="muted">Lance une analyse pour afficher la recommandation.</p> : null}

          {decisionResult ? (
            <div className="decision-grid">
              <div className="card">
                <h3>Voie selectionnee</h3>
                <p>
                  <strong>{decisionResult.road.designation}</strong> ({decisionResult.road.roadCode})
                </p>
                <p>{decisionResult.road.sapCode}</p>
                <p>
                  {decisionResult.road.startLabel}
                  {" -> "}
                  {decisionResult.road.endLabel}
                </p>
              </div>

              <div className="card">
                <h3>Degradation</h3>
                <p>{decisionResult.degradation.name}</p>
                <p>
                  <strong>Cause probable:</strong> {decisionResult.probableCause}
                </p>
              </div>

              <div className="card">
                <h3>Deflexion</h3>
                <p>
                  <strong>D:</strong> {decisionResult.deflection.value ?? "non renseigne"}
                </p>
                <p>
                  <strong>Etat:</strong> {decisionResult.deflection.severity}
                </p>
                <p>
                  <strong>Intervention:</strong> {decisionResult.deflection.recommendation}
                </p>
              </div>

              <div className="card card--full">
                <h3>Solution maintenance</h3>
                <p>{decisionResult.maintenanceSolution}</p>
              </div>

              <div className="card card--full">
                <h3>Assainissement</h3>
                <p>{decisionResult.drainage.recommendation}</p>
              </div>

              <div
                className={`card card--full ${
                  isToDetermineIntervention(decisionResult.contextualIntervention) ? "card--warning" : ""
                }`}
              >
                <h3>Intervention contextuelle troncon</h3>
                <p>{decisionResult.contextualIntervention || "a determiner (A D)"}</p>
              </div>

              <div className="card card--full">
                <h3>Causes connues (catalogue)</h3>
                <ul>
                  {decisionResult.degradation.causes.length > 0 ? (
                    decisionResult.degradation.causes.map((cause, index) => (
                      <li key={`${decisionResult.degradation.id}-${index}`}>{cause}</li>
                    ))
                  ) : (
                    <li>Aucune cause detaillee en base pour cette degradation.</li>
                  )}
                </ul>
              </div>
            </div>
          ) : null}
        </section>
      </main>
    );
  }

  function renderCatalogueView() {
    return (
      <main className="workspace workspace--full">
        <section className="panel table-panel table-panel--full">
          <h2>Catalogue des voies</h2>
          <p className="muted">Reference complete des voies par secteur SAP.</p>

          <div className="table-toolbar table-toolbar--triple">
            <select value={selectedSap} onChange={(event) => setSelectedSap(event.target.value)}>
              <option value="">Tous les secteurs</option>
              {sapSectors.map((sector) => (
                <option key={sector.code} value={sector.code}>
                  {sector.code}
                </option>
              ))}
            </select>
            <input
              value={roadSearch}
              onChange={(event) => setRoadSearch(event.target.value)}
              placeholder="Recherche code/designation/debut/fin"
            />
            <span className="muted">{roads.length} voie(s)</span>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Action</th>
                  <th>SAP</th>
                  <th>Code</th>
                  <th>Designation</th>
                  <th>Debut</th>
                  <th>Fin</th>
                  <th>Longueur (m)</th>
                  <th>Largeur (m)</th>
                  <th>Revetement</th>
                  <th>Etat chaussee</th>
                  <th>Type caniveaux</th>
                  <th>Description assain.</th>
                </tr>
              </thead>
              <tbody>
                {roads.map((road) => (
                  <tr key={road.id}>
                    <td>
                      <button
                        className="row-action"
                        type="button"
                        onClick={() => {
                          setSelectedRoadId(road.id);
                          setActiveView("decision");
                        }}
                      >
                        Evaluer
                      </button>
                    </td>
                    <td>{road.sapCode}</td>
                    <td>{road.roadCode}</td>
                    <td>{road.designation}</td>
                    <td>{road.startLabel}</td>
                    <td>{road.endLabel}</td>
                    <td>{road.lengthM ?? "-"}</td>
                    <td>{road.widthM ?? "-"}</td>
                    <td>{road.surfaceType || "-"}</td>
                    <td>{road.pavementState || "-"}</td>
                    <td>{road.drainageType || "-"}</td>
                    <td>{road.drainageState || "-"}</td>
                  </tr>
                ))}
                {roads.length === 0 ? (
                  <tr>
                    <td colSpan={12}>Aucune voie trouvee.</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>
        </section>
      </main>
    );
  }

  function renderDegradationsView() {
    return (
      <main className="workspace workspace--full">
        <section className="panel table-panel table-panel--full">
          <h2>Catalogue des degradations</h2>
          <p className="muted">Liste des degradations, causes probables et solution de maintenance.</p>

          <div className="table-toolbar table-toolbar--triple">
            <input
              value={degradationSearch}
              onChange={(event) => setDegradationSearch(event.target.value)}
              placeholder="Recherche degradation ou cause"
            />
            <span className="muted">{filteredDegradations.length} degradation(s)</span>
            <span className="muted">{degradations.length} total</span>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Action</th>
                  <th>Degradation</th>
                  <th>Nb causes</th>
                  <th>Cause principale</th>
                  <th>Source solution</th>
                  <th>Solution maintenance</th>
                </tr>
              </thead>
              <tbody>
                {filteredDegradations.map((item) => (
                  <tr key={item.id} className={selectedDegradationId === item.id ? "is-selected" : ""}>
                    <td>
                      <div className="row-buttons">
                        <button
                          className="row-action"
                          type="button"
                          onClick={() => {
                            setSelectedDegradationId(item.id);
                          }}
                        >
                          Configurer
                        </button>
                        <button
                          className="row-action"
                          type="button"
                          onClick={() => {
                            setSelectedDegradationId(item.id);
                            setActiveView("decision");
                          }}
                        >
                          Utiliser
                        </button>
                      </div>
                    </td>
                    <td>{item.name}</td>
                    <td>{item.causes.length}</td>
                    <td>{item.causes[0] || "-"}</td>
                    <td>{toSolutionSourceLabel(item.solutionSource)}</td>
                    <td>{item.solution}</td>
                  </tr>
                ))}
                {filteredDegradations.length === 0 ? (
                  <tr>
                    <td colSpan={6}>Aucune degradation trouvee.</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>

          {selectedDegradation ? (
            <div className="card card--spaced">
              <h3>Causes detaillees: {selectedDegradation.name}</h3>
              <p>
                <strong>Source solution:</strong> {toSolutionSourceLabel(selectedDegradation.solutionSource)}
              </p>
              <ul>
                {selectedDegradation.causes.length > 0 ? (
                  selectedDegradation.causes.map((cause, index) => <li key={`${selectedDegradation.id}-cause-${index}`}>{cause}</li>)
                ) : (
                  <li>Aucune cause detaillee pour cette degradation.</li>
                )}
              </ul>

              <h3>Parametrage solution de maintenance</h3>
              <label htmlFor="template-key">Modele de solution</label>
              <select
                id="template-key"
                value={selectedTemplateKey}
                onChange={(event) => setSelectedTemplateKey(event.target.value)}
                disabled={isSolutionBusy}
              >
                <option value="">Selectionner un modele</option>
                {solutionTemplates.map((template) => (
                  <option key={template.templateKey} value={template.templateKey}>
                    {template.title}
                  </option>
                ))}
              </select>

              <div className="editor-actions">
                <button className="row-action" type="button" onClick={handleAssignSolutionTemplate} disabled={isSolutionBusy}>
                  Appliquer modele
                </button>
              </div>

              <label htmlFor="solution-override">Solution personnalisee</label>
              <textarea
                id="solution-override"
                className="input-textarea"
                value={solutionDraft}
                onChange={(event) => setSolutionDraft(event.target.value)}
                placeholder="Saisir la solution de maintenance specifique a cette degradation"
                rows={5}
                disabled={isSolutionBusy}
              />

              <div className="editor-actions">
                <button className="primary" type="button" onClick={handleSaveSolutionOverride} disabled={isSolutionBusy}>
                  Enregistrer personnalisation
                </button>
                <button className="row-action row-action--danger" type="button" onClick={handleClearSolutionOverride} disabled={isSolutionBusy}>
                  Retirer personnalisation
                </button>
              </div>
            </div>
          ) : null}

          <div className="card card--spaced">
            <h3>Moteur assainissement (regles)</h3>
            <p className="muted">Regles utilisees automatiquement pendant l'evaluation de decision.</p>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Actions</th>
                    <th>Ordre</th>
                    <th>Operateur</th>
                    <th>Pattern</th>
                    <th>Question decideur</th>
                    <th>Alerte</th>
                    <th>Active</th>
                    <th>Recommandation</th>
                  </tr>
                </thead>
                <tbody>
                  {drainageRules.map((rule) => (
                    <tr key={rule.id} className={editingDrainageRuleId === rule.id ? "is-selected" : ""}>
                      <td>
                        <div className="row-buttons">
                          <button className="row-action" type="button" onClick={() => handleEditDrainageRule(rule)}>
                            Editer
                          </button>
                          <button className="row-action row-action--danger" type="button" onClick={() => handleDeleteDrainageRule(rule.id)}>
                            Suppr
                          </button>
                        </div>
                      </td>
                      <td>{rule.ruleOrder}</td>
                      <td>{rule.matchOperator}</td>
                      <td>{rule.pattern || "-"}</td>
                      <td>{rule.askRequired ? "Oui" : "Non"}</td>
                      <td>{rule.needsAttention ? "Oui" : "Non"}</td>
                      <td>{rule.isActive ? "Oui" : "Non"}</td>
                      <td>{rule.recommendation}</td>
                    </tr>
                  ))}
                  {drainageRules.length === 0 ? (
                    <tr>
                      <td colSpan={8}>Aucune regle assainissement.</td>
                    </tr>
                  ) : null}
                </tbody>
              </table>
            </div>

            <h3>{editingDrainageRuleId ? `Edition regle #${editingDrainageRuleId}` : "Nouvelle regle assainissement"}</h3>
            <div className="cells-grid">
              <div className="cell-field">
                <label htmlFor="dr-rule-order">Rule order</label>
                <input
                  id="dr-rule-order"
                  type="number"
                  min={1}
                  value={drainageRuleOrder}
                  onChange={(event) => setDrainageRuleOrder(event.target.value)}
                  disabled={isDrainageRuleBusy}
                />
              </div>

              <div className="cell-field">
                <label htmlFor="dr-rule-operator">Operateur</label>
                <select
                  id="dr-rule-operator"
                  value={drainageRuleOperator}
                  onChange={(event) => setDrainageRuleOperator(event.target.value as DrainageRule["matchOperator"])}
                  disabled={isDrainageRuleBusy}
                >
                  {DRAINAGE_OPERATORS.map((operator) => (
                    <option key={operator} value={operator}>
                      {operator}
                    </option>
                  ))}
                </select>
              </div>

              <div className="cell-field">
                <label htmlFor="dr-rule-pattern">Pattern</label>
                <input
                  id="dr-rule-pattern"
                  value={drainageRulePattern}
                  onChange={(event) => setDrainageRulePattern(event.target.value)}
                  placeholder={drainageRuleOperator === "ALWAYS" ? "Non utilise avec ALWAYS" : "Ex: OBSTR"}
                  disabled={isDrainageRuleBusy || drainageRuleOperator === "ALWAYS"}
                />
              </div>
            </div>

            <label className="checkbox-row" htmlFor="dr-rule-ask-required">
              <input
                id="dr-rule-ask-required"
                type="checkbox"
                checked={drainageRuleAskRequired}
                onChange={(event) => setDrainageRuleAskRequired(event.target.checked)}
                disabled={isDrainageRuleBusy}
              />
              Regle appliquee seulement si "Interroger explicitement l'assainissement" est coche
            </label>

            <label className="checkbox-row" htmlFor="dr-rule-needs-attention">
              <input
                id="dr-rule-needs-attention"
                type="checkbox"
                checked={drainageRuleNeedsAttention}
                onChange={(event) => setDrainageRuleNeedsAttention(event.target.checked)}
                disabled={isDrainageRuleBusy}
              />
              Declenche une alerte assainissement
            </label>

            <label className="checkbox-row" htmlFor="dr-rule-active">
              <input
                id="dr-rule-active"
                type="checkbox"
                checked={drainageRuleActive}
                onChange={(event) => setDrainageRuleActive(event.target.checked)}
                disabled={isDrainageRuleBusy}
              />
              Regle active
            </label>

            <label htmlFor="dr-rule-reco">Recommandation</label>
            <textarea
              id="dr-rule-reco"
              className="input-textarea"
              value={drainageRuleRecommendation}
              onChange={(event) => setDrainageRuleRecommendation(event.target.value)}
              placeholder="Texte de recommandation assainissement"
              rows={3}
              disabled={isDrainageRuleBusy}
            />

            <div className="editor-actions">
              <button className="primary" type="button" onClick={handleSaveDrainageRule} disabled={isDrainageRuleBusy}>
                {isDrainageRuleBusy ? "Enregistrement..." : "Enregistrer regle"}
              </button>
              <button className="row-action" type="button" onClick={resetDrainageRuleEditor} disabled={isDrainageRuleBusy}>
                Reinitialiser
              </button>
              <button
                className="row-action row-action--danger"
                type="button"
                onClick={() => handleDeleteDrainageRule()}
                disabled={isDrainageRuleBusy || !editingDrainageRuleId}
              >
                Supprimer
              </button>
            </div>
          </div>
        </section>
      </main>
    );
  }

  function renderHistoryView() {
    return (
      <main className="workspace workspace--full">
        <section className="panel table-panel table-panel--full">
          <h2>Historique et reporting</h2>
          <p className="muted">Historique des decisions calculees avec export CSV.</p>

          <div className="table-toolbar table-toolbar--penta">
            <select value={historySap} onChange={(event) => setHistorySap(event.target.value)}>
              <option value="">Tous les secteurs</option>
              {sapSectors.map((sector) => (
                <option key={sector.code} value={sector.code}>
                  {sector.code}
                </option>
              ))}
            </select>
            <input
              value={historySearch}
              onChange={(event) => setHistorySearch(event.target.value)}
              placeholder="Recherche voie/degradation/cause"
            />
            <button className="row-action" type="button" onClick={() => loadHistory()} disabled={isHistoryLoading}>
              Actualiser
            </button>
            <button className="primary" type="button" onClick={exportHistoryCsv}>
              Export CSV
            </button>
            <button className="row-action row-action--danger" type="button" onClick={handleClearHistory} disabled={isBusy}>
              Vider
            </button>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Date</th>
                  <th>SAP</th>
                  <th>Voie</th>
                  <th>Degradation</th>
                  <th>Cause probable</th>
                  <th>D</th>
                  <th>Severite</th>
                  <th>Intervention</th>
                  <th>Assainissement</th>
                </tr>
              </thead>
              <tbody>
                {historyRows.map((row) => (
                  <tr key={row.id}>
                    <td>{fmtDate(row.createdAt)}</td>
                    <td>{row.sapCode}</td>
                    <td>
                      {row.roadCode} - {row.roadDesignation}
                    </td>
                    <td>{row.degradationName}</td>
                    <td>{row.probableCause}</td>
                    <td>{row.deflectionValue ?? "-"}</td>
                    <td>{row.deflectionSeverity}</td>
                    <td>{row.deflectionRecommendation}</td>
                    <td>{row.drainageRecommendation}</td>
                  </tr>
                ))}
                {historyRows.length === 0 ? (
                  <tr>
                    <td colSpan={9}>{isHistoryLoading ? "Chargement..." : "Aucun historique disponible."}</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>
        </section>
      </main>
    );
  }

  function renderSheetView() {
    const isFeuil3Table = activeSheet?.name === "Feuil3";

    return (
      <main className="workspace">
        <section className="panel editor-panel">
          <h2>{activeSheet?.title ?? "Feuille"}</h2>
          <p className="muted">{activeSheet?.description ?? "Selectionne une feuille"}</p>

          <div className="cells-grid">
            {activeColumns.map((column) => (
              <div className="cell-field" key={column}>
                <label htmlFor={`cell-${column}`}>{getColumnLabel(activeSheet, column)}</label>
                <input
                  id={`cell-${column}`}
                  value={draftCells[column] ?? ""}
                  onChange={(event) =>
                    setDraftCells((prev) => ({
                      ...prev,
                      [column]: event.target.value
                    }))
                  }
                />
              </div>
            ))}
          </div>

          <div className="editor-actions">
            <button className="primary" type="button" onClick={handleSaveRow} disabled={isBusy || !activeSheet}>
              {editingRowId ? "Enregistrer" : "Ajouter"}
            </button>
            <button className="row-action" type="button" onClick={resetDraft} disabled={isBusy}>
              Reinitialiser
            </button>
            <button
              className="row-action row-action--danger"
              type="button"
              onClick={() => handleDeleteRow()}
              disabled={isBusy || !editingRowId}
            >
              Supprimer
            </button>
          </div>
        </section>

        <section className="panel table-panel">
          <div className="table-toolbar">
            <input
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Rechercher dans la feuille active"
            />
            <span className="muted">
              {activeSheet ? `${status?.sheetCounts?.[activeSheet.name] ?? 0} lignes en base` : "Aucune feuille"}
            </span>
          </div>

          <div className="table-wrap">
            <table className={isFeuil3Table ? "table--feuil3" : undefined}>
              <thead>
                {isFeuil3Table ? (
                  <>
                    <tr>
                      <th rowSpan={3}>Actions</th>
                      <th rowSpan={3}>id</th>
                      <th rowSpan={3}>{getColumnLabel(activeSheet, "A")}</th>
                      <th rowSpan={3}>{getColumnLabel(activeSheet, "B")}</th>
                      <th rowSpan={3}>{getColumnLabel(activeSheet, "C")}</th>
                      <th rowSpan={3}>{getColumnLabel(activeSheet, "D")}</th>
                      <th rowSpan={3}>{getColumnLabel(activeSheet, "E")}</th>
                      <th rowSpan={3}>{getColumnLabel(activeSheet, "F")}</th>
                      <th colSpan={2}>CHAUSSEE (en %)</th>
                      <th colSpan={2}>ASSAINISSEMENT</th>
                      <th rowSpan={3}>{getColumnLabel(activeSheet, "K")}</th>
                      <th rowSpan={3}>{getColumnLabel(activeSheet, "L")}</th>
                    </tr>
                    <tr>
                      <th>{getColumnLabel(activeSheet, "G")}</th>
                      <th>{getColumnLabel(activeSheet, "H")}</th>
                      <th>Type</th>
                      <th>Etat</th>
                    </tr>
                    <tr>
                      <th>&nbsp;</th>
                      <th>&nbsp;</th>
                      <th>caniveaux</th>
                      <th>description</th>
                    </tr>
                  </>
                ) : (
                  <tr>
                    <th>Actions</th>
                    <th>id</th>
                    {activeColumns.map((column) => (
                      <th key={`head-${column}`}>
                        {getColumnLabel(activeSheet, column)}
                      </th>
                    ))}
                  </tr>
                )}
              </thead>
              <tbody>
                {rows.map((row) => (
                  <tr key={row.id} className={editingRowId === row.id ? "is-selected" : ""}>
                    <td>
                      <div className="row-buttons">
                        <button className="row-action" type="button" onClick={() => handleEdit(row)}>
                          Editer
                        </button>
                        <button className="row-action row-action--danger" type="button" onClick={() => handleDeleteRow(row.id)}>
                          Suppr
                        </button>
                      </div>
                    </td>
                    <td>{row.id}</td>
                    {activeColumns.map((column) => {
                      const rawValue = String(row[column] ?? "");
                      const isFeuil3Intervention = activeSheet?.name === "Feuil3" && column === "L";
                      const value = isFeuil3Intervention && !rawValue.trim() ? "a determiner (A D)" : rawValue;
                      return (
                        <td
                          key={`${row.id}-${column}`}
                          title={value}
                          className={isFeuil3Intervention && isToDetermineIntervention(value) ? "cell-warning" : ""}
                        >
                          {value}
                        </td>
                      );
                    })}
                  </tr>
                ))}
                {rows.length === 0 ? (
                  <tr>
                    <td colSpan={activeColumns.length + 2}>{isLoadingRows ? "Chargement..." : "Aucune ligne."}</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>
        </section>
      </main>
    );
  }

  if (!hasElectronBridge) {
    return (
      <div className="bridge-error-shell">
        <div className="bridge-error-card">
          <img className="hero-logo" src="/logo-pad.png" alt="Logo Port Autonome de Douala" />
          <h1>Bridge Electron indisponible</h1>
          <p>Lance l'application avec `npm run dev` puis ouvre la fenetre Electron.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero__brand">
          <img className="hero-logo" src="/logo-pad.png" alt="Logo Port Autonome de Douala" />
          <div>
            <h1>PAD Outil de Maintenance Routiere</h1>
            <p>Pilotez la maintenance routiere du PAD avec des decisions rapides et fiables.</p>
          </div>
        </div>

        <div className="hero__status">
          <span className="pill">Total lignes: {status?.totalRows ?? "-"}</span>
          <span className="pill">Voies: {roads.length}</span>
          <span className="pill">Degradations: {degradations.length}</span>
          <span className="pill">Historique: {status?.decisionHistoryCount ?? 0}</span>
          <span className="pill">Dernier import: {status?.lastImportAt ?? "-"}</span>
          <span className={`pill ${integrityReport?.status === "WARNING" ? "pill--warn" : "pill--ok"}`}>
            Coherence: {integrityReport?.status === "WARNING" ? "a verifier" : "OK"} ({integrityIssueTotal})
          </span>
        </div>

        <div className="hero__actions">
          <input
            className="hero-input hero-input--path"
            value={importPath}
            onChange={(event) => setImportPath(event.target.value)}
            placeholder="Chemin du fichier Excel"
          />
          <button className="icon-btn" type="button" onClick={handlePickImportPath} disabled={isBusy}>
            Parcourir
          </button>
          <button className="icon-btn" type="button" onClick={handleImport} disabled={isBusy}>
            Importer
          </button>
          <button className="icon-btn" type="button" onClick={handleRefresh} disabled={isBusy}>
            Actualiser
          </button>
        </div>

        {error ? <p className="hero__error">{error}</p> : null}
        {notice ? <p className="hero__notice">{notice}</p> : null}
        {integrityReport?.issues?.length ? (
          <div className="integrity-alert">
            <strong>Audit coherence:</strong>{" "}
            {integrityReport.issues
              .slice(0, 3)
              .map((issue) => `${issue.message} (${issue.count})`)
              .join(" | ")}
          </div>
        ) : null}

        <nav className="hero__nav">
          <button className={activeView === "decision" ? "active" : ""} type="button" onClick={() => setActiveView("decision")}>
            Aide decision
          </button>
          <button className={activeView === "catalogue" ? "active" : ""} type="button" onClick={() => setActiveView("catalogue")}>
            Catalogue
          </button>
          <button
            className={activeView === "degradations" ? "active" : ""}
            type="button"
            onClick={() => setActiveView("degradations")}
          >
            Degradations
          </button>
          <button className={activeView === "history" ? "active" : ""} type="button" onClick={() => setActiveView("history")}>
            Historique
          </button>
          {definitions.map((sheet) => (
            <button
              key={sheet.name}
              className={activeView === `sheet:${sheet.name}` ? "active" : ""}
              type="button"
              onClick={() => setActiveView(`sheet:${sheet.name}`)}
            >
              {sheet.name}
            </button>
          ))}
        </nav>
      </header>

      {activeView === "decision" ? renderDecisionView() : null}
      {activeView === "catalogue" ? renderCatalogueView() : null}
      {activeView === "degradations" ? renderDegradationsView() : null}
      {activeView === "history" ? renderHistoryView() : null}
      {activeView.startsWith("sheet:") ? renderSheetView() : null}

      <footer className="app-footer">
        <span>{appName}</span>
        <span>Version {appVersion}</span>
      </footer>
    </div>
  );
}
