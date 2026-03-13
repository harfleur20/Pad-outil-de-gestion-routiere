import { useCallback, useEffect, useMemo, useState } from "react";
import { padApi } from "./lib/pad-api";
import { FolderOpen, Gauge, Pencil, RefreshCw, Settings2, Trash2, Upload } from "lucide-react";
import type {
  DataStatus,
  DecisionHistoryItem,
  DecisionResult,
  DrainageRule,
  DegradationItem,
  MaintenanceInterventionItem,
  MaintenanceInterventionPayload,
  MaintenanceInterventionStatus,
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
const MAX_MAINTENANCE_ROWS = 500;
const DRAINAGE_OPERATORS: Array<DrainageRule["matchOperator"]> = ["CONTAINS", "EQUALS", "REGEX", "ALWAYS"];
const MAINTENANCE_STATUSES: MaintenanceInterventionStatus[] = ["PREVU", "EN_COURS", "TERMINE"];
const MAINTENANCE_TYPE_SUGGESTIONS = [
  "Entretien préventif",
  "Curage caniveaux",
  "Bouchage nids de poule",
  "Reprise de surface de chaussée",
  "Reprofilage",
  "Rechargement",
  "Renforcement léger",
  "Renforcement lourd",
  "Réhabilitation"
];

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

type DeflectionPreview = {
  value: string;
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
  return "Opération impossible.";
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

function getDisplayRowNumber(row: Pick<SheetRow, "rowNo" | "id">) {
  const value = Number(row.rowNo);
  if (Number.isFinite(value) && value > 0) {
    return value;
  }
  return row.id;
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
    return "Personnalisée";
  }
  if (source === "TEMPLATE") {
    return "Modèle";
  }
  return "À paramétrer";
}

function toMaintenanceStatusLabel(status: MaintenanceInterventionStatus | string) {
  if (status === "PREVU") {
    return "Prévu";
  }
  if (status === "EN_COURS") {
    return "En cours";
  }
  if (status === "TERMINE") {
    return "Terminé";
  }
  return toDisplay(status);
}

function toDeflectionSeverityLabel(value: unknown) {
  const normalized = normalizeLabel(value);
  if (normalized === "NON RENSEIGNE") {
    return "NON RENSEIGNÉ";
  }
  if (normalized === "TRES FORT") {
    return "TRÈS FORT";
  }
  return toDisplay(value);
}

function toDeflectionRecommendationLabel(value: unknown) {
  const text = toDisplay(value);
  const normalized = normalizeLabel(text);
  if (normalized === "RENFORCEMENT LEGER") {
    return "RENFORCEMENT LÉGER";
  }
  if (normalized.startsWith("REHABILITATION")) {
    return text.replace(/REHABILITATION/gi, "RÉHABILITATION");
  }
  return text;
}

function formatAmount(value: number | null) {
  if (!Number.isFinite(value ?? null)) {
    return "-";
  }
  return `${Number(value).toLocaleString()} FCFA`;
}

function isToDetermineIntervention(value: unknown) {
  const normalized = normalizeLabel(value);
  if (!normalized) {
    return true;
  }
  return /A\s*DETERMINER|\(A\s*D\)/.test(normalized);
}

function classifyDeflectionPreview(value: string): DeflectionPreview {
  const numeric = Number(String(value).replace(",", "."));
  if (!Number.isFinite(numeric)) {
    return {
      value: "-",
      severity: "NON RENSEIGNE",
      recommendation: "Renseigner D pour proposer le niveau d'intervention."
    };
  }

  if (numeric < 60) {
    return { value: String(numeric), severity: "FAIBLE", recommendation: "PAS D'ENTRETIEN" };
  }
  if (numeric < 80) {
    return { value: String(numeric), severity: "MOYEN", recommendation: "RENFORCEMENT LEGER" };
  }
  if (numeric < 90) {
    return { value: String(numeric), severity: "FORT", recommendation: "RENFORCEMENT LOURD" };
  }
  return {
    value: String(numeric),
    severity: "TRÈS FORT",
    recommendation: "REHABILITATION COUCHE DE ROULEMENT ET DE BASE"
  };
}

function buildFeuil4Snapshot(rows: SheetRow[]): Feuil4Snapshot | null {
  if (rows.length === 0) {
    return null;
  }

  const rowByA = (prefix: string) =>
    rows.find((row) => normalizeLabel(row.A).startsWith(normalizeLabel(prefix))) ?? null;

  const domainRow = rowByA("Domaine");
  const sapSectorRow = rowByA("Secteur d'Activités Portuaires");
  const roadRow = rowByA("Blvd ou Rue");
  const pkHeaderRow = rowByA("Pk début");
  const observationRow = rowByA("Entrée observation faite sur la chaussée");
  const causeRow = rowByA("Entrée la cause");
  const deflectionRow =
    rows.find((row) => normalizeLabel(row.B) === normalizeLabel("Valeur de la déflexion")) ?? null;
  const severityRow =
    rows.find((row) => ["FAIBLE", "MOYEN", "FORT", "TRÈS FORT"].includes(normalizeLabel(row.C))) ?? null;
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
  const appName = window.padApp?.appName || "PAD Maintenance Routière";
  const appVersion = window.padApp?.appVersion || "0.0.0";
  const [status, setStatus] = useState<DataStatus | null>(null);
  const [definitions, setDefinitions] = useState<SheetDefinition[]>([]);
  const [activeView, setActiveView] = useState<string>("decision");

  const [rows, setRows] = useState<SheetRow[]>([]);
  const [search, setSearch] = useState("");
  const [importPath, setImportPath] = useState(DEFAULT_IMPORT_PATH);
  const [draftCells, setDraftCells] = useState<Partial<Record<SheetColumnKey, string>>>({});
  const [editingRowId, setEditingRowId] = useState<number | null>(null);

  const [sapSectors, setSapSectors] = useState<SapSector[]>([]);
  const [allRoads, setAllRoads] = useState<RoadCatalogItem[]>([]);
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
  const [maintenanceRows, setMaintenanceRows] = useState<MaintenanceInterventionItem[]>([]);
  const [maintenancePreviewRows, setMaintenancePreviewRows] = useState<MaintenanceInterventionItem[]>([]);
  const [maintenanceSap, setMaintenanceSap] = useState("");
  const [maintenanceRoadFilterId, setMaintenanceRoadFilterId] = useState<number | "">("");
  const [maintenanceStatusFilter, setMaintenanceStatusFilter] = useState<MaintenanceInterventionStatus | "">("");
  const [maintenanceSearch, setMaintenanceSearch] = useState("");
  const [editingMaintenanceId, setEditingMaintenanceId] = useState<number | null>(null);
  const [maintenanceRoadId, setMaintenanceRoadId] = useState<number | "">("");
  const [maintenanceDegradationCode, setMaintenanceDegradationCode] = useState("");
  const [maintenanceType, setMaintenanceType] = useState("");
  const [maintenanceStatus, setMaintenanceStatus] = useState<MaintenanceInterventionStatus>("PREVU");
  const [maintenanceDate, setMaintenanceDate] = useState("");
  const [maintenanceCompletionDate, setMaintenanceCompletionDate] = useState("");
  const [maintenanceStateBefore, setMaintenanceStateBefore] = useState("");
  const [maintenanceStateAfter, setMaintenanceStateAfter] = useState("");
  const [maintenanceDeflectionBefore, setMaintenanceDeflectionBefore] = useState("");
  const [maintenanceDeflectionAfter, setMaintenanceDeflectionAfter] = useState("");
  const [maintenanceSolutionApplied, setMaintenanceSolutionApplied] = useState("");
  const [maintenanceContractorName, setMaintenanceContractorName] = useState("");
  const [maintenanceObservation, setMaintenanceObservation] = useState("");
  const [maintenanceCostAmount, setMaintenanceCostAmount] = useState("");
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
  const [isMaintenanceLoading, setIsMaintenanceLoading] = useState(false);
  const [isMaintenanceBusy, setIsMaintenanceBusy] = useState(false);
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
    () => allRoads.find((road) => road.id === selectedRoadId) ?? null,
    [allRoads, selectedRoadId]
  );
  const selectedDegradation = useMemo(
    () => degradations.find((item) => item.id === selectedDegradationId) ?? null,
    [degradations, selectedDegradationId]
  );
  const maintenanceSelectedRoad = useMemo(
    () => allRoads.find((road) => road.id === maintenanceRoadId) ?? null,
    [allRoads, maintenanceRoadId]
  );
  const maintenanceFilterRoadOptions = useMemo(() => {
    if (!maintenanceSap) {
      return allRoads;
    }
    return allRoads.filter((road) => road.sapCode === maintenanceSap);
  }, [allRoads, maintenanceSap]);
  const latestMaintenance = maintenancePreviewRows[0] ?? null;
  const filteredDegradations = useMemo(() => {
    const searchTerm = degradationSearch.trim().toLowerCase();
    if (!searchTerm) {
      return degradations;
    }
    return degradations.filter((item) => {
      return item.name.toLowerCase().includes(searchTerm) || item.causes.join(" ").toLowerCase().includes(searchTerm);
    });
  }, [degradationSearch, degradations]);
  const dynamicFeuil4Snapshot = useMemo(() => {
    if (!selectedRoad) {
      return null;
    }

    const fallback = feuil4Snapshot;
  const preview = classifyDeflectionPreview(deflectionValue || fallback?.deflectionValue || "");
  const decisionSeverity = toDeflectionSeverityLabel(decisionResult?.deflection?.severity || preview.severity);
  const decisionRecommendation = toDeflectionRecommendationLabel(
    decisionResult?.deflection?.recommendation || preview.recommendation
  );
    const selectedRoadLabel = [selectedRoad.roadCode, selectedRoad.designation].filter(Boolean).join(" - ");
    const hasCauses = Boolean(selectedDegradation && selectedDegradation.causes.length > 0);

    return {
      domain: fallback?.domain || "Port de Douala Bonaberi",
      sapSector: selectedRoad.sapCode || fallback?.sapSector || "-",
      roadLabel: selectedRoadLabel || fallback?.roadLabel || "-",
      roadMatch: "VRAI",
      pkStart: selectedRoad.startLabel || fallback?.pkStart || "-",
      pkEnd: selectedRoad.endLabel || fallback?.pkEnd || "-",
      observation: selectedDegradation?.name || fallback?.observation || "-",
      causeMatch: selectedDegradation ? (hasCauses ? "VRAI" : "FAUX") : fallback?.causeMatch || "-",
      deflectionValue: decisionResult?.deflection?.value != null ? String(decisionResult.deflection.value) : preview.value,
      severity: decisionSeverity || toDeflectionSeverityLabel(fallback?.severity) || "-",
      recommendation: decisionRecommendation || toDeflectionRecommendationLabel(fallback?.recommendation) || "-"
    } satisfies Feuil4Snapshot;
  }, [decisionResult, deflectionValue, feuil4Snapshot, selectedDegradation, selectedRoad]);

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

  const loadAllRoads = useCallback(async () => {
    const roadList = await padApi.listRoads();
    setAllRoads(roadList);
  }, []);

  const loadRows = useCallback(async () => {
    if (
      !activeSheetName ||
      activeView === "decision" ||
      activeView === "catalogue" ||
      activeView === "degradations" ||
      activeView === "maintenance" ||
      activeView === "history"
    ) {
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

  const loadMaintenanceRows = useCallback(async () => {
    setIsMaintenanceLoading(true);
    try {
      const items = await padApi.listMaintenanceInterventions({
        sapCode: maintenanceSap || undefined,
        roadId: maintenanceRoadFilterId || undefined,
        status: maintenanceStatusFilter || undefined,
        search: maintenanceSearch.trim() || undefined,
        limit: MAX_MAINTENANCE_ROWS
      });
      setMaintenanceRows(items);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsMaintenanceLoading(false);
    }
  }, [maintenanceRoadFilterId, maintenanceSap, maintenanceSearch, maintenanceStatusFilter]);

  const loadMaintenancePreview = useCallback(async (roadId?: number | "") => {
    const normalizedRoadId = Number(roadId);
    if (!Number.isFinite(normalizedRoadId) || normalizedRoadId <= 0) {
      setMaintenancePreviewRows([]);
      return;
    }

    try {
      const items = await padApi.listMaintenanceInterventions({
        roadId: normalizedRoadId,
        limit: 3
      });
      setMaintenancePreviewRows(items);
    } catch {
      setMaintenancePreviewRows([]);
    }
  }, []);

  const loadSolutionTemplates = useCallback(async () => {
    const templates = await padApi.listSolutionTemplates();
    setSolutionTemplates(templates);
  }, []);

  const loadDrainageRules = useCallback(async () => {
    const items = await padApi.listDrainageRules();
    setDrainageRules(items);
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

  const resetMaintenanceForm = useCallback(() => {
    setEditingMaintenanceId(null);
    setMaintenanceRoadId("");
    setMaintenanceDegradationCode("");
    setMaintenanceType("");
    setMaintenanceStatus("PREVU");
    setMaintenanceDate("");
    setMaintenanceCompletionDate("");
    setMaintenanceStateBefore("");
    setMaintenanceStateAfter("");
    setMaintenanceDeflectionBefore("");
    setMaintenanceDeflectionAfter("");
    setMaintenanceSolutionApplied("");
    setMaintenanceContractorName("");
    setMaintenanceObservation("");
    setMaintenanceCostAmount("");
  }, []);

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
          loadAllRoads(),
          loadRoads(),
          loadHistory(),
          loadMaintenanceRows(),
          loadFeuil4Snapshot(),
          loadSolutionTemplates(),
          loadDrainageRules()
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
    loadAllRoads,
    loadDrainageRules,
    loadFeuil4Snapshot,
    loadHistory,
    loadMaintenanceRows,
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
    if (!hasElectronBridge) {
      return;
    }
    loadMaintenanceRows().catch((err) => setError(toErrorMessage(err)));
  }, [hasElectronBridge, loadMaintenanceRows]);

  useEffect(() => {
    if (!hasElectronBridge) {
      return;
    }
    loadMaintenancePreview(selectedRoadId).catch((err) => setError(toErrorMessage(err)));
  }, [hasElectronBridge, loadMaintenancePreview, selectedRoadId]);

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
    if (
      activeView === "decision" ||
      activeView === "catalogue" ||
      activeView === "degradations" ||
      activeView === "maintenance" ||
      activeView === "history"
    ) {
      return;
    }
    resetDraft();
  }, [activeSheetName, activeView, resetDraft]);

  useEffect(() => {
    if (
      !hasElectronBridge ||
      activeView === "decision" ||
      activeView === "catalogue" ||
      activeView === "degradations" ||
      activeView === "maintenance" ||
      activeView === "history"
    ) {
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
        loadAllRoads(),
        loadRoads(),
        loadHistory(),
        loadMaintenanceRows(),
        loadFeuil4Snapshot(),
        loadSolutionTemplates(),
        loadDrainageRules()
      ]);
      if (activeView.startsWith("sheet:")) {
        await loadRows();
      }
      setNotice("Import Excel terminé.");
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
        loadAllRoads(),
        loadRoads(),
        loadHistory(),
        loadMaintenanceRows(),
        loadFeuil4Snapshot(),
        loadSolutionTemplates(),
        loadDrainageRules()
      ]);
      if (activeView.startsWith("sheet:")) {
        await loadRows();
      }
      setError("");
      setNotice("Données actualisées.");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleEvaluateDecision() {
    if (!selectedRoadId) {
      setError("Sélectionne une voie.");
      return;
    }
    if (!selectedDegradationId) {
      setError("Sélectionne une dégradation.");
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
      setNotice("Décision calculée et enregistrée dans l'historique.");
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
    setNotice(`Edition de la ligne ${getDisplayRowNumber(row)}.`);
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
        const currentRow = rows.find((row) => row.id === editingRowId);
        setNotice(`Ligne ${currentRow ? getDisplayRowNumber(currentRow) : editingRowId} mise à jour.`);
      } else {
        await padApi.createSheetRow(activeSheet.name, payload);
        setNotice("Nouvelle ligne ajoutee.");
      }

      await Promise.all([refreshStatus(), loadRows()]);
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
      const row = rows.find((item) => item.id === id);
      await padApi.deleteSheetRow(activeSheet.name, id);
      await Promise.all([refreshStatus(), loadRows()]);
      if (activeSheet.name === "Feuil4") {
        await loadFeuil4Snapshot();
      }
      if (editingRowId === id) {
        resetDraft();
      }
      setNotice(`Ligne ${row ? getDisplayRowNumber(row) : id} supprimée.`);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleClearHistory() {
    if (!window.confirm("Supprimer tout l'historique des décisions ?")) {
      return;
    }

    setIsBusy(true);
    try {
      await padApi.clearDecisionHistory();
      await Promise.all([refreshStatus(), loadHistory()]);
      setNotice("Historique vidé.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleAssignSolutionTemplate() {
    if (!selectedDegradation) {
      setError("Sélectionne une dégradation.");
      return;
    }
    if (!selectedTemplateKey) {
      setError("Sélectionne un modèle de solution.");
      return;
    }

    setIsSolutionBusy(true);
    try {
      await padApi.assignTemplateToDegradation(selectedDegradation.code, selectedTemplateKey);
      await Promise.all([refreshDecisionCatalogs(), loadSolutionTemplates()]);
      setNotice("Modèle de solution appliqué à la dégradation.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsSolutionBusy(false);
    }
  }

  async function handleSaveSolutionOverride() {
    if (!selectedDegradation) {
      setError("Sélectionne une dégradation.");
      return;
    }
    if (!solutionDraft.trim()) {
      setError("Saisis une solution personnalisée.");
      return;
    }

    setIsSolutionBusy(true);
    try {
      await padApi.setDegradationSolutionOverride(selectedDegradation.code, solutionDraft.trim());
      await refreshDecisionCatalogs();
      setNotice("Solution personnalisée enregistrée.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsSolutionBusy(false);
    }
  }

  async function handleClearSolutionOverride() {
    if (!selectedDegradation) {
      setError("Sélectionne une dégradation.");
      return;
    }

    setIsSolutionBusy(true);
    try {
      await padApi.clearDegradationSolutionOverride(selectedDegradation.code);
      await refreshDecisionCatalogs();
      setNotice("Solution personnalisée retirée (retour au modèle).");
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
    setNotice(`Édition règle assainissement #${rule.id}`);
  }

  async function handleSaveDrainageRule() {
    const parsedOrder = Number(drainageRuleOrder);
    if (!Number.isFinite(parsedOrder) || parsedOrder <= 0) {
      setError("Ordre de règle invalide (nombre > 0 attendu).");
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
      setNotice("Règle assainissement enregistrée.");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsDrainageRuleBusy(false);
    }
  }

  async function handleDeleteDrainageRule(ruleId?: number) {
    const id = Number(ruleId ?? editingDrainageRuleId);
    if (!Number.isFinite(id) || id <= 0) {
      setError("Sélectionne une règle assainissement à supprimer.");
      return;
    }

    if (!window.confirm("Supprimer cette règle assainissement ?")) {
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
      setNotice(`Règle assainissement #${id} supprimée.`);
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsDrainageRuleBusy(false);
    }
  }

  function handleEditMaintenance(intervention: MaintenanceInterventionItem) {
    setEditingMaintenanceId(intervention.id);
    setMaintenanceRoadId(intervention.roadId ?? "");
    setMaintenanceDegradationCode(intervention.degradationCode || "");
    setMaintenanceType(intervention.interventionType || "");
    setMaintenanceStatus(intervention.status);
    setMaintenanceDate(intervention.interventionDate || "");
    setMaintenanceCompletionDate(intervention.completionDate || "");
    setMaintenanceStateBefore(intervention.stateBefore || "");
    setMaintenanceStateAfter(intervention.stateAfter || "");
    setMaintenanceDeflectionBefore(
      intervention.deflectionBefore != null ? String(intervention.deflectionBefore) : ""
    );
    setMaintenanceDeflectionAfter(
      intervention.deflectionAfter != null ? String(intervention.deflectionAfter) : ""
    );
    setMaintenanceSolutionApplied(intervention.solutionApplied || "");
    setMaintenanceContractorName(intervention.contractorName || "");
    setMaintenanceObservation(intervention.observation || "");
    setMaintenanceCostAmount(intervention.costAmount != null ? String(intervention.costAmount) : "");
    setError("");
    setNotice(`Édition entretien #${intervention.id}`);
  }

  async function handleSaveMaintenance() {
    const parsedRoadId = Number(maintenanceRoadId);
    if (!Number.isFinite(parsedRoadId) || parsedRoadId <= 0) {
      setError("Sélectionne une voie pour l'entretien.");
      return;
    }
    if (!maintenanceType.trim()) {
      setError("Renseigne le type d'entretien.");
      return;
    }
    if (!maintenanceDate) {
      setError("Renseigne la date d'intervention.");
      return;
    }

    const payload: MaintenanceInterventionPayload = {
      id: editingMaintenanceId ?? undefined,
      roadId: parsedRoadId,
      degradationCode: maintenanceDegradationCode || undefined,
      interventionType: maintenanceType.trim(),
      status: maintenanceStatus,
      interventionDate: maintenanceDate,
      completionDate: maintenanceCompletionDate || undefined,
      stateBefore: maintenanceStateBefore.trim() || undefined,
      stateAfter: maintenanceStateAfter.trim() || undefined,
      deflectionBefore:
        maintenanceDeflectionBefore.trim() !== "" ? Number(maintenanceDeflectionBefore) : undefined,
      deflectionAfter:
        maintenanceDeflectionAfter.trim() !== "" ? Number(maintenanceDeflectionAfter) : undefined,
      solutionApplied: maintenanceSolutionApplied.trim() || undefined,
      contractorName: maintenanceContractorName.trim() || undefined,
      observation: maintenanceObservation.trim() || undefined,
      costAmount: maintenanceCostAmount.trim() !== "" ? Number(maintenanceCostAmount) : undefined
    };

    setIsMaintenanceBusy(true);
    try {
      await padApi.upsertMaintenanceIntervention(payload);
      await Promise.all([
        loadMaintenanceRows(),
        loadMaintenancePreview(selectedRoadId),
        refreshStatus()
      ]);
      resetMaintenanceForm();
      setError("");
      setNotice("Entretien enregistré.");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsMaintenanceBusy(false);
    }
  }

  async function handleDeleteMaintenance(interventionId?: number) {
    const id = Number(interventionId ?? editingMaintenanceId);
    if (!Number.isFinite(id) || id <= 0) {
      setError("Sélectionne un entretien à supprimer.");
      return;
    }
    if (!window.confirm("Supprimer cette intervention d'entretien ?")) {
      return;
    }

    setIsMaintenanceBusy(true);
    try {
      await padApi.deleteMaintenanceIntervention(id);
      await Promise.all([
        loadMaintenanceRows(),
        loadMaintenancePreview(selectedRoadId),
        refreshStatus()
      ]);
      if (editingMaintenanceId === id) {
        resetMaintenanceForm();
      }
      setError("");
      setNotice(`Entretien #${id} supprimé.`);
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsMaintenanceBusy(false);
    }
  }

  function exportHistoryCsv() {
    if (historyRows.length === 0) {
      setError("Aucune ligne d'historique à exporter.");
      return;
    }

    const headers = [
      "Date",
      "SAP",
      "Code voie",
      "Désignation voie",
      "Début",
      "Fin",
      "Dégradation",
      "Cause probable",
      "Solution maintenance",
      "Intervention contextuelle",
      "Déflexion D",
      "Sévérité",
      "Recommandation déflexion",
      "Alerte assainissement",
      "Recommandation assainissement"
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
          toDeflectionSeverityLabel(row.deflectionSeverity),
          toDeflectionRecommendationLabel(row.deflectionRecommendation),
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
          <h2>Aide à la décision maintenance</h2>
          <p className="muted">Sélectionne une voie, une dégradation et la valeur de déflexion (optionnelle).</p>

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
            <option value="">Sélectionner une voie</option>
            {roads.map((road) => (
              <option key={road.id} value={road.id}>
                {road.sapCode || "SAP?"} | {road.roadCode} | {road.designation}
              </option>
            ))}
          </select>

          <label htmlFor="degradation">Dégradation</label>
          <select
            id="degradation"
            value={selectedDegradationId}
            onChange={(event) => setSelectedDegradationId(event.target.value ? Number(event.target.value) : "")}
          >
            <option value="">Sélectionner une dégradation</option>
            {degradations.map((item) => (
              <option key={item.id} value={item.id}>
                {item.name}
              </option>
            ))}
          </select>

          <label htmlFor="deflection">Valeur de déflexion D</label>
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
              <h3 className="card__title--spaced">Résumé voie</h3>
              <p>
                <strong>SAP:</strong> {selectedRoad.sapCode || "-"}
              </p>
              <p>
                <strong>PK/borne début:</strong> {selectedRoad.startLabel || "-"}
              </p>
              <p>
                <strong>PK/borne fin:</strong> {selectedRoad.endLabel || "-"}
              </p>
              <p>
                <strong>Longueur:</strong> {selectedRoad.lengthM ?? "-"} m
              </p>
              <p>
                <strong>État chaussée:</strong> {selectedRoad.pavementState || "-"}
              </p>
              <p>
                <strong>État courant suivi:</strong> {latestMaintenance?.stateAfter || selectedRoad.pavementState || "-"}
              </p>
              <p>
                <strong>Assainissement (type caniveaux / description):</strong> {selectedRoad.drainageType || "-"} /{" "}
                {selectedRoad.drainageState || "-"}
              </p>
            </div>
          ) : null}

          {selectedRoad ? (
            <div className="card">
              <h3>Dernier entretien</h3>
              {latestMaintenance ? (
                <>
                  <p>
                    <strong>Date:</strong> {latestMaintenance.interventionDate}
                  </p>
                  <p>
                    <strong>Statut:</strong> {toMaintenanceStatusLabel(latestMaintenance.status)}
                  </p>
                  <p>
                    <strong>Type:</strong> {latestMaintenance.interventionType || "-"}
                  </p>
                  <p>
                    <strong>État après:</strong> {latestMaintenance.stateAfter || "-"}
                  </p>
                  <p>
                    <strong>Solution appliquée:</strong> {latestMaintenance.solutionApplied || "-"}
                  </p>
                </>
              ) : (
                <p className="muted">Aucun entretien encore enregistré pour cette voie.</p>
              )}
            </div>
          ) : null}

          {dynamicFeuil4Snapshot ? (
            <div className="card">
              <h3 className="card__title--spaced">Cockpit Decisionnel PAD</h3>
              <p>
                <strong>Domaine:</strong> {dynamicFeuil4Snapshot.domain}
              </p>
              <p>
                <strong>SAP:</strong> {dynamicFeuil4Snapshot.sapSector}
              </p>
              <p>
                <strong>Blvd/Rue:</strong> {dynamicFeuil4Snapshot.roadLabel} ({dynamicFeuil4Snapshot.roadMatch})
              </p>
              <p>
                <strong>PK début/fin:</strong> {dynamicFeuil4Snapshot.pkStart} / {dynamicFeuil4Snapshot.pkEnd}
              </p>
              <p>
                <strong>Observation:</strong> {dynamicFeuil4Snapshot.observation}
              </p>
              <p>
                <strong>Validation cause:</strong> {dynamicFeuil4Snapshot.causeMatch}
              </p>
              <p>
                <strong>D:</strong> {dynamicFeuil4Snapshot.deflectionValue} | <strong>État:</strong>{" "}
                {dynamicFeuil4Snapshot.severity}
              </p>
              <p>
                <strong>Intervention:</strong> {dynamicFeuil4Snapshot.recommendation}
              </p>
            </div>
          ) : null}
        </section>

        <section className="panel decision-output">
          <h2>Résultat automatique</h2>
          {!decisionResult ? <p className="muted">Lance une analyse pour afficher la recommandation.</p> : null}

          {decisionResult ? (
            <div className="decision-grid">
              <div className="card">
                <h3>Voie sélectionnée</h3>
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
                <h3>Dégradation</h3>
                <p>{decisionResult.degradation.name}</p>
                <p>
                  <strong>Cause probable:</strong> {decisionResult.probableCause}
                </p>
              </div>

              <div className="card">
                <h3>Déflexion</h3>
                <p>
                  <strong>D:</strong> {decisionResult.deflection.value ?? "non renseigné"}
                </p>
                <p>
                  <strong>État:</strong> {toDeflectionSeverityLabel(decisionResult.deflection.severity)}
                </p>
                <p>
                  <strong>Intervention:</strong> {toDeflectionRecommendationLabel(decisionResult.deflection.recommendation)}
                </p>
              </div>

              <div
                className={`card card--full ${
                  decisionResult.degradation.solutionSource === "MISSING" ? "card--solution-missing" : ""
                }`}
              >
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
                <h3>Intervention contextuelle tronçon</h3>
                <p>{decisionResult.contextualIntervention || "à déterminer (A D)"}</p>
              </div>

              <div className="card card--full">
                <h3>Causes connues (catalogue)</h3>
                <ul>
                  {decisionResult.degradation.causes.length > 0 ? (
                    decisionResult.degradation.causes.map((cause, index) => (
                      <li key={`${decisionResult.degradation.id}-${index}`}>{cause}</li>
                    ))
                  ) : (
                    <li>Aucune cause détaillée en base pour cette dégradation.</li>
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
          <p className="muted">Référence complète des voies par secteur SAP.</p>

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
                  <th>Début</th>
                  <th>Fin</th>
                  <th>Longueur (m)</th>
                  <th>Largeur (m)</th>
                  <th>Revetement</th>
                  <th>État chaussée</th>
                  <th>Type caniveaux</th>
                  <th>Description assain.</th>
                </tr>
              </thead>
              <tbody>
                {roads.map((road) => (
                  <tr key={road.id}>
                    <td>
                      <button
                        className="row-action row-action--evaluate row-action--with-icon"
                        type="button"
                        onClick={() => {
                          setSelectedRoadId(road.id);
                          setActiveView("decision");
                        }}
                      >
                        <Gauge size={15} aria-hidden="true" />
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
                    <td colSpan={12}>Aucune voie trouvée.</td>
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
          <h2>Catalogue des dégradations</h2>
          <p className="muted">Liste des dégradations, causes probables et solution de maintenance.</p>

          <div className="table-toolbar table-toolbar--triple">
            <input
              value={degradationSearch}
              onChange={(event) => setDegradationSearch(event.target.value)}
              placeholder="Recherche dégradation ou cause"
            />
            <span className="muted">{filteredDegradations.length} dégradation(s)</span>
            <span className="muted">{degradations.length} total</span>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Action</th>
                  <th>Dégradation</th>
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
                          className="row-action row-action--configure row-action--with-icon"
                          type="button"
                          onClick={() => {
                            setSelectedDegradationId(item.id);
                          }}
                        >
                          <Settings2 size={15} aria-hidden="true" />
                          Configurer
                        </button>
                        <button
                          className="row-action row-action--use"
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
                    <td colSpan={6}>Aucune dégradation trouvée.</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>

          {selectedDegradation ? (
            <div className="card card--spaced">
              <h3>Causes détaillées: {selectedDegradation.name}</h3>
              <p>
                <strong>Source solution:</strong> {toSolutionSourceLabel(selectedDegradation.solutionSource)}
              </p>
              <ul>
                {selectedDegradation.causes.length > 0 ? (
                  selectedDegradation.causes.map((cause, index) => <li key={`${selectedDegradation.id}-cause-${index}`}>{cause}</li>)
                ) : (
                  <li>Aucune cause détaillée pour cette dégradation.</li>
                )}
              </ul>

              <h3>Paramétrage solution de maintenance</h3>
              <label htmlFor="template-key">Modèle de solution</label>
              <select
                id="template-key"
                value={selectedTemplateKey}
                onChange={(event) => setSelectedTemplateKey(event.target.value)}
                disabled={isSolutionBusy}
              >
                <option value="">Sélectionner un modèle</option>
                {solutionTemplates.map((template) => (
                  <option key={template.templateKey} value={template.templateKey}>
                    {template.title}
                  </option>
                ))}
              </select>

              <div className="editor-actions">
                <button className="row-action" type="button" onClick={handleAssignSolutionTemplate} disabled={isSolutionBusy}>
                  Appliquer le modèle
                </button>
              </div>

              <label htmlFor="solution-override">Solution personnalisée</label>
              <textarea
                id="solution-override"
                className="input-textarea"
                value={solutionDraft}
                onChange={(event) => setSolutionDraft(event.target.value)}
                placeholder="Saisir la solution de maintenance spécifique à cette dégradation"
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
            <h3>Moteur assainissement (règles)</h3>
            <p className="muted">Règles utilisées automatiquement pendant l'évaluation de décision.</p>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Actions</th>
                    <th>Ordre</th>
                    <th>Opérateur</th>
                    <th>Pattern</th>
                    <th>Question décideur</th>
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
                          <button
                            className="row-action row-action--icon"
                            type="button"
                            onClick={() => handleEditDrainageRule(rule)}
                            title="Éditer"
                            aria-label="Éditer"
                          >
                            <Pencil size={16} aria-hidden="true" />
                          </button>
                          <button
                            className="row-action row-action--danger row-action--icon"
                            type="button"
                            onClick={() => handleDeleteDrainageRule(rule.id)}
                            title="Supprimer"
                            aria-label="Supprimer"
                          >
                            <Trash2 size={16} aria-hidden="true" />
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
                      <td colSpan={8}>Aucune règle assainissement.</td>
                    </tr>
                  ) : null}
                </tbody>
              </table>
            </div>

            <h3 className="drainage-editor-title">
              {editingDrainageRuleId ? `Édition règle #${editingDrainageRuleId}` : "Nouvelle règle assainissement"}
            </h3>
            <div className="cells-grid">
              <div className="cell-field">
                <label htmlFor="dr-rule-order">Ordre de règle</label>
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
                <label htmlFor="dr-rule-operator">Opérateur</label>
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
                  placeholder={drainageRuleOperator === "ALWAYS" ? "Non utilisé avec ALWAYS" : "Ex: OBSTR"}
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
              Règle appliquée seulement si "Interroger explicitement l'assainissement" est coché
            </label>

            <label className="checkbox-row" htmlFor="dr-rule-needs-attention">
              <input
                id="dr-rule-needs-attention"
                type="checkbox"
                checked={drainageRuleNeedsAttention}
                onChange={(event) => setDrainageRuleNeedsAttention(event.target.checked)}
                disabled={isDrainageRuleBusy}
              />
              Déclenche une alerte assainissement
            </label>

            <label className="checkbox-row" htmlFor="dr-rule-active">
              <input
                id="dr-rule-active"
                type="checkbox"
                checked={drainageRuleActive}
                onChange={(event) => setDrainageRuleActive(event.target.checked)}
                disabled={isDrainageRuleBusy}
              />
              Règle active
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
                {isDrainageRuleBusy ? "Enregistrement..." : "Enregistrer règle"}
              </button>
              <button className="row-action" type="button" onClick={resetDrainageRuleEditor} disabled={isDrainageRuleBusy}>
                Réinitialiser
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

  function renderMaintenanceView() {
    return (
      <main className="workspace">
        <section className="panel editor-panel">
          <h2>Suivi des entretiens</h2>
          <p className="muted">Enregistre les interventions réalisées pour garder un état dynamique du réseau.</p>

          <label htmlFor="maintenance-road">Voie</label>
          <select
            id="maintenance-road"
            value={maintenanceRoadId}
            onChange={(event) => setMaintenanceRoadId(event.target.value ? Number(event.target.value) : "")}
            disabled={isMaintenanceBusy}
          >
            <option value="">Sélectionner une voie</option>
            {allRoads.map((road) => (
              <option key={road.id} value={road.id}>
                {road.sapCode || "SAP?"} | {road.roadCode} | {road.designation}
              </option>
            ))}
          </select>

          <label htmlFor="maintenance-degradation">Dégradation concernée</label>
          <select
            id="maintenance-degradation"
            value={maintenanceDegradationCode}
            onChange={(event) => setMaintenanceDegradationCode(event.target.value)}
            disabled={isMaintenanceBusy}
          >
            <option value="">Aucune dégradation précise</option>
            {degradations.map((item) => (
              <option key={item.code} value={item.code}>
                {item.name}
              </option>
            ))}
          </select>

          <label htmlFor="maintenance-type">Type d'entretien</label>
          <input
            id="maintenance-type"
            list="maintenance-type-options"
            value={maintenanceType}
            onChange={(event) => setMaintenanceType(event.target.value)}
            placeholder="Ex: Curage caniveaux"
            disabled={isMaintenanceBusy}
          />
          <datalist id="maintenance-type-options">
            {MAINTENANCE_TYPE_SUGGESTIONS.map((item) => (
              <option key={item} value={item} />
            ))}
          </datalist>

          <label htmlFor="maintenance-status">Statut</label>
          <select
            id="maintenance-status"
            value={maintenanceStatus}
            onChange={(event) => setMaintenanceStatus(event.target.value as MaintenanceInterventionStatus)}
            disabled={isMaintenanceBusy}
          >
            {MAINTENANCE_STATUSES.map((item) => (
              <option key={item} value={item}>
                {toMaintenanceStatusLabel(item)}
              </option>
            ))}
          </select>

          <div className="maintenance-form-grid">
            <div className="cell-field">
              <label htmlFor="maintenance-date">Date d'intervention</label>
              <input
                id="maintenance-date"
                type="date"
                value={maintenanceDate}
                onChange={(event) => setMaintenanceDate(event.target.value)}
                disabled={isMaintenanceBusy}
              />
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-completion-date">Date de clôture</label>
              <input
                id="maintenance-completion-date"
                type="date"
                value={maintenanceCompletionDate}
                onChange={(event) => setMaintenanceCompletionDate(event.target.value)}
                disabled={isMaintenanceBusy}
              />
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-state-before">État avant</label>
              <input
                id="maintenance-state-before"
                value={maintenanceStateBefore}
                onChange={(event) => setMaintenanceStateBefore(event.target.value)}
                placeholder="État observé avant intervention"
                disabled={isMaintenanceBusy}
              />
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-state-after">État après</label>
              <input
                id="maintenance-state-after"
                value={maintenanceStateAfter}
                onChange={(event) => setMaintenanceStateAfter(event.target.value)}
                placeholder="État après travaux"
                disabled={isMaintenanceBusy}
              />
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-deflection-before">Déflexion avant (D)</label>
              <input
                id="maintenance-deflection-before"
                type="number"
                value={maintenanceDeflectionBefore}
                onChange={(event) => setMaintenanceDeflectionBefore(event.target.value)}
                disabled={isMaintenanceBusy}
              />
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-deflection-after">Déflexion après (D)</label>
              <input
                id="maintenance-deflection-after"
                type="number"
                value={maintenanceDeflectionAfter}
                onChange={(event) => setMaintenanceDeflectionAfter(event.target.value)}
                disabled={isMaintenanceBusy}
              />
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-contractor">Prestataire / équipe</label>
              <input
                id="maintenance-contractor"
                value={maintenanceContractorName}
                onChange={(event) => setMaintenanceContractorName(event.target.value)}
                placeholder="Ex: Équipe interne"
                disabled={isMaintenanceBusy}
              />
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-cost">Coût estimé</label>
              <input
                id="maintenance-cost"
                type="number"
                value={maintenanceCostAmount}
                onChange={(event) => setMaintenanceCostAmount(event.target.value)}
                placeholder="FCFA"
                disabled={isMaintenanceBusy}
              />
            </div>
          </div>

          <label htmlFor="maintenance-solution">Solution appliquée</label>
          <textarea
            id="maintenance-solution"
            className="input-textarea"
            value={maintenanceSolutionApplied}
            onChange={(event) => setMaintenanceSolutionApplied(event.target.value)}
            placeholder="Solution de maintenance effectivement appliquée"
            rows={4}
            disabled={isMaintenanceBusy}
          />

          <label htmlFor="maintenance-observation">Observation</label>
          <textarea
            id="maintenance-observation"
            className="input-textarea"
            value={maintenanceObservation}
            onChange={(event) => setMaintenanceObservation(event.target.value)}
            placeholder="Commentaires de terrain, resultat, contraintes..."
            rows={4}
            disabled={isMaintenanceBusy}
          />

          <div className="editor-actions">
            <button className="primary" type="button" onClick={handleSaveMaintenance} disabled={isMaintenanceBusy}>
              {isMaintenanceBusy ? "Enregistrement..." : "Enregistrer entretien"}
            </button>
            <button className="row-action" type="button" onClick={resetMaintenanceForm} disabled={isMaintenanceBusy}>
              Réinitialiser
            </button>
            <button
              className="row-action row-action--danger"
              type="button"
              onClick={() => handleDeleteMaintenance()}
              disabled={isMaintenanceBusy || !editingMaintenanceId}
            >
              Supprimer
            </button>
          </div>

          {maintenanceSelectedRoad ? (
            <div className="card card--spaced">
              <h3>Voie suivie</h3>
              <p>
                <strong>{maintenanceSelectedRoad.designation}</strong> ({maintenanceSelectedRoad.roadCode})
              </p>
              <p>
                <strong>SAP:</strong> {maintenanceSelectedRoad.sapCode || "-"}
              </p>
              <p>
                <strong>État importé:</strong> {maintenanceSelectedRoad.pavementState || "-"}
              </p>
            </div>
          ) : null}
        </section>

        <section className="panel table-panel">
          <h2>Historique des entretiens</h2>
          <p className="muted">Suivi chronologique des interventions prévues, en cours et terminées.</p>

          <div className="table-toolbar table-toolbar--penta">
            <select value={maintenanceSap} onChange={(event) => setMaintenanceSap(event.target.value)}>
              <option value="">Tous les secteurs</option>
              {sapSectors.map((sector) => (
                <option key={sector.code} value={sector.code}>
                  {sector.code}
                </option>
              ))}
            </select>
            <select
              value={maintenanceRoadFilterId}
              onChange={(event) => setMaintenanceRoadFilterId(event.target.value ? Number(event.target.value) : "")}
            >
              <option value="">Toutes les voies</option>
              {maintenanceFilterRoadOptions.map((road) => (
                <option key={`filter-${road.id}`} value={road.id}>
                  {road.roadCode} | {road.designation}
                </option>
              ))}
            </select>
            <select
              value={maintenanceStatusFilter}
              onChange={(event) =>
                setMaintenanceStatusFilter(event.target.value as MaintenanceInterventionStatus | "")
              }
            >
              <option value="">Tous statuts</option>
              {MAINTENANCE_STATUSES.map((status) => (
                <option key={status} value={status}>
                  {toMaintenanceStatusLabel(status)}
                </option>
              ))}
            </select>
            <input
              value={maintenanceSearch}
              onChange={(event) => setMaintenanceSearch(event.target.value)}
              placeholder="Recherche voie, type, solution, observation"
            />
            <button
              className="row-action row-action--with-icon"
              type="button"
              onClick={() => loadMaintenanceRows()}
              disabled={isMaintenanceLoading}
            >
              <RefreshCw size={15} aria-hidden="true" />
              <span>Actualiser</span>
            </button>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Actions</th>
                  <th>Date</th>
                  <th>Statut</th>
                  <th>SAP</th>
                  <th>Voie</th>
                  <th>Type</th>
                  <th>Dégradation</th>
                  <th>État avant</th>
                  <th>État après</th>
                  <th>Solution appliquée</th>
                  <th>Prestataire</th>
                  <th>Coût</th>
                </tr>
              </thead>
              <tbody>
                {maintenanceRows.map((item) => (
                  <tr key={item.id} className={editingMaintenanceId === item.id ? "is-selected" : ""}>
                    <td>
                      <div className="row-buttons">
                        <button
                          className="row-action row-action--icon"
                          type="button"
                          onClick={() => handleEditMaintenance(item)}
                          title="Éditer"
                          aria-label="Éditer"
                        >
                          <Pencil size={16} aria-hidden="true" />
                        </button>
                        <button
                          className="row-action row-action--danger row-action--icon"
                          type="button"
                          onClick={() => handleDeleteMaintenance(item.id)}
                          title="Supprimer"
                          aria-label="Supprimer"
                        >
                          <Trash2 size={16} aria-hidden="true" />
                        </button>
                      </div>
                    </td>
                    <td>{item.interventionDate}</td>
                    <td>
                      <span className={`status-chip status-chip--${item.status.toLowerCase()}`}>
                        {toMaintenanceStatusLabel(item.status)}
                      </span>
                    </td>
                    <td>{item.sapCode || "-"}</td>
                    <td>
                      {item.roadCode} - {item.roadDesignation}
                    </td>
                    <td>{item.interventionType}</td>
                    <td>{item.degradationName || "-"}</td>
                    <td>{item.stateBefore || "-"}</td>
                    <td>{item.stateAfter || "-"}</td>
                    <td>{item.solutionApplied || "-"}</td>
                    <td>{item.contractorName || "-"}</td>
                    <td>{formatAmount(item.costAmount)}</td>
                  </tr>
                ))}
                {maintenanceRows.length === 0 ? (
                  <tr>
                    <td colSpan={12}>{isMaintenanceLoading ? "Chargement..." : "Aucun entretien enregistré."}</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
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
          <p className="muted">Historique des décisions calculées avec export CSV.</p>

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
              placeholder="Recherche voie/dégradation/cause"
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
                  <th>Dégradation</th>
                  <th>Cause probable</th>
                  <th>D</th>
                  <th>Sévérité</th>
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
                    <td>{toDeflectionSeverityLabel(row.deflectionSeverity)}</td>
                    <td>{toDeflectionRecommendationLabel(row.deflectionRecommendation)}</td>
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
          <p className="muted">{activeSheet?.description ?? "Sélectionne une feuille"}</p>

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
              Réinitialiser
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
                      <th rowSpan={3}>N° ligne</th>
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
                    <th>N° ligne</th>
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
                        <button
                          className="row-action row-action--icon"
                          type="button"
                          onClick={() => handleEdit(row)}
                          title="Éditer"
                          aria-label="Éditer"
                        >
                          <Pencil size={16} aria-hidden="true" />
                        </button>
                        <button
                          className="row-action row-action--danger row-action--icon"
                          type="button"
                          onClick={() => handleDeleteRow(row.id)}
                          title="Supprimer"
                          aria-label="Supprimer"
                        >
                          <Trash2 size={16} aria-hidden="true" />
                        </button>
                      </div>
                    </td>
                    <td>{getDisplayRowNumber(row)}</td>
                    {activeColumns.map((column) => {
                      const rawValue = String(row[column] ?? "");
                      const isFeuil3Intervention = activeSheet?.name === "Feuil3" && column === "L";
                      const value = isFeuil3Intervention && !rawValue.trim() ? "à déterminer (A D)" : rawValue;
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
            <h1>PAD Outil de Maintenance Routière</h1>
            <p>Pilotez la maintenance routière du PAD avec des décisions rapides et fiables.</p>
          </div>
        </div>

        <div className="hero__status">
          <span className="pill">Total lignes: {status?.totalRows ?? "-"}</span>
          <span className="pill">Voies: {roads.length}</span>
          <span className="pill">Dégradations: {degradations.length}</span>
          <span className="pill">Historique: {status?.decisionHistoryCount ?? 0}</span>
          <span className="pill">Dernier import: {status?.lastImportAt ?? "-"}</span>
        </div>

        <div className="hero__actions">
          <input
            className="hero-input hero-input--path"
            value={importPath}
            onChange={(event) => setImportPath(event.target.value)}
            placeholder="Chemin du fichier Excel"
          />
          <button className="icon-btn" type="button" onClick={handlePickImportPath} disabled={isBusy}>
            <FolderOpen size={16} aria-hidden="true" />
            <span>Parcourir</span>
          </button>
          <button className="icon-btn" type="button" onClick={handleImport} disabled={isBusy}>
            <Upload size={16} aria-hidden="true" />
            <span>Importer</span>
          </button>
          <button className="icon-btn" type="button" onClick={handleRefresh} disabled={isBusy}>
            <RefreshCw size={16} aria-hidden="true" />
            <span>Actualiser</span>
          </button>
        </div>

        {error ? <p className="hero__error">{error}</p> : null}
        {notice ? <p className="hero__notice">{notice}</p> : null}
        <nav className="hero__nav">
          <button className={activeView === "decision" ? "active" : ""} type="button" onClick={() => setActiveView("decision")}>
            Aide à la décision
          </button>
          <button className={activeView === "catalogue" ? "active" : ""} type="button" onClick={() => setActiveView("catalogue")}>
            Catalogue
          </button>
          <button
            className={activeView === "degradations" ? "active" : ""}
            type="button"
            onClick={() => setActiveView("degradations")}
          >
            Dégradations
          </button>
          <button className={activeView === "maintenance" ? "active" : ""} type="button" onClick={() => setActiveView("maintenance")}>
            Suivi
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
      {activeView === "maintenance" ? renderMaintenanceView() : null}
      {activeView === "history" ? renderHistoryView() : null}
      {activeView.startsWith("sheet:") ? renderSheetView() : null}

      <footer className="app-footer">
        <span>{appName}</span>
        <span>Version {appVersion}</span>
      </footer>
    </div>
  );
}
