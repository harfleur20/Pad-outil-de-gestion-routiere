import { useCallback, useEffect, useMemo, useRef, useState, type ChangeEvent } from "react";
import { padApi } from "./lib/pad-api";
import {
  Activity,
  BarChart3,
  BookOpen,
  CalendarRange,
  ChevronDown,
  ChevronUp,
  CircleHelp,
  ClipboardPlus,
  Calculator,
  DatabaseBackup,
  Eye,
  ExternalLink,
  FileSpreadsheet,
  FolderOpen,
  Gauge,
  History,
  Layers3,
  Map as MapIcon,
  Route,
  Paperclip,
  Pencil,
  Plus,
  Printer,
  RefreshCw,
  ShieldCheck,
  TriangleAlert,
  Trash2,
  Upload,
  X
} from "lucide-react";
import type {
  DataIntegrityReport,
  DataStatus,
  DashboardSummary,
  DecisionHistoryItem,
  DecisionResult,
  DegradationItem,
  ImportPreview,
  MeasurementCampaignItem,
  MaintenanceInterventionItem,
  MaintenanceInterventionPayload,
  MaintenanceInterventionStatus,
  MaintenanceSolutionTemplate,
  RoadCatalogItem,
  RoadSectionItem,
  RoadMeasurementItem,
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

function getEditableColumns(sheet: SheetDefinition | null): SheetColumnKey[] {
  if (!sheet) {
    return [];
  }
  if (sheet.name === "Feuil2") {
    return ["A", "B", "C", "D", "E", "F", "G"];
  }
  if (sheet.name === "Feuil6") {
    return ["A", "B", "C", "D", "E", "F", "G"];
  }
  return sheet.columns;
}

function getSheetFieldPlaceholder(sheetName: string | undefined, column: SheetColumnKey) {
  if (sheetName === "Feuil2") {
    if (column === "A") return "Calculé automatiquement";
    if (column === "B") return "Ex: 1_7";
    if (column === "C") return "Ex: Rue 01 ou Bvd 02";
    if (column === "D") return "Ex: Rue des Archives";
    if (column === "E") return "Ex: Dangote";
    if (column === "F") return "Ex: Quai 60";
    if (column === "G") return "Ex: 700";
  }
  if (sheetName === "Feuil3") {
    if (column === "A") return "Ex: Rue 01";
    if (column === "B") return "Ex: Rue des Archives";
    if (column === "C") return "Ex: Dangote";
    if (column === "D") return "Ex: Quai 60";
    if (column === "E") return "Ex: 700";
    if (column === "F") return "Ex: 7";
    if (column === "G") return "Ex: BB";
    if (column === "H") return "Ex: Bon, Moyen, Mauvais";
    if (column === "I") return "Ex: E,F ou C";
    if (column === "J") return "Ex: Obstrué par déchets et sable";
    if (column === "K") return "Ex: 1,5";
    if (column === "L") return "Ex: Curage de caniveaux";
  }
  if (sheetName === "Feuil5") {
    if (column === "A") return "Ex: 1";
    if (column === "B") return "Ex: 1_7";
    if (column === "C") return "Ex: Rue 01";
    if (column === "D") return "Ex: Rue des Archives";
    if (column === "E") return "Ex: Dangote";
    if (column === "F") return "Ex: Quai 60";
    if (column === "G") return "Ex: 700";
    if (column === "H") return "Ex: 7";
    if (column === "I") return "Ex: BB";
    if (column === "J") return "Ex: Bon, Moyen, Mauvais";
    if (column === "K") return "Ex: E,F ou C";
    if (column === "L") return "Ex: Bon ou Obstrué";
    if (column === "M") return "Ex: 1,5";
    if (column === "N") return "Ex: 1";
    if (column === "O") return "Ex: 1";
    if (column === "P") return "Ex: Stationnement poids lourds";
  }
  if (sheetName === "Feuil6") {
    if (column === "A") return "Ex: SAP4";
    if (column === "B") return "Ex: Rue, Boulevard ou Avenue";
    if (column === "C") return "Ex: Rue.01";
    if (column === "D") return "Ex: 700";
    if (column === "E") return "Ex: Rue des Archives";
    if (column === "F") return "Ex: Dangote à Quai 60";
    if (column === "G") return "Ex: Nom choisi à cause de l'activité de la zone";
  }
  if (sheetName === "Feuil4") {
    if (column === "A") return "Ex: Domaine";
    if (column === "B") return "Ex: Port de Douala Bonabéri";
    if (column === "C") return "Ex: FORT";
    if (column === "D") return "Ex: Renforcement lourd";
    if (column === "E") return "Ex: Observations";
    if (column === "F") return "Ex: VRAI ou FAUX";
  }
  if (sheetName === "Feuil7") {
    if (column === "A") return "Ex: Chaussée";
    if (column === "B") return "Ex: D1";
    if (column === "C") return "Ex: Nids de poule";
    if (column === "D") return "Ex: Dégradations de surface";
    if (column === "E") return "Ex: Déformation";
    if (column === "F") return "Ex: Observation terrain";
    if (column === "G") return "Ex: Présence d'eau stagnante";
  }
  return "";
}

function getSheetFieldHelpText(sheetName: string | undefined, column: SheetColumnKey) {
  if (sheetName === "Feuil2") {
    if (column === "A") {
      return "Le tronçon est déduit automatiquement depuis le numéro de section. Exemple : 1_7 donne le tronçon 7.";
    }
    if (column === "B") {
      return "Écrivez le numéro de section sous la forme SAP_tronçon. Exemple : 1_7 = SAP1, tronçon 7. Le préfixe SAP doit correspondre au SAP de la voie.";
    }
    if (column === "C") {
      return "Écrivez le code de la voie, par exemple Rue 01 ou Bvd 02.";
    }
    if (column === "G") {
      return "Écrivez la longueur en mètres, par exemple 700.";
    }
  }
  if (sheetName === "Feuil3") {
    if (column === "A") return "Code court de la voie.";
    if (column === "B") return "Nom complet de la voie.";
    if (column === "C") return "Nom du lieu où la section commence. Ce n'est pas une date.";
    if (column === "D") return "Nom du lieu où la section se termine. Ce n'est pas une date.";
    if (column === "E") return "Longueur de la section en mètres.";
    if (column === "F") return "Largeur minimale en mètres.";
    if (column === "G") return "Type de revêtement observé sur la chaussée.";
    if (column === "H") return "État général de la chaussée.";
    if (column === "I") return "Type de caniveaux ou d'assainissement.";
    if (column === "J") return "Décrivez l'état de l'assainissement avec des mots simples.";
    if (column === "K") return "Largeur minimale des trottoirs en mètres.";
    if (column === "L") return "Travaux ou entretien à prévoir sur cette voie.";
  }
  if (sheetName === "Feuil5") {
    if (column === "A") return "Le tronçon est repris automatiquement depuis le numéro de section.";
    if (column === "B") return "Numéro de section sous la forme SAP_tronçon, par exemple 1_7 ou 4_12. Le préfixe SAP doit correspondre au SAP de la voie.";
    if (column === "C") return "Code court de la voie.";
    if (column === "D") return "Nom complet de la voie.";
    if (column === "E") return "Nom du lieu où la section commence. Ce n'est pas une date.";
    if (column === "F") return "Nom du lieu où la section se termine. Ce n'est pas une date.";
    if (column === "G") return "Longueur de la section en mètres.";
    if (column === "H") return "Largeur minimale en mètres.";
    if (column === "K") return "Type de caniveaux ou d'assainissement.";
    if (column === "L") return "État de l'assainissement.";
    if (column === "M") return "Largeur minimale des trottoirs en mètres.";
    if (column === "N") return "Valeur de stationnement du côté gauche.";
    if (column === "O") return "Valeur de stationnement du côté droit.";
    if (column === "P") return "Autre information utile sur le stationnement ou l'occupation latérale.";
  }
  if (sheetName === "Feuil6") {
    if (column === "A") return "Choisissez le SAP de rattachement de cette voie, par exemple SAP1 ou SAP4.";
    if (column === "B") return "Choisissez le type de voie : Rue, Boulevard ou Avenue.";
    if (column === "C") return "Code court de la voie.";
    if (column === "D") return "Longueur totale en mètres.";
    if (column === "E") return "Nom proposé pour cette voie.";
    if (column === "F") return "Écrivez le début et la fin sous la forme Dangote à Quai 60.";
    if (column === "G") return "Expliquez simplement pourquoi ce nom a été choisi.";
  }
  if (sheetName === "Feuil4") {
    if (column === "A") return "Nom de l'information affichée dans le programme d'évaluation.";
    if (column === "B") return "Valeur à afficher ou à utiliser pour cette ligne.";
    if (column === "C") return "Résultat de l'évaluation de l'état de la chaussée.";
    if (column === "D") return "Travaux ou entretien recommandés.";
    if (column === "E") return "Zone de contrôle ou zone d'observation.";
    if (column === "F") return "Résultat calculé ou indicateur simple, par exemple VRAI ou FAUX.";
  }
  if (sheetName === "Feuil7") {
    if (column === "A") return "Grande catégorie de la dégradation.";
    if (column === "B") return "Code ou référence courte de la dégradation.";
    if (column === "C") return "Nom de la dégradation observée.";
    if (column === "D") return "Famille principale de la dégradation.";
    if (column === "E") return "Sous-famille ou précision complémentaire.";
    if (column === "F") return "Observation utile pour mieux comprendre cette dégradation.";
    if (column === "G") return "Cause probable de cette dégradation.";
  }
  return "";
}

function getRequiredSheetColumns(sheetName: string | undefined): SheetColumnKey[] {
  if (sheetName === "Feuil2") {
    return ["B", "C"];
  }
  if (sheetName === "Feuil3") {
    return ["A", "F", "G", "H", "I", "J", "L"];
  }
  if (sheetName === "Feuil5") {
    return ["C", "H", "K", "L", "M"];
  }
  if (sheetName === "Feuil6") {
    return ["A", "B", "C", "D", "E", "F", "G"];
  }
  if (sheetName === "Feuil7") {
    return ["A", "B", "C", "G"];
  }
  if (sheetName === "Feuil4") {
    return ["A"];
  }
  return [];
}

function isSheetFieldRequired(sheetName: string | undefined, column: SheetColumnKey) {
  return getRequiredSheetColumns(sheetName).includes(column);
}

function getSheetFieldRequiredMessage(sheetName: string | undefined, column: SheetColumnKey) {
  if (sheetName === "Feuil2") {
    if (column === "B") return "Veuillez renseigner le numéro de section.";
    if (column === "C") return "Veuillez choisir la voie concernée.";
  }
  if (sheetName === "Feuil3") {
    if (column === "A") return "Veuillez choisir la voie à diagnostiquer.";
    if (column === "F") return "Veuillez renseigner la largeur minimale côté façade.";
    if (column === "G") return "Veuillez renseigner le type de revêtement.";
    if (column === "H") return "Veuillez renseigner l'état de la chaussée.";
    if (column === "I") return "Veuillez renseigner le type de caniveaux.";
    if (column === "J") return "Veuillez renseigner l'état de l'assainissement.";
    if (column === "L") return "Veuillez renseigner l'intervention à prévoir.";
  }
  if (sheetName === "Feuil5") {
    if (column === "C") return "Veuillez choisir la voie à compléter.";
    if (column === "H") return "Veuillez renseigner la largeur minimale côté façade.";
    if (column === "K") return "Veuillez renseigner le type d'assainissement.";
    if (column === "L") return "Veuillez renseigner l'état de l'assainissement.";
    if (column === "M") return "Veuillez renseigner la largeur minimale des trottoirs.";
  }
  if (sheetName === "Feuil6") {
    if (column === "A") return "Veuillez choisir le SAP.";
    if (column === "B") return "Veuillez renseigner le type de voie.";
    if (column === "C") return "Veuillez renseigner le code de la voie.";
    if (column === "D") return "Veuillez renseigner le linéaire en mètres.";
    if (column === "E") return "Veuillez renseigner le nom proposé.";
    if (column === "F") return "Veuillez renseigner le début et la fin de la voie.";
    if (column === "G") return "Veuillez renseigner la justification.";
  }
  if (sheetName === "Feuil7") {
    if (column === "A") return "Veuillez renseigner la catégorie.";
    if (column === "B") return "Veuillez renseigner la référence.";
    if (column === "C") return "Veuillez renseigner le nom de la dégradation.";
    if (column === "G") return "Veuillez renseigner la cause probable.";
  }
  if (sheetName === "Feuil4") {
    if (column === "A") return "Veuillez renseigner le nom de la ligne.";
  }
  return "Veuillez remplir ce champ.";
}

function setSheetFieldError(
  errors: Partial<Record<SheetColumnKey, string>>,
  column: SheetColumnKey,
  message: string
) {
  if (!errors[column]) {
    errors[column] = message;
  }
}

function validateSheetDraft(
  sheet: SheetDefinition | null,
  cells: Partial<Record<SheetColumnKey, string>>
) {
  const fieldErrors: Partial<Record<SheetColumnKey, string>> = {};
  let formError = "";

  if (!sheet) {
    return { fieldErrors, formError };
  }

  for (const column of getRequiredSheetColumns(sheet.name)) {
    if (!String(cells[column] ?? "").trim()) {
      setSheetFieldError(fieldErrors, column, getSheetFieldRequiredMessage(sheet.name, column));
    }
  }

  if (sheet.name === "Feuil2") {
    const sectionNo = String(cells.B ?? "").trim();
    const lengthM = parseNumberValue(cells.G);
    if (sectionNo && !/^[1-9][0-9]*_[1-9][0-9]*$/.test(sectionNo)) {
      setSheetFieldError(
        fieldErrors,
        "B",
        "Écrivez par exemple 1_7 ou 4_12. Le nombre avant _ indique le SAP et le nombre après _ devient le tronçon."
      );
    }
    if (String(cells.G ?? "").trim() && (!Number.isFinite(lengthM) || Number(lengthM) <= 0)) {
      setSheetFieldError(fieldErrors, "G", "La longueur doit être un nombre supérieur à 0.");
    }
  }

  if (sheet.name === "Feuil3") {
    const lengthM = parseNumberValue(cells.E);
    const facadeWidthM = parseNumberValue(cells.F);
    const sidewalkWidthM = parseNumberValue(cells.K);
    if (String(cells.E ?? "").trim() && (!Number.isFinite(lengthM) || Number(lengthM) <= 0)) {
      setSheetFieldError(fieldErrors, "E", "La longueur doit être un nombre supérieur à 0.");
    }
    if (String(cells.F ?? "").trim() && (!Number.isFinite(facadeWidthM) || Number(facadeWidthM) <= 0)) {
      setSheetFieldError(fieldErrors, "F", "La largeur côté façade doit être un nombre supérieur à 0.");
    }
    if (String(cells.K ?? "").trim() && (!Number.isFinite(sidewalkWidthM) || Number(sidewalkWidthM) < 0)) {
      setSheetFieldError(fieldErrors, "K", "La largeur des trottoirs doit être un nombre positif ou nul.");
    }
  }

  if (sheet.name === "Feuil5") {
    const numericColumns: Array<[SheetColumnKey, string, boolean]> = [
      ["H", "La largeur côté façade doit être un nombre supérieur à 0.", true],
      ["M", "La largeur des trottoirs doit être un nombre positif ou nul.", false],
      ["N", "La valeur du stationnement à gauche doit être un nombre positif ou nul.", false],
      ["O", "La valeur du stationnement à droite doit être un nombre positif ou nul.", false]
    ];

    for (const [column, message, strictPositive] of numericColumns) {
      const rawValue = String(cells[column] ?? "").trim();
      const numericValue = parseNumberValue(cells[column]);
      if (!rawValue) {
        continue;
      }
      if (!Number.isFinite(numericValue) || (strictPositive ? Number(numericValue) <= 0 : Number(numericValue) < 0)) {
        setSheetFieldError(fieldErrors, column, message);
      }
    }
  }

  if (sheet.name === "Feuil6") {
    const sapCode = parseFeuil6SapMarker(cells.A);
    const linearM = parseNumberValue(cells.D);
    const bounds = splitBoundsValue(cells.F);
    if (String(cells.A ?? "").trim() && !sapCode) {
      setSheetFieldError(fieldErrors, "A", "Veuillez saisir un SAP valide, par exemple SAP1 ou SAP5.");
    }
    if (String(cells.D ?? "").trim() && (!Number.isFinite(linearM) || Number(linearM) <= 0)) {
      setSheetFieldError(fieldErrors, "D", "Le linéaire doit être un nombre supérieur à 0.");
    }
    if (!bounds.startLabel || !bounds.endLabel) {
      setSheetFieldError(fieldErrors, "F", "Veuillez renseigner le début et la fin de la voie.");
    }
  }

  if (sheet.name === "Feuil7") {
    if (!String(cells.G ?? "").trim()) {
      setSheetFieldError(fieldErrors, "G", "Veuillez expliquer la cause probable de cette dégradation.");
    }
  }

  if (sheet.name === "Feuil4") {
    const hasValue = ["B", "C", "D", "E", "F"].some((column) => String(cells[column as SheetColumnKey] ?? "").trim());
    if (!hasValue) {
      setSheetFieldError(fieldErrors, "B", "Veuillez remplir au moins une valeur utile sur cette ligne.");
      formError = "Veuillez remplir au moins une valeur utile sur cette ligne.";
    }
  }

  if (!formError && Object.keys(fieldErrors).length > 0) {
    formError = Object.values(fieldErrors)[0] || "Veuillez corriger les champs obligatoires.";
  }

  return { fieldErrors, formError };
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
  if (source === "DERIVED") {
    return "Importée";
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

function normalizeComparisonText(value: unknown) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function describeRoadState(value: unknown): {
  label: string;
  tone: "ok" | "warning" | "danger" | "info" | "neutral";
  family: "technique" | "travaux" | "exploitation" | "autre";
  helper: string;
} | null {
  const raw = String(value ?? "").trim();
  if (!raw || raw === "-") {
    return null;
  }

  const normalized = normalizeLabel(raw);
  if (!normalized) {
    return null;
  }

  if (normalized.includes("EN COURS") && normalized.includes("AMENAG")) {
    return {
      label: "En cours d'aménagement",
      tone: "info",
      family: "travaux",
      helper: "Travaux ou aménagement en cours"
    };
  }

  if (normalized.includes("NON AMENAG")) {
    return {
      label: "Non aménagée",
      tone: "neutral",
      family: "exploitation",
      helper: "Section non aménagée"
    };
  }

  if (normalized === "RAS") {
    return {
      label: "RAS",
      tone: "neutral",
      family: "exploitation",
      helper: "Aucun signalement explicite"
    };
  }

  if (normalized.startsWith("BON")) {
    return {
      label: "Bon",
      tone: "ok",
      family: "technique",
      helper: "État technique favorable"
    };
  }

  if (normalized.startsWith("MOY")) {
    return {
      label: "Moyen",
      tone: "warning",
      family: "technique",
      helper: "État technique à surveiller"
    };
  }

  if (normalized.startsWith("MAU")) {
    return {
      label: "Mauvais",
      tone: "danger",
      family: "technique",
      helper: "État technique dégradé"
    };
  }

  return {
    label: raw,
    tone: "neutral",
    family: "autre",
    helper: "État renseigné"
  };
}

function renderRoadState(value: unknown, showHelper = false) {
  const descriptor = describeRoadState(value);
  if (!descriptor) {
    return toDisplay(value);
  }

  return (
    <span className={`road-state road-state--${showHelper ? "detailed" : "compact"}`}>
      <span className={`status-pill status-pill--${descriptor.tone}`}>{descriptor.label}</span>
      {showHelper ? <small>{descriptor.helper}</small> : null}
    </span>
  );
}

function compareRecommendations(referenceValue: unknown, computedValue: unknown): {
  tone: "ok" | "warning" | "danger" | "info" | "neutral";
  label: string;
  message: string;
} {
  const reference = String(referenceValue ?? "").trim();
  const computed = String(computedValue ?? "").trim();
  const normalizedReference = normalizeLabel(reference);
  const normalizedComputed = normalizeLabel(computed);
  const referenceKnown = Boolean(reference && !isToDetermineIntervention(reference));
  const computedKnown = Boolean(computed && computed !== "-");

  if (!referenceKnown && !computedKnown) {
    return {
      tone: "neutral",
      label: "Alignement incomplet",
      message: "Ni le référentiel ni le calcul n'apportent encore une intervention exploitable."
    };
  }

  if (!referenceKnown) {
    return {
      tone: "info",
      label: "Référentiel à compléter",
      message: "Le calcul propose une orientation, mais aucune intervention fiable n'est encore renseignée dans le référentiel."
    };
  }

  if (!computedKnown) {
    return {
      tone: "warning",
      label: "Calcul à compléter",
      message: "Le référentiel indique déjà une intervention, mais le calcul n'apporte pas encore de recommandation exploitable."
    };
  }

  if (normalizedReference === normalizedComputed) {
    return {
      tone: "ok",
      label: "Référentiel cohérent",
      message: "L'intervention connue dans le référentiel est cohérente avec la recommandation calculée."
    };
  }

  if (
    normalizedReference.includes(normalizedComputed) ||
    normalizedComputed.includes(normalizedReference)
  ) {
    return {
      tone: "warning",
      label: "Formulations à rapprocher",
      message: "Le référentiel et le calcul pointent dans la même direction, mais les formulations ne sont pas strictement identiques."
    };
  }

  return {
    tone: "danger",
    label: "Divergence à arbitrer",
    message: "L'intervention connue dans le référentiel diverge de la recommandation calculée. Un arbitrage métier est nécessaire."
  };
}

function formatMeasurementNumber(value: number | null | undefined) {
  if (!Number.isFinite(value ?? null)) {
    return "-";
  }
  return Number(value).toLocaleString("fr-FR");
}

function scrollPageTop() {
  if (typeof window === "undefined") {
    return;
  }
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function parseNumberValue(value: unknown) {
  const normalized = String(value ?? "").replace(/\s+/g, "").replace(",", ".");
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : null;
}

function normalizeRoadCompareKey(value: unknown) {
  return String(value ?? "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/BOULEVARD/g, "BVD")
    .replace(/AVENUE/g, "AV")
    .replace(/[().,;:_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseFeuil2SapCode(row: SheetRow) {
  const candidates = [row.I, row.H, row.B, row.A].map((value) => String(value ?? "").trim());
  for (const candidate of candidates) {
    const explicit = candidate.match(/SAP\s*([1-9][0-9]?)/i);
    if (explicit) {
      return `SAP${explicit[1]}`;
    }
    const sectionPrefix = candidate.match(/^([1-9][0-9]?)_/);
    if (sectionPrefix) {
      return `SAP${sectionPrefix[1]}`;
    }
  }
  return "";
}

function deriveTronconNoFromSectionNo(sectionNo: unknown) {
  const match = String(sectionNo ?? "").trim().match(/^[1-9][0-9]?_([1-9][0-9]*)$/);
  return match ? String(Number(match[1])) : "";
}

function parseFeuil6SapMarker(value: unknown) {
  const text = String(value ?? "").trim();
  const explicit = text.match(/SAP\s*([1-9][0-9]?)/i);
  if (explicit) {
    return `SAP${explicit[1]}`;
  }
  return "";
}

function trimPreviewText(value: unknown, maxLength = 180) {
  const text = String(value ?? "").trim();
  if (!text) {
    return "";
  }
  if (text.length <= maxLength) {
    return text;
  }
  return `${text.slice(0, Math.max(0, maxLength - 1)).trim()}…`;
}

function getDegradationPrimaryCause(item: DegradationItem | null) {
  return item?.causes.find((cause) => String(cause ?? "").trim()) || "";
}

function getDegradationOtherCauses(item: DegradationItem | null) {
  if (!item) {
    return [];
  }
  const primaryCause = getDegradationPrimaryCause(item);
  return item.causes.filter((cause) => String(cause ?? "").trim() && cause !== primaryCause);
}

function parseMultilineValues(value: string) {
  return [...new Set(String(value || "").split(/\n+/).map((item) => item.replace(/\s+/g, " ").trim()).filter(Boolean))];
}

function formatDegradationCausePreview(item: DegradationItem) {
  const causes = item.causes.map((cause) => String(cause ?? "").trim()).filter(Boolean);
  if (causes.length === 0) {
    return "Aucune cause renseignée";
  }
  if (causes.length === 1) {
    return causes[0];
  }
  const preview = causes.slice(0, 2).join(" ; ");
  return causes.length > 2 ? `${preview} + ${causes.length - 2} autre(s)` : preview;
}

function isMissingMaintenanceSolutionText(value: string) {
  return normalizeComparisonText(value) === normalizeComparisonText("Solution a parametrer dans le catalogue de maintenance.");
}

function deriveDegradationReferenceTreatment(item: Pick<DegradationItem, "solution" | "preventiveCriterion" | "treatmentDetails">) {
  const explicitSolution = String(item.solution || "").trim();
  if (explicitSolution && !isMissingMaintenanceSolutionText(explicitSolution)) {
    return explicitSolution;
  }

  const preventiveCriterion = String(item.preventiveCriterion || "").trim();
  const detailLines = String(item.treatmentDetails || "")
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  const firstDetail = detailLines[0] || "";
  const preventiveKey = normalizeComparisonText(preventiveCriterion);
  const genericCriterion =
    /LE TRAITEMENT COMPREND/.test(preventiveKey) ||
    /CRITERE PREVENTIF/.test(preventiveKey) ||
    /PHASES/.test(preventiveKey) ||
    /ETAPES/.test(preventiveKey) ||
    /:$/.test(preventiveCriterion);

  if (preventiveCriterion) {
    return genericCriterion && firstDetail ? firstDetail : preventiveCriterion;
  }
  if (firstDetail) {
    return firstDetail;
  }
  return explicitSolution;
}

function getTreatmentPhaseLines(value: string) {
  const lines = String(value || "")
    .split(/\n+/)
    .map((line) =>
      line
        .replace(/\s+/g, " ")
        .replace(/^[•\-\u2022]+\s*/, "")
        .trim()
    )
    .filter(Boolean);

  if (lines.length === 0) {
    return [];
  }

  const firstLine = normalizeComparisonText(lines[0]);
  if (/LE TRAITEMENT COMPREND|PHASES|ETAPES/.test(firstLine)) {
    return lines.slice(1);
  }

  return lines;
}

function getTreatmentPhaseLabel(phaseCount: number) {
  if (phaseCount <= 0) {
    return "";
  }
  if (phaseCount === 1) {
    return "Le traitement comprend 1 phase :";
  }
  if (phaseCount <= 3) {
    return `Le traitement comprend ${phaseCount} phases :`;
  }
  return "Le traitement comprend plusieurs phases :";
}

function getDegradationReferenceTreatment(item: DegradationItem) {
  return trimPreviewText(deriveDegradationReferenceTreatment(item), 220) || "-";
}

function getSheetDisplayName(sheetName: string) {
  if (sheetName === "Feuil1") {
    return "Campagne";
  }
  if (sheetName === "Feuil2") {
    return "Sections";
  }
  if (sheetName === "Feuil3") {
    return "Diagnostic";
  }
  if (sheetName === "Feuil4") {
    return "Évaluation";
  }
  if (sheetName === "Feuil5") {
    return "Compléments";
  }
  if (sheetName === "Feuil6") {
    return "Voies";
  }
  if (sheetName === "Feuil7") {
    return "Causes";
  }
  return sheetName;
}

function getSheetPrintSubtitle(sheetName: string) {
  if (sheetName === "Feuil1") {
    return "Feuille de campagnes de mesures de déflexion rattachées aux voies réelles du réseau.";
  }
  if (sheetName === "Feuil2") {
    return "Référentiel des sections du réseau, groupé par SAP, utilisé pour structurer les voies et leurs bornes.";
  }
  if (sheetName === "Feuil3") {
    return "Diagnostic technique des sections : état de la chaussée, assainissement et intervention à prévoir.";
  }
  if (sheetName === "Feuil4") {
    return "Tableau d'évaluation des observations, causes probables et décisions de maintenance.";
  }
  if (sheetName === "Feuil5") {
    return "Compléments latéraux des sections : largeurs utiles, trottoirs, stationnement et contexte d'assainissement.";
  }
  if (sheetName === "Feuil6") {
    return "Répertoire codifié central des voies : type, code, début, fin et justification.";
  }
  if (sheetName === "Feuil7") {
    return "Catalogue des dégradations, causes probables et informations utiles à la décision.";
  }
  return "Impression du tableau de la feuille active, avec les colonnes et lignes actuellement affichées.";
}

function renderSheetNavIcon(sheetName: string) {
  if (sheetName === "Feuil1") {
    return <CalendarRange size={16} aria-hidden="true" />;
  }
  if (sheetName === "Feuil2") {
    return <Route size={16} aria-hidden="true" />;
  }
  if (sheetName === "Feuil3") {
    return <Activity size={16} aria-hidden="true" />;
  }
  if (sheetName === "Feuil4") {
    return <Calculator size={16} aria-hidden="true" />;
  }
  if (sheetName === "Feuil5") {
    return <Layers3 size={16} aria-hidden="true" />;
  }
  if (sheetName === "Feuil6") {
    return <MapIcon size={16} aria-hidden="true" />;
  }
  if (sheetName === "Feuil7") {
    return <CircleHelp size={16} aria-hidden="true" />;
  }
  return <FileSpreadsheet size={16} aria-hidden="true" />;
}

function compareSapCodes(left: string, right: string) {
  const leftMatch = String(left ?? "").match(/([0-9]+)/);
  const rightMatch = String(right ?? "").match(/([0-9]+)/);
  if (leftMatch && rightMatch) {
    const diff = Number(leftMatch[1]) - Number(rightMatch[1]);
    if (diff !== 0) {
      return diff;
    }
  }
  return String(left ?? "").localeCompare(String(right ?? ""), "fr-FR");
}

function inferRoadTypeFromCode(roadCode: string) {
  const normalized = normalizeRoadCompareKey(roadCode);
  if (normalized.startsWith("BVD") || normalized.startsWith("BOULEVARD")) {
    return "Boulevard";
  }
  if (normalized.startsWith("AV") || normalized.startsWith("AVENUE")) {
    return "Avenue";
  }
  if (normalized.startsWith("RUE")) {
    return "Rue";
  }
  return "";
}

function parseRoadSectionSourcePayload(sourcePayload: string) {
  try {
    const parsed = JSON.parse(sourcePayload || "{}");
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function getRoadSectionSourceRowNo(section: RoadSectionItem, sheetName: string) {
  const parsed = parseRoadSectionSourcePayload(section.sourcePayload);
  const rowNo = Number(parsed?.sources?.[sheetName]?.rowNo);
  if (Number.isFinite(rowNo) && rowNo > 0) {
    return rowNo;
  }
  if (section.sourceSheet === sheetName && Number(section.sourceRowNo) > 0) {
    return Number(section.sourceRowNo);
  }
  return 0;
}

function getRoadSectionDisplayOrder(section: RoadSectionItem, sourceRow: SheetRow | null) {
  if (sourceRow) {
    return getDisplayRowNumber(sourceRow);
  }
  if (Number(section.sourceRowNo) > 0) {
    return Number(section.sourceRowNo);
  }
  return section.id;
}

function toSheetCellValue(value: string | number | null | undefined) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value);
}

function splitBoundsValue(value: string | null | undefined) {
  const text = String(value ?? "").trim();
  if (!text) {
    return { startLabel: "", endLabel: "" };
  }
  const itineraryMatch = text.match(/^(?:de\s+)?(.+?)\s+[aà]\s+(.+)$/i);
  if (itineraryMatch) {
    return {
      startLabel: String(itineraryMatch[1] ?? "").trim(),
      endLabel: String(itineraryMatch[2] ?? "").trim()
    };
  }
  const slashMatch = text.match(/^(.+?)\s*\/\s*(.+)$/);
  if (slashMatch) {
    return {
      startLabel: String(slashMatch[1] ?? "").trim(),
      endLabel: String(slashMatch[2] ?? "").trim()
    };
  }
  return { startLabel: text, endLabel: "" };
}

function formatBoundsValue(startLabel: string | null | undefined, endLabel: string | null | undefined) {
  const start = String(startLabel ?? "").trim();
  const end = String(endLabel ?? "").trim();
  if (start && end) {
    return `${start} à ${end}`;
  }
  return start || end;
}

function resolveRoadFromFeuil2Row(row: SheetRow, roads: RoadCatalogItem[]) {
  const codeKey = normalizeRoadCompareKey(row.C);
  const designationKey = normalizeRoadCompareKey(row.D);
  const startKey = normalizeRoadCompareKey(row.E);
  const endKey = normalizeRoadCompareKey(row.F);
  const sapKey = normalizeRoadCompareKey(parseFeuil2SapCode(row));

  return (
    roads.find((road) => {
      const roadCode = normalizeRoadCompareKey(road.roadCode);
      const roadDesignation = normalizeRoadCompareKey(road.designation);
      const roadStart = normalizeRoadCompareKey(road.startLabel);
      const roadEnd = normalizeRoadCompareKey(road.endLabel);
      const roadSap = normalizeRoadCompareKey(road.sapCode);

      if (codeKey && roadCode === codeKey) {
        return true;
      }

      if (designationKey && roadDesignation === designationKey && (!sapKey || roadSap === sapKey)) {
        return true;
      }

      return Boolean(
        designationKey &&
          startKey &&
          endKey &&
          roadDesignation === designationKey &&
          roadStart === startKey &&
          roadEnd === endKey
      );
    }) ?? null
  );
}

function resolveRoadFromFeuil6Row(row: SheetRow, roads: RoadCatalogItem[]) {
  const codeKey = normalizeRoadCompareKey(row.C);
  const designationKey = normalizeRoadCompareKey(row.E);
  const bounds = splitBoundsValue(row.F);
  const itineraryKey = normalizeRoadCompareKey(formatBoundsValue(bounds.startLabel, bounds.endLabel));

  return (
    roads.find((road) => {
      const roadCode = normalizeRoadCompareKey(road.roadCode);
      const roadDesignation = normalizeRoadCompareKey(road.designation);
      const roadItinerary = normalizeRoadCompareKey(formatBoundsValue(road.startLabel, road.endLabel) || road.itinerary);

      if (codeKey && roadCode === codeKey) {
        return true;
      }
      if (designationKey && roadDesignation === designationKey) {
        return true;
      }
      return Boolean(itineraryKey && roadItinerary === itineraryKey);
    }) ?? null
  );
}

function resolveRoadFromFeuil3Row(row: SheetRow, roads: RoadCatalogItem[]) {
  const codeKey = normalizeRoadCompareKey(row.A);
  const designationKey = normalizeRoadCompareKey(row.B);
  const startKey = normalizeRoadCompareKey(row.C);
  const endKey = normalizeRoadCompareKey(row.D);

  return (
    roads.find((road) => {
      const roadCode = normalizeRoadCompareKey(road.roadCode);
      const roadDesignation = normalizeRoadCompareKey(road.designation);
      const roadStart = normalizeRoadCompareKey(road.startLabel);
      const roadEnd = normalizeRoadCompareKey(road.endLabel);

      if (codeKey && roadCode === codeKey) {
        return true;
      }
      if (designationKey && roadDesignation === designationKey && startKey && endKey) {
        return roadStart === startKey && roadEnd === endKey;
      }
      if (designationKey && roadDesignation === designationKey) {
        return true;
      }
      return false;
    }) ?? null
  );
}

function hasFeuil5ParkingValue(value: unknown) {
  const normalized = normalizeLabel(value);
  return Boolean(normalized && !["-", "0", "NON", "FAUX"].includes(normalized));
}

function formatAmount(value: number | null) {
  if (!Number.isFinite(value ?? null)) {
    return "-";
  }
  return `${Number(value).toLocaleString()} FCFA`;
}

function formatBytes(value: number | null | undefined) {
  const bytes = Number(value);
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "-";
  }
  if (bytes < 1024) {
    return `${bytes} o`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} Ko`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} Mo`;
}

function getFileNameFromPath(value: string) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }
  const parts = text.split(/[\\/]/);
  return parts[parts.length - 1] || text;
}

function uniqueValues(values: Array<string | null | undefined>) {
  return [...new Set(values.map((value) => String(value ?? "").trim()).filter(Boolean))];
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

function buildFallbackRoadFromHistory(row: DecisionHistoryItem): RoadCatalogItem {
  return {
    id: row.roadId ?? 0,
    roadKey: `${row.roadCode}|${row.roadDesignation}|${row.startLabel}|${row.endLabel}`,
    roadCode: row.roadCode || "",
    designation: row.roadDesignation || "",
    sapCode: row.sapCode || "",
    startLabel: row.startLabel || "",
    endLabel: row.endLabel || "",
    lengthM: null,
    widthM: null,
    surfaceType: "",
    pavementState: "",
    drainageType: "",
    drainageState: "",
    sidewalkMinM: null,
    parkingLeft: "",
    parkingRight: "",
    parkingOther: "",
    itinerary: [row.startLabel, row.endLabel].filter(Boolean).join(" -> "),
    justification: "",
    interventionHint: row.contextualIntervention || ""
  };
}

function buildFallbackDegradationFromHistory(row: DecisionHistoryItem): DegradationItem {
  return {
    id: 0,
    code: "",
    name: row.degradationName || "",
    causes: row.probableCause ? [row.probableCause] : [],
    solution: row.maintenanceSolution || "",
    solutionSource: row.maintenanceSolution ? "OVERRIDE" : "MISSING",
    templateKey: null,
    preventiveCriterion: row.maintenanceSolution || "",
    treatmentDetails: row.maintenanceSolution || ""
  };
}

export default function App() {
  const hasElectronBridge = Boolean(window.padApp);
  const appName = window.padApp?.appName || "PAD Maintenance Routière";
  const appVersion = window.padApp?.appVersion || "0.0.0";
  const appLogoSrc = `${import.meta.env.BASE_URL}logo-pad.png`;
  const [status, setStatus] = useState<DataStatus | null>(null);
  const [integrityReport, setIntegrityReport] = useState<DataIntegrityReport | null>(null);
  const [dashboardSummary, setDashboardSummary] = useState<DashboardSummary | null>(null);
  const [definitions, setDefinitions] = useState<SheetDefinition[]>([]);
  const [activeView, setActiveView] = useState<string>("dashboard");

  const [rows, setRows] = useState<SheetRow[]>([]);
  const [search, setSearch] = useState("");
  const [importPath, setImportPath] = useState(DEFAULT_IMPORT_PATH);
  const [importPreview, setImportPreview] = useState<ImportPreview | null>(null);
  const [draftCells, setDraftCells] = useState<Partial<Record<SheetColumnKey, string>>>({});
  const [draftFieldErrors, setDraftFieldErrors] = useState<Partial<Record<SheetColumnKey, string>>>({});
  const [draftFormError, setDraftFormError] = useState("");
  const [editingRowId, setEditingRowId] = useState<number | null>(null);

  const [sapSectors, setSapSectors] = useState<SapSector[]>([]);
  const [allRoads, setAllRoads] = useState<RoadCatalogItem[]>([]);
  const [allRoadSections, setAllRoadSections] = useState<RoadSectionItem[]>([]);
  const [roads, setRoads] = useState<RoadCatalogItem[]>([]);
  const [measurementCampaigns, setMeasurementCampaigns] = useState<MeasurementCampaignItem[]>([]);
  const [measurementRows, setMeasurementRows] = useState<RoadMeasurementItem[]>([]);
  const [selectedMeasurementCampaignKey, setSelectedMeasurementCampaignKey] = useState("");
  const [isMeasurementCampaignModalOpen, setIsMeasurementCampaignModalOpen] = useState(false);
  const [isMeasurementRowModalOpen, setIsMeasurementRowModalOpen] = useState(false);
  const [editingMeasurementCampaignId, setEditingMeasurementCampaignId] = useState<number | null>(null);
  const [measurementCampaignRoadId, setMeasurementCampaignRoadId] = useState<number | "">("");
  const [measurementCampaignSectionLabel, setMeasurementCampaignSectionLabel] = useState("");
  const [measurementCampaignStartLabel, setMeasurementCampaignStartLabel] = useState("");
  const [measurementCampaignEndLabel, setMeasurementCampaignEndLabel] = useState("");
  const [measurementCampaignDate, setMeasurementCampaignDate] = useState("");
  const [measurementCampaignPkStartM, setMeasurementCampaignPkStartM] = useState("");
  const [measurementCampaignPkEndM, setMeasurementCampaignPkEndM] = useState("");
  const [measurementCampaignFieldErrors, setMeasurementCampaignFieldErrors] = useState<Record<string, string>>({});
  const [editingMeasurementRowId, setEditingMeasurementRowId] = useState<number | null>(null);
  const [measurementPkLabel, setMeasurementPkLabel] = useState("");
  const [measurementPkM, setMeasurementPkM] = useState("");
  const [measurementLectureLeft, setMeasurementLectureLeft] = useState("");
  const [measurementLectureAxis, setMeasurementLectureAxis] = useState("");
  const [measurementLectureRight, setMeasurementLectureRight] = useState("");
  const [measurementDeflectionLeft, setMeasurementDeflectionLeft] = useState("");
  const [measurementDeflectionAxis, setMeasurementDeflectionAxis] = useState("");
  const [measurementDeflectionRight, setMeasurementDeflectionRight] = useState("");
  const [measurementDeflectionAvg, setMeasurementDeflectionAvg] = useState("");
  const [measurementStdDev, setMeasurementStdDev] = useState("");
  const [measurementDeflectionDc, setMeasurementDeflectionDc] = useState("");
  const [measurementRowFieldErrors, setMeasurementRowFieldErrors] = useState<Record<string, string>>({});
  const [degradations, setDegradations] = useState<DegradationItem[]>([]);
  const [feuil2SapFilter, setFeuil2SapFilter] = useState("");
  const [feuil3SapFilter, setFeuil3SapFilter] = useState("");
  const [feuil5SapFilter, setFeuil5SapFilter] = useState("");
  const [feuil6SapFilter, setFeuil6SapFilter] = useState("");
  const [selectedSap, setSelectedSap] = useState("");
  const [roadSearch, setRoadSearch] = useState("");
  const [selectedRoadId, setSelectedRoadId] = useState<number | "">("");
  const [selectedDegradationId, setSelectedDegradationId] = useState<number | "">("");
  const [deflectionValue, setDeflectionValue] = useState("");
  const [askDrainage, setAskDrainage] = useState(true);
  const [decisionFieldErrors, setDecisionFieldErrors] = useState<Record<string, string>>({});
  const [decisionFormError, setDecisionFormError] = useState("");
  const [decisionResult, setDecisionResult] = useState<DecisionResult | null>(null);
  const [degradationSearch, setDegradationSearch] = useState("");
  const [editingDegradationId, setEditingDegradationId] = useState<number | null>(null);
  const [degradationNameDraft, setDegradationNameDraft] = useState("");
  const [degradationCausesDraft, setDegradationCausesDraft] = useState("");
  const [degradationPreventiveCriterionDraft, setDegradationPreventiveCriterionDraft] = useState("");
  const [degradationTreatmentDetailsDraft, setDegradationTreatmentDetailsDraft] = useState("");
  const [degradationFieldErrors, setDegradationFieldErrors] = useState<Record<string, string>>({});
  const [degradationFormError, setDegradationFormError] = useState("");
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
  const [maintenanceResponsibleName, setMaintenanceResponsibleName] = useState("");
  const [maintenanceAttachmentPath, setMaintenanceAttachmentPath] = useState("");
  const [maintenanceObservation, setMaintenanceObservation] = useState("");
  const [maintenanceCostAmount, setMaintenanceCostAmount] = useState("");
  const [feuil4Snapshot, setFeuil4Snapshot] = useState<Feuil4Snapshot | null>(null);
  const [solutionTemplates, setSolutionTemplates] = useState<MaintenanceSolutionTemplate[]>([]);
  const [selectedTemplateKey, setSelectedTemplateKey] = useState("");
  const [solutionDraft, setSolutionDraft] = useState("");
  const [solutionFieldErrors, setSolutionFieldErrors] = useState<Record<string, string>>({});
  const [solutionFormError, setSolutionFormError] = useState("");
  const [shouldScrollToDegradationEditor, setShouldScrollToDegradationEditor] = useState(false);
  const [isDegradationEditorHighlighted, setIsDegradationEditorHighlighted] = useState(false);
  const [maintenanceFieldErrors, setMaintenanceFieldErrors] = useState<Record<string, string>>({});
  const [maintenanceFormError, setMaintenanceFormError] = useState("");

  const [isBusy, setIsBusy] = useState(false);
  const [isLoadingRows, setIsLoadingRows] = useState(false);
  const [isDecisionBusy, setIsDecisionBusy] = useState(false);
  const [isMeasurementLoading, setIsMeasurementLoading] = useState(false);
  const [isMeasurementBusy, setIsMeasurementBusy] = useState(false);
  const [isDegradationBusy, setIsDegradationBusy] = useState(false);
  const [isHistoryLoading, setIsHistoryLoading] = useState(false);
  const [isMaintenanceLoading, setIsMaintenanceLoading] = useState(false);
  const [isMaintenanceBusy, setIsMaintenanceBusy] = useState(false);
  const [isPreviewingImport, setIsPreviewingImport] = useState(false);
  const [isBackupBusy, setIsBackupBusy] = useState(false);
  const [isReportingBusy, setIsReportingBusy] = useState(false);
  const [isSolutionBusy, setIsSolutionBusy] = useState(false);
  const [isAttachmentBusy, setIsAttachmentBusy] = useState(false);
  const supportsMaintenanceAttachments =
    typeof window.padApp?.maintenance?.pickAttachment === "function" &&
    typeof window.padApp?.maintenance?.openAttachment === "function";
  const [isImportAssistantCollapsed, setIsImportAssistantCollapsed] = useState(false);
  const [isIntegrityAlertDismissed, setIsIntegrityAlertDismissed] = useState(false);
  const [shouldScrollToIntegritySection, setShouldScrollToIntegritySection] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [pendingSheetDraft, setPendingSheetDraft] = useState<{
    sheetName: string;
    cells: Partial<Record<SheetColumnKey, string>>;
    notice: string;
  } | null>(null);
  const hasNotifiedAppReadyRef = useRef(false);
  const degradationEditorRef = useRef<HTMLDivElement | null>(null);
  const integritySectionRef = useRef<HTMLDivElement | null>(null);
  const sheetEditorRef = useRef<HTMLElement | null>(null);
  const measurementCampaignEditorRef = useRef<HTMLDivElement | null>(null);
  const measurementRowEditorRef = useRef<HTMLDivElement | null>(null);

  const activeSheetName = activeView.startsWith("sheet:") ? activeView.replace("sheet:", "") : "";
  const activeSheet = useMemo(
    () => definitions.find((sheet) => sheet.name === activeSheetName) ?? null,
    [definitions, activeSheetName]
  );
  const activeColumns = activeSheet?.columns ?? [];
  const editableColumns = useMemo(() => getEditableColumns(activeSheet), [activeSheet]);
  const selectedRoad = useMemo(
    () => allRoads.find((road) => road.id === selectedRoadId) ?? null,
    [allRoads, selectedRoadId]
  );
  const selectedMeasurementCampaign = useMemo(
    () => measurementCampaigns.find((item) => item.campaignKey === selectedMeasurementCampaignKey) ?? null,
    [measurementCampaigns, selectedMeasurementCampaignKey]
  );
  const decisionMeasurementCampaigns = useMemo(() => {
    if (selectedRoadId === "") {
      return measurementCampaigns;
    }
    return measurementCampaigns.filter((item) => item.roadId === selectedRoadId);
  }, [measurementCampaigns, selectedRoadId]);
  const selectedDegradation = useMemo(
    () => degradations.find((item) => item.id === selectedDegradationId) ?? null,
    [degradations, selectedDegradationId]
  );
  const degradationDraftCauses = useMemo(() => parseMultilineValues(degradationCausesDraft), [degradationCausesDraft]);
  const degradationEditorReferenceTreatment = useMemo(() => {
    const draftValue = trimPreviewText(
      deriveDegradationReferenceTreatment({
        solution: selectedDegradation?.solution || "",
        preventiveCriterion: degradationPreventiveCriterionDraft,
        treatmentDetails: degradationTreatmentDetailsDraft
      }),
      220
    );
    return draftValue || "-";
  }, [degradationPreventiveCriterionDraft, degradationTreatmentDetailsDraft, selectedDegradation]);
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
  const centralSapCodes = useMemo(
    () =>
      uniqueValues([
        ...sapSectors.map((item) => item.code),
        ...allRoads.map((item) => item.sapCode)
      ])
        .filter(Boolean)
        .sort(compareSapCodes),
    [allRoads, sapSectors]
  );
  const latestMaintenance = maintenancePreviewRows[0] ?? null;
  const latestDecisionCampaign = useMemo(() => {
    if (decisionMeasurementCampaigns.length === 0) {
      return null;
    }
    return [...decisionMeasurementCampaigns].sort((a, b) => {
      const dateA = Date.parse(a.measurementDate || "");
      const dateB = Date.parse(b.measurementDate || "");
      if (Number.isFinite(dateA) && Number.isFinite(dateB) && dateA !== dateB) {
        return dateB - dateA;
      }
      return b.id - a.id;
    })[0];
  }, [decisionMeasurementCampaigns]);
  const decisionSummary = useMemo(() => {
    if (!decisionResult) {
      return null;
    }

    const hasDeflection = decisionResult.deflection.value != null;
    const hasCause = Boolean(String(decisionResult.probableCause || "").trim());
    const hasSolution =
      Boolean(String(decisionResult.maintenanceSolution || "").trim()) &&
      decisionResult.degradation.solutionSource !== "MISSING";
    const hasTreatmentDetails = Boolean(String(decisionResult.degradation.treatmentDetails || "").trim());
    const hasContextualIntervention = Boolean(
      decisionResult.contextualIntervention && !isToDetermineIntervention(decisionResult.contextualIntervention)
    );
    const warnings: string[] = [];
    let confidenceScore = 0;

    if (hasDeflection) {
      confidenceScore += 2;
    } else {
      warnings.push("Décision calculée sans valeur D mesurée.");
    }

    if (hasCause) {
      confidenceScore += 1;
    } else {
      warnings.push("Cause probable non renseignée.");
    }

    if (hasSolution) {
      confidenceScore += 1;
    } else {
      warnings.push("Solution de maintenance encore à préciser dans le catalogue.");
    }

    if (!hasTreatmentDetails) {
      warnings.push("Le mode opératoire détaillé n'est pas encore renseigné pour cette dégradation.");
    }

    if (askDrainage) {
      confidenceScore += 1;
    } else {
      warnings.push("Assainissement non interrogé explicitement dans cette analyse.");
    }

    if (decisionResult.drainage.needsAttention) {
      warnings.push("L'assainissement demande une vigilance particulière.");
    } else {
      confidenceScore += 1;
    }

    if (!hasContextualIntervention) {
      warnings.push("Intervention contextuelle tronçon encore à préciser.");
    }

    const confidence =
      confidenceScore >= 5
        ? { label: "Décision robuste", tone: "ok" as const }
        : confidenceScore >= 3
          ? { label: "Décision à confirmer", tone: "warning" as const }
          : { label: "Décision fragile", tone: "alert" as const };

    const finalRecommendation =
      (hasContextualIntervention ? decisionResult.contextualIntervention : "") ||
      decisionResult.maintenanceSolution ||
      toDeflectionRecommendationLabel(decisionResult.deflection.recommendation) ||
      "À confirmer";
    const alignmentStatus = compareRecommendations(
      hasContextualIntervention ? decisionResult.contextualIntervention : "",
      decisionResult.maintenanceSolution || toDeflectionRecommendationLabel(decisionResult.deflection.recommendation)
    );

    if (alignmentStatus.tone === "danger" || alignmentStatus.tone === "warning") {
      warnings.push(alignmentStatus.message);
    }

    const rationale = [
      hasDeflection
        ? `La valeur de déflexion place la chaussée au niveau ${toDeflectionSeverityLabel(
            decisionResult.deflection.severity
          ).toUpperCase()} et oriente vers ${toDeflectionRecommendationLabel(decisionResult.deflection.recommendation).toLowerCase()}.`
        : "Aucune valeur de déflexion n'a été fournie : la recommandation est calculée à partir des autres règles disponibles.",
      hasSolution
        ? `La dégradation ${decisionResult.degradation.name.toLowerCase()} active la solution de maintenance ${decisionResult.maintenanceSolution.toLowerCase()}.`
        : `La dégradation ${decisionResult.degradation.name.toLowerCase()} a été reconnue, mais sa solution de maintenance reste à préciser dans le catalogue.`,
      hasTreatmentDetails
        ? "Le traitement détaillé de la dégradation est disponible pour l'exploitation terrain."
        : "Le traitement détaillé de la dégradation reste à compléter.",
      askDrainage
        ? `L'assainissement a été interrogé explicitement : ${decisionResult.drainage.recommendation}${
            decisionResult.drainage.needsAttention ? " Cette règle renforce la vigilance sur la décision finale." : "."
          }`
        : "L'assainissement n'a pas été interrogé explicitement dans ce calcul ; la décision doit être lue avec prudence.",
      hasContextualIntervention
        ? `Le contexte du tronçon confirme une intervention de type ${String(
            decisionResult.contextualIntervention || ""
          ).toLowerCase()}.`
        : "Le contexte du tronçon n'apporte pas encore d'intervention complémentaire fiable."
    ];

    return {
      confidence,
      finalRecommendation,
      warnings,
      rationale
    };
  }, [askDrainage, decisionResult]);
  const decisionTreatment = useMemo(() => {
    if (!decisionResult) {
      return null;
    }

    const preventiveCriterion =
      String(decisionResult.degradation.preventiveCriterion || "").trim() ||
      String(decisionResult.maintenanceSolution || "").trim() ||
      String(decisionSummary?.finalRecommendation || "").trim();
    const treatmentDetails =
      String(decisionResult.degradation.treatmentDetails || "").trim() ||
      String(decisionResult.maintenanceSolution || "").trim();
    const otherCauses = (decisionResult.degradation.causes || []).filter((cause) => {
      return normalizeComparisonText(cause) !== normalizeComparisonText(decisionResult.probableCause || "");
    });
    const phaseLines = getTreatmentPhaseLines(treatmentDetails);

    return {
      preventiveCriterion,
      treatmentDetails,
      otherCauses,
      phaseLines,
      phaseLabel: getTreatmentPhaseLabel(phaseLines.length)
    };
  }, [decisionResult, decisionSummary]);
  const filteredDegradations = useMemo(() => {
    const searchTerm = degradationSearch.trim().toLowerCase();
    if (!searchTerm) {
      return degradations;
    }
    return degradations.filter((item) => {
      return (
        item.name.toLowerCase().includes(searchTerm) ||
        item.causes.join(" ").toLowerCase().includes(searchTerm) ||
        item.preventiveCriterion.toLowerCase().includes(searchTerm) ||
        item.treatmentDetails.toLowerCase().includes(searchTerm)
      );
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
  const feuil2Sections = useMemo(() => {
    if (activeSheetName !== "Feuil2") {
      return [];
    }

    const normalizedSearch = search.trim().toLowerCase();
    return allRoadSections
      .filter((section) => {
        if (!normalizedSearch) {
          return true;
        }
        return [
          section.tronconNo,
          section.sectionNo,
          section.roadCode,
          section.designation,
          section.startLabel,
          section.endLabel,
          section.sapCode
        ]
          .join(" ")
          .toLowerCase()
          .includes(normalizedSearch);
      })
      .map((section) => {
        const linkedRoad =
          (section.roadId ? allRoads.find((road) => road.id === section.roadId) : null) ||
          allRoads.find((road) => road.roadKey === section.roadKey) ||
          null;
        const sourceRowNo = getRoadSectionSourceRowNo(section, "Feuil2");
        const sourceRow = sourceRowNo ? rows.find((row) => row.rowNo === sourceRowNo) ?? null : null;
        return {
          section,
          sourceRow,
          linkedRoad,
          sapCode: section.sapCode || linkedRoad?.sapCode || "",
          tronconNo: deriveTronconNoFromSectionNo(section.sectionNo) || section.tronconNo,
          sectionNo: section.sectionNo,
          roadLabel: section.roadCode,
          designation: section.designation,
          startLabel: section.startLabel,
          endLabel: section.endLabel,
          lengthM: section.lengthM
        };
      })
      .filter((item) => item.roadLabel || item.designation || item.sapCode);
  }, [activeSheetName, allRoadSections, allRoads, rows, search]);
  const feuil2SapOptions = useMemo(
    () =>
      uniqueValues([...centralSapCodes, ...feuil2Sections.map((item) => item.sapCode)])
        .filter(Boolean)
        .sort(compareSapCodes),
    [centralSapCodes, feuil2Sections]
  );
  const feuil2Groups = useMemo(() => {
    const source = feuil2Sections.filter((item) => !feuil2SapFilter || item.sapCode === feuil2SapFilter);
    const grouped = new Map<
      string,
      {
        sapCode: string;
        rows: typeof source;
        totalLengthM: number;
      }
    >();

    for (const item of source) {
      const key = item.sapCode || "SAP?";
      if (!grouped.has(key)) {
        grouped.set(key, { sapCode: key, rows: [], totalLengthM: 0 });
      }
      const group = grouped.get(key)!;
      group.rows.push(item);
      group.totalLengthM += item.lengthM ?? 0;
    }

    return [...grouped.values()]
      .sort((a, b) => compareSapCodes(a.sapCode, b.sapCode))
      .map((group) => ({
        ...group,
        rows: group.rows.sort(
          (left, right) =>
            getRoadSectionDisplayOrder(left.section, left.sourceRow) -
            getRoadSectionDisplayOrder(right.section, right.sourceRow)
        )
      }));
  }, [feuil2SapFilter, feuil2Sections]);
  const feuil6DirectoryRows = useMemo(() => {
    if (activeSheetName !== "Feuil6") {
      return [];
    }

    const normalizedSearch = search.trim().toLowerCase();
    return allRoads
      .filter((road) => {
        if (!normalizedSearch) {
          return true;
        }
        return [
          road.sapCode,
          road.roadCode,
          road.designation,
          road.itinerary,
          road.startLabel,
          road.endLabel,
          road.justification
        ]
          .join(" ")
          .toLowerCase()
          .includes(normalizedSearch);
      })
      .map((road) => {
        const roadBounds = splitBoundsValue(road.itinerary || formatBoundsValue(road.startLabel, road.endLabel));
        const sourceRow =
          rows.find((row) => {
            const matchedRoad = resolveRoadFromFeuil6Row(row, [road]);
            return matchedRoad?.id === road.id;
          }) ?? null;
        return {
          sourceRow,
          sapCode: road.sapCode,
          roadType: inferRoadTypeFromCode(road.roadCode),
          roadCode: road.roadCode,
          linearM: road.lengthM,
          proposedName: road.designation,
          itinerary: road.itinerary || (road.startLabel && road.endLabel ? `${road.startLabel} à ${road.endLabel}` : ""),
          startLabel: road.startLabel || roadBounds.startLabel,
          endLabel: road.endLabel || roadBounds.endLabel,
          justification: road.justification,
          linkedRoad: road
        };
      });
  }, [activeSheetName, allRoads, rows, search]);
  const feuil6SapOptions = useMemo(
    () =>
      uniqueValues([...centralSapCodes, ...feuil6DirectoryRows.map((item) => item.sapCode)])
        .filter(Boolean)
        .sort(compareSapCodes),
    [centralSapCodes, feuil6DirectoryRows]
  );
  const feuil6Groups = useMemo(() => {
    const source = feuil6DirectoryRows.filter((item) => !feuil6SapFilter || item.sapCode === feuil6SapFilter);
    const grouped = new Map<
      string,
      {
        sapCode: string;
        rows: typeof source;
        totalLinearM: number;
      }
    >();

    for (const item of source) {
      const key = item.sapCode || "SAP?";
      if (!grouped.has(key)) {
        grouped.set(key, { sapCode: key, rows: [], totalLinearM: 0 });
      }
      const group = grouped.get(key)!;
      group.rows.push(item);
      group.totalLinearM += item.linearM ?? 0;
    }

    return [...grouped.values()].sort((a, b) => compareSapCodes(a.sapCode, b.sapCode));
  }, [feuil6DirectoryRows, feuil6SapFilter]);
  const feuil3Profiles = useMemo(() => {
    if (activeSheetName !== "Feuil3") {
      return [];
    }

    const normalizedSearch = search.trim().toLowerCase();
    return allRoadSections
      .filter((section) => {
        if (!normalizedSearch) {
          return true;
        }
        return [
          section.sapCode,
          section.roadCode,
          section.designation,
          section.startLabel,
          section.endLabel,
          section.surfaceType,
          section.pavementState,
          section.drainageType,
          section.drainageState,
          section.interventionHint
        ]
          .join(" ")
          .toLowerCase()
          .includes(normalizedSearch);
      })
      .map((section) => {
        const linkedRoad =
          (section.roadId ? allRoads.find((road) => road.id === section.roadId) : null) ||
          allRoads.find((road) => road.roadKey === section.roadKey) ||
          null;
        const sourceRowNo = getRoadSectionSourceRowNo(section, "Feuil3");
        const sourceRow = sourceRowNo ? rows.find((row) => row.rowNo === sourceRowNo) ?? null : null;
        return {
          section,
          sourceRow,
          linkedRoad,
          sapCode: section.sapCode || linkedRoad?.sapCode || "",
          roadLabel: section.roadCode || linkedRoad?.roadCode || "",
          designation: section.designation || linkedRoad?.designation || "",
          startLabel: section.startLabel || linkedRoad?.startLabel || "",
          endLabel: section.endLabel || linkedRoad?.endLabel || "",
          lengthM: section.lengthM ?? linkedRoad?.lengthM ?? null,
          facadeWidthM: section.widthM ?? linkedRoad?.widthM ?? null,
          surfaceType: section.surfaceType || linkedRoad?.surfaceType || "",
          pavementState: section.pavementState || linkedRoad?.pavementState || "",
          drainageType: section.drainageType || linkedRoad?.drainageType || "",
          drainageState: section.drainageState || linkedRoad?.drainageState || "",
          sidewalkMinM: section.sidewalkMinM ?? linkedRoad?.sidewalkMinM ?? null,
          interventionHint: section.interventionHint || linkedRoad?.interventionHint || ""
        };
      })
      .filter((item) => item.roadLabel || item.designation || item.surfaceType || item.interventionHint);
  }, [activeSheetName, allRoadSections, allRoads, rows, search]);
  const feuil3SapOptions = useMemo(
    () =>
      uniqueValues([...centralSapCodes, ...feuil3Profiles.map((item) => item.sapCode)])
        .filter(Boolean)
        .sort(compareSapCodes),
    [centralSapCodes, feuil3Profiles]
  );
  const feuil3Groups = useMemo(() => {
    const source = feuil3Profiles.filter((item) => !feuil3SapFilter || item.sapCode === feuil3SapFilter);
    const grouped = new Map<
      string,
      {
        sapCode: string;
        rows: typeof source;
      }
    >();

    for (const item of source) {
      const key = item.sapCode || "SAP?";
      if (!grouped.has(key)) {
        grouped.set(key, { sapCode: key, rows: [] });
      }
      grouped.get(key)!.rows.push(item);
    }

    return [...grouped.values()]
      .sort((a, b) => compareSapCodes(a.sapCode, b.sapCode))
      .map((group) => ({
        ...group,
        rows: group.rows.sort(
          (left, right) =>
            getRoadSectionDisplayOrder(left.section, left.sourceRow) -
            getRoadSectionDisplayOrder(right.section, right.sourceRow)
        )
      }));
  }, [feuil3Profiles, feuil3SapFilter]);
  const feuil3PendingInterventions = useMemo(
    () => feuil3Profiles.filter((item) => isToDetermineIntervention(item.interventionHint)).length,
    [feuil3Profiles]
  );
  const feuil3DrainageAlerts = useMemo(
    () =>
      feuil3Profiles.filter((item) => {
        const normalized = normalizeLabel(item.drainageState);
        return Boolean(normalized && !["-", "BON"].includes(normalized));
      }).length,
    [feuil3Profiles]
  );
  const feuil3SpecialStateCounts = useMemo(() => {
    return feuil3Profiles.reduce(
      (acc, item) => {
        const descriptor = describeRoadState(item.pavementState);
        if (!descriptor) {
          return acc;
        }
        if (descriptor.family === "travaux") {
          acc.inProgress += 1;
        } else if (descriptor.label === "Non aménagée") {
          acc.notBuilt += 1;
        } else if (descriptor.label === "RAS") {
          acc.ras += 1;
        }
        return acc;
      },
      { inProgress: 0, notBuilt: 0, ras: 0 }
    );
  }, [feuil3Profiles]);
  const feuil5Profiles = useMemo(() => {
    if (activeSheetName !== "Feuil5") {
      return [];
    }

    const normalizedSearch = search.trim().toLowerCase();
    return allRoadSections
      .filter((section) => {
        if (!normalizedSearch) {
          return true;
        }
        return [
          section.sapCode,
          section.tronconNo,
          section.sectionNo,
          section.roadCode,
          section.designation,
          section.startLabel,
          section.endLabel,
          section.surfaceType,
          section.pavementState,
          section.drainageType,
          section.drainageState
        ]
          .join(" ")
          .toLowerCase()
          .includes(normalizedSearch);
      })
      .map((section) => {
        const linkedRoad =
          (section.roadId ? allRoads.find((road) => road.id === section.roadId) : null) ||
          allRoads.find((road) => road.roadKey === section.roadKey) ||
          null;
        const sourceRowNo = getRoadSectionSourceRowNo(section, "Feuil5");
        const sourceRow = sourceRowNo ? rows.find((row) => row.rowNo === sourceRowNo) ?? null : null;
        return {
          section,
          sourceRow,
          linkedRoad,
          sapCode: section.sapCode || linkedRoad?.sapCode || "",
          tronconNo: deriveTronconNoFromSectionNo(section.sectionNo) || section.tronconNo,
          sectionNo: section.sectionNo,
          roadLabel: section.roadCode || linkedRoad?.roadCode || "",
          designation: section.designation || linkedRoad?.designation || "",
          startLabel: section.startLabel || linkedRoad?.startLabel || "",
          endLabel: section.endLabel || linkedRoad?.endLabel || "",
          lengthM: section.lengthM ?? linkedRoad?.lengthM ?? null,
          facadeWidthM: section.widthM ?? linkedRoad?.widthM ?? null,
          surfaceType: section.surfaceType || linkedRoad?.surfaceType || "",
          pavementState: section.pavementState || linkedRoad?.pavementState || "",
          drainageType: section.drainageType || linkedRoad?.drainageType || "",
          drainageState: section.drainageState || linkedRoad?.drainageState || "",
          sidewalkMinM: section.sidewalkMinM ?? linkedRoad?.sidewalkMinM ?? null,
          parkingLeft: linkedRoad?.parkingLeft || "",
          parkingRight: linkedRoad?.parkingRight || "",
          parkingOther: linkedRoad?.parkingOther || ""
        };
      })
      .filter((item) => item.roadLabel || item.designation || item.surfaceType || item.drainageState);
  }, [activeSheetName, allRoadSections, allRoads, rows, search]);
  const feuil5SapOptions = useMemo(
    () =>
      uniqueValues([...centralSapCodes, ...feuil5Profiles.map((item) => item.sapCode)])
        .filter(Boolean)
        .sort(compareSapCodes),
    [centralSapCodes, feuil5Profiles]
  );
  const feuil5Groups = useMemo(() => {
    const source = feuil5Profiles.filter((item) => !feuil5SapFilter || item.sapCode === feuil5SapFilter);
    const grouped = new Map<
      string,
      {
        sapCode: string;
        rows: typeof source;
      }
    >();

    for (const item of source) {
      const key = item.sapCode || "SAP?";
      if (!grouped.has(key)) {
        grouped.set(key, { sapCode: key, rows: [] });
      }
      grouped.get(key)!.rows.push(item);
    }

    return [...grouped.values()]
      .sort((a, b) => compareSapCodes(a.sapCode, b.sapCode))
      .map((group) => ({
        ...group,
        rows: group.rows.sort(
          (left, right) =>
            getRoadSectionDisplayOrder(left.section, left.sourceRow) -
            getRoadSectionDisplayOrder(right.section, right.sourceRow)
        )
      }));
  }, [feuil5Profiles, feuil5SapFilter]);
  const feuil5ParkingCount = useMemo(
    () =>
      feuil5Profiles.filter(
        (item) =>
          hasFeuil5ParkingValue(item.parkingLeft) ||
          hasFeuil5ParkingValue(item.parkingRight) ||
          hasFeuil5ParkingValue(item.parkingOther)
      ).length,
    [feuil5Profiles]
  );
  const feuil5DrainageWatchCount = useMemo(
    () =>
      feuil5Profiles.filter((item) => {
        const normalized = normalizeLabel(item.drainageState);
        return Boolean(normalized && !["-", "BON"].includes(normalized));
      }).length,
    [feuil5Profiles]
  );

  const refreshStatus = useCallback(async () => {
    const nextStatus = await padApi.getDataStatus();
    setStatus(nextStatus);
    if (nextStatus.lastImportPath) {
      setImportPath(nextStatus.lastImportPath);
    }
  }, []);

  const loadIntegrityReport = useCallback(async () => {
    const report = await padApi.getDataIntegrityReport();
    setIntegrityReport(report);
  }, []);

  const loadDashboardSummary = useCallback(async () => {
    const summary = await padApi.getDashboardSummary();
    setDashboardSummary(summary);
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
    return roadList;
  }, []);

  const loadAllRoadSections = useCallback(async () => {
    const sectionList = await padApi.listRoadSections();
    setAllRoadSections(sectionList);
    return sectionList;
  }, []);

  const loadMeasurementCampaigns = useCallback(async () => {
    setIsMeasurementLoading(true);
    try {
      const items = await padApi.listMeasurementCampaigns({ limit: 500 });
      setMeasurementCampaigns(items);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsMeasurementLoading(false);
    }
  }, []);

  const loadMeasurementRows = useCallback(async (campaignKey?: string) => {
    const targetKey = String(campaignKey || "").trim();
    if (!targetKey) {
      setMeasurementRows([]);
      return;
    }

    setIsMeasurementLoading(true);
    try {
      const items = await padApi.listRoadMeasurements({
        campaignKey: targetKey,
        limit: 5000
      });
      setMeasurementRows(items);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsMeasurementLoading(false);
    }
  }, []);

  const loadRows = useCallback(async () => {
    if (
      !activeSheetName ||
      activeView === "dashboard" ||
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
    setDraftCells(createEmptyCells(editableColumns));
    setDraftFieldErrors({});
    setDraftFormError("");
  }, [editableColumns]);

  const clearDraftFieldError = useCallback((column: SheetColumnKey) => {
    setDraftFieldErrors((prev) => {
      if (!prev[column]) {
        return prev;
      }
      const next = { ...prev };
      delete next[column];
      return next;
    });
    setDraftFormError("");
  }, []);

  const buildDraftRow = useCallback(
    (cells: Partial<Record<SheetColumnKey, string>>) =>
      ({
        id: editingRowId ?? 0,
        rowNo: 0,
        ...cells
      }) as SheetRow,
    [editingRowId]
  );

  const resolveDraftRoadMatch = useCallback(
    (sheetName: string | undefined, cells: Partial<Record<SheetColumnKey, string>>) => {
      const draftRow = buildDraftRow(cells);
      if (sheetName === "Feuil2" || sheetName === "Feuil5") {
        return resolveRoadFromFeuil2Row(draftRow, allRoads);
      }
      if (sheetName === "Feuil3") {
        return resolveRoadFromFeuil3Row(draftRow, allRoads);
      }
      if (sheetName === "Feuil6") {
        return resolveRoadFromFeuil6Row(draftRow, allRoads);
      }
      return null;
    },
    [allRoads, buildDraftRow]
  );

  const suggestSectionForSap = useCallback(
    (sapCode: string) => {
      const sapNumber = String(sapCode ?? "").match(/([0-9]+)/)?.[1] || "";
      if (!sapNumber) {
        return { tronconNo: "", sectionNo: "" };
      }
      const prefix = `${sapNumber}_`;
      const currentIndexes = allRoadSections
        .map((section) => String(section.sectionNo ?? "").trim())
        .filter((value) => value.startsWith(prefix))
        .map((value) => Number(value.split("_")[1] || 0))
        .filter((value) => Number.isFinite(value) && value > 0);
      const nextIndex = currentIndexes.length > 0 ? Math.max(...currentIndexes) + 1 : 1;
      return {
        tronconNo: String(nextIndex),
        sectionNo: `${sapNumber}_${nextIndex}`
      };
    },
    [allRoadSections]
  );

  const autofillDraftFromRoad = useCallback(
    (
      sheetName: string | undefined,
      cells: Partial<Record<SheetColumnKey, string>>,
      road: RoadCatalogItem
    ): Partial<Record<SheetColumnKey, string>> => {
      const next = { ...cells };

      if (sheetName === "Feuil2") {
        next.C = road.roadCode;
        next.D = road.designation;
        next.E = road.startLabel || next.E || "";
        next.F = road.endLabel || next.F || "";
        next.G = road.lengthM !== null && road.lengthM !== undefined ? String(road.lengthM) : next.G || "";

        const existingRow = rows.find((row) => {
          if (editingRowId && row.id === editingRowId) {
            return false;
          }
          const rowRoad = resolveRoadFromFeuil2Row(row, allRoads);
          return rowRoad?.id === road.id;
        });

        if (existingRow) {
          next.B = String(existingRow.B ?? "");
          next.A = deriveTronconNoFromSectionNo(existingRow.B) || String(existingRow.A ?? "");
        } else {
          const suggestion = suggestSectionForSap(road.sapCode);
          next.A = suggestion.tronconNo;
          next.B = suggestion.sectionNo;
        }
      }

      if (sheetName === "Feuil5") {
        next.C = road.roadCode;
        next.D = road.designation;
        next.E = road.startLabel || next.E || "";
        next.F = road.endLabel || next.F || "";
        next.G = road.lengthM !== null && road.lengthM !== undefined ? String(road.lengthM) : next.G || "";
        next.H = road.widthM !== null && road.widthM !== undefined ? String(road.widthM) : next.H || "";
        next.I = road.surfaceType || next.I || "";
        next.J = road.pavementState || next.J || "";
        next.K = road.drainageType || next.K || "";
        next.L = road.drainageState || next.L || "";
        next.M = road.sidewalkMinM !== null && road.sidewalkMinM !== undefined ? String(road.sidewalkMinM) : next.M || "";
        next.N = road.parkingLeft || next.N || "";
        next.O = road.parkingRight || next.O || "";
        next.P = road.parkingOther || next.P || "";

        const existingRow = rows.find((row) => {
          if (editingRowId && row.id === editingRowId) {
            return false;
          }
          const rowRoad = resolveRoadFromFeuil2Row(row, allRoads);
          return rowRoad?.id === road.id;
        });

        if (existingRow) {
          next.B = String(existingRow.B ?? "");
          next.A = deriveTronconNoFromSectionNo(existingRow.B) || String(existingRow.A ?? "");
        } else {
          const suggestion = suggestSectionForSap(road.sapCode);
          next.A = suggestion.tronconNo;
          next.B = suggestion.sectionNo;
        }
      }

      if (sheetName === "Feuil3") {
        next.A = road.roadCode;
        next.B = road.designation;
        next.C = road.startLabel || next.C || "";
        next.D = road.endLabel || next.D || "";
        next.E = road.lengthM !== null && road.lengthM !== undefined ? String(road.lengthM) : next.E || "";
        next.F = road.widthM !== null && road.widthM !== undefined ? String(road.widthM) : next.F || "";
        next.G = road.surfaceType || next.G || "";
        next.H = road.pavementState || next.H || "";
        next.I = road.drainageType || next.I || "";
        next.J = road.drainageState || next.J || "";
        next.K = road.sidewalkMinM !== null && road.sidewalkMinM !== undefined ? String(road.sidewalkMinM) : next.K || "";
        next.L = road.interventionHint || next.L || "";
      }

      if (sheetName === "Feuil6") {
        next.A = road.sapCode || next.A || "";
        next.B = inferRoadTypeFromCode(road.roadCode) || next.B || "";
        next.C = road.roadCode;
        next.D = road.lengthM !== null && road.lengthM !== undefined ? String(road.lengthM) : next.D || "";
        next.E = road.designation;
        next.F = road.itinerary || (road.startLabel && road.endLabel ? `${road.startLabel} à ${road.endLabel}` : next.F || "");
        next.G = road.justification || next.G || "";
      }

      return next;
    },
    [allRoads, editingRowId, rows, suggestSectionForSap]
  );

  const validateDraftDuplicates = useCallback(
    (
      sheetName: string | undefined,
      cells: Partial<Record<SheetColumnKey, string>>
    ): { fieldErrors: Partial<Record<SheetColumnKey, string>>; formError: string } => {
      const fieldErrors: Partial<Record<SheetColumnKey, string>> = {};
      const comparableRows = rows.filter((row) => !editingRowId || row.id !== editingRowId);
      const currentEditingRow = editingRowId ? rows.find((row) => row.id === editingRowId) ?? null : null;
      const currentEditingRowNo = currentEditingRow?.rowNo ?? 0;
      const draftRoad =
        sheetName === "Feuil2" || sheetName === "Feuil5" || sheetName === "Feuil3" || sheetName === "Feuil6"
          ? resolveDraftRoadMatch(sheetName, cells)
          : null;

      if (sheetName === "Feuil2" || sheetName === "Feuil5") {
        const sectionSapCode = parseFeuil2SapCode({
          A: cells.A ?? "",
          B: cells.B ?? ""
        } as SheetRow);
        const roadSapCode = draftRoad?.sapCode || "";
        if (sectionSapCode && roadSapCode && sectionSapCode !== roadSapCode) {
          fieldErrors.B = `Cette voie est classée dans ${toDisplay(roadSapCode)}. Pour utiliser le numéro de section ${toDisplay(cells.B)}, changez d'abord le SAP de la voie dans Voies.`;
          return { fieldErrors, formError: fieldErrors.B };
        }

        const sectionNo = normalizeLabel(cells.B);
        const startKey = normalizeLabel(cells.E);
        const endKey = normalizeLabel(cells.F);
        const duplicateSection = allRoadSections.find((section) => {
          const sheetRowNo = getRoadSectionSourceRowNo(section, sheetName);
          if (!sheetRowNo || (currentEditingRowNo && sheetRowNo === currentEditingRowNo)) {
            return false;
          }
          const sameRoad = draftRoad
            ? (section.roadId ? section.roadId === draftRoad.id : normalizeLabel(section.roadKey) === normalizeLabel(draftRoad.roadKey))
            : normalizeRoadCompareKey(section.roadCode) === normalizeRoadCompareKey(cells.C);
          if (!sameRoad) {
            return false;
          }
          if (sectionNo && normalizeLabel(section.sectionNo) === sectionNo) {
            return true;
          }
          return Boolean(startKey && endKey && normalizeLabel(section.startLabel) === startKey && normalizeLabel(section.endLabel) === endKey);
        });

        if (duplicateSection) {
          fieldErrors.B = `Cette section existe déjà pour cette voie sous le numéro ${toDisplay(duplicateSection.sectionNo)}. Ouvrez la ligne existante si vous voulez la compléter.`;
          return { fieldErrors, formError: fieldErrors.B };
        }
      }

      if (sheetName === "Feuil3") {
        const duplicateSection = allRoadSections.find((section) => {
          const sheetRowNo = getRoadSectionSourceRowNo(section, "Feuil3");
          if (!sheetRowNo || (currentEditingRowNo && sheetRowNo === currentEditingRowNo)) {
            return false;
          }
          if (draftRoad) {
            return section.roadId ? section.roadId === draftRoad.id : normalizeLabel(section.roadKey) === normalizeLabel(draftRoad.roadKey);
          }
          return (
            normalizeRoadCompareKey(section.roadCode) === normalizeRoadCompareKey(cells.A) ||
            normalizeRoadCompareKey(section.designation) === normalizeRoadCompareKey(cells.B)
          );
        });

        if (duplicateSection) {
          fieldErrors.A = `Un diagnostic existe déjà pour cette voie (${toDisplay(duplicateSection.roadCode)}). Ouvrez la ligne existante si vous voulez le mettre à jour.`;
          return { fieldErrors, formError: fieldErrors.A };
        }
      }

      if (sheetName === "Feuil6") {
        const codeKey = normalizeRoadCompareKey(cells.C);
        const designationKey = normalizeRoadCompareKey(cells.E);
        const bounds = splitBoundsValue(cells.F);
        const itineraryKey = normalizeRoadCompareKey(formatBoundsValue(bounds.startLabel, bounds.endLabel));

        const duplicateByCode = allRoads.find(
          (road) => codeKey && normalizeRoadCompareKey(road.roadCode) === codeKey && (!draftRoad || road.id !== draftRoad.id)
        );
        if (duplicateByCode) {
          fieldErrors.C = `Une voie avec le code ${toDisplay(duplicateByCode.roadCode)} existe déjà dans le logiciel. Ouvrez la voie existante si vous souhaitez la consulter ou la modifier.`;
          return { fieldErrors, formError: fieldErrors.C };
        }

        const duplicateByDesignation = allRoads.find(
          (road) =>
            designationKey &&
            normalizeRoadCompareKey(road.designation) === designationKey &&
            (!draftRoad || road.id !== draftRoad.id)
        );
        if (duplicateByDesignation) {
          fieldErrors.E = `Une voie nommée ${toDisplay(duplicateByDesignation.designation)} existe déjà dans le logiciel. Ouvrez la voie existante au lieu d'en créer une nouvelle.`;
          return { fieldErrors, formError: fieldErrors.E };
        }

        const duplicateByItinerary = allRoads.find((road) => {
          const roadItinerary = normalizeRoadCompareKey(formatBoundsValue(road.startLabel, road.endLabel) || road.itinerary);
          return itineraryKey && roadItinerary === itineraryKey && (!draftRoad || road.id !== draftRoad.id);
        });
        if (duplicateByItinerary) {
          fieldErrors.F = `Une voie existe déjà avec ce même début et cette même fin (${toDisplay(duplicateByItinerary.designation)}). Vérifiez la voie existante avant d'en créer une nouvelle.`;
          return { fieldErrors, formError: fieldErrors.F };
        }
      }

      if (sheetName === "Feuil7") {
        const referenceKey = normalizeLabel(cells.B);
        const degradationKey = normalizeRoadCompareKey(cells.C);
        const duplicate = comparableRows.find(
          (row) =>
            (referenceKey && normalizeLabel(row.B) === referenceKey) ||
            (degradationKey && normalizeRoadCompareKey(row.C) === degradationKey)
        );
        if (duplicate) {
          if (referenceKey && normalizeLabel(duplicate.B) === referenceKey) {
            fieldErrors.B = `La référence ${toDisplay(duplicate.B)} existe déjà dans la liste. Ouvrez la ligne existante si vous voulez la compléter.`;
            return { fieldErrors, formError: fieldErrors.B };
          }
          fieldErrors.C = `La dégradation ${toDisplay(duplicate.C)} existe déjà dans la liste. Ouvrez la ligne existante si vous voulez la compléter.`;
          return { fieldErrors, formError: fieldErrors.C };
        }
      }

      if (sheetName === "Feuil4") {
        const labelKey = normalizeRoadCompareKey(cells.A);
        const duplicate = comparableRows.find((row) => labelKey && normalizeRoadCompareKey(row.A) === labelKey);
        if (duplicate) {
          fieldErrors.A = `Le libellé ${toDisplay(duplicate.A)} existe déjà dans le programme d'évaluation. Ouvrez la ligne existante si vous voulez le modifier.`;
          return { fieldErrors, formError: fieldErrors.A };
        }
      }

      return { fieldErrors, formError: "" };
    },
    [allRoadSections, allRoads, editingRowId, resolveDraftRoadMatch, rows]
  );

  const handleDraftCellChange = useCallback(
    (column: SheetColumnKey, value: string) => {
      clearDraftFieldError(column);
      setDraftFormError("");
      setDraftCells((prev) => {
        let next = {
          ...prev,
          [column]: value
        };

        if ((activeSheetName === "Feuil2" || activeSheetName === "Feuil5") && column === "B") {
          next.A = deriveTronconNoFromSectionNo(value);
        }

        const isRoadSelectorField =
          (activeSheetName === "Feuil3" && (column === "A" || column === "B")) ||
          (activeSheetName === "Feuil5" && (column === "C" || column === "D")) ||
          (activeSheetName === "Feuil6" && (column === "C" || column === "E"));

        if (isRoadSelectorField) {
          const matchedRoad = resolveDraftRoadMatch(activeSheetName, next);
          if (matchedRoad) {
            next = autofillDraftFromRoad(activeSheetName, next, matchedRoad);
          }
        }

        return next;
      });
    },
    [activeSheetName, autofillDraftFromRoad, clearDraftFieldError, resolveDraftRoadMatch]
  );

  useEffect(() => {
    setDraftFieldErrors({});
    setDraftFormError("");
  }, [activeSheetName]);

  useEffect(() => {
    if (!pendingSheetDraft || activeSheetName !== pendingSheetDraft.sheetName) {
      return;
    }

    setEditingRowId(null);
    setDraftCells({
      ...createEmptyCells(editableColumns),
      ...pendingSheetDraft.cells
    });
    setDraftFieldErrors({});
    setDraftFormError("");
    setNotice(pendingSheetDraft.notice);
    setError("");
    setPendingSheetDraft(null);
    sheetEditorRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  }, [activeSheetName, editableColumns, pendingSheetDraft]);

  const resetMeasurementCampaignForm = useCallback(
    (roadId?: number | "") => {
      setEditingMeasurementCampaignId(null);
      setMeasurementCampaignFieldErrors({});
      setMeasurementCampaignRoadId(
        roadId === "" || (Number.isFinite(Number(roadId)) && Number(roadId) > 0) ? roadId ?? "" : ""
      );
      setMeasurementCampaignSectionLabel("");
      setMeasurementCampaignStartLabel("");
      setMeasurementCampaignEndLabel("");
      setMeasurementCampaignDate("");
      setMeasurementCampaignPkStartM("");
      setMeasurementCampaignPkEndM("");
    },
    []
  );

  const resetMeasurementRowForm = useCallback(() => {
    setEditingMeasurementRowId(null);
    setMeasurementRowFieldErrors({});
    setMeasurementPkLabel("");
    setMeasurementPkM("");
    setMeasurementLectureLeft("");
    setMeasurementLectureAxis("");
    setMeasurementLectureRight("");
    setMeasurementDeflectionLeft("");
    setMeasurementDeflectionAxis("");
    setMeasurementDeflectionRight("");
    setMeasurementDeflectionAvg("");
    setMeasurementStdDev("");
    setMeasurementDeflectionDc("");
  }, []);

  const clearMeasurementCampaignFieldError = useCallback((field: string) => {
    setMeasurementCampaignFieldErrors((current) => {
      if (!current[field]) {
        return current;
      }
      const next = { ...current };
      delete next[field];
      return next;
    });
  }, []);

  const clearMeasurementRowFieldError = useCallback((field: string) => {
    setMeasurementRowFieldErrors((current) => {
      if (!current[field]) {
        return current;
      }
      const next = { ...current };
      delete next[field];
      return next;
    });
  }, []);

  const clearSolutionFieldError = useCallback((field: string) => {
    setSolutionFieldErrors((current) => {
      if (!current[field]) {
        return current;
      }
      const next = { ...current };
      delete next[field];
      return next;
    });
    setSolutionFormError("");
    setError("");
  }, []);

  const clearMaintenanceFieldError = useCallback((field: string) => {
    setMaintenanceFieldErrors((current) => {
      if (!current[field]) {
        return current;
      }
      const next = { ...current };
      delete next[field];
      return next;
    });
    setMaintenanceFormError("");
    setError("");
  }, []);

  const clearDecisionFieldError = useCallback((field: string) => {
    setDecisionFieldErrors((current) => {
      if (!current[field]) {
        return current;
      }
      const next = { ...current };
      delete next[field];
      return next;
    });
    setDecisionFormError("");
  }, []);

  const validateDecisionForm = useCallback(() => {
    const nextErrors: Record<string, string> = {};
    const rawDeflection = String(deflectionValue ?? "").trim();
    const parsedDeflection = rawDeflection ? Number(deflectionValue) : null;

    if (!selectedRoadId) {
      nextErrors.roadId = "Veuillez choisir une voie.";
    }
    if (!selectedDegradationId) {
      nextErrors.degradationId = "Veuillez choisir une dégradation.";
    }
    if (rawDeflection && !Number.isFinite(parsedDeflection)) {
      nextErrors.deflectionValue = "Veuillez saisir un nombre valide pour la déflexion.";
    }

    setDecisionFieldErrors(nextErrors);
    setDecisionFormError(Object.values(nextErrors)[0] || "");
    return nextErrors;
  }, [deflectionValue, selectedDegradationId, selectedRoadId]);

  const validateMeasurementCampaignForm = useCallback(() => {
    const nextErrors: Record<string, string> = {};
    const roadId = Number(measurementCampaignRoadId);
    const pkStart = measurementCampaignPkStartM === "" ? null : Number(measurementCampaignPkStartM);
    const pkEnd = measurementCampaignPkEndM === "" ? null : Number(measurementCampaignPkEndM);

    if (!Number.isFinite(roadId) || roadId <= 0) {
      nextErrors.roadId = "Veuillez choisir la voie concernée.";
    }
    if (!measurementCampaignSectionLabel.trim()) {
      nextErrors.sectionLabel = "Veuillez renseigner le nom du tronçon.";
    }
    if (!measurementCampaignStartLabel.trim()) {
      nextErrors.startLabel = "Veuillez renseigner le point de départ.";
    }
    if (!measurementCampaignEndLabel.trim()) {
      nextErrors.endLabel = "Veuillez renseigner le point d'arrivée.";
    }
    if (!measurementCampaignDate.trim()) {
      nextErrors.measurementDate = "Veuillez renseigner la date de mesure.";
    }
    if ((pkStart === null) !== (pkEnd === null)) {
      if (pkStart === null) {
        nextErrors.pkStartM = "Veuillez renseigner le PK début.";
      }
      if (pkEnd === null) {
        nextErrors.pkEndM = "Veuillez renseigner le PK fin.";
      }
    }
    if (pkStart !== null && !Number.isFinite(pkStart)) {
      nextErrors.pkStartM = "Veuillez saisir un nombre valide pour le PK début.";
    }
    if (pkEnd !== null && !Number.isFinite(pkEnd)) {
      nextErrors.pkEndM = "Veuillez saisir un nombre valide pour le PK fin.";
    }
    if (pkStart !== null && pkEnd !== null && Number.isFinite(pkStart) && Number.isFinite(pkEnd) && pkStart > pkEnd) {
      nextErrors.pkEndM = "Le PK fin doit être supérieur ou égal au PK début.";
    }

    setMeasurementCampaignFieldErrors(nextErrors);
    return nextErrors;
  }, [
    measurementCampaignDate,
    measurementCampaignEndLabel,
    measurementCampaignPkEndM,
    measurementCampaignPkStartM,
    measurementCampaignRoadId,
    measurementCampaignSectionLabel,
    measurementCampaignStartLabel
  ]);

  const validateMeasurementRowForm = useCallback(() => {
    const nextErrors: Record<string, string> = {};
    const pkMeters = measurementPkM === "" ? null : Number(measurementPkM);
    const hasValue = [
      measurementLectureLeft,
      measurementLectureAxis,
      measurementLectureRight,
      measurementDeflectionLeft,
      measurementDeflectionAxis,
      measurementDeflectionRight,
      measurementDeflectionAvg,
      measurementStdDev,
      measurementDeflectionDc
    ].some((value) => String(value).trim() !== "");

    if (!measurementPkLabel.trim()) {
      nextErrors.pkLabel = "Veuillez renseigner le PK affiché.";
    }
    if (measurementPkM === "") {
      nextErrors.pkM = "Veuillez renseigner le PK en mètres.";
    } else if (!Number.isFinite(pkMeters)) {
      nextErrors.pkM = "Veuillez saisir un nombre valide pour le PK en mètres.";
    }
    if (!hasValue) {
      nextErrors.values = "Veuillez renseigner au moins une valeur de mesure.";
    }

    setMeasurementRowFieldErrors(nextErrors);
    return nextErrors;
  }, [
    measurementDeflectionAvg,
    measurementDeflectionAxis,
    measurementDeflectionDc,
    measurementDeflectionLeft,
    measurementDeflectionRight,
    measurementLectureAxis,
    measurementLectureLeft,
    measurementLectureRight,
    measurementPkLabel,
    measurementPkM,
    measurementStdDev
  ]);

  const validateSolutionTemplateForm = useCallback(() => {
    const nextErrors: Record<string, string> = {};
    if (!selectedTemplateKey) {
      nextErrors.templateKey = "Veuillez choisir un modèle de solution.";
    }
    setSolutionFieldErrors(nextErrors);
    return nextErrors;
  }, [selectedTemplateKey]);

  const validateSolutionOverrideForm = useCallback(() => {
    const nextErrors: Record<string, string> = {};
    if (!solutionDraft.trim()) {
      nextErrors.solutionDraft = "Veuillez saisir une solution personnalisée.";
    }
    setSolutionFieldErrors(nextErrors);
    return nextErrors;
  }, [solutionDraft]);

  const validateDegradationForm = useCallback(() => {
    const nextErrors: Record<string, string> = {};
    if (!degradationNameDraft.trim()) {
      nextErrors.name = "Veuillez renseigner le nom de la dégradation.";
    }
    if (degradationDraftCauses.length === 0) {
      nextErrors.causes = "Ajoutez au moins une cause potentielle, une par ligne.";
    }
    setDegradationFieldErrors(nextErrors);
    return nextErrors;
  }, [degradationDraftCauses.length, degradationNameDraft]);

  const validateMaintenanceForm = useCallback(() => {
    const nextErrors: Record<string, string> = {};
    const parsedRoadId = Number(maintenanceRoadId);
    const deflectionBefore =
      String(maintenanceDeflectionBefore).trim() === "" ? null : Number(maintenanceDeflectionBefore);
    const deflectionAfter = String(maintenanceDeflectionAfter).trim() === "" ? null : Number(maintenanceDeflectionAfter);
    const costAmount = String(maintenanceCostAmount).trim() === "" ? null : Number(maintenanceCostAmount);

    if (!Number.isFinite(parsedRoadId) || parsedRoadId <= 0) {
      nextErrors.roadId = "Veuillez choisir la voie concernée.";
    }
    if (!maintenanceType.trim()) {
      nextErrors.type = "Veuillez renseigner le type d'entretien.";
    }
    if (!maintenanceDate) {
      nextErrors.interventionDate = "Veuillez renseigner la date prévue.";
    }
    if (!maintenanceStateBefore.trim()) {
      nextErrors.stateBefore = "Veuillez renseigner l'état avant.";
    }
    if (!maintenanceResponsibleName.trim()) {
      nextErrors.responsibleName = "Veuillez renseigner le responsable PAD.";
    }
    if (String(maintenanceCostAmount).trim() === "") {
      nextErrors.costAmount = "Veuillez renseigner le coût estimé.";
    }
    if (maintenanceStatus === "TERMINE" && !maintenanceCompletionDate) {
      nextErrors.completionDate = "Veuillez renseigner la date réelle ou de clôture.";
    }
    if (maintenanceDate && maintenanceCompletionDate && maintenanceCompletionDate < maintenanceDate) {
      nextErrors.completionDate = "La date réelle ne peut pas être antérieure à la date prévue.";
    }
    if (deflectionBefore !== null && !Number.isFinite(deflectionBefore)) {
      nextErrors.deflectionBefore = "Veuillez saisir un nombre valide pour la déflexion avant.";
    }
    if (deflectionAfter !== null && !Number.isFinite(deflectionAfter)) {
      nextErrors.deflectionAfter = "Veuillez saisir un nombre valide pour la déflexion après.";
    }
    if (costAmount !== null && (!Number.isFinite(costAmount) || costAmount < 0)) {
      nextErrors.costAmount = "Veuillez saisir un coût valide supérieur ou égal à 0.";
    }

    setMaintenanceFieldErrors(nextErrors);
    return nextErrors;
  }, [
    maintenanceCompletionDate,
    maintenanceCostAmount,
    maintenanceDate,
    maintenanceDeflectionAfter,
    maintenanceDeflectionBefore,
    maintenanceRoadId,
    maintenanceResponsibleName,
    maintenanceStateBefore,
    maintenanceStatus,
    maintenanceType
  ]);

  const resetMaintenanceForm = useCallback(() => {
    setEditingMaintenanceId(null);
    setMaintenanceFieldErrors({});
    setMaintenanceFormError("");
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
    setMaintenanceResponsibleName("");
    setMaintenanceAttachmentPath("");
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
          loadIntegrityReport(),
          loadDashboardSummary(),
          refreshDecisionCatalogs(),
          loadAllRoads(),
          loadAllRoadSections(),
          loadRoads(),
          loadMeasurementCampaigns(),
          loadHistory(),
          loadMaintenanceRows(),
          loadFeuil4Snapshot(),
          loadSolutionTemplates()
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
          if (!hasNotifiedAppReadyRef.current && typeof window.padApp?.lifecycle?.notifyReady === "function") {
            hasNotifiedAppReadyRef.current = true;
            window.padApp.lifecycle.notifyReady();
          }
        }
      }
    }

    bootstrap();

    return () => {
      cancelled = true;
    };
  }, [
    hasElectronBridge,
    loadDashboardSummary,
    loadAllRoads,
    loadAllRoadSections,
    loadFeuil4Snapshot,
    loadHistory,
    loadIntegrityReport,
    loadMaintenanceRows,
    loadMeasurementCampaigns,
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
    loadMeasurementCampaigns().catch((err) => setError(toErrorMessage(err)));
  }, [hasElectronBridge, loadMeasurementCampaigns]);

  useEffect(() => {
    if (isMeasurementBusy || isMeasurementLoading) {
      return;
    }
    const availableCampaigns =
      activeView === "decision"
        ? selectedRoadId === ""
          ? measurementCampaigns
          : decisionMeasurementCampaigns
        : measurementCampaigns;
    if (availableCampaigns.length === 0) {
      setSelectedMeasurementCampaignKey("");
      setMeasurementRows([]);
      return;
    }

    const stillAvailable = availableCampaigns.some((item) => item.campaignKey === selectedMeasurementCampaignKey);
    if (!stillAvailable && activeView === "decision") {
      setSelectedMeasurementCampaignKey(availableCampaigns[0].campaignKey);
      return;
    }
    if (!stillAvailable && selectedMeasurementCampaignKey) {
      setSelectedMeasurementCampaignKey("");
    }
  }, [activeView, decisionMeasurementCampaigns, isMeasurementBusy, isMeasurementLoading, measurementCampaigns, selectedMeasurementCampaignKey, selectedRoadId]);

  useEffect(() => {
    if (!hasElectronBridge) {
      return;
    }
    loadMeasurementRows(selectedMeasurementCampaignKey).catch((err) => setError(toErrorMessage(err)));
  }, [hasElectronBridge, loadMeasurementRows, selectedMeasurementCampaignKey]);

  useEffect(() => {
    if (!selectedMeasurementCampaign) {
      resetMeasurementCampaignForm(selectedRoadId);
      resetMeasurementRowForm();
      return;
    }

    setEditingMeasurementCampaignId(selectedMeasurementCampaign.id);
    setMeasurementCampaignRoadId(selectedMeasurementCampaign.roadId ?? "");
    setMeasurementCampaignSectionLabel(selectedMeasurementCampaign.sectionLabel || "");
    setMeasurementCampaignStartLabel(selectedMeasurementCampaign.startLabel || "");
    setMeasurementCampaignEndLabel(selectedMeasurementCampaign.endLabel || "");
    setMeasurementCampaignDate(selectedMeasurementCampaign.measurementDate || "");
    setMeasurementCampaignPkStartM(
      selectedMeasurementCampaign.pkStartM == null ? "" : String(selectedMeasurementCampaign.pkStartM)
    );
    setMeasurementCampaignPkEndM(selectedMeasurementCampaign.pkEndM == null ? "" : String(selectedMeasurementCampaign.pkEndM));
    resetMeasurementRowForm();
  }, [resetMeasurementCampaignForm, resetMeasurementRowForm, selectedMeasurementCampaign, selectedRoadId]);

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
    if (selectedDegradationId === "") {
      return;
    }
    const exists = degradations.some((item) => item.id === selectedDegradationId);
    if (!exists) {
      setSelectedDegradationId("");
    }
  }, [degradations, selectedDegradationId]);

  useEffect(() => {
    if (!selectedDegradation) {
      setSelectedTemplateKey("");
      setSolutionDraft("");
      setSolutionFieldErrors({});
      setSolutionFormError("");
      return;
    }
    setEditingDegradationId(selectedDegradation.id);
    setDegradationNameDraft(selectedDegradation.name || "");
    setDegradationCausesDraft(selectedDegradation.causes.join("\n"));
    setDegradationPreventiveCriterionDraft(selectedDegradation.preventiveCriterion || "");
    setDegradationTreatmentDetailsDraft(selectedDegradation.treatmentDetails || "");
    setDegradationFieldErrors({});
    setDegradationFormError("");
    setSelectedTemplateKey(selectedDegradation.templateKey ?? "");
    setSolutionDraft(selectedDegradation.solution || "");
    setSolutionFieldErrors({});
    setSolutionFormError("");
  }, [selectedDegradation]);

  useEffect(() => {
    if (!shouldScrollToDegradationEditor || !degradationEditorRef.current) {
      return;
    }

    let timeoutId = 0;
    const rafId = window.requestAnimationFrame(() => {
      degradationEditorRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      setIsDegradationEditorHighlighted(true);
      setShouldScrollToDegradationEditor(false);
      timeoutId = window.setTimeout(() => setIsDegradationEditorHighlighted(false), 1800);
    });

    return () => {
      window.cancelAnimationFrame(rafId);
      if (timeoutId) {
        window.clearTimeout(timeoutId);
      }
    };
  }, [shouldScrollToDegradationEditor]);

  useEffect(() => {
    if (!shouldScrollToIntegritySection || activeView !== "dashboard" || !integritySectionRef.current) {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      integritySectionRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      setShouldScrollToIntegritySection(false);
    }, 60);

    return () => window.clearTimeout(timeoutId);
  }, [activeView, shouldScrollToIntegritySection]);

  useEffect(() => {
    if (dashboardSummary?.integrity?.status !== "OK") {
      setIsIntegrityAlertDismissed(false);
    }
  }, [dashboardSummary?.integrity?.status]);

  useEffect(() => {
    if (
      activeView === "dashboard" ||
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
      activeView === "dashboard" ||
      activeView === "decision" ||
      activeView === "catalogue" ||
      activeView === "degradations" ||
      activeView === "maintenance" ||
      activeView === "history" ||
      activeSheetName === "Feuil1"
    ) {
      return;
    }
    loadRows();
  }, [hasElectronBridge, activeSheetName, activeView, loadRows]);

  useEffect(() => {
    if (activeView === "sheet:Feuil7") {
      setActiveView("degradations");
    }
  }, [activeView]);

  async function handleImport() {
    setIsBusy(true);
    try {
      const nextStatus = await padApi.importFromExcel(importPath.trim() || undefined);
      setStatus(nextStatus);
      await Promise.all([
        loadIntegrityReport(),
        loadDashboardSummary(),
        refreshDecisionCatalogs(),
        loadAllRoads(),
        loadAllRoadSections(),
        loadRoads(),
        loadMeasurementCampaigns(),
        loadHistory(),
        loadMaintenanceRows(),
        loadFeuil4Snapshot(),
        loadSolutionTemplates()
      ]);
      if (activeView.startsWith("sheet:") && activeSheetName !== "Feuil1") {
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

  async function handlePreviewImport() {
    setIsPreviewingImport(true);
    try {
      const preview = await padApi.previewExcelImport(importPath.trim() || undefined);
      setImportPreview(preview);
      setError("");
      setNotice("Aperçu d'import généré.");
      setActiveView("dashboard");
    } catch (err) {
      setImportPreview(null);
      setError(toErrorMessage(err));
    } finally {
      setIsPreviewingImport(false);
    }
  }

  async function handleExportBackup() {
    setIsBackupBusy(true);
    try {
      const result = await padApi.exportBackup();
      if (result?.filePath) {
        setNotice(`Sauvegarde exportée: ${result.filePath}`);
      }
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBackupBusy(false);
    }
  }

  async function handleRestoreBackup() {
    if (!window.confirm("Restaurer une sauvegarde remplacera les données locales actuelles. Continuer ?")) {
      return;
    }

    setIsBackupBusy(true);
    try {
      const result = await padApi.restoreBackup();
      if (result) {
        setStatus(result);
        await Promise.all([
          loadIntegrityReport(),
          loadDashboardSummary(),
          refreshDecisionCatalogs(),
          loadAllRoads(),
          loadAllRoadSections(),
          loadRoads(),
          loadMeasurementCampaigns(),
          loadHistory(),
          loadMaintenanceRows(),
          loadFeuil4Snapshot(),
          loadSolutionTemplates()
        ]);
        if (activeView.startsWith("sheet:") && activeSheetName !== "Feuil1") {
          await loadRows();
        }
        setNotice("Sauvegarde restaurée.");
        setError("");
      }
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBackupBusy(false);
    }
  }

  async function handleExportHistoryXlsx() {
    setIsReportingBusy(true);
    try {
      const result = await padApi.exportDecisionHistoryXlsx();
      if (result?.filePath) {
        setNotice(`Historique exporté: ${result.filePath}`);
      }
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsReportingBusy(false);
    }
  }

  async function handleExportMaintenanceXlsx() {
    setIsReportingBusy(true);
    try {
      const result = await padApi.exportMaintenanceHistoryXlsx();
      if (result?.filePath) {
        setNotice(`Entretiens exportés: ${result.filePath}`);
      }
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsReportingBusy(false);
    }
  }

  async function exportCurrentPdf(suggestedName: string) {
    if (!window.padApp?.printing?.exportCurrentViewPdf) {
      window.print();
      return;
    }
    try {
      const result = await padApi.exportCurrentViewPdf(suggestedName);
      if (result?.filePath) {
        setNotice(`PDF généré : ${result.filePath}`);
      }
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    }
  }

  function handlePrintDecision() {
    void exportCurrentPdf("fiche-decision-pad");
  }

  function handleOpenDecisionHistory() {
    setActiveView("history");
    setNotice("La décision calculée est enregistrée automatiquement dans l'historique.");
    setError("");
    scrollPageTop();
  }

  function handleReviewHistoryDecision(row: DecisionHistoryItem) {
    const matchedRoad =
      (row.roadId ? allRoads.find((road) => road.id === row.roadId) : null) ||
      allRoads.find(
        (road) =>
          normalizeLabel(road.roadCode) === normalizeLabel(row.roadCode) &&
          normalizeLabel(road.designation) === normalizeLabel(row.roadDesignation)
      ) ||
      null;
    const matchedDegradation =
      degradations.find((item) => normalizeLabel(item.name) === normalizeLabel(row.degradationName)) || null;

    setSelectedRoadId(matchedRoad?.id ?? "");
    setSelectedDegradationId(matchedDegradation?.id ?? "");
    setSelectedMeasurementCampaignKey("");
    setMeasurementRows([]);
    setDeflectionValue(row.deflectionValue == null ? "" : String(row.deflectionValue));
    setAskDrainage(Boolean(String(row.drainageRecommendation || "").trim()));
    setDecisionFieldErrors({});
    setDecisionFormError("");
    setDecisionResult({
      road: matchedRoad || buildFallbackRoadFromHistory(row),
      degradation: matchedDegradation || buildFallbackDegradationFromHistory(row),
      probableCause: row.probableCause || "",
      maintenanceSolution: row.maintenanceSolution || "",
      contextualIntervention: row.contextualIntervention || null,
      deflection: {
        value: row.deflectionValue,
        severity: row.deflectionSeverity || "",
        recommendation: row.deflectionRecommendation || ""
      },
      drainage: {
        needsAttention: Boolean(row.drainageNeedsAttention),
        recommendation: row.drainageRecommendation || "",
        ruleId: null
      }
    });
    setActiveView("decision");
    setError("");
    setNotice("Décision historique rechargée dans l'aide à la décision.");
    scrollPageTop();
  }

  function handlePrepareMaintenanceFromDecision() {
    if (!decisionResult) {
      return;
    }
    resetMaintenanceForm();
    setMaintenanceRoadId(decisionResult.road.id);
    setMaintenanceDegradationCode(decisionResult.degradation.name || "");
    setMaintenanceType(
      (!isToDetermineIntervention(decisionResult.contextualIntervention) && decisionResult.contextualIntervention) ||
        toDeflectionRecommendationLabel(decisionResult.deflection.recommendation) ||
        decisionResult.maintenanceSolution ||
        ""
    );
    setMaintenanceStatus("PREVU");
    setMaintenanceDate(new Date().toISOString().slice(0, 10));
    setMaintenanceStateBefore(decisionResult.road.pavementState || "");
    setMaintenanceDeflectionBefore(
      decisionResult.deflection.value == null ? "" : String(decisionResult.deflection.value)
    );
    setMaintenanceSolutionApplied(decisionResult.maintenanceSolution || "");
    setMaintenanceObservation(
      [
        `Cause probable: ${decisionResult.probableCause || "-"}`,
        `Assainissement: ${decisionResult.drainage.recommendation || "-"}`,
        `Décision: ${
          (!isToDetermineIntervention(decisionResult.contextualIntervention) && decisionResult.contextualIntervention) ||
          toDeflectionRecommendationLabel(decisionResult.deflection.recommendation) ||
          "-"
        }`
      ].join(" | ")
    );
    setActiveView("maintenance");
    setNotice("Fiche d'entretien préremplie à partir du résultat automatique.");
    setError("");
    scrollPageTop();
  }

  function handlePrintActiveSheet() {
    if (!activeSheet) {
      return;
    }
    void exportCurrentPdf(activeSheet.title || activeSheet.name || "impression-feuille");
  }

  function renderSheetPrintButton(sheetTitle?: string, disabled?: boolean, disabledReason?: string) {
    const label = `Imprimer ${sheetTitle || "la feuille"}`;
    return (
      <button
        className="row-action row-action--print row-action--icon row-action--icon-sm"
        type="button"
        onClick={handlePrintActiveSheet}
        title={disabled ? disabledReason || label : label}
        aria-label={disabled ? disabledReason || label : label}
        disabled={Boolean(disabled)}
      >
        <Printer size={15} aria-hidden="true" />
      </button>
    );
  }

  function renderStandardSheetPrintHeader(sheetTitle?: string, sheetSubtitle?: string) {
    return (
      <>
        <div className="print-sheet-header">
          <div className="print-sheet-header__brand">
            <img className="print-sheet-header__logo" src={appLogoSrc} alt="Logo Port Autonome de Douala" />
            <div>
              <strong>{appName}</strong>
              <div>Pilotez la maintenance routière du PAD avec des décisions rapides et fiables.</div>
            </div>
          </div>
        </div>
        {sheetTitle || sheetSubtitle ? (
          <div className="print-sheet-heading">
            {sheetTitle ? <h2 className="print-sheet-header__title">{sheetTitle}</h2> : null}
            {sheetSubtitle ? <p className="print-sheet-header__subtitle">{sheetSubtitle}</p> : null}
          </div>
        ) : null}
      </>
    );
  }

  function renderStandardSheetPrintFooter() {
    return (
      <div className="print-sheet-footer">
        <span className="print-sheet-footer__left">{appName}</span>
        <span className="print-sheet-footer__page" aria-hidden="true" />
        <span className="print-sheet-footer__right">Version {appVersion}</span>
      </div>
    );
  }

  async function handleRefresh() {
    setIsBusy(true);
    try {
      await Promise.all([
        refreshStatus(),
        loadIntegrityReport(),
        loadDashboardSummary(),
        refreshDecisionCatalogs(),
        loadAllRoads(),
        loadAllRoadSections(),
        loadRoads(),
        loadMeasurementCampaigns(),
        loadHistory(),
        loadMaintenanceRows(),
        loadFeuil4Snapshot(),
        loadSolutionTemplates()
      ]);
      if (activeView.startsWith("sheet:") && activeSheetName !== "Feuil1") {
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
    const fieldErrors = validateDecisionForm();
    if (Object.keys(fieldErrors).length > 0) {
      setError(Object.values(fieldErrors)[0] || "Veuillez corriger les champs obligatoires.");
      scrollPageTop();
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
      await Promise.all([refreshStatus(), loadHistory(), loadDashboardSummary()]);
      setDecisionFieldErrors({});
      setDecisionFormError("");
      setError("");
      setNotice("Décision calculée et enregistrée dans l'historique.");
      scrollPageTop();
    } catch (err) {
      setError(toErrorMessage(err));
      scrollPageTop();
    } finally {
      setIsDecisionBusy(false);
    }
  }

  function handleManageIntegrityFromAlert() {
    setActiveView("dashboard");
    setShouldScrollToIntegritySection(true);
  }

  function handleDecisionCampaignSelection(campaignKey: string) {
    const normalizedKey = String(campaignKey || "").trim();
    setSelectedMeasurementCampaignKey(normalizedKey);
    setDecisionResult(null);
    if (!normalizedKey) {
      setMeasurementRows([]);
      return;
    }

    const campaign = measurementCampaigns.find((item) => item.campaignKey === normalizedKey) ?? null;
    if (campaign?.roadId) {
      setSelectedRoadId(campaign.roadId);
    }
    setError("");
    clearDecisionFieldError("roadId");
  }

  function handleUseMeasurementCampaign(campaign: MeasurementCampaignItem) {
    if (campaign.roadId) {
      setSelectedRoadId(campaign.roadId);
    }
    if (campaign.sapCode) {
      setSelectedSap(campaign.sapCode);
    }
    clearDecisionFieldError("roadId");
    setSelectedMeasurementCampaignKey(campaign.campaignKey);
    setDecisionResult(null);
    setActiveView("decision");
    setNotice("Campagne Feuil1 chargée dans l'aide à la décision. Choisis une ligne PK pour injecter D.");
    setError("");
  }

  function handleUseMeasurementInDecision(measurement: RoadMeasurementItem) {
    if (measurement.roadId) {
      setSelectedRoadId(measurement.roadId);
    }
    if (measurement.campaignKey) {
      setSelectedMeasurementCampaignKey(measurement.campaignKey);
    }
    if (measurement.deflectionDc != null) {
      setDeflectionValue(String(measurement.deflectionDc));
      clearDecisionFieldError("deflectionValue");
    }
    clearDecisionFieldError("roadId");
    setDecisionResult(null);
    setActiveView("decision");
    setNotice(`Mesure PK ${measurement.pkLabel || "-"} injectée dans D pour l'analyse.`);
    setError("");
  }

  function handleStartNewMeasurementCampaign() {
    setSelectedMeasurementCampaignKey("");
    setMeasurementRows([]);
    resetMeasurementCampaignForm(selectedRoadId);
    resetMeasurementRowForm();
    setIsMeasurementCampaignModalOpen(true);
    setIsMeasurementRowModalOpen(false);
    setNotice("Préparation d'une nouvelle campagne de mesure.");
    setError("");
  }

  function handleEditMeasurementCampaign() {
    if (!selectedMeasurementCampaign) {
      setError("Choisis d'abord une campagne à modifier.");
      setNotice("");
      return;
    }
    setMeasurementCampaignFieldErrors({});
    setIsMeasurementCampaignModalOpen(true);
    setIsMeasurementRowModalOpen(false);
    setError("");
  }

  function handleEditMeasurementRow(measurement: RoadMeasurementItem) {
    if (measurement.campaignKey && measurement.campaignKey !== selectedMeasurementCampaignKey) {
      setSelectedMeasurementCampaignKey(measurement.campaignKey);
    }
    setEditingMeasurementRowId(measurement.id);
    setMeasurementPkLabel(measurement.pkLabel || "");
    setMeasurementPkM(measurement.pkM == null ? "" : String(measurement.pkM));
    setMeasurementLectureLeft(measurement.lectureLeft == null ? "" : String(measurement.lectureLeft));
    setMeasurementLectureAxis(measurement.lectureAxis == null ? "" : String(measurement.lectureAxis));
    setMeasurementLectureRight(measurement.lectureRight == null ? "" : String(measurement.lectureRight));
    setMeasurementDeflectionLeft(measurement.deflectionLeft == null ? "" : String(measurement.deflectionLeft));
    setMeasurementDeflectionAxis(measurement.deflectionAxis == null ? "" : String(measurement.deflectionAxis));
    setMeasurementDeflectionRight(measurement.deflectionRight == null ? "" : String(measurement.deflectionRight));
    setMeasurementDeflectionAvg(measurement.deflectionAvg == null ? "" : String(measurement.deflectionAvg));
    setMeasurementStdDev(measurement.stdDev == null ? "" : String(measurement.stdDev));
    setMeasurementDeflectionDc(measurement.deflectionDc == null ? "" : String(measurement.deflectionDc));
    setMeasurementRowFieldErrors({});
    setIsMeasurementRowModalOpen(true);
    setNotice(`Modification de la ligne PK ${measurement.pkLabel || "-"}.`);
    setError("");
  }

  async function handleSaveMeasurementCampaign() {
    const fieldErrors = validateMeasurementCampaignForm();
    if (Object.keys(fieldErrors).length > 0) {
      setError(Object.values(fieldErrors)[0] || "Veuillez corriger les champs obligatoires.");
      setNotice("");
      setIsMeasurementCampaignModalOpen(true);
      return;
    }
    const roadId = Number(measurementCampaignRoadId);

    setIsMeasurementBusy(true);
    try {
      const saved = await padApi.upsertMeasurementCampaign({
        id: editingMeasurementCampaignId || undefined,
        roadId,
        sectionLabel: measurementCampaignSectionLabel,
        startLabel: measurementCampaignStartLabel,
        endLabel: measurementCampaignEndLabel,
        measurementDate: measurementCampaignDate,
        pkStartM: measurementCampaignPkStartM,
        pkEndM: measurementCampaignPkEndM
      });

      if (!saved) {
        throw new Error("Campagne de mesure non enregistrée.");
      }

      await Promise.all([refreshStatus(), loadDashboardSummary(), loadIntegrityReport(), loadMeasurementCampaigns()]);
      setSelectedRoadId(roadId);
      setSelectedMeasurementCampaignKey(saved.campaignKey);
      await loadMeasurementRows(saved.campaignKey);
      setIsMeasurementCampaignModalOpen(false);
      setNotice(
        `${editingMeasurementCampaignId ? "Campagne mise à jour." : "Campagne créée."} Sélectionne ou ajoute maintenant les lignes PK.`
      );
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
      setNotice("");
      setIsMeasurementCampaignModalOpen(true);
    } finally {
      setIsMeasurementBusy(false);
    }
  }

  async function handleDeleteMeasurementCampaign() {
    const campaign = selectedMeasurementCampaign;
    if (!campaign) {
      return;
    }
    if (!window.confirm("Supprimer cette campagne et toutes ses lignes PK ?")) {
      return;
    }

    setIsMeasurementBusy(true);
    try {
      await padApi.deleteMeasurementCampaign(campaign.id);
      await Promise.all([refreshStatus(), loadDashboardSummary(), loadIntegrityReport(), loadMeasurementCampaigns()]);
      setSelectedMeasurementCampaignKey("");
      setMeasurementRows([]);
      resetMeasurementCampaignForm(selectedRoadId);
      resetMeasurementRowForm();
      setIsMeasurementCampaignModalOpen(false);
      setIsMeasurementRowModalOpen(false);
      setNotice("Campagne supprimée.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
      setNotice("");
    } finally {
      setIsMeasurementBusy(false);
    }
  }

  function handleStartNewMeasurementRow() {
    if (!selectedMeasurementCampaignKey) {
      setError("Choisis ou crée d'abord une campagne de mesure.");
      setNotice("");
      return;
    }
    resetMeasurementRowForm();
    setIsMeasurementRowModalOpen(true);
    setNotice("Préparation d'une nouvelle ligne PK.");
    setError("");
  }

  async function handleSaveMeasurementRow() {
    if (!selectedMeasurementCampaignKey) {
      setError("Choisis ou crée d'abord une campagne de mesure.");
      setNotice("");
      return;
    }
    const fieldErrors = validateMeasurementRowForm();
    if (Object.keys(fieldErrors).length > 0) {
      setError(Object.values(fieldErrors)[0] || "Veuillez corriger les champs obligatoires.");
      setNotice("");
      setIsMeasurementRowModalOpen(true);
      return;
    }

    setIsMeasurementBusy(true);
    try {
      const saved = await padApi.upsertRoadMeasurement({
        id: editingMeasurementRowId || undefined,
        campaignKey: selectedMeasurementCampaignKey,
        pkLabel: measurementPkLabel,
        pkM: measurementPkM,
        lectureLeft: measurementLectureLeft,
        lectureAxis: measurementLectureAxis,
        lectureRight: measurementLectureRight,
        deflectionLeft: measurementDeflectionLeft,
        deflectionAxis: measurementDeflectionAxis,
        deflectionRight: measurementDeflectionRight,
        deflectionAvg: measurementDeflectionAvg,
        stdDev: measurementStdDev,
        deflectionDc: measurementDeflectionDc
      });

      if (!saved) {
        throw new Error("Ligne de mesure non enregistrée.");
      }

      await Promise.all([
        refreshStatus(),
        loadDashboardSummary(),
        loadIntegrityReport(),
        loadMeasurementCampaigns(),
        loadMeasurementRows(selectedMeasurementCampaignKey)
      ]);
      resetMeasurementRowForm();
      setIsMeasurementRowModalOpen(false);
      setNotice(`${editingMeasurementRowId ? "Ligne PK mise à jour." : "Ligne PK ajoutée."}`);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
      setNotice("");
      setIsMeasurementRowModalOpen(true);
    } finally {
      setIsMeasurementBusy(false);
    }
  }

  async function handleDeleteMeasurementRow(measurementId: number) {
    if (!window.confirm("Supprimer cette ligne PK ?")) {
      return;
    }

    setIsMeasurementBusy(true);
    try {
      await padApi.deleteRoadMeasurement(measurementId);
      await Promise.all([
        refreshStatus(),
        loadDashboardSummary(),
        loadIntegrityReport(),
        loadMeasurementCampaigns(),
        loadMeasurementRows(selectedMeasurementCampaignKey)
      ]);
      if (editingMeasurementRowId === measurementId) {
        resetMeasurementRowForm();
        setIsMeasurementRowModalOpen(false);
      }
      setNotice("Ligne PK supprimée.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
      setNotice("");
    } finally {
      setIsMeasurementBusy(false);
    }
  }

  function handleUseCentralRoad(
    road: RoadCatalogItem | null,
    missingMessage: string,
    successMessage: string,
    nextSapCode?: string
  ) {
    if (!road) {
      setError(missingMessage);
      return;
    }

    setSelectedRoadId(road.id);
    if (nextSapCode || road.sapCode) {
      setSelectedSap(nextSapCode || road.sapCode);
    }
    setDecisionResult(null);
    setActiveView("decision");
    setNotice(successMessage);
    setError("");
  }

  function handleEditSourceRow(row: SheetRow | null, missingMessage: string) {
    if (!row) {
      setError(missingMessage);
      return;
    }
    handleEdit(row);
  }

  function handleDeleteSourceRow(row: SheetRow | null, missingMessage: string) {
    if (!row) {
      setError(missingMessage);
      return;
    }
    void handleDeleteRow(row.id);
  }

  async function handleCreateSourceRowAndEdit(
    sheetName: "Feuil2" | "Feuil3" | "Feuil5" | "Feuil6",
    payload: SheetRowPayload,
    successMessage: string
  ) {
    setIsBusy(true);
    try {
      const createdRow = await padApi.createSheetRow(sheetName, payload);
      await Promise.all([
        refreshStatus(),
        loadRows(),
        loadDashboardSummary(),
        loadIntegrityReport(),
        refreshDecisionCatalogs(),
        loadAllRoads(),
        loadAllRoadSections(),
        loadRoads(),
        loadMeasurementCampaigns()
      ]);
      handleEdit(createdRow);
      setNotice(successMessage);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  function handleUseFeuil2Section(row: SheetRow) {
    const matchedRoad = resolveRoadFromFeuil2Row(row, allRoads);
    handleUseCentralRoad(
      matchedRoad,
      "Aucune voie normalisée correspondante n'a été trouvée pour cette section.",
      `Section ${toDisplay(row.C)} chargée dans l'aide à la décision.`
    );
  }

  function handleUseFeuil6Road(row: SheetRow) {
    const matchedRoad = resolveRoadFromFeuil6Row(row, allRoads);
    handleUseCentralRoad(
      matchedRoad,
      "Aucune voie normalisée correspondante n'a été trouvée pour cette entrée du répertoire.",
      `Voie ${matchedRoad?.roadCode || ""} chargée depuis le répertoire codifié.`
    );
  }

  function buildFeuil3SourcePayload(item: {
    roadLabel: string;
    designation: string;
    startLabel: string;
    endLabel: string;
    lengthM: number | null;
    facadeWidthM: number | null;
    surfaceType: string;
    pavementState: string;
    drainageType: string;
    drainageState: string;
    sidewalkMinM: number | null;
    interventionHint: string;
  }): SheetRowPayload {
    return {
      A: toSheetCellValue(item.roadLabel),
      B: toSheetCellValue(item.designation),
      C: toSheetCellValue(item.startLabel),
      D: toSheetCellValue(item.endLabel),
      E: toSheetCellValue(item.lengthM),
      F: toSheetCellValue(item.facadeWidthM),
      G: toSheetCellValue(item.surfaceType),
      H: toSheetCellValue(item.pavementState),
      I: toSheetCellValue(item.drainageType),
      J: toSheetCellValue(item.drainageState),
      K: toSheetCellValue(item.sidewalkMinM),
      L: toSheetCellValue(item.interventionHint || "à déterminer (A D)")
    };
  }

  function buildFeuil2SourcePayload(item: {
    tronconNo: string;
    sectionNo: string;
    roadLabel: string;
    designation: string;
    startLabel: string;
    endLabel: string;
    lengthM: number | null;
  }): SheetRowPayload {
    const tronconNo = deriveTronconNoFromSectionNo(item.sectionNo) || item.tronconNo;
    return {
      A: toSheetCellValue(tronconNo),
      B: toSheetCellValue(item.sectionNo),
      C: toSheetCellValue(item.roadLabel),
      D: toSheetCellValue(item.designation),
      E: toSheetCellValue(item.startLabel),
      F: toSheetCellValue(item.endLabel),
      G: toSheetCellValue(item.lengthM)
    };
  }

  function buildFeuil5SourcePayload(item: {
    tronconNo: string;
    sectionNo: string;
    roadLabel: string;
    designation: string;
    startLabel: string;
    endLabel: string;
    lengthM: number | null;
    facadeWidthM: number | null;
    surfaceType: string;
    pavementState: string;
    drainageType: string;
    drainageState: string;
    sidewalkMinM: number | null;
    parkingLeft: string;
    parkingRight: string;
    parkingOther: string;
  }): SheetRowPayload {
    const tronconNo = deriveTronconNoFromSectionNo(item.sectionNo) || item.tronconNo;
    return {
      A: toSheetCellValue(tronconNo),
      B: toSheetCellValue(item.sectionNo),
      C: toSheetCellValue(item.roadLabel),
      D: toSheetCellValue(item.designation),
      E: toSheetCellValue(item.startLabel),
      F: toSheetCellValue(item.endLabel),
      G: toSheetCellValue(item.lengthM),
      H: toSheetCellValue(item.facadeWidthM),
      I: toSheetCellValue(item.surfaceType),
      J: toSheetCellValue(item.pavementState),
      K: toSheetCellValue(item.drainageType),
      L: toSheetCellValue(item.drainageState),
      M: toSheetCellValue(item.sidewalkMinM),
      N: toSheetCellValue(item.parkingLeft),
      O: toSheetCellValue(item.parkingRight),
      P: toSheetCellValue(item.parkingOther)
    };
  }

  function buildFeuil6SourcePayload(item: {
    sapCode: string;
    roadType: string;
    roadCode: string;
    linearM: number | null;
    proposedName: string;
    itinerary: string;
    justification: string;
  }): SheetRowPayload {
    return {
      A: toSheetCellValue(item.sapCode),
      B: toSheetCellValue(item.roadType),
      C: toSheetCellValue(item.roadCode),
      D: toSheetCellValue(item.linearM),
      E: toSheetCellValue(item.proposedName),
      F: toSheetCellValue(item.itinerary),
      G: toSheetCellValue(item.justification)
    };
  }

  function hydrateFeuil5DraftCells(cells: Partial<Record<SheetColumnKey, string>>) {
    const next = { ...cells };
    const draftRow = buildDraftRow(next);
    const matchedRoad = resolveRoadFromFeuil2Row(draftRow, allRoads);

    if (matchedRoad) {
      next.C = next.C || matchedRoad.roadCode || "";
      next.D = next.D || matchedRoad.designation || "";
      next.E = next.E || matchedRoad.startLabel || "";
      next.F = next.F || matchedRoad.endLabel || "";
      next.G = next.G || toSheetCellValue(matchedRoad.lengthM);
      next.H = next.H || toSheetCellValue(matchedRoad.widthM);
      next.I = next.I || matchedRoad.surfaceType || "";
      next.J = next.J || matchedRoad.pavementState || "";
      next.K = next.K || matchedRoad.drainageType || "";
      next.L = next.L || matchedRoad.drainageState || "";
      next.M = next.M || toSheetCellValue(matchedRoad.sidewalkMinM);
      next.N = next.N || matchedRoad.parkingLeft || "";
      next.O = next.O || matchedRoad.parkingRight || "";
      next.P = next.P || matchedRoad.parkingOther || "";
    }

    const matchedSection = allRoadSections.find((section) => {
      if (normalizeLabel(section.sectionNo) !== normalizeLabel(next.B)) {
        return false;
      }
      if (matchedRoad?.id && section.roadId) {
        return matchedRoad.id === section.roadId;
      }
      return normalizeLabel(section.roadCode) === normalizeLabel(next.C);
    });

    if (matchedSection) {
      next.A = deriveTronconNoFromSectionNo(next.B) || next.A || matchedSection.tronconNo || "";
      next.B = next.B || matchedSection.sectionNo || "";
      next.C = next.C || matchedSection.roadCode || "";
      next.D = next.D || matchedSection.designation || "";
      next.E = next.E || matchedSection.startLabel || "";
      next.F = next.F || matchedSection.endLabel || "";
      next.G = next.G || toSheetCellValue(matchedSection.lengthM);
      next.H = next.H || toSheetCellValue(matchedSection.widthM);
      next.I = next.I || matchedSection.surfaceType || "";
      next.J = next.J || matchedSection.pavementState || "";
      next.K = next.K || matchedSection.drainageType || "";
      next.L = next.L || matchedSection.drainageState || "";
      next.M = next.M || toSheetCellValue(matchedSection.sidewalkMinM);
    }

    if (next.B) {
      next.A = deriveTronconNoFromSectionNo(next.B) || next.A;
    }

    return next;
  }

  function handleEdit(row: SheetRow) {
    if (!activeSheet) {
      return;
    }

    let nextCells = createEmptyCells(editableColumns);
    for (const column of editableColumns) {
      nextCells[column] = String(row[column] ?? "");
    }
    if (activeSheet.name === "Feuil2" || activeSheet.name === "Feuil5") {
      nextCells.A = deriveTronconNoFromSectionNo(nextCells.B) || nextCells.A;
    }
    if (activeSheet.name === "Feuil5") {
      nextCells = hydrateFeuil5DraftCells(nextCells);
    }

    setEditingRowId(row.id);
    setDraftCells(nextCells);
    setDraftFieldErrors({});
    setDraftFormError("");
    setNotice(`Édition de la ligne ${getDisplayRowNumber(row)}.`);
    setError("");
    sheetEditorRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function handleStartPrefilledFeuil5Draft(payload: SheetRowPayload, successMessage: string) {
    if (!activeSheet || activeSheet.name !== "Feuil5") {
      return;
    }

    let nextCells = createEmptyCells(editableColumns);
    for (const column of editableColumns) {
      nextCells[column] = String(payload[column] ?? "");
    }
    nextCells = hydrateFeuil5DraftCells(nextCells);

    setEditingRowId(null);
    setDraftCells(nextCells);
    setDraftFieldErrors({});
    setDraftFormError("");
    setNotice(successMessage);
    setError("");
    sheetEditorRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function handleStartNewRow() {
    resetDraft();
    setNotice("Saisie d'une nouvelle ligne.");
    setError("");
    sheetEditorRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function handleSaveRow() {
    if (!activeSheet) {
      return;
    }

    const draftValidation = validateSheetDraft(activeSheet, draftCells);
    const duplicateValidation = validateDraftDuplicates(activeSheet.name, draftCells);
    const mergedFieldErrors = {
      ...draftValidation.fieldErrors,
      ...duplicateValidation.fieldErrors
    };
    const draftError = draftValidation.formError || duplicateValidation.formError;

    if (draftError) {
      setDraftFieldErrors(mergedFieldErrors);
      setDraftFormError(draftError);
      setError(draftError);
      setNotice("");
      scrollPageTop();
      return;
    }

    setDraftFieldErrors({});
    setDraftFormError("");
    setIsBusy(true);
    try {
      const payload = toPayload(editableColumns, draftCells);
      let createdRow: SheetRow | null = null;
      const inferredSapCode =
        activeSheet.name === "Feuil2"
          ? parseFeuil2SapCode({
              A: draftCells.A ?? "",
              B: draftCells.B ?? "",
              H: draftCells.H ?? "",
              I: draftCells.I ?? ""
            } as SheetRow)
          : "";
      const isNewSap = Boolean(
        inferredSapCode && !sapSectors.some((item) => normalizeLabel(item.code) === normalizeLabel(inferredSapCode))
      );
      if (editingRowId) {
        await padApi.updateSheetRow(activeSheet.name, editingRowId, payload);
        const currentRow = rows.find((row) => row.id === editingRowId);
        setNotice(
          `Ligne ${currentRow ? getDisplayRowNumber(currentRow) : editingRowId} mise à jour.${
            isNewSap ? ` ${inferredSapCode} sera ajouté automatiquement à la liste des SAP.` : ""
          }`
        );
      } else {
        createdRow = await padApi.createSheetRow(activeSheet.name, payload);
        setNotice(`Nouvelle ligne ajoutée.${isNewSap ? ` ${inferredSapCode} sera ajouté automatiquement à la liste des SAP.` : ""}`);
      }

      const [, , , , , refreshedRoads] = await Promise.all([
        refreshStatus(),
        loadRows(),
        loadDashboardSummary(),
        loadIntegrityReport(),
        refreshDecisionCatalogs(),
        loadAllRoads(),
        loadAllRoadSections(),
        loadRoads(),
        loadMeasurementCampaigns()
      ]);
      if (activeSheet.name === "Feuil4") {
        await loadFeuil4Snapshot();
      }
      resetDraft();
      setError("");

      if (activeSheet.name === "Feuil6" && !editingRowId && createdRow) {
        const createdRoad = resolveRoadFromFeuil6Row(createdRow, refreshedRoads ?? []);
        if (createdRoad) {
          const sectionDraft = autofillDraftFromRoad(
            "Feuil2",
            createEmptyCells(["A", "B", "C", "D", "E", "F", "G"]),
            createdRoad
          );
          setPendingSheetDraft({
            sheetName: "Feuil2",
            cells: sectionDraft,
            notice: `Voie ${toDisplay(createdRoad.roadCode)} créée. Vous pouvez maintenant enregistrer sa première section.`
          });
          setActiveView("sheet:Feuil2");
        }
      }
    } catch (err) {
      const message = toErrorMessage(err);
      setDraftFormError(message);
      setError(message);
      scrollPageTop();
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
      await Promise.all([
        refreshStatus(),
        loadRows(),
        loadDashboardSummary(),
        loadIntegrityReport(),
        refreshDecisionCatalogs(),
        loadAllRoads(),
        loadAllRoadSections(),
        loadRoads(),
        loadMeasurementCampaigns()
      ]);
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
      await Promise.all([refreshStatus(), loadHistory(), loadDashboardSummary(), loadIntegrityReport()]);
      setNotice("Historique vidé.");
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsBusy(false);
    }
  }

  function resetDegradationEditor() {
    setEditingDegradationId(null);
    setSelectedDegradationId("");
    setDegradationNameDraft("");
    setDegradationCausesDraft("");
    setDegradationPreventiveCriterionDraft("");
    setDegradationTreatmentDetailsDraft("");
    setDegradationFieldErrors({});
    setDegradationFormError("");
  }

  function handleStartNewDegradation() {
    resetDegradationEditor();
    setShouldScrollToDegradationEditor(true);
  }

  function handleEditDegradation(item: DegradationItem) {
    setSelectedDegradationId(item.id);
    setShouldScrollToDegradationEditor(true);
  }

  async function handleSaveDegradation() {
    const fieldErrors = validateDegradationForm();
    if (Object.keys(fieldErrors).length > 0) {
      const message = Object.values(fieldErrors)[0] || "Veuillez corriger les champs obligatoires.";
      setDegradationFormError(message);
      setError(message);
      scrollPageTop();
      return;
    }

    setIsDegradationBusy(true);
    try {
      const saved = await padApi.upsertDegradation({
        id: editingDegradationId ?? undefined,
        name: degradationNameDraft.trim(),
        causes: degradationDraftCauses,
        preventiveCriterion: degradationPreventiveCriterionDraft.trim(),
        treatmentDetails: degradationTreatmentDetailsDraft.trim()
      });
      await refreshDecisionCatalogs();
      if (saved) {
        setSelectedDegradationId(saved.id);
        setEditingDegradationId(saved.id);
        setDegradationNameDraft(saved.name || "");
        setDegradationCausesDraft(saved.causes.join("\n"));
        setDegradationPreventiveCriterionDraft(saved.preventiveCriterion || "");
        setDegradationTreatmentDetailsDraft(saved.treatmentDetails || "");
      }
      setDegradationFieldErrors({});
      setDegradationFormError("");
      setNotice(editingDegradationId ? "Dégradation mise à jour." : "Dégradation créée.");
      setError("");
      window.requestAnimationFrame(() => {
        scrollPageTop();
      });
    } catch (err) {
      const message = toErrorMessage(err);
      setDegradationFormError(message);
      setError(message);
      scrollPageTop();
    } finally {
      setIsDegradationBusy(false);
    }
  }

  async function handleDeleteDegradation(degradationId?: number) {
    const targetId = Number(degradationId ?? editingDegradationId ?? 0);
    if (!Number.isFinite(targetId) || targetId <= 0) {
      setError("Sélectionne d'abord une dégradation à supprimer.");
      return;
    }

    const target = degradations.find((item) => item.id === targetId) ?? null;
    const label = target?.name || "cette dégradation";
    if (!window.confirm(`Supprimer ${label} du catalogue ?`)) {
      return;
    }

    setIsDegradationBusy(true);
    try {
      await padApi.deleteDegradation(targetId);
      await refreshDecisionCatalogs();
      if (selectedDegradationId === targetId || editingDegradationId === targetId) {
        resetDegradationEditor();
      }
      setNotice(`Dégradation supprimée : ${label}.`);
      setError("");
    } catch (err) {
      const message = toErrorMessage(err);
      setDegradationFormError(message);
      setError(message);
      scrollPageTop();
    } finally {
      setIsDegradationBusy(false);
    }
  }

  async function handleAssignSolutionTemplate() {
    if (!selectedDegradation) {
      setError("Sélectionne une dégradation.");
      scrollPageTop();
      return;
    }
    const fieldErrors = validateSolutionTemplateForm();
    if (Object.keys(fieldErrors).length > 0) {
      const message = Object.values(fieldErrors)[0] || "Veuillez corriger les champs obligatoires.";
      setSolutionFormError(message);
      setError(message);
      scrollPageTop();
      return;
    }

    setIsSolutionBusy(true);
    try {
      await padApi.assignTemplateToDegradation(selectedDegradation.code, selectedTemplateKey);
      await Promise.all([refreshDecisionCatalogs(), loadSolutionTemplates()]);
      setNotice("Modèle de solution appliqué à la dégradation.");
      setSolutionFieldErrors({});
      setSolutionFormError("");
      setError("");
    } catch (err) {
      const message = toErrorMessage(err);
      setSolutionFormError(message);
      setError(message);
      scrollPageTop();
    } finally {
      setIsSolutionBusy(false);
    }
  }

  async function handleSaveSolutionOverride() {
    if (!selectedDegradation) {
      setError("Sélectionne une dégradation.");
      scrollPageTop();
      return;
    }
    const fieldErrors = validateSolutionOverrideForm();
    if (Object.keys(fieldErrors).length > 0) {
      const message = Object.values(fieldErrors)[0] || "Veuillez corriger les champs obligatoires.";
      setSolutionFormError(message);
      setError(message);
      scrollPageTop();
      return;
    }

    setIsSolutionBusy(true);
    try {
      await padApi.setDegradationSolutionOverride(selectedDegradation.code, solutionDraft.trim());
      await refreshDecisionCatalogs();
      setNotice("Solution personnalisée enregistrée.");
      setSolutionFieldErrors({});
      setSolutionFormError("");
      setError("");
    } catch (err) {
      const message = toErrorMessage(err);
      setSolutionFormError(message);
      setError(message);
      scrollPageTop();
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
      setSolutionFieldErrors({});
      setSolutionFormError("");
      setError("");
    } catch (err) {
      const message = toErrorMessage(err);
      setSolutionFormError(message);
      setError(message);
      scrollPageTop();
    } finally {
      setIsSolutionBusy(false);
    }
  }

  function handleEditMaintenance(intervention: MaintenanceInterventionItem) {
    setEditingMaintenanceId(intervention.id);
    setMaintenanceFieldErrors({});
    setMaintenanceFormError("");
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
    setMaintenanceResponsibleName(intervention.responsibleName || "");
    setMaintenanceAttachmentPath(intervention.attachmentPath || "");
    setMaintenanceObservation(intervention.observation || "");
    setMaintenanceCostAmount(intervention.costAmount != null ? String(intervention.costAmount) : "");
    setError("");
    setNotice(`Édition entretien #${intervention.id}`);
    scrollPageTop();
  }

  async function handleSaveMaintenance() {
    const fieldErrors = validateMaintenanceForm();
    if (Object.keys(fieldErrors).length > 0) {
      const message = Object.values(fieldErrors)[0] || "Veuillez corriger les champs obligatoires.";
      setMaintenanceFormError(message);
      setError(message);
      scrollPageTop();
      return;
    }
    const parsedRoadId = Number(maintenanceRoadId);

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
      responsibleName: maintenanceResponsibleName.trim() || undefined,
      attachmentPath: maintenanceAttachmentPath.trim() || undefined,
      observation: maintenanceObservation.trim() || undefined,
      costAmount: maintenanceCostAmount.trim() !== "" ? Number(maintenanceCostAmount) : undefined
    };

    setIsMaintenanceBusy(true);
    try {
      await padApi.upsertMaintenanceIntervention(payload);
      await Promise.all([
        loadMaintenanceRows(),
        loadMaintenancePreview(selectedRoadId),
        refreshStatus(),
        loadDashboardSummary()
      ]);
      resetMaintenanceForm();
      setMaintenanceFieldErrors({});
      setMaintenanceFormError("");
      setError("");
      setNotice("Entretien enregistré.");
    } catch (err) {
      const message = toErrorMessage(err);
      setMaintenanceFormError(message);
      setError(message);
      scrollPageTop();
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
        refreshStatus(),
        loadDashboardSummary()
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

  async function handlePickMaintenanceAttachment() {
    setIsAttachmentBusy(true);
    try {
      const result = await padApi.pickMaintenanceAttachment();
      if (result) {
        setMaintenanceAttachmentPath(result.storedPath);
        setError("");
        setNotice(`Pièce jointe ajoutée: ${result.fileName} (${formatBytes(result.size)})`);
      }
    } catch (err) {
      setError(toErrorMessage(err));
    } finally {
      setIsAttachmentBusy(false);
    }
  }

  async function handleOpenMaintenanceAttachment(targetPath?: string) {
    const filePath = String(targetPath || maintenanceAttachmentPath || "").trim();
    if (!filePath) {
      setError("Aucune pièce jointe à ouvrir.");
      return;
    }

    try {
      await padApi.openMaintenanceAttachment(filePath);
      setError("");
    } catch (err) {
      setError(toErrorMessage(err));
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
    setNotice("Export CSV généré.");
  }

  function getSheetFieldSuggestions(sheetName: string | undefined, column: SheetColumnKey) {
    const sapCodes = uniqueValues(sapSectors.map((item) => item.code));
    const roadCodes = uniqueValues(allRoads.map((item) => item.roadCode));
    const roadDesignations = uniqueValues(allRoads.map((item) => item.designation));
    const surfaceTypes = uniqueValues([...allRoads.map((item) => item.surfaceType), "BB", "Mixte", "BB/Pavés", "Pavés"]);
    const pavementStates = uniqueValues([
      ...allRoads.map((item) => item.pavementState),
      "Bon",
      "Moy",
      "Moyen",
      "Mau",
      "Mauvais",
      "RAS",
      "En cours d'aménagement",
      "Non aménagée"
    ]);
    const drainageTypes = uniqueValues([...allRoads.map((item) => item.drainageType), "E,F", "C", "-"]);
    const drainageStates = uniqueValues([
      ...allRoads.map((item) => item.drainageState),
      "Bon",
      "Moy",
      "Mau",
      "Obstrué",
      "-"
    ]);
    const interventionHints = uniqueValues([
      ...allRoads.map((item) => item.interventionHint),
      "À déterminer (A D)",
      ...MAINTENANCE_TYPE_SUGGESTIONS
    ]);

    if (sheetName === "Feuil2") {
      if (column === "C") return roadCodes;
      if (column === "D") return roadDesignations;
      if (column === "H" || column === "I") return sapCodes;
    }

    if (sheetName === "Feuil3") {
      if (column === "A") return roadCodes;
      if (column === "B") return roadDesignations;
      if (column === "G") return surfaceTypes;
      if (column === "H") return pavementStates;
      if (column === "I") return drainageTypes;
      if (column === "J") return drainageStates;
      if (column === "L") return interventionHints;
    }

    if (sheetName === "Feuil5") {
      if (column === "C") return roadCodes;
      if (column === "D") return roadDesignations;
      if (column === "I") return surfaceTypes;
      if (column === "J") return pavementStates;
      if (column === "K") return drainageTypes;
      if (column === "L") return drainageStates;
    }

    if (sheetName === "Feuil6") {
      if (column === "B") return ["Rue", "Boulevard", "Avenue"];
      if (column === "C") return roadCodes;
      if (column === "E") return roadDesignations;
      if (column === "F") return uniqueValues(allRoads.map((item) => `${item.startLabel} à ${item.endLabel}`));
    }

    if (sheetName === "Feuil7" && column === "C") {
      return uniqueValues(degradations.map((item) => item.name));
    }

    return [];
  }

  function renderDashboardView() {
    const summary = dashboardSummary;
    const integrity = summary?.integrity || integrityReport;

    return (
      <main className="workspace workspace--full">
        <section className="panel table-panel table-panel--full">
          <h2>Pilotage & données</h2>
          <p className="muted">
            Tableau de bord décideur, assistant d’import, contrôle de cohérence et protection des données.
          </p>

          <div className="kpi-grid">
            <div className="kpi-card">
              <span className="kpi-card__label">Voies</span>
              <strong>{summary?.totals.roads ?? "-"}</strong>
            </div>
            <div className="kpi-card">
              <span className="kpi-card__label">Dégradations</span>
              <strong>{summary?.totals.degradations ?? "-"}</strong>
            </div>
            <div className="kpi-card">
              <span className="kpi-card__label">Entretiens en cours</span>
              <strong>{summary?.totals.pendingMaintenance ?? 0}</strong>
            </div>
            <div className="kpi-card">
              <span className="kpi-card__label">Entretiens terminés</span>
              <strong>{summary?.totals.completedMaintenance ?? 0}</strong>
            </div>
            <div className="kpi-card">
              <span className="kpi-card__label">Total frais entretien</span>
              <strong>{formatAmount(summary?.totals.estimatedBudget ?? null)}</strong>
            </div>
            <div className="kpi-card">
              <span className="kpi-card__label">Assainissement urgent</span>
              <strong>{summary?.totals.urgentDrainage ?? 0}</strong>
            </div>
          </div>

          <div className="dashboard-grid">
            <div className="card card--spaced card--full">
              <div className="dashboard-card__header">
                <div>
                  <h3>Assistant d’import Excel</h3>
                  <p className="muted">Prévisualise le fichier, contrôle les feuilles et estime le contenu avant import.</p>
                </div>
                <div className="row-buttons">
                  <button
                    className="row-action row-action--with-icon row-action--preview"
                    type="button"
                    onClick={handlePreviewImport}
                    disabled={isPreviewingImport || isBusy}
                  >
                    <Eye size={15} aria-hidden="true" />
                    <span>{isPreviewingImport ? "Analyse..." : "Prévisualiser"}</span>
                  </button>
                  <button className="primary icon-btn" type="button" onClick={handleImport} disabled={isBusy}>
                    <Upload size={16} aria-hidden="true" />
                    <span>Importer</span>
                  </button>
                  <button
                    className="row-action row-action--with-icon"
                    type="button"
                    onClick={() => setIsImportAssistantCollapsed((current) => !current)}
                  >
                    {isImportAssistantCollapsed ? <ChevronDown size={15} aria-hidden="true" /> : <ChevronUp size={15} aria-hidden="true" />}
                    <span>{isImportAssistantCollapsed ? "Déplier" : "Rétracter"}</span>
                  </button>
                </div>
              </div>

              {isImportAssistantCollapsed ? (
                <p className="muted">Bloc replié. Déplie-le pour afficher l’aperçu du fichier et les contrôles détaillés.</p>
              ) : importPreview ? (
                <>
                  <div className="dashboard-meta">
                    <span className={`status-pill ${importPreview.ready ? "status-pill--ok" : "status-pill--warning"}`}>
                      {importPreview.ready ? "Prêt à importer" : "Import à vérifier"}
                    </span>
                    <span>{importPreview.totals.rows} ligne(s) détectée(s)</span>
                    <span>{importPreview.totals.roads} voie(s) estimée(s)</span>
                    <span>{importPreview.totals.degradations} dégradation(s) estimée(s)</span>
                  </div>

                  <div className="table-wrap">
                    <table>
                      <thead>
                        <tr>
                          <th>Feuille</th>
                          <th>Présente</th>
                          <th>Lignes utiles</th>
                          <th>Colonnes attendues</th>
                        </tr>
                      </thead>
                      <tbody>
                        {importPreview.sheetPreviews.map((sheet) => (
                          <tr key={sheet.name}>
                            <td>
                              <strong>{getSheetDisplayName(sheet.name)}</strong>
                              <div className="muted">{sheet.title}</div>
                            </td>
                            <td>{sheet.present ? "Oui" : "Non"}</td>
                            <td>{sheet.rowCount}</td>
                            <td>{sheet.expectedColumns}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {importPreview.warnings.length > 0 ? (
                    <ul className="issue-list">
                      {importPreview.warnings.map((warning) => (
                        <li key={warning}>{warning}</li>
                      ))}
                    </ul>
                  ) : (
                    <p className="muted">Aucune alerte structurelle détectée sur le fichier prévisualisé.</p>
                  )}
                </>
              ) : (
                <p className="muted">Aucun aperçu chargé. Utilise “Prévisualiser” pour auditer un fichier Excel avant import.</p>
              )}
            </div>

            <div ref={integritySectionRef} className="card card--spaced">
              <div className="dashboard-card__header">
                <h3>Recette & cohérence</h3>
                <ShieldCheck size={18} aria-hidden="true" />
              </div>
              <p className="muted">
                {integrity?.status === "OK"
                  ? "La base est cohérente sur les contrôles actuellement implémentés."
                  : "Des points de cohérence métier restent à vérifier."}
              </p>
              {integrity?.issues?.length ? (
                <ul className="issue-list">
                  {integrity.issues.map((issue) => (
                    <li key={issue.code}>
                      <strong>{issue.level}</strong> · {issue.message} ({issue.count})
                    </li>
                  ))}
                </ul>
              ) : (
                <p className="muted">Aucune incohérence détectée.</p>
              )}
            </div>

            <div className="card card--spaced">
              <div className="dashboard-card__header">
                <h3>Protection des données</h3>
                <DatabaseBackup size={18} aria-hidden="true" />
              </div>
              <p className="muted">Sauvegarde JSON et restauration rapide de la base locale.</p>
              <div className="editor-actions">
                <button
                  className="row-action row-action--with-icon row-action--save"
                  type="button"
                  onClick={handleExportBackup}
                  disabled={isBackupBusy}
                >
                  <DatabaseBackup size={15} aria-hidden="true" />
                  <span>Sauvegarder</span>
                </button>
                <button
                  className="row-action row-action--with-icon row-action--restore"
                  type="button"
                  onClick={handleRestoreBackup}
                  disabled={isBackupBusy}
                >
                  <FolderOpen size={15} aria-hidden="true" />
                  <span>Restaurer</span>
                </button>
              </div>
            </div>

            <div className="card card--spaced">
              <div className="dashboard-card__header">
                <h3>Entretiens par statut</h3>
                <RefreshCw size={18} aria-hidden="true" />
              </div>
              <ul className="count-list">
                {(summary?.maintenanceByStatus ?? []).map((item) => (
                  <li key={item.label}>
                    <span>{toMaintenanceStatusLabel(item.label)}</span>
                    <strong>{item.count}</strong>
                  </li>
                ))}
              </ul>
            </div>

            <div className="card card--spaced">
              <div className="dashboard-card__header">
                <h3>Réseau par SAP</h3>
                <BarChart3 size={18} aria-hidden="true" />
              </div>
              <ul className="count-list">
                {(summary?.roadsBySap ?? []).map((item) => (
                  <li key={item.label}>
                    <span>{item.label}</span>
                    <strong>{item.count}</strong>
                  </li>
                ))}
              </ul>
            </div>

            <div className="card card--spaced">
              <div className="dashboard-card__header">
                <h3>État de la chaussée</h3>
                <Gauge size={18} aria-hidden="true" />
              </div>
              <ul className="count-list">
                {(summary?.roadsByState ?? []).map((item) => (
                  <li key={item.label}>
                    <span>{item.label}</span>
                    <strong>{item.count}</strong>
                  </li>
                ))}
              </ul>
            </div>

            <div className="card card--spaced">
              <div className="dashboard-card__header">
                <h3>Dégradations les plus fréquentes</h3>
                <BarChart3 size={18} aria-hidden="true" />
              </div>
              <ul className="count-list">
                {(summary?.topDegradations ?? []).map((item) => (
                  <li key={item.label}>
                    <span>{item.label}</span>
                    <strong>{item.count}</strong>
                  </li>
                ))}
              </ul>
            </div>

            <div className="card card--spaced card--full">
              <div className="dashboard-card__header">
                <h3>Suivi récent</h3>
                <FileSpreadsheet size={18} aria-hidden="true" />
              </div>
              {(summary?.recentMaintenance ?? []).length > 0 ? (
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Date</th>
                        <th>Voie</th>
                        <th>Type</th>
                        <th>Statut</th>
                        <th>Coût</th>
                      </tr>
                    </thead>
                    <tbody>
                      {summary?.recentMaintenance.map((item) => (
                        <tr key={item.id}>
                          <td>{item.interventionDate}</td>
                          <td>
                            {item.roadCode} - {item.roadDesignation}
                          </td>
                          <td>{item.interventionType}</td>
                          <td>{toMaintenanceStatusLabel(item.status)}</td>
                          <td>{formatAmount(item.costAmount)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <p className="muted">Aucun entretien récent enregistré.</p>
              )}
            </div>
          </div>
        </section>
      </main>
    );
  }

  function renderDecisionView() {
    return (
      <main className="workspace workspace--decision">
        <section className="panel decision-form">
          <h2>Aide à la décision maintenance</h2>
          <p className="muted">Sélectionne une voie, une dégradation et la valeur de déflexion (optionnelle).</p>

          {decisionFormError ? <p className="modal-feedback modal-feedback--error">{decisionFormError}</p> : null}

          <div className="cell-field">
            <label htmlFor="sap">Secteur SAP</label>
            <select id="sap" value={selectedSap} onChange={(event) => setSelectedSap(event.target.value)}>
              <option value="">Tous les secteurs</option>
              {sapSectors.map((sector) => (
                <option key={sector.code} value={sector.code}>
                  {sector.code}
                </option>
              ))}
            </select>
            <p className="field-help">Filtre les voies proposées dans la liste ci-dessous.</p>
          </div>

          <div className="cell-field">
            <label htmlFor="road-search">Recherche voie</label>
            <input
              id="road-search"
              value={roadSearch}
              onChange={(event) => setRoadSearch(event.target.value)}
              placeholder="Code ou désignation"
            />
            <p className="field-help">Recherche rapide par code ou par nom de voie.</p>
          </div>

          <div className={`cell-field${decisionFieldErrors.roadId ? " cell-field--error" : ""}`}>
            <label htmlFor="road">
              Voie <span className="field-label__required"> *</span>
            </label>
            <select
              id="road"
              className={decisionFieldErrors.roadId ? "cell-field--error" : undefined}
              value={selectedRoadId}
              onChange={(event) => {
                setSelectedRoadId(event.target.value ? Number(event.target.value) : "");
                clearDecisionFieldError("roadId");
              }}
            >
              <option value="">Sélectionner une voie</option>
              {roads.map((road) => (
                <option key={road.id} value={road.id}>
                  {road.sapCode || "SAP?"} | {road.roadCode} | {road.designation}
                </option>
              ))}
            </select>
            {decisionFieldErrors.roadId ? <p className="field-error">{decisionFieldErrors.roadId}</p> : null}
            <p className="field-help">Choisissez la voie exacte sur laquelle vous voulez calculer la décision.</p>
          </div>

          <div className="cell-field">
            <label htmlFor="measurement-campaign">Campagne</label>
            <select
              id="measurement-campaign"
              value={selectedMeasurementCampaignKey}
              onChange={(event) => handleDecisionCampaignSelection(event.target.value)}
            >
              <option value="">Sélectionner une campagne</option>
              {decisionMeasurementCampaigns.map((campaign) => (
                <option key={campaign.campaignKey} value={campaign.campaignKey}>
                  {campaign.roadCode || "Voie"} | {campaign.designation} | {campaign.measurementDate || "Sans date"}
                </option>
              ))}
            </select>
            <p className="field-help">
              Optionnel. Si une campagne existe, vous pouvez réutiliser directement une valeur de déflexion mesurée.
            </p>
          </div>

          {selectedMeasurementCampaign ? (
            <div className="card measurement-card">
              <div className="dashboard-card__header">
                <div>
                  <h3>Campagne Feuil1 liée</h3>
                  <p className="muted">Mesures réelles de déflexion disponibles pour la voie sélectionnée.</p>
                </div>
                <button
                  className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                  type="button"
                  onClick={() => handleUseMeasurementCampaign(selectedMeasurementCampaign)}
                >
                  <Gauge size={15} aria-hidden="true" />
                  <span>Charger</span>
                </button>
              </div>

              <div className="measurement-summary__grid">
                <div className="measurement-summary__item">
                  <span>Date</span>
                  <strong>{selectedMeasurementCampaign.measurementDate || "-"}</strong>
                </div>
                <div className="measurement-summary__item">
                  <span>Tronçon</span>
                  <strong>{selectedMeasurementCampaign.sectionLabel || "-"}</strong>
                </div>
                <div className="measurement-summary__item">
                  <span>PK début / fin</span>
                  <strong>
                    {selectedMeasurementCampaign.startLabel || "-"} / {selectedMeasurementCampaign.endLabel || "-"}
                  </strong>
                </div>
                <div className="measurement-summary__item">
                  <span>Mesures</span>
                  <strong>{selectedMeasurementCampaign.measurementCount}</strong>
                </div>
                <div className="measurement-summary__item">
                  <span>Dc max</span>
                  <strong>{formatMeasurementNumber(selectedMeasurementCampaign.maxDeflectionDc)}</strong>
                </div>
                <div className="measurement-summary__item">
                  <span>Dc moyen</span>
                  <strong>{formatMeasurementNumber(selectedMeasurementCampaign.avgDeflectionDc)}</strong>
                </div>
              </div>

              <div className="table-wrap measurement-picker">
                <table className="table--measurements table--measurements-compact">
                  <thead>
                    <tr>
                      <th className="col-actions">Action</th>
                      <th>PK</th>
                      <th>Dc</th>
                      <th>Defl.Brute.Moy</th>
                      <th>Écart type</th>
                    </tr>
                  </thead>
                  <tbody>
                    {measurementRows.map((measurement) => (
                      <tr key={measurement.id}>
                        <td className="col-actions">
                          <button
                            className="row-action row-action--use row-action--with-icon row-action--compact"
                            type="button"
                            onClick={() => handleUseMeasurementInDecision(measurement)}
                          >
                            <Gauge size={14} aria-hidden="true" />
                            <span>Utiliser D</span>
                          </button>
                        </td>
                        <td>{measurement.pkLabel || "-"}</td>
                        <td>{formatMeasurementNumber(measurement.deflectionDc)}</td>
                        <td>{formatMeasurementNumber(measurement.deflectionAvg)}</td>
                        <td>{formatMeasurementNumber(measurement.stdDev)}</td>
                      </tr>
                    ))}
                    {measurementRows.length === 0 ? (
                      <tr>
                        <td colSpan={5}>{isMeasurementLoading ? "Chargement..." : "Aucune mesure PK disponible."}</td>
                      </tr>
                    ) : null}
                  </tbody>
                </table>
              </div>
            </div>
          ) : null}

          <div className={`cell-field${decisionFieldErrors.degradationId ? " cell-field--error" : ""}`}>
            <label className="field-label--spaced" htmlFor="degradation">
              Dégradation <span className="field-label__required"> *</span>
            </label>
            <select
              id="degradation"
              className={decisionFieldErrors.degradationId ? "cell-field--error" : undefined}
              value={selectedDegradationId}
              onChange={(event) => {
                setSelectedDegradationId(event.target.value ? Number(event.target.value) : "");
                clearDecisionFieldError("degradationId");
              }}
            >
              <option value="">Sélectionner une dégradation</option>
              {degradations.map((item) => (
                <option key={item.id} value={item.id}>
                  {item.name}
                </option>
              ))}
            </select>
            {decisionFieldErrors.degradationId ? <p className="field-error">{decisionFieldErrors.degradationId}</p> : null}
            <p className="field-help">Choisissez la dégradation observée sur la voie.</p>
          </div>

          <div className={`cell-field${decisionFieldErrors.deflectionValue ? " cell-field--error" : ""}`}>
            <label htmlFor="deflection">Valeur de déflexion D</label>
            <input
              id="deflection"
              type="number"
              className={decisionFieldErrors.deflectionValue ? "cell-field--error" : undefined}
              value={deflectionValue}
              onChange={(event) => {
                setDeflectionValue(event.target.value);
                clearDecisionFieldError("deflectionValue");
              }}
              placeholder="Ex: 80"
            />
            {decisionFieldErrors.deflectionValue ? <p className="field-error">{decisionFieldErrors.deflectionValue}</p> : null}
            <p className="field-help">Optionnel. Laissez vide si vous voulez analyser sans valeur D.</p>
          </div>

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
                <strong>État chaussée:</strong> {renderRoadState(selectedRoad.pavementState, true)}
              </p>
              <p>
                <strong>État courant suivi:</strong> {renderRoadState(latestMaintenance?.stateAfter || selectedRoad.pavementState, true)}
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
                    <strong>État après:</strong> {renderRoadState(latestMaintenance.stateAfter, true)}
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

        </section>

        <section className="panel decision-output decision-print-view">
          <div className="dashboard-card__header">
            <div>
              <h2>Résultat automatique</h2>
              <p className="muted">Analyse calculée à partir de la voie, de la dégradation, de la déflexion et des règles métier.</p>
            </div>
            <button
              className="row-action row-action--with-icon row-action--nowrap row-action--compact row-action--print"
              type="button"
              onClick={handlePrintDecision}
              disabled={!decisionResult}
            >
              <Printer size={15} aria-hidden="true" />
              <span>Imprimer la fiche</span>
            </button>
          </div>
          {!decisionResult ? <p className="muted">Lance une analyse pour afficher la recommandation.</p> : null}

          {decisionResult ? (
            <>
              <div className="print-sheet-header">
                <div className="print-sheet-header__brand">
                  <img className="print-sheet-header__logo" src={appLogoSrc} alt="Logo Port Autonome de Douala" />
                  <div>
                    <strong>{appName}</strong>
                    <div>Fiche d'aide à la décision de maintenance routière</div>
                  </div>
                </div>
              </div>

              <div className="decision-grid">
                <div className="card decision-hero card--full">
                  <div className="decision-hero__eyebrow">Décision recommandée</div>
                  <h3 className="decision-hero__title">{decisionSummary?.finalRecommendation || "À confirmer"}</h3>
                  <div className="decision-hero__badges">
                    <span className="decision-badge decision-badge--neutral">
                      Sévérité : {toDeflectionSeverityLabel(decisionResult.deflection.severity)}
                    </span>
                    <span className={`decision-badge decision-badge--${decisionSummary?.confidence.tone || "neutral"}`}>
                      {decisionSummary?.confidence.label || "Décision calculée"}
                    </span>
                    {decisionResult.drainage.needsAttention ? (
                      <span className="decision-badge decision-badge--alert">Assainissement à surveiller</span>
                    ) : (
                      <span className="decision-badge decision-badge--ok">Assainissement maîtrisé</span>
                    )}
                  </div>
                  <p className="decision-hero__cause">
                    <strong>Cause probable :</strong> {decisionResult.probableCause || "Non renseignée"}
                  </p>
                  <p className="decision-hero__cause">
                    <strong>Critère préventif :</strong>{" "}
                    {decisionTreatment?.preventiveCriterion || decisionSummary?.finalRecommendation || "À préciser"}
                  </p>
                  {decisionTreatment?.otherCauses.length ? (
                    <div className="decision-hero__cause">
                      <strong>Autres causes recensées :</strong>
                      <ul className="decision-list decision-list--compact">
                        {decisionTreatment.otherCauses.map((cause, index) => (
                          <li key={`decision-other-cause-${index}`}>{cause}</li>
                        ))}
                      </ul>
                    </div>
                  ) : null}
                </div>

                <div className="card decision-card--with-bottom-gap">
                  <h3>Voie analysée</h3>
                  <p>
                    <strong>{decisionResult.road.designation}</strong> ({decisionResult.road.roadCode})
                  </p>
                  <p>
                    <strong>SAP :</strong> {decisionResult.road.sapCode || "-"}
                  </p>
                  <p>
                    <strong>Début / fin :</strong> {decisionResult.road.startLabel || "-"} / {decisionResult.road.endLabel || "-"}
                  </p>
                  <p>
                    <strong>État connu :</strong> {renderRoadState(decisionResult.road.pavementState, true)}
                  </p>
                  <p>
                    <strong>Validation cause :</strong> {dynamicFeuil4Snapshot?.causeMatch || "-"}
                  </p>
                </div>

                <div className="card decision-card--with-bottom-gap">
                  <h3>Dégradation et déflexion</h3>
                  <p>
                    <strong>Dégradation :</strong> {decisionResult.degradation.name}
                  </p>
                  <p>
                    <strong>D :</strong> {decisionResult.deflection.value ?? "non renseigné"}
                  </p>
                  <p>
                    <strong>Niveau :</strong> {toDeflectionSeverityLabel(decisionResult.deflection.severity)}
                  </p>
                  <p>
                    <strong>Orientation déflexion :</strong>{" "}
                    {toDeflectionRecommendationLabel(decisionResult.deflection.recommendation)}
                  </p>
                </div>

                <div className="card card--full">
                  <h3>Comment traiter cette dégradation</h3>
                  {decisionTreatment?.phaseLines.length ? (
                    <div className="decision-treatment-details">
                      <p className="decision-treatment-details__label">{decisionTreatment.phaseLabel}</p>
                      <ul className="decision-list">
                        {decisionTreatment.phaseLines.map((item, index) => (
                          <li key={`decision-treatment-${index}`}>{item}</li>
                        ))}
                      </ul>
                    </div>
                  ) : (
                    <p className="muted">Le mode opératoire détaillé n'est pas encore renseigné pour cette dégradation.</p>
                  )}
                </div>

                <div className="card card--full">
                  <h3>Règles appliquées</h3>
                  <ul className="decision-list">
                    {(decisionSummary?.rationale || []).map((item, index) => (
                      <li key={`decision-rationale-${index}`}>{item}</li>
                    ))}
                  </ul>
                </div>

                <div className="card">
                  <h3>Comparaison terrain</h3>
                  {latestDecisionCampaign ? (
                    <div className="decision-compare">
                      <div>
                        <span>Dernière campagne Feuil1</span>
                        <strong>{latestDecisionCampaign.measurementDate || "-"}</strong>
                        <small>
                          Dc max {formatMeasurementNumber(latestDecisionCampaign.maxDeflectionDc)} / Dc moyen{" "}
                          {formatMeasurementNumber(latestDecisionCampaign.avgDeflectionDc)}
                        </small>
                      </div>
                    </div>
                  ) : null}
                  {latestMaintenance ? (
                    <div className="decision-compare">
                      <div>
                        <span>Dernier entretien</span>
                        <strong>{latestMaintenance.interventionDate || "-"}</strong>
                        <small>
                          {toMaintenanceStatusLabel(latestMaintenance.status)} · {latestMaintenance.stateAfter || "État après non renseigné"}
                        </small>
                      </div>
                    </div>
                  ) : null}
                  {!latestDecisionCampaign && !latestMaintenance ? (
                    <p className="muted">Aucune campagne récente ni entretien enregistré pour comparer l'évolution.</p>
                  ) : null}
                </div>

                <div className="card">
                  <h3>Actions</h3>
                  <p className="muted">La décision est déjà enregistrée dans l'historique après calcul.</p>
                  <div className="decision-actions">
                    <button className="row-action row-action--with-icon row-action--compact" type="button" onClick={handleOpenDecisionHistory}>
                      <History size={15} aria-hidden="true" />
                      <span>Historique</span>
                    </button>
                    <button
                      className="row-action row-action--with-icon row-action--compact"
                      type="button"
                      onClick={handlePrepareMaintenanceFromDecision}
                    >
                      <ClipboardPlus size={15} aria-hidden="true" />
                      <span>Créer un entretien</span>
                    </button>
                  </div>
                </div>

                <div className="card card--full">
                  <h3>Justification de la voie</h3>
                  {String(decisionResult.road.justification || "").trim() ? (
                    <p>{decisionResult.road.justification}</p>
                  ) : (
                    <p className="muted">Aucune justification détaillée n'est enregistrée pour cette voie.</p>
                  )}
                </div>
              </div>

              <div className="print-sheet-footer">
                <span className="print-sheet-footer__left">{appName}</span>
                <span className="print-sheet-footer__page" aria-hidden="true" />
                <span className="print-sheet-footer__right">Version {appVersion}</span>
              </div>
            </>
          ) : null}
        </section>
      </main>
    );
  }

  function renderCatalogueView() {
    return (
      <main className="workspace workspace--full">
        <section className="panel table-panel table-panel--full sheet-print-view degradation-print-view">
          {renderStandardSheetPrintHeader(
            "Catalogue des voies",
            "Référence complète des voies par secteur SAP."
          )}
          <div className="dashboard-card__header">
            <div>
              <h2>Catalogue des voies</h2>
              <p className="muted">Référence complète des voies par secteur SAP.</p>
            </div>
            <div className="sheet-header-actions">
              <button
                className="row-action row-action--print row-action--icon row-action--icon-sm"
                type="button"
                onClick={() => void exportCurrentPdf("catalogue-des-voies-pad")}
                title="Imprimer le catalogue des voies"
                aria-label="Imprimer le catalogue des voies"
              >
                <Printer size={15} aria-hidden="true" />
              </button>
            </div>
          </div>

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
              placeholder="Recherche code/désignation/début/fin"
            />
            <span className="muted">{roads.length} voie(s)</span>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th className="col-actions">Action</th>
                  <th>SAP</th>
                  <th>Code</th>
                  <th>Désignation</th>
                  <th>Début</th>
                  <th>Fin</th>
                  <th>Longueur (m)</th>
                  <th>Largeur (m)</th>
                  <th>Revêtement</th>
                  <th>État chaussée</th>
                  <th>Type caniveaux</th>
                  <th>Description assain.</th>
                </tr>
              </thead>
              <tbody>
                {roads.map((road) => (
                  <tr key={road.id}>
                    <td className="col-actions">
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
                    <td>{renderRoadState(road.pavementState, true)}</td>
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

          {renderStandardSheetPrintFooter()}
        </section>
      </main>
    );
  }

  function renderDegradationsView() {
    const editorPrimaryCause = degradationDraftCauses[0] || "";
    const editorOtherCauses = degradationDraftCauses.slice(1);
    const editorTreatmentLines = getTreatmentPhaseLines(degradationTreatmentDetailsDraft);
    const editorTreatmentLabel = getTreatmentPhaseLabel(editorTreatmentLines.length);
    const editorTitle = editingDegradationId ? degradationNameDraft || "Dégradation sélectionnée" : "Nouvelle dégradation";

    return (
      <main className="workspace workspace--full">
        <section className="panel table-panel table-panel--full sheet-print-view degradation-print-view">
          {renderStandardSheetPrintHeader(
            "Dégradations et traitements",
            "Catalogue unifié des dégradations, causes potentielles et solutions de maintenance."
          )}
          <div className="dashboard-card__header">
            <div>
              <h2>Dégradations et traitements</h2>
              <p className="muted">
                Une seule vue métier regroupe la dégradation, ses causes potentielles et la façon de la traiter.
              </p>
            </div>
            <div className="sheet-header-actions">
              <div className="measurement-toolbar__meta">
                <span className="pill">Dégradations: {filteredDegradations.length}</span>
                <span className="pill">
                  Causes détaillées: {filteredDegradations.reduce((total, item) => total + item.causes.length, 0)}
                </span>
              </div>
              <button
                className="row-action row-action--save row-action--with-icon"
                type="button"
                onClick={handleStartNewDegradation}
              >
                <Plus size={15} aria-hidden="true" />
                Nouvelle dégradation
              </button>
              <button
                className="row-action row-action--print row-action--icon row-action--icon-sm"
                type="button"
                onClick={() => void exportCurrentPdf("degradations-et-traitements-pad")}
                title="Imprimer le catalogue des dégradations"
                aria-label="Imprimer le catalogue des dégradations"
              >
                <Printer size={15} aria-hidden="true" />
              </button>
            </div>
          </div>

          <div className="table-toolbar">
            <input
              value={degradationSearch}
              onChange={(event) => setDegradationSearch(event.target.value)}
              placeholder="Rechercher une dégradation, une cause ou un traitement"
            />
            <span className="muted">{degradations.length} ligne(s) métier</span>
          </div>

          <div className="table-wrap">
            <table className="table--degradation-catalog">
              <thead>
                <tr>
                  <th className="col-actions">Actions</th>
                  <th>Dégradations</th>
                  <th>Causes potentielles</th>
                  <th>Solutions de maintenance</th>
                </tr>
              </thead>
              <tbody>
                {filteredDegradations.map((item) => {
                  const primaryCause = getDegradationPrimaryCause(item);
                  const otherCauses = getDegradationOtherCauses(item);
                  const referenceTreatment = getDegradationReferenceTreatment(item);
                  const treatmentLines = getTreatmentPhaseLines(item.treatmentDetails || item.preventiveCriterion || "");
                  const treatmentLabel = getTreatmentPhaseLabel(treatmentLines.length);

                  return (
                    <tr
                      key={item.id}
                      className={`${selectedDegradationId === item.id ? "is-selected " : ""}is-clickable`}
                      onClick={() => setSelectedDegradationId(item.id)}
                    >
                      <td className="col-actions">
                        <div className="row-buttons">
                          <button
                            className="row-action row-action--evaluate row-action--icon"
                            type="button"
                            onClick={(event) => {
                              event.stopPropagation();
                              setSelectedDegradationId(item.id);
                              setActiveView("decision");
                            }}
                            title="Utiliser dans l'aide à la décision"
                            aria-label={`Utiliser ${item.name} dans l'aide à la décision`}
                          >
                            <Gauge size={16} aria-hidden="true" />
                          </button>
                          <button
                            className="row-action row-action--icon"
                            type="button"
                            onClick={(event) => {
                              event.stopPropagation();
                              handleEditDegradation(item);
                            }}
                            title="Éditer"
                            aria-label={`Éditer ${item.name}`}
                          >
                            <Pencil size={16} aria-hidden="true" />
                          </button>
                          <button
                            className="row-action row-action--danger row-action--icon"
                            type="button"
                            onClick={(event) => {
                              event.stopPropagation();
                              void handleDeleteDegradation(item.id);
                            }}
                            title="Supprimer"
                            aria-label={`Supprimer ${item.name}`}
                          >
                            <Trash2 size={16} aria-hidden="true" />
                          </button>
                        </div>
                      </td>
                      <td>
                        <div className="degradation-table-cell">
                          <strong>{item.name}</strong>
                          <span className="degradation-table-meta">{toSolutionSourceLabel(item.solutionSource)}</span>
                        </div>
                      </td>
                      <td>
                        <div className="table-cell-stack">
                          <strong>{primaryCause || "Aucune cause renseignée"}</strong>
                          {otherCauses.length > 0 ? (
                            <ul className="table-cell-list">
                              {otherCauses.map((cause) => (
                                <li key={`${item.id}-${cause}`}>{cause}</li>
                              ))}
                            </ul>
                          ) : (
                            <span>Aucune autre cause recensée.</span>
                          )}
                        </div>
                      </td>
                      <td>
                        <div className="table-cell-stack table-cell-stack--solution">
                          {treatmentLines.length > 0 ? (
                            <>
                              {treatmentLabel ? <span className="table-cell-stack__primary">{treatmentLabel}</span> : null}
                              <ul className="table-cell-list table-cell-list--muted">
                                {treatmentLines.map((line, index) => (
                                  <li key={`${item.id}-phase-${index}`}>{line}</li>
                                ))}
                              </ul>
                            </>
                          ) : (
                            <>
                              <span className="table-cell-stack__primary">
                                {referenceTreatment || "Aucune procédure détaillée disponible."}
                              </span>
                              <span>{trimPreviewText(item.preventiveCriterion || item.treatmentDetails || "", 260) || "Aucune procédure détaillée disponible."}</span>
                            </>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
                {filteredDegradations.length === 0 ? (
                  <tr>
                    <td colSpan={4}>Aucune dégradation trouvée.</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>

          <div
            ref={degradationEditorRef}
            className={`card card--spaced degradation-editor-card${isDegradationEditorHighlighted ? " card--highlighted" : ""}`}
          >
            <div className="dashboard-card__header">
              <div>
                <h3>{editingDegradationId ? `Fiche métier : ${editorTitle}` : "Nouvelle dégradation"}</h3>
                <p className="muted">
                  Modifie directement la dégradation, ses causes, son critère préventif et la procédure de traitement.
                </p>
              </div>
              <div className="row-buttons row-buttons--wrap">
                <button
                  className="row-action row-action--use row-action--with-icon"
                  type="button"
                  onClick={() => setActiveView("decision")}
                  disabled={!selectedDegradation}
                >
                  <Gauge size={15} aria-hidden="true" />
                  Utiliser dans l'aide à la décision
                </button>
                <button className="row-action" type="button" onClick={handleStartNewDegradation}>
                  Nouvelle ligne
                </button>
                <button className="row-action row-action--save" type="button" onClick={() => void handleSaveDegradation()} disabled={isDegradationBusy}>
                  {isDegradationBusy ? "Enregistrement..." : "Enregistrer"}
                </button>
                {editingDegradationId ? (
                  <button
                    className="row-action row-action--danger"
                    type="button"
                    onClick={() => void handleDeleteDegradation()}
                    disabled={isDegradationBusy}
                  >
                    Supprimer
                  </button>
                ) : null}
              </div>
            </div>

            {degradationFormError ? <p className="alert alert--danger">{degradationFormError}</p> : null}

            <div className="measurement-summary__grid degradation-summary-grid">
              <div className="measurement-summary__item">
                <span>Dégradation</span>
                <strong>{degradationNameDraft.trim() || "-"}</strong>
              </div>
              <div className="measurement-summary__item">
                <span>Causes potentielles</span>
                <strong>{degradationDraftCauses.length}</strong>
              </div>
              <div className="measurement-summary__item">
                <span>Traitement de référence</span>
                <strong>{degradationEditorReferenceTreatment}</strong>
              </div>
              <div className="measurement-summary__item">
                <span>Source</span>
                <strong>{selectedDegradation ? toSolutionSourceLabel(selectedDegradation.solutionSource) : "Saisie manuelle"}</strong>
              </div>
            </div>

            <div className="degradation-editor-grid">
              <label className="cell-field">
                <span>Dégradation *</span>
                <input
                  value={degradationNameDraft}
                  onChange={(event) => setDegradationNameDraft(event.target.value)}
                  placeholder="Ex: Nids de poule"
                />
                {degradationFieldErrors.name ? <p className="field-error">{degradationFieldErrors.name}</p> : null}
                <p className="field-help">Nom métier de la dégradation tel qu’il doit apparaître dans le catalogue.</p>
              </label>

              <label className="cell-field cell-field--wide">
                <span>Causes potentielles *</span>
                <textarea
                  rows={6}
                  value={degradationCausesDraft}
                  onChange={(event) => setDegradationCausesDraft(event.target.value)}
                  placeholder={"Une cause par ligne\nEx: Défaut localisé de la couche de roulement"}
                />
                {degradationFieldErrors.causes ? <p className="field-error">{degradationFieldErrors.causes}</p> : null}
                <p className="field-help">Ajoute une cause par ligne. La première ligne sera utilisée comme cause probable par défaut.</p>
              </label>

              <label className="cell-field">
                <span>Critère préventif</span>
                <textarea
                  rows={4}
                  value={degradationPreventiveCriterionDraft}
                  onChange={(event) => setDegradationPreventiveCriterionDraft(event.target.value)}
                  placeholder="Ex: Le traitement comprend quatre phases :"
                />
                <p className="field-help">Phrase courte qui annonce l’orientation de traitement.</p>
              </label>

              <label className="cell-field">
                <span>Comment traiter cette dégradation</span>
                <textarea
                  rows={7}
                  value={degradationTreatmentDetailsDraft}
                  onChange={(event) => setDegradationTreatmentDetailsDraft(event.target.value)}
                  placeholder={"Une étape par ligne\nEx: On délimite d'abord la zone à réparer"}
                />
                <p className="field-help">Procédure détaillée. Écris une étape par ligne pour faciliter la lecture dans l’aide à la décision.</p>
              </label>
            </div>

            <div className="degradation-fiche-grid">
              <div className="degradation-fiche-panel">
                <h4>Causes potentielles</h4>
                {editorPrimaryCause ? (
                  <div className="degradation-cause-primary">
                    <span className="degradation-cause-primary__label">Cause probable</span>
                    <strong>{editorPrimaryCause}</strong>
                  </div>
                ) : (
                  <p className="muted">Aucune cause détaillée n'est encore renseignée pour cette dégradation.</p>
                )}

                {editorOtherCauses.length > 0 ? (
                  <ul className="degradation-cause-list">
                    {editorOtherCauses.map((cause, index) => (
                      <li key={`editor-cause-${index}`}>{cause}</li>
                    ))}
                  </ul>
                ) : editorPrimaryCause ? (
                  <p className="muted">Aucune autre cause complémentaire n'est encore recensée.</p>
                ) : null}
              </div>

              <div className="degradation-fiche-panel">
                <h4>Solutions de maintenance</h4>
                <p>
                  <strong>Traitement de référence :</strong> {degradationEditorReferenceTreatment}
                </p>
                <p>
                  <strong>Critère préventif :</strong>{" "}
                  {degradationPreventiveCriterionDraft.trim() || "Aucun critère préventif détaillé n'est encore renseigné."}
                </p>
              </div>

              <div className="degradation-fiche-panel degradation-fiche-panel--full">
                <h4>Comment traiter cette dégradation</h4>
                <div className="decision-treatment-details">
                  {editorTreatmentLines.length > 0 ? (
                    <>
                      <p className="decision-treatment-details__label">{editorTreatmentLabel}</p>
                      <ul className="decision-list">
                        {editorTreatmentLines.map((item, index) => (
                          <li key={`editor-treatment-${index}`}>{item}</li>
                        ))}
                      </ul>
                    </>
                  ) : (
                    <p>Aucun mode opératoire détaillé n'est encore renseigné pour cette dégradation.</p>
                  )}
                </div>
              </div>
            </div>
          </div>

          {renderStandardSheetPrintFooter()}
        </section>
      </main>
    );
  }

  function renderMaintenanceView() {
    const hasMaintenanceAttachment = maintenanceAttachmentPath.trim().length > 0;

    return (
      <main className="workspace">
        <section className="panel editor-panel">
          <h2>Suivi des entretiens</h2>
          <p className="muted">Enregistre les interventions réalisées pour garder un état dynamique du réseau.</p>
          {maintenanceFormError ? <p className="modal-feedback modal-feedback--error">{maintenanceFormError}</p> : null}

          <label htmlFor="maintenance-road">
            Voie <span className="field-label__required">*</span>
          </label>
          <select
            id="maintenance-road"
            className={maintenanceFieldErrors.roadId ? "cell-field--error" : undefined}
            value={maintenanceRoadId}
            onChange={(event) => {
              setMaintenanceRoadId(event.target.value ? Number(event.target.value) : "");
              clearMaintenanceFieldError("roadId");
            }}
            disabled={isMaintenanceBusy}
          >
            <option value="">Sélectionner une voie</option>
            {allRoads.map((road) => (
              <option key={road.id} value={road.id}>
                {road.sapCode || "SAP?"} | {road.roadCode} | {road.designation}
              </option>
            ))}
          </select>
          {maintenanceFieldErrors.roadId ? <p className="field-error">{maintenanceFieldErrors.roadId}</p> : null}

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

          <label htmlFor="maintenance-type">
            Type d'entretien <span className="field-label__required">*</span>
          </label>
          <input
            id="maintenance-type"
            list="maintenance-type-options"
            className={maintenanceFieldErrors.type ? "cell-field--error" : undefined}
            value={maintenanceType}
            onChange={(event) => {
              setMaintenanceType(event.target.value);
              clearMaintenanceFieldError("type");
            }}
            placeholder="Ex: Curage caniveaux"
            disabled={isMaintenanceBusy}
          />
          {maintenanceFieldErrors.type ? <p className="field-error">{maintenanceFieldErrors.type}</p> : null}
          <datalist id="maintenance-type-options">
            {MAINTENANCE_TYPE_SUGGESTIONS.map((item) => (
              <option key={item} value={item} />
            ))}
          </datalist>

          <label htmlFor="maintenance-status">Statut</label>
          <select
            id="maintenance-status"
            value={maintenanceStatus}
            onChange={(event) => {
              setMaintenanceStatus(event.target.value as MaintenanceInterventionStatus);
              clearMaintenanceFieldError("completionDate");
            }}
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
              <label htmlFor="maintenance-date">
                Date prévue <span className="field-label__required">*</span>
              </label>
              <input
                id="maintenance-date"
                type="date"
                className={maintenanceFieldErrors.interventionDate ? "cell-field--error" : undefined}
                value={maintenanceDate}
                onChange={(event) => {
                  setMaintenanceDate(event.target.value);
                  clearMaintenanceFieldError("interventionDate");
                  clearMaintenanceFieldError("completionDate");
                }}
                disabled={isMaintenanceBusy}
              />
              {maintenanceFieldErrors.interventionDate ? (
                <p className="field-error">{maintenanceFieldErrors.interventionDate}</p>
              ) : null}
            </div>

            <div className={`cell-field${maintenanceFieldErrors.completionDate ? " cell-field--error" : ""}`}>
              <label htmlFor="maintenance-completion-date">
                Date réelle / clôture{maintenanceStatus === "TERMINE" ? <span className="field-label__required"> *</span> : null}
              </label>
              <input
                id="maintenance-completion-date"
                type="date"
                className={maintenanceFieldErrors.completionDate ? "cell-field--error" : undefined}
                value={maintenanceCompletionDate}
                onChange={(event) => {
                  setMaintenanceCompletionDate(event.target.value);
                  clearMaintenanceFieldError("completionDate");
                }}
                disabled={isMaintenanceBusy}
              />
              {maintenanceFieldErrors.completionDate ? (
                <p className="field-error">{maintenanceFieldErrors.completionDate}</p>
              ) : null}
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-state-before">
                État avant <span className="field-label__required">*</span>
              </label>
              <input
                id="maintenance-state-before"
                className={maintenanceFieldErrors.stateBefore ? "cell-field--error" : undefined}
                value={maintenanceStateBefore}
                onChange={(event) => {
                  setMaintenanceStateBefore(event.target.value);
                  clearMaintenanceFieldError("stateBefore");
                }}
                placeholder="État observé avant intervention"
                disabled={isMaintenanceBusy}
              />
              {maintenanceFieldErrors.stateBefore ? <p className="field-error">{maintenanceFieldErrors.stateBefore}</p> : null}
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
                className={maintenanceFieldErrors.deflectionBefore ? "cell-field--error" : undefined}
                value={maintenanceDeflectionBefore}
                onChange={(event) => {
                  setMaintenanceDeflectionBefore(event.target.value);
                  clearMaintenanceFieldError("deflectionBefore");
                }}
                disabled={isMaintenanceBusy}
              />
              {maintenanceFieldErrors.deflectionBefore ? (
                <p className="field-error">{maintenanceFieldErrors.deflectionBefore}</p>
              ) : null}
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-deflection-after">Déflexion après (D)</label>
              <input
                id="maintenance-deflection-after"
                type="number"
                className={maintenanceFieldErrors.deflectionAfter ? "cell-field--error" : undefined}
                value={maintenanceDeflectionAfter}
                onChange={(event) => {
                  setMaintenanceDeflectionAfter(event.target.value);
                  clearMaintenanceFieldError("deflectionAfter");
                }}
                disabled={isMaintenanceBusy}
              />
              {maintenanceFieldErrors.deflectionAfter ? (
                <p className="field-error">{maintenanceFieldErrors.deflectionAfter}</p>
              ) : null}
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
              <label htmlFor="maintenance-responsible">
                Responsable PAD <span className="field-label__required">*</span>
              </label>
              <input
                id="maintenance-responsible"
                className={maintenanceFieldErrors.responsibleName ? "cell-field--error" : undefined}
                value={maintenanceResponsibleName}
                onChange={(event) => {
                  setMaintenanceResponsibleName(event.target.value);
                  clearMaintenanceFieldError("responsibleName");
                }}
                placeholder="Ex: Chef section voirie"
                disabled={isMaintenanceBusy}
              />
              {maintenanceFieldErrors.responsibleName ? (
                <p className="field-error">{maintenanceFieldErrors.responsibleName}</p>
              ) : null}
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-attachment">Pièce jointe / photo</label>
              <input
                id="maintenance-attachment"
                value={maintenanceAttachmentPath}
                onChange={(event) => setMaintenanceAttachmentPath(event.target.value)}
                placeholder="Aucune pièce jointe sélectionnée"
                disabled={isMaintenanceBusy || isAttachmentBusy}
                readOnly
              />
              <div className="row-buttons row-buttons--compact">
                <button
                  className="row-action row-action--with-icon row-action--compact"
                  type="button"
                  onClick={handlePickMaintenanceAttachment}
                  disabled={isMaintenanceBusy || isAttachmentBusy || !supportsMaintenanceAttachments}
                  title={
                    supportsMaintenanceAttachments
                      ? "Joindre un fichier"
                      : "Redémarre l'application Electron pour activer les pièces jointes"
                  }
                >
                  <Paperclip size={15} aria-hidden="true" />
                  <span>{isAttachmentBusy ? "Ajout..." : "Joindre"}</span>
                </button>
                <button
                  className="row-action row-action--with-icon row-action--restore row-action--compact"
                  type="button"
                  onClick={() => handleOpenMaintenanceAttachment()}
                  disabled={
                    isMaintenanceBusy || isAttachmentBusy || !supportsMaintenanceAttachments || !hasMaintenanceAttachment
                  }
                  title={
                    !supportsMaintenanceAttachments
                      ? "Redémarre l'application Electron pour activer les pièces jointes"
                      : hasMaintenanceAttachment
                        ? "Ouvrir la pièce jointe"
                        : "Aucune pièce jointe disponible"
                  }
                >
                  <ExternalLink size={15} aria-hidden="true" />
                  <span>Ouvrir</span>
                </button>
                <button
                  className="row-action row-action--danger row-action--icon row-action--icon-sm"
                  type="button"
                  onClick={() => setMaintenanceAttachmentPath("")}
                  title="Retirer la pièce jointe"
                  aria-label="Retirer la pièce jointe"
                  disabled={
                    isMaintenanceBusy || isAttachmentBusy || !supportsMaintenanceAttachments || !hasMaintenanceAttachment
                  }
                >
                  <X size={15} aria-hidden="true" />
                </button>
              </div>
              <p className="field-help field-help--compact">
                {supportsMaintenanceAttachments
                  ? "Formats acceptés: PNG, JPEG, WEBP, PDF, DOC, DOCX, XLS, XLSX, TXT. Taille maximale: 2 Mo."
                  : "Pièces jointes indisponibles dans cette session. Ferme puis relance l'application Electron."}
              </p>
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-cost">
                Coût estimé <span className="field-label__required">*</span>
              </label>
              <input
                id="maintenance-cost"
                type="number"
                className={maintenanceFieldErrors.costAmount ? "cell-field--error" : undefined}
                value={maintenanceCostAmount}
                onChange={(event) => {
                  setMaintenanceCostAmount(event.target.value);
                  clearMaintenanceFieldError("costAmount");
                }}
                placeholder="FCFA"
                disabled={isMaintenanceBusy}
              />
              {maintenanceFieldErrors.costAmount ? <p className="field-error">{maintenanceFieldErrors.costAmount}</p> : null}
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

          <div className="table-toolbar table-toolbar--hexa">
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
            <button
              className="row-action row-action--with-icon"
              type="button"
              onClick={handleExportMaintenanceXlsx}
              disabled={isReportingBusy}
            >
              <FileSpreadsheet size={15} aria-hidden="true" />
              <span>Export XLSX</span>
            </button>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th className="col-actions">Actions</th>
                  <th>Date</th>
                  <th>Statut</th>
                  <th>SAP</th>
                  <th>Voie</th>
                  <th>Type</th>
                  <th>Dégradation</th>
                  <th>État avant</th>
                  <th>État après</th>
                  <th>D avant</th>
                  <th>D après</th>
                  <th>Solution appliquée</th>
                  <th>Prestataire</th>
                  <th>Responsable</th>
                  <th>Pièce jointe</th>
                  <th>Coût</th>
                </tr>
              </thead>
              <tbody>
                {maintenanceRows.map((item) => (
                  <tr key={item.id} className={editingMaintenanceId === item.id ? "is-selected" : ""}>
                    <td className="col-actions">
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
                    <td>{item.deflectionBefore != null ? String(item.deflectionBefore) : "-"}</td>
                    <td>{item.deflectionAfter != null ? String(item.deflectionAfter) : "-"}</td>
                    <td>{item.solutionApplied || "-"}</td>
                    <td>{item.contractorName || "-"}</td>
                    <td>{item.responsibleName || "-"}</td>
                    <td>
                      {item.attachmentPath ? (
                        <button
                          className="row-action row-action--with-icon row-action--restore"
                          type="button"
                          onClick={() => handleOpenMaintenanceAttachment(item.attachmentPath)}
                          title={item.attachmentPath}
                        >
                          <ExternalLink size={14} aria-hidden="true" />
                          <span>{getFileNameFromPath(item.attachmentPath)}</span>
                        </button>
                      ) : (
                        "-"
                      )}
                    </td>
                    <td>{formatAmount(item.costAmount)}</td>
                  </tr>
                ))}
                {maintenanceRows.length === 0 ? (
                  <tr>
                    <td colSpan={16}>{isMaintenanceLoading ? "Chargement..." : "Aucun entretien enregistré."}</td>
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
          <p className="muted">Historique des décisions calculées avec exports CSV et Excel.</p>

          <div className="table-toolbar table-toolbar--hexa">
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
            <button
              className="row-action row-action--with-icon"
              type="button"
              onClick={() => loadHistory()}
              disabled={isHistoryLoading}
            >
              <RefreshCw size={15} aria-hidden="true" />
              <span>Actualiser</span>
            </button>
            <button className="primary row-action--with-icon" type="button" onClick={exportHistoryCsv}>
              <FileSpreadsheet size={15} aria-hidden="true" />
              <span>Export CSV</span>
            </button>
            <button
              className="row-action row-action--with-icon"
              type="button"
              onClick={handleExportHistoryXlsx}
              disabled={isReportingBusy}
            >
              <FileSpreadsheet size={15} aria-hidden="true" />
              <span>Export XLSX</span>
            </button>
            <button className="row-action row-action--danger" type="button" onClick={handleClearHistory} disabled={isBusy}>
              Vider
            </button>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th className="col-actions">Action</th>
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
                    <td className="col-actions">
                      <button
                        className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                        type="button"
                        onClick={() => handleReviewHistoryDecision(row)}
                        title="Revoir dans l'aide à la décision"
                        aria-label="Revoir dans l'aide à la décision"
                      >
                        <Gauge size={14} aria-hidden="true" />
                        <span>Revoir</span>
                      </button>
                    </td>
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
                    <td colSpan={10}>{isHistoryLoading ? "Chargement..." : "Aucun historique disponible."}</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>
        </section>
      </main>
    );
  }

  function renderFeuil1View() {
    const campaign = selectedMeasurementCampaign;
    const campaignLabel = [campaign?.roadCode, campaign?.designation].filter(Boolean).join(" - ") || "Campagne sélectionnée";
    const canCreateMeasurementRow = Boolean(selectedMeasurementCampaignKey);

    return (
      <main className="workspace workspace--full">
        <section className="panel table-panel table-panel--full sheet-print-view">
          {renderStandardSheetPrintHeader(
            activeSheet ? getSheetDisplayName(activeSheet.name) : "Campagne",
            activeSheet ? getSheetPrintSubtitle(activeSheet.name) : getSheetPrintSubtitle("Feuil1")
          )}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet ? getSheetDisplayName(activeSheet.name) : "Campagne"}</h2>
              <p className="muted">
                Feuille de campagnes de mesures de déflexion rattachées aux voies réelles du réseau.
              </p>
            </div>
            <div className="sheet-header-actions">
              <button
                className="row-action row-action--evaluate row-action--with-icon"
                type="button"
                onClick={() => (campaign ? handleUseMeasurementCampaign(campaign) : undefined)}
                disabled={!campaign}
              >
                <Gauge size={16} aria-hidden="true" />
                <span>Utiliser dans l'aide à la décision</span>
              </button>
              {renderSheetPrintButton(
                activeSheet ? getSheetDisplayName(activeSheet.name) : "Campagne",
                !campaign,
                "Sélectionnez d'abord une campagne avant d'imprimer."
              )}
            </div>
          </div>

          <div className="measurement-actionbar">
            <div className="measurement-toolbar__field">
              <label htmlFor="feuil1-campaign">Campagne de mesure</label>
              <select
                id="feuil1-campaign"
                value={selectedMeasurementCampaignKey}
                onChange={(event) => handleDecisionCampaignSelection(event.target.value)}
              >
                <option value="">Sélectionner une campagne</option>
                {measurementCampaigns.map((item) => (
                  <option key={item.campaignKey} value={item.campaignKey}>
                    {item.roadCode || "Voie"} | {item.designation} | {item.measurementDate || "Sans date"}
                  </option>
                ))}
              </select>
              <p className="field-help">
                Choisis une campagne existante pour voir ses mesures, ou crée-en une nouvelle.
              </p>
            </div>
            <div className="measurement-toolbar__meta">
              <span className="pill">Campagnes: {measurementCampaigns.length}</span>
              <span className="pill">Mesures: {campaign?.measurementCount ?? 0}</span>
            </div>
            <div className="measurement-actionbar__buttons">
              <button
                className="row-action row-action--save row-action--with-icon row-action--nowrap"
                type="button"
                onClick={handleStartNewMeasurementCampaign}
                disabled={isMeasurementBusy}
              >
                <Plus size={16} aria-hidden="true" />
                <span>Nouvelle campagne</span>
              </button>
              <button
                className="row-action row-action--configure row-action--with-icon row-action--nowrap"
                type="button"
                onClick={handleEditMeasurementCampaign}
                disabled={isMeasurementBusy || !campaign}
              >
                <Pencil size={16} aria-hidden="true" />
                <span>Modifier campagne</span>
              </button>
              <button
                className="row-action row-action--evaluate row-action--with-icon row-action--nowrap"
                type="button"
                onClick={handleStartNewMeasurementRow}
                disabled={isMeasurementBusy || !canCreateMeasurementRow}
              >
                <Plus size={16} aria-hidden="true" />
                <span>Nouvelle ligne PK</span>
              </button>
            </div>
          </div>

          {campaign ? (
            <>
              <div className="card measurement-summary">
                <h3>Campagne active</h3>
                <div className="measurement-summary__grid">
                  <div className="measurement-summary__item">
                    <span>Voie</span>
                    <strong>{campaignLabel}</strong>
                  </div>
                  <div className="measurement-summary__item">
                    <span>SAP</span>
                    <strong>{campaign.sapCode || "-"}</strong>
                  </div>
                  <div className="measurement-summary__item">
                    <span>Date</span>
                    <strong>{campaign.measurementDate || "-"}</strong>
                  </div>
                  <div className="measurement-summary__item">
                    <span>Tronçon</span>
                    <strong>{campaign.sectionLabel || "-"}</strong>
                  </div>
                  <div className="measurement-summary__item">
                    <span>PK début / fin</span>
                    <strong>
                      {campaign.startLabel || "-"} / {campaign.endLabel || "-"}
                    </strong>
                  </div>
                  <div className="measurement-summary__item">
                    <span>Intervalle PK</span>
                    <strong>
                      {formatMeasurementNumber(campaign.pkStartM)} m / {formatMeasurementNumber(campaign.pkEndM)} m
                    </strong>
                  </div>
                  <div className="measurement-summary__item">
                    <span>Dc max</span>
                    <strong>{formatMeasurementNumber(campaign.maxDeflectionDc)}</strong>
                  </div>
                  <div className="measurement-summary__item">
                    <span>Dc moyen</span>
                    <strong>{formatMeasurementNumber(campaign.avgDeflectionDc)}</strong>
                  </div>
                </div>
              </div>

              <div className="table-wrap measurement-table-wrap">
                <table className="table--measurements">
                  <thead>
                    <tr>
                      <th className="col-actions" rowSpan={2}>Actions</th>
                      <th rowSpan={2}>PK</th>
                      <th colSpan={3}>Lecture comparateur 1/100mm</th>
                      <th rowSpan={2}>PK</th>
                      <th colSpan={6}>Déflexion 1/100mm</th>
                    </tr>
                    <tr>
                      <th>Gauche</th>
                      <th>Axe</th>
                      <th>Droit</th>
                      <th>Gauche</th>
                      <th>Axe</th>
                      <th>Droit</th>
                      <th>Defl.Brute.Moy</th>
                      <th>Écart type</th>
                      <th>Déflexion caractéristique Dc</th>
                    </tr>
                  </thead>
                  <tbody>
                    {measurementRows.map((measurement) => (
                      <tr key={measurement.id}>
                        <td className="col-actions">
                          <div className="row-buttons row-buttons--compact row-buttons--wrap">
                            <button
                              className="row-action row-action--use row-action--with-icon row-action--compact"
                              type="button"
                              onClick={() => handleUseMeasurementInDecision(measurement)}
                            >
                              <Gauge size={14} aria-hidden="true" />
                              <span>Utiliser D</span>
                            </button>
                            <button
                              className="row-action row-action--icon row-action--icon-sm"
                              type="button"
                              onClick={() => handleEditMeasurementRow(measurement)}
                              title="Éditer"
                              aria-label="Éditer"
                            >
                              <Pencil size={15} aria-hidden="true" />
                            </button>
                            <button
                              className="row-action row-action--danger row-action--icon row-action--icon-sm"
                              type="button"
                              onClick={() => handleDeleteMeasurementRow(measurement.id)}
                              title="Supprimer"
                              aria-label="Supprimer"
                            >
                              <Trash2 size={15} aria-hidden="true" />
                            </button>
                          </div>
                        </td>
                        <td>{measurement.pkLabel || "-"}</td>
                        <td>{formatMeasurementNumber(measurement.lectureLeft)}</td>
                        <td>{formatMeasurementNumber(measurement.lectureAxis)}</td>
                        <td>{formatMeasurementNumber(measurement.lectureRight)}</td>
                        <td>{measurement.pkLabel || "-"}</td>
                        <td>{formatMeasurementNumber(measurement.deflectionLeft)}</td>
                        <td>{formatMeasurementNumber(measurement.deflectionAxis)}</td>
                        <td>{formatMeasurementNumber(measurement.deflectionRight)}</td>
                        <td>{formatMeasurementNumber(measurement.deflectionAvg)}</td>
                        <td>{formatMeasurementNumber(measurement.stdDev)}</td>
                        <td>{formatMeasurementNumber(measurement.deflectionDc)}</td>
                      </tr>
                    ))}
                    {measurementRows.length === 0 ? (
                      <tr>
                        <td colSpan={12}>{isMeasurementLoading ? "Chargement..." : "Aucune mesure disponible."}</td>
                      </tr>
                    ) : null}
                  </tbody>
                </table>
              </div>
            </>
          ) : (
            <div className="card">
              <p className="muted">
                Sélectionne une campagne Feuil1 pour afficher la voie, la date, le tronçon et les mesures PK.
              </p>
            </div>
          )}
          {renderStandardSheetPrintFooter()}
        </section>

        {isMeasurementCampaignModalOpen ? (
          <div className="modal-backdrop" role="presentation" onClick={() => setIsMeasurementCampaignModalOpen(false)}>
            <div
              ref={measurementCampaignEditorRef}
              className="modal-card modal-card--wide"
              role="dialog"
              aria-modal="true"
              aria-labelledby="measurement-campaign-modal-title"
              onClick={(event) => event.stopPropagation()}
            >
              <div className="modal-card__header">
                <div>
                  <h3 id="measurement-campaign-modal-title">
                    {editingMeasurementCampaignId ? "Modifier la campagne" : "Nouvelle campagne"}
                  </h3>
                  <p className="muted">
                    Renseigne la voie, la date, le tronçon et les bornes de la campagne de mesures.
                  </p>
                </div>
                <button
                  className="row-action row-action--icon"
                  type="button"
                  onClick={() => setIsMeasurementCampaignModalOpen(false)}
                  aria-label="Fermer"
                  title="Fermer"
                >
                  <X size={16} aria-hidden="true" />
                </button>
              </div>

              {error ? <p className="modal-feedback modal-feedback--error">{error}</p> : null}
              {!error && notice ? <p className="modal-feedback modal-feedback--notice">{notice}</p> : null}

              <div className="maintenance-form-grid">
                <div className={`cell-field${measurementCampaignFieldErrors.roadId ? " cell-field--error" : ""}`}>
                  <label htmlFor="measurement-road">
                    Voie concernée <span className="field-label__required">*</span>
                  </label>
                  <select
                    id="measurement-road"
                    required
                    value={measurementCampaignRoadId}
                    onChange={(event) => {
                      clearMeasurementCampaignFieldError("roadId");
                      setMeasurementCampaignRoadId(event.target.value ? Number(event.target.value) : "");
                    }}
                    disabled={isMeasurementBusy}
                  >
                    <option value="">Choisir une voie</option>
                    {allRoads.map((road) => (
                      <option key={"measurement-road-" + road.id} value={road.id}>
                        {road.sapCode || "SAP?"} | {road.roadCode} | {road.designation}
                      </option>
                    ))}
                  </select>
                  <p className="field-help">Choisis la voie exacte sur laquelle les mesures ont été faites.</p>
                  {measurementCampaignFieldErrors.roadId ? <p className="field-error">{measurementCampaignFieldErrors.roadId}</p> : null}
                </div>
                <div className={`cell-field${measurementCampaignFieldErrors.measurementDate ? " cell-field--error" : ""}`}>
                  <label htmlFor="measurement-date">
                    Date de mesure <span className="field-label__required">*</span>
                  </label>
                  <input
                    id="measurement-date"
                    type="date"
                    required
                    value={measurementCampaignDate}
                    onChange={(event) => {
                      clearMeasurementCampaignFieldError("measurementDate");
                      setMeasurementCampaignDate(event.target.value);
                    }}
                    disabled={isMeasurementBusy}
                  />
                  <p className="field-help">Date du relevé de déflexion.</p>
                  {measurementCampaignFieldErrors.measurementDate ? <p className="field-error">{measurementCampaignFieldErrors.measurementDate}</p> : null}
                </div>
                <div className={`cell-field${measurementCampaignFieldErrors.sectionLabel ? " cell-field--error" : ""}`}>
                  <label htmlFor="measurement-section">
                    Nom du tronçon <span className="field-label__required">*</span>
                  </label>
                  <input
                    id="measurement-section"
                    required
                    value={measurementCampaignSectionLabel}
                    onChange={(event) => {
                      clearMeasurementCampaignFieldError("sectionLabel");
                      setMeasurementCampaignSectionLabel(event.target.value);
                    }}
                    placeholder="Ex: Rue du Port de Pêche (tronçon SCDP - DAP)"
                    disabled={isMeasurementBusy}
                  />
                  <p className="field-help">Écris le nom complet du tronçon ou de la portion de voie mesurée.</p>
                  {measurementCampaignFieldErrors.sectionLabel ? <p className="field-error">{measurementCampaignFieldErrors.sectionLabel}</p> : null}
                </div>
                <div className={`cell-field${measurementCampaignFieldErrors.startLabel ? " cell-field--error" : ""}`}>
                  <label htmlFor="measurement-start-label">
                    Point de départ <span className="field-label__required">*</span>
                  </label>
                  <input
                    id="measurement-start-label"
                    required
                    value={measurementCampaignStartLabel}
                    onChange={(event) => {
                      clearMeasurementCampaignFieldError("startLabel");
                      setMeasurementCampaignStartLabel(event.target.value);
                    }}
                    placeholder="Ex: SCDP"
                    disabled={isMeasurementBusy}
                  />
                  <p className="field-help">Nom du lieu où la campagne commence.</p>
                  {measurementCampaignFieldErrors.startLabel ? <p className="field-error">{measurementCampaignFieldErrors.startLabel}</p> : null}
                </div>
                <div className={`cell-field${measurementCampaignFieldErrors.endLabel ? " cell-field--error" : ""}`}>
                  <label htmlFor="measurement-end-label">
                    Point d'arrivée <span className="field-label__required">*</span>
                  </label>
                  <input
                    id="measurement-end-label"
                    required
                    value={measurementCampaignEndLabel}
                    onChange={(event) => {
                      clearMeasurementCampaignFieldError("endLabel");
                      setMeasurementCampaignEndLabel(event.target.value);
                    }}
                    placeholder="Ex: DAP"
                    disabled={isMeasurementBusy}
                  />
                  <p className="field-help">Nom du lieu où la campagne se termine.</p>
                  {measurementCampaignFieldErrors.endLabel ? <p className="field-error">{measurementCampaignFieldErrors.endLabel}</p> : null}
                </div>
                <div className={`cell-field${measurementCampaignFieldErrors.pkStartM ? " cell-field--error" : ""}`}>
                  <label htmlFor="measurement-pk-start">PK début (mètres)</label>
                  <input
                    id="measurement-pk-start"
                    type="number"
                    step="0.001"
                    value={measurementCampaignPkStartM}
                    onChange={(event) => {
                      clearMeasurementCampaignFieldError("pkStartM");
                      setMeasurementCampaignPkStartM(event.target.value);
                    }}
                    placeholder="Ex: 1100"
                    disabled={isMeasurementBusy}
                  />
                  <p className="field-help">Exemple : 1+100 devient 1100.</p>
                  {measurementCampaignFieldErrors.pkStartM ? <p className="field-error">{measurementCampaignFieldErrors.pkStartM}</p> : null}
                </div>
                <div className={`cell-field${measurementCampaignFieldErrors.pkEndM ? " cell-field--error" : ""}`}>
                  <label htmlFor="measurement-pk-end">PK fin (mètres)</label>
                  <input
                    id="measurement-pk-end"
                    type="number"
                    step="0.001"
                    value={measurementCampaignPkEndM}
                    onChange={(event) => {
                      clearMeasurementCampaignFieldError("pkEndM");
                      setMeasurementCampaignPkEndM(event.target.value);
                    }}
                    placeholder="Ex: 1730"
                    disabled={isMeasurementBusy}
                  />
                  <p className="field-help">Exemple : 1+730 devient 1730.</p>
                  {measurementCampaignFieldErrors.pkEndM ? <p className="field-error">{measurementCampaignFieldErrors.pkEndM}</p> : null}
                </div>
              </div>

              <div className="modal-card__actions">
                <button className="primary" type="button" onClick={handleSaveMeasurementCampaign} disabled={isMeasurementBusy}>
                  {isMeasurementBusy ? "Enregistrement..." : editingMeasurementCampaignId ? "Mettre à jour la campagne" : "Enregistrer la campagne"}
                </button>
                <button
                  className="row-action row-action--restore row-action--with-icon"
                  type="button"
                  onClick={() => resetMeasurementCampaignForm(selectedRoadId)}
                  disabled={isMeasurementBusy}
                >
                  <RefreshCw size={15} aria-hidden="true" />
                  <span>Réinitialiser</span>
                </button>
                <button
                  className="row-action row-action--danger"
                  type="button"
                  onClick={handleDeleteMeasurementCampaign}
                  disabled={isMeasurementBusy || !selectedMeasurementCampaign}
                >
                  Supprimer
                </button>
              </div>
            </div>
          </div>
        ) : null}

        {isMeasurementRowModalOpen ? (
          <div className="modal-backdrop" role="presentation" onClick={() => setIsMeasurementRowModalOpen(false)}>
            <div
              ref={measurementRowEditorRef}
              className="modal-card modal-card--wide"
              role="dialog"
              aria-modal="true"
              aria-labelledby="measurement-row-modal-title"
              onClick={(event) => event.stopPropagation()}
            >
              <div className="modal-card__header">
                <div>
                  <h3 id="measurement-row-modal-title">
                    {editingMeasurementRowId ? "Modifier la ligne PK" : "Nouvelle ligne PK"}
                  </h3>
                  <p className="muted">
                    Saisis les lectures comparateur et les valeurs de déflexion pour un PK précis.
                  </p>
                </div>
                <button
                  className="row-action row-action--icon"
                  type="button"
                  onClick={() => setIsMeasurementRowModalOpen(false)}
                  aria-label="Fermer"
                  title="Fermer"
                >
                  <X size={16} aria-hidden="true" />
                </button>
              </div>

              {error ? <p className="modal-feedback modal-feedback--error">{error}</p> : null}
              {!error && notice ? <p className="modal-feedback modal-feedback--notice">{notice}</p> : null}

              {selectedMeasurementCampaignKey ? (
                <>
                  <p className="field-help">
                    Campagne active : <strong>{campaignLabel}</strong>
                  </p>
                  {measurementRowFieldErrors.values ? <p className="modal-feedback modal-feedback--error">{measurementRowFieldErrors.values}</p> : null}
                  <div className="maintenance-form-grid measurement-entry-grid">
                    <div className={`cell-field${measurementRowFieldErrors.pkLabel ? " cell-field--error" : ""}`}>
                      <label htmlFor="measurement-pk-label">
                        PK affiché <span className="field-label__required">*</span>
                      </label>
                      <input
                        id="measurement-pk-label"
                        required
                        value={measurementPkLabel}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("pkLabel");
                          setMeasurementPkLabel(event.target.value);
                        }}
                        placeholder="Ex: 0.000 ou 50.000"
                        disabled={isMeasurementBusy}
                      />
                      <p className="field-help">Valeur affichée dans le tableau.</p>
                      {measurementRowFieldErrors.pkLabel ? <p className="field-error">{measurementRowFieldErrors.pkLabel}</p> : null}
                    </div>
                    <div className={`cell-field${measurementRowFieldErrors.pkM ? " cell-field--error" : ""}`}>
                      <label htmlFor="measurement-pk-m">
                        PK en mètres <span className="field-label__required">*</span>
                      </label>
                      <input
                        id="measurement-pk-m"
                        type="number"
                        step="0.001"
                        required
                        value={measurementPkM}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("pkM");
                          setMeasurementPkM(event.target.value);
                        }}
                        placeholder="Ex: 50"
                        disabled={isMeasurementBusy}
                      />
                      <p className="field-help">Sert au classement automatique des PK.</p>
                      {measurementRowFieldErrors.pkM ? <p className="field-error">{measurementRowFieldErrors.pkM}</p> : null}
                    </div>
                    <div className="cell-field">
                      <label htmlFor="measurement-lecture-left">Lecture comparateur - Gauche</label>
                      <input
                        id="measurement-lecture-left"
                        type="number"
                        step="0.001"
                        value={measurementLectureLeft}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("values");
                          setMeasurementLectureLeft(event.target.value);
                        }}
                        disabled={isMeasurementBusy}
                      />
                    </div>
                    <div className="cell-field">
                      <label htmlFor="measurement-lecture-axis">Lecture comparateur - Axe</label>
                      <input
                        id="measurement-lecture-axis"
                        type="number"
                        step="0.001"
                        value={measurementLectureAxis}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("values");
                          setMeasurementLectureAxis(event.target.value);
                        }}
                        disabled={isMeasurementBusy}
                      />
                    </div>
                    <div className="cell-field">
                      <label htmlFor="measurement-lecture-right">Lecture comparateur - Droit</label>
                      <input
                        id="measurement-lecture-right"
                        type="number"
                        step="0.001"
                        value={measurementLectureRight}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("values");
                          setMeasurementLectureRight(event.target.value);
                        }}
                        disabled={isMeasurementBusy}
                      />
                    </div>
                    <div className="cell-field">
                      <label htmlFor="measurement-deflection-left">Déflexion - Gauche</label>
                      <input
                        id="measurement-deflection-left"
                        type="number"
                        step="0.001"
                        value={measurementDeflectionLeft}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("values");
                          setMeasurementDeflectionLeft(event.target.value);
                        }}
                        disabled={isMeasurementBusy}
                      />
                    </div>
                    <div className="cell-field">
                      <label htmlFor="measurement-deflection-axis">Déflexion - Axe</label>
                      <input
                        id="measurement-deflection-axis"
                        type="number"
                        step="0.001"
                        value={measurementDeflectionAxis}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("values");
                          setMeasurementDeflectionAxis(event.target.value);
                        }}
                        disabled={isMeasurementBusy}
                      />
                    </div>
                    <div className="cell-field">
                      <label htmlFor="measurement-deflection-right">Déflexion - Droit</label>
                      <input
                        id="measurement-deflection-right"
                        type="number"
                        step="0.001"
                        value={measurementDeflectionRight}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("values");
                          setMeasurementDeflectionRight(event.target.value);
                        }}
                        disabled={isMeasurementBusy}
                      />
                    </div>
                    <div className="cell-field">
                      <label htmlFor="measurement-deflection-avg">Déflexion brute moyenne</label>
                      <input
                        id="measurement-deflection-avg"
                        type="number"
                        step="0.001"
                        value={measurementDeflectionAvg}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("values");
                          setMeasurementDeflectionAvg(event.target.value);
                        }}
                        disabled={isMeasurementBusy}
                      />
                    </div>
                    <div className="cell-field">
                      <label htmlFor="measurement-std-dev">Écart type</label>
                      <input
                        id="measurement-std-dev"
                        type="number"
                        step="0.001"
                        value={measurementStdDev}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("values");
                          setMeasurementStdDev(event.target.value);
                        }}
                        disabled={isMeasurementBusy}
                      />
                    </div>
                    <div className="cell-field">
                      <label htmlFor="measurement-deflection-dc">Déflexion caractéristique Dc</label>
                      <input
                        id="measurement-deflection-dc"
                        type="number"
                        step="0.001"
                        value={measurementDeflectionDc}
                        onChange={(event) => {
                          clearMeasurementRowFieldError("values");
                          setMeasurementDeflectionDc(event.target.value);
                        }}
                        disabled={isMeasurementBusy}
                      />
                      <p className="field-help">Cette valeur peut être injectée dans le champ D de l'aide à la décision.</p>
                    </div>
                  </div>

                  <div className="modal-card__actions">
                    <button className="primary" type="button" onClick={handleSaveMeasurementRow} disabled={isMeasurementBusy}>
                      {isMeasurementBusy ? "Enregistrement..." : editingMeasurementRowId ? "Mettre à jour la ligne" : "Enregistrer la ligne"}
                    </button>
                    <button
                      className="row-action row-action--restore row-action--with-icon"
                      type="button"
                      onClick={resetMeasurementRowForm}
                      disabled={isMeasurementBusy}
                    >
                      <RefreshCw size={15} aria-hidden="true" />
                      <span>Réinitialiser</span>
                    </button>
                    <button
                      className="row-action row-action--danger"
                      type="button"
                      onClick={() => (editingMeasurementRowId ? handleDeleteMeasurementRow(editingMeasurementRowId) : undefined)}
                      disabled={isMeasurementBusy || !editingMeasurementRowId}
                    >
                      Supprimer
                    </button>
                  </div>
                </>
              ) : (
                <p className="muted">Crée ou sélectionne d'abord une campagne pour saisir les lignes PK.</p>
              )}
            </div>
          </div>
        ) : null}
      </main>
    );
  }


  function renderSheetEditorPanel() {
    if (!activeSheet || activeSheet.name === "Feuil1") {
      return null;
    }

    const editorTitle =
      activeSheet.name === "Feuil2"
        ? "Référentiel central des sections"
        : activeSheet.name === "Feuil6"
          ? "Référentiel central des voies"
          : activeSheet.title;
    const editorHelp =
      activeSheet.name === "Feuil2"
        ? "Choisissez d'abord une voie existante. Les informations d'identité sont reprises automatiquement ; vous ne saisissez ici que les informations de section."
        : activeSheet.name === "Feuil6"
          ? "Ajoute ou modifie ici la voie maître utilisée dans tout le système. Choisissez ou saisissez le SAP, puis renseignez séparément le début et la fin de la voie."
          : activeSheet.name === "Feuil3"
            ? "Choisissez d'abord une voie existante. Les informations d'identité sont reprises automatiquement ; vous ne saisissez ici que le diagnostic technique."
          : activeSheet.name === "Feuil5"
            ? "Choisissez d'abord une voie existante. Les informations d'identité sont reprises automatiquement ; vous ne saisissez ici que les compléments utiles."
          : "Les champs marqués d'un * sont obligatoires.";

    const renderEditorField = (column: SheetColumnKey) => {
      const suggestions = getSheetFieldSuggestions(activeSheet?.name, column);
      const inputId = `cell-${column}`;
      const datalistId = suggestions.length > 0 ? `cell-suggestions-${activeSheet?.name ?? "sheet"}-${column}` : undefined;
      const useTextarea =
        (activeSheet?.name === "Feuil3" && ["H", "J", "L"].includes(column)) ||
        (activeSheet?.name === "Feuil5" && column === "P");
      const fieldPlaceholder = getSheetFieldPlaceholder(activeSheet?.name, column) || suggestions[0] || "Entrez une valeur";
      const fieldHelp = getSheetFieldHelpText(activeSheet?.name, column);

      const commonProps = {
        id: inputId,
        value: draftCells[column] ?? "",
        required: isSheetFieldRequired(activeSheet.name, column),
        onChange: (event: ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
          handleDraftCellChange(column, event.target.value);
        }
      };

      return (
        <div className={`cell-field${draftFieldErrors[column] ? " cell-field--error" : ""}`} key={column}>
          <label htmlFor={inputId}>
            {getColumnLabel(activeSheet, column)}
            {isSheetFieldRequired(activeSheet.name, column) ? <span className="field-label__required"> *</span> : null}
          </label>
          {useTextarea ? (
            <textarea
              {...commonProps}
              className={`input-textarea${draftFieldErrors[column] ? " cell-field--error" : ""}`}
              rows={column === "L" ? 4 : 3}
              placeholder={fieldPlaceholder}
            />
          ) : (
            <>
              <input
                {...commonProps}
                className={draftFieldErrors[column] ? "cell-field--error" : undefined}
                list={datalistId}
                placeholder={fieldPlaceholder}
              />
              {datalistId ? (
                <datalist id={datalistId}>
                  {suggestions.map((item) => (
                    <option key={`${column}-${item}`} value={item} />
                  ))}
                </datalist>
              ) : null}
            </>
          )}
          {draftFieldErrors[column] ? <p className="field-error">{draftFieldErrors[column]}</p> : null}
          {fieldHelp ? <p className="field-help">{fieldHelp}</p> : null}
        </div>
      );
    };

    if (activeSheet.name === "Feuil2") {
      const selectedRoad = resolveDraftRoadMatch("Feuil2", draftCells);
      const selectedRoadId = selectedRoad?.id ?? "";

      return (
        <section ref={sheetEditorRef} className="panel editor-panel editor-panel--sheet">
          <h2>{editorTitle}</h2>
          <p className="field-help">{editorHelp}</p>
          {draftFormError ? <p className="modal-feedback modal-feedback--error">{draftFormError}</p> : null}

          <div className={`cell-field${draftFieldErrors.C ? " cell-field--error" : ""}`}>
            <label htmlFor="feuil2-road-selector">
              Voie concernée <span className="field-label__required"> *</span>
            </label>
            <select
              id="feuil2-road-selector"
              className={draftFieldErrors.C ? "cell-field--error" : undefined}
              value={selectedRoadId}
              onChange={(event) => {
                const nextRoadId = event.target.value ? Number(event.target.value) : 0;
                clearDraftFieldError("C");
                setDraftFormError("");
                setDraftCells((prev) => {
                  if (!nextRoadId) {
                    return {
                      ...prev,
                      A: "",
                      B: "",
                      C: "",
                      D: "",
                      E: "",
                      F: "",
                      G: ""
                    };
                  }
                  const road = allRoads.find((item) => item.id === nextRoadId);
                  if (!road) {
                    return prev;
                  }
                  return autofillDraftFromRoad(
                    "Feuil2",
                    {
                      ...prev,
                      C: road.roadCode,
                      D: road.designation
                    },
                    road
                  );
                });
              }}
            >
              <option value="">Choisir une voie existante</option>
              {allRoads.map((road) => (
                <option key={`feuil2-road-${road.id}`} value={road.id}>
                  {road.sapCode || "SAP?"} | {road.roadCode} | {road.designation}
                </option>
              ))}
            </select>
            {draftFieldErrors.C ? <p className="field-error">{draftFieldErrors.C}</p> : null}
            <p className="field-help">
              Choisissez une voie déjà présente dans le référentiel. Le code, la désignation, les bornes et la
              longueur sont repris automatiquement.
            </p>
          </div>

          {selectedRoad || draftCells.C || draftCells.D ? (
            <div className="card card--spaced feuil5-identity-card">
              <h3>Informations reprises automatiquement</h3>
              <div className="feuil5-identity-grid">
                <div className="feuil5-identity-item">
                  <span>SAP</span>
                  <strong>{toDisplay(selectedRoad?.sapCode || "")}</strong>
                </div>
                <div className="feuil5-identity-item">
                  <span>Voie</span>
                  <strong>{toDisplay(draftCells.C)}</strong>
                </div>
                <div className="feuil5-identity-item feuil5-identity-item--wide">
                  <span>Désignation</span>
                  <strong>{toDisplay(draftCells.D)}</strong>
                </div>
                <div className="feuil5-identity-item feuil5-identity-item--wide">
                  <span>Début / fin</span>
                  <strong>{`${toDisplay(draftCells.E)} / ${toDisplay(draftCells.F)}`}</strong>
                </div>
                <div className="feuil5-identity-item">
                  <span>Longueur (m)</span>
                  <strong>{toDisplay(draftCells.G)}</strong>
                </div>
              </div>
            </div>
          ) : null}

          <div className="cells-grid">
            <div className="cell-field">
              <label htmlFor="feuil2-troncon-auto">N° tronçon</label>
              <input id="feuil2-troncon-auto" value={draftCells.A ?? ""} readOnly placeholder="Calculé automatiquement" />
              <p className="field-help">Le tronçon est repris automatiquement depuis la partie après `_` du numéro de section.</p>
            </div>
            {renderEditorField("B")}
          </div>

          <div className="editor-actions">
            <button className="primary" type="button" onClick={handleSaveRow} disabled={isBusy || !activeSheet}>
              {editingRowId ? "Enregistrer" : "Ajouter"}
            </button>
            <button className="row-action" type="button" onClick={handleStartNewRow} disabled={isBusy}>
              Nouvelle ligne
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
      );
    }

    if (activeSheet.name === "Feuil3") {
      const selectedRoad = resolveDraftRoadMatch("Feuil3", draftCells);
      const selectedRoadId = selectedRoad?.id ?? "";

      return (
        <section ref={sheetEditorRef} className="panel editor-panel editor-panel--sheet">
          <h2>{editorTitle}</h2>
          <p className="field-help">{editorHelp}</p>
          {draftFormError ? <p className="modal-feedback modal-feedback--error">{draftFormError}</p> : null}

          <div className={`cell-field${draftFieldErrors.A ? " cell-field--error" : ""}`}>
            <label htmlFor="feuil3-road-selector">
              Voie concernée <span className="field-label__required"> *</span>
            </label>
            <select
              id="feuil3-road-selector"
              className={draftFieldErrors.A ? "cell-field--error" : undefined}
              value={selectedRoadId}
              onChange={(event) => {
                const nextRoadId = event.target.value ? Number(event.target.value) : 0;
                clearDraftFieldError("A");
                setDraftFormError("");
                setDraftCells((prev) => {
                  if (!nextRoadId) {
                    return {
                      ...prev,
                      A: "",
                      B: "",
                      C: "",
                      D: "",
                      E: ""
                    };
                  }
                  const road = allRoads.find((item) => item.id === nextRoadId);
                  if (!road) {
                    return prev;
                  }
                  return autofillDraftFromRoad(
                    "Feuil3",
                    {
                      ...prev,
                      A: road.roadCode,
                      B: road.designation
                    },
                    road
                  );
                });
              }}
            >
              <option value="">Choisir une voie existante</option>
              {allRoads.map((road) => (
                <option key={`feuil3-road-${road.id}`} value={road.id}>
                  {road.sapCode || "SAP?"} | {road.roadCode} | {road.designation}
                </option>
              ))}
            </select>
            {draftFieldErrors.A ? <p className="field-error">{draftFieldErrors.A}</p> : null}
            <p className="field-help">
              Choisissez une voie déjà présente dans le référentiel. La désignation, les bornes et la longueur sont
              reprises automatiquement.
            </p>
          </div>

          {selectedRoad || draftCells.A || draftCells.B ? (
            <div className="card card--spaced feuil5-identity-card">
              <h3>Informations reprises automatiquement</h3>
              <div className="feuil5-identity-grid">
                <div className="feuil5-identity-item">
                  <span>SAP</span>
                  <strong>{toDisplay(selectedRoad?.sapCode || "")}</strong>
                </div>
                <div className="feuil5-identity-item">
                  <span>Voie</span>
                  <strong>{toDisplay(draftCells.A)}</strong>
                </div>
                <div className="feuil5-identity-item feuil5-identity-item--wide">
                  <span>Désignation</span>
                  <strong>{toDisplay(draftCells.B)}</strong>
                </div>
                <div className="feuil5-identity-item feuil5-identity-item--wide">
                  <span>Début / fin</span>
                  <strong>{`${toDisplay(draftCells.C)} / ${toDisplay(draftCells.D)}`}</strong>
                </div>
                <div className="feuil5-identity-item">
                  <span>Longueur (m)</span>
                  <strong>{toDisplay(draftCells.E)}</strong>
                </div>
              </div>
            </div>
          ) : null}

          <div className="card card--spaced">
            <h3>Périmètre du diagnostic</h3>
            <p className="muted">
              Cette fiche porte l'état de la chaussée, l'assainissement observé et l'intervention à prévoir. Les
              compléments latéraux comme le stationnement et les trottoirs se renseignent dans l'onglet Compléments.
            </p>
          </div>

          <div className="cells-grid">
            {(["F", "G", "H", "I", "J", "K", "L"] as SheetColumnKey[]).map((column) => renderEditorField(column))}
          </div>

          <div className="editor-actions">
            <button className="primary" type="button" onClick={handleSaveRow} disabled={isBusy || !activeSheet}>
              {editingRowId ? "Enregistrer" : "Ajouter"}
            </button>
            <button className="row-action" type="button" onClick={handleStartNewRow} disabled={isBusy}>
              Nouvelle ligne
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
      );
    }

    if (activeSheet.name === "Feuil5") {
      const selectedRoad = resolveDraftRoadMatch("Feuil5", draftCells);
      const selectedRoadId = selectedRoad?.id ?? "";

      return (
        <section ref={sheetEditorRef} className="panel editor-panel editor-panel--sheet">
          <h2>{editorTitle}</h2>
          <p className="field-help">{editorHelp}</p>
          {draftFormError ? <p className="modal-feedback modal-feedback--error">{draftFormError}</p> : null}

          <div className={`cell-field${draftFieldErrors.C || draftFieldErrors.D ? " cell-field--error" : ""}`}>
            <label htmlFor="feuil5-road-selector">
              Voie concernée <span className="field-label__required"> *</span>
            </label>
            <select
              id="feuil5-road-selector"
              className={draftFieldErrors.C || draftFieldErrors.D ? "cell-field--error" : undefined}
              value={selectedRoadId}
              onChange={(event) => {
                const nextRoadId = event.target.value ? Number(event.target.value) : 0;
                clearDraftFieldError("C");
                clearDraftFieldError("D");
                setDraftFormError("");
                setDraftCells((prev) => {
                  if (!nextRoadId) {
                    return {
                      ...prev,
                      A: "",
                      B: "",
                      C: "",
                      D: "",
                      E: "",
                      F: "",
                      G: ""
                    };
                  }
                  const road = allRoads.find((item) => item.id === nextRoadId);
                  if (!road) {
                    return prev;
                  }
                  return autofillDraftFromRoad(
                    "Feuil5",
                    {
                      ...prev,
                      C: road.roadCode,
                      D: road.designation
                    },
                    road
                  );
                });
              }}
            >
              <option value="">Choisir une voie existante</option>
              {allRoads.map((road) => (
                <option key={`feuil5-road-${road.id}`} value={road.id}>
                  {road.sapCode || "SAP?"} | {road.roadCode} | {road.designation}
                </option>
              ))}
            </select>
            {draftFieldErrors.C || draftFieldErrors.D ? (
              <p className="field-error">{draftFieldErrors.C || draftFieldErrors.D}</p>
            ) : null}
            <p className="field-help">
              Choisissez une voie déjà présente dans le référentiel. Le tronçon, la section, les bornes et la longueur
              sont repris automatiquement.
            </p>
          </div>

          {selectedRoad || draftCells.C || draftCells.D ? (
            <div className="card card--spaced feuil5-identity-card">
              <h3>Informations reprises automatiquement</h3>
              <div className="feuil5-identity-grid">
                <div className="feuil5-identity-item">
                  <span>SAP</span>
                  <strong>{toDisplay(selectedRoad?.sapCode || "")}</strong>
                </div>
                <div className="feuil5-identity-item">
                  <span>N° tronçon</span>
                  <strong>{toDisplay(draftCells.A)}</strong>
                </div>
                <div className="feuil5-identity-item">
                  <span>N° section</span>
                  <strong>{toDisplay(draftCells.B)}</strong>
                </div>
                <div className="feuil5-identity-item">
                  <span>Voie</span>
                  <strong>{toDisplay(draftCells.C)}</strong>
                </div>
                <div className="feuil5-identity-item feuil5-identity-item--wide">
                  <span>Désignation</span>
                  <strong>{toDisplay(draftCells.D)}</strong>
                </div>
                <div className="feuil5-identity-item feuil5-identity-item--wide">
                  <span>Début / fin</span>
                  <strong>{`${toDisplay(draftCells.E)} / ${toDisplay(draftCells.F)}`}</strong>
                </div>
                <div className="feuil5-identity-item">
                  <span>Longueur (m)</span>
                  <strong>{toDisplay(draftCells.G)}</strong>
                </div>
              </div>
            </div>
          ) : null}

          <div className="card card--spaced">
            <h3>Ce que vous renseignez ici</h3>
            <p className="muted">
              Cette fiche complète la section avec les largeurs utiles, le contexte d'assainissement, les trottoirs et
              le stationnement. Le revêtement et l'état de la chaussée restent pilotés par l'onglet Diagnostic.
            </p>
          </div>

          <div className="cells-grid">
            {(["H", "K", "L", "M", "N", "O", "P"] as SheetColumnKey[]).map((column) => renderEditorField(column))}
          </div>

          <div className="editor-actions">
            <button className="primary" type="button" onClick={handleSaveRow} disabled={isBusy || !activeSheet}>
              {editingRowId ? "Enregistrer" : "Ajouter"}
            </button>
            <button className="row-action" type="button" onClick={handleStartNewRow} disabled={isBusy}>
              Nouvelle ligne
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
      );
    }

    if (activeSheet.name === "Feuil6") {
      const bounds = splitBoundsValue(draftCells.F);
      const sapOptions = centralSapCodes.length > 0 ? centralSapCodes : ["SAP1", "SAP2", "SAP3", "SAP4"];
      const sapListId = sapOptions.length > 0 ? "feuil6-sap-suggestions" : undefined;
      const startSuggestions = uniqueValues(allRoads.map((item) => item.startLabel)).filter(Boolean);
      const endSuggestions = uniqueValues(allRoads.map((item) => item.endLabel)).filter(Boolean);
      const startListId = startSuggestions.length > 0 ? "feuil6-start-suggestions" : undefined;
      const endListId = endSuggestions.length > 0 ? "feuil6-end-suggestions" : undefined;

      const updateFeuil6Bounds = (part: "start" | "end", value: string) => {
        clearDraftFieldError("F");
        setDraftFormError("");
        setDraftCells((prev) => {
          const currentBounds = splitBoundsValue(prev.F);
          const nextStart = part === "start" ? value : currentBounds.startLabel;
          const nextEnd = part === "end" ? value : currentBounds.endLabel;
          return {
            ...prev,
            F: formatBoundsValue(nextStart, nextEnd)
          };
        });
      };

      return (
        <section ref={sheetEditorRef} className="panel editor-panel editor-panel--sheet">
          <h2>{editorTitle}</h2>
          <p className="field-help">{editorHelp}</p>
          {draftFormError ? <p className="modal-feedback modal-feedback--error">{draftFormError}</p> : null}

          <div className="cells-grid">
            <div className={`cell-field${draftFieldErrors.A ? " cell-field--error" : ""}`}>
              <label htmlFor="feuil6-sap-input">
                SAP <span className="field-label__required"> *</span>
              </label>
              <input
                id="feuil6-sap-input"
                value={draftCells.A ?? ""}
                className={draftFieldErrors.A ? "cell-field--error" : undefined}
                list={sapListId}
                placeholder="Ex: SAP5"
                onChange={(event) => handleDraftCellChange("A", event.target.value)}
              />
              {sapListId ? (
                <datalist id={sapListId}>
                  {sapOptions.map((sapCode) => (
                    <option key={`feuil6-sap-${sapCode}`} value={sapCode} />
                  ))}
                </datalist>
              ) : null}
              {draftFieldErrors.A ? <p className="field-error">{draftFieldErrors.A}</p> : null}
              <p className="field-help">SAP de rattachement de la voie. Vous pouvez choisir un SAP existant ou saisir un nouveau code, par exemple SAP5.</p>
            </div>

            {(["B", "C", "D", "E"] as SheetColumnKey[]).map((column) => renderEditorField(column))}

            <div className={`cell-field${draftFieldErrors.F ? " cell-field--error" : ""}`}>
              <label htmlFor="feuil6-start-label">
                Début <span className="field-label__required"> *</span>
              </label>
              <input
                id="feuil6-start-label"
                value={bounds.startLabel}
                onChange={(event) => updateFeuil6Bounds("start", event.target.value)}
                className={draftFieldErrors.F ? "cell-field--error" : undefined}
                list={startListId}
                placeholder="Ex: Dangote"
              />
              {startListId ? (
                <datalist id={startListId}>
                  {startSuggestions.map((item) => (
                    <option key={`feuil6-start-${item}`} value={item} />
                  ))}
                </datalist>
              ) : null}
              <p className="field-help">Nom du lieu où la voie commence.</p>
            </div>

            <div className={`cell-field${draftFieldErrors.F ? " cell-field--error" : ""}`}>
              <label htmlFor="feuil6-end-label">
                Fin <span className="field-label__required"> *</span>
              </label>
              <input
                id="feuil6-end-label"
                value={bounds.endLabel}
                onChange={(event) => updateFeuil6Bounds("end", event.target.value)}
                className={draftFieldErrors.F ? "cell-field--error" : undefined}
                list={endListId}
                placeholder="Ex: Quai 60"
              />
              {endListId ? (
                <datalist id={endListId}>
                  {endSuggestions.map((item) => (
                    <option key={`feuil6-end-${item}`} value={item} />
                  ))}
                </datalist>
              ) : null}
              {draftFieldErrors.F ? <p className="field-error">{draftFieldErrors.F}</p> : null}
              <p className="field-help">Nom du lieu où la voie se termine.</p>
            </div>

            {renderEditorField("G")}
          </div>

          <div className="editor-actions">
            <button className="primary" type="button" onClick={handleSaveRow} disabled={isBusy || !activeSheet}>
              {editingRowId ? "Enregistrer" : "Ajouter"}
            </button>
            <button className="row-action" type="button" onClick={handleStartNewRow} disabled={isBusy}>
              Nouvelle ligne
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
      );
    }

    if (activeSheet.name === "Feuil4") {
      return (
        <section ref={sheetEditorRef} className="panel editor-panel editor-panel--sheet">
          <h2>{editorTitle}</h2>
          <p className="field-help">
            Renseignez ici les lignes utiles au programme d'évaluation. Commencez par le libellé, puis la valeur
            affichée et les résultats calculés si nécessaire.
          </p>
          {draftFormError ? <p className="modal-feedback modal-feedback--error">{draftFormError}</p> : null}

          <div className="card card--spaced">
            <h3>Identification de la ligne</h3>
            <div className="cells-grid">
              {(["A", "B"] as SheetColumnKey[]).map((column) => renderEditorField(column))}
            </div>
          </div>

          <div className="card card--spaced">
            <h3>Résultat de l'évaluation</h3>
            <div className="cells-grid">
              {(["C", "D", "F"] as SheetColumnKey[]).map((column) => renderEditorField(column))}
            </div>
          </div>

          <div className="card card--spaced">
            <h3>Zone et observation</h3>
            <div className="cells-grid">
              {(["E"] as SheetColumnKey[]).map((column) => renderEditorField(column))}
            </div>
          </div>

          <div className="editor-actions">
            <button className="primary" type="button" onClick={handleSaveRow} disabled={isBusy || !activeSheet}>
              {editingRowId ? "Enregistrer" : "Ajouter"}
            </button>
            <button className="row-action" type="button" onClick={handleStartNewRow} disabled={isBusy}>
              Nouvelle ligne
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
      );
    }

    if (activeSheet.name === "Feuil7") {
      const causeRowsCount = rows.filter((row) => String(row.G ?? "").trim()).length;
      const uniqueDegradationCount = new Set(rows.map((row) => normalizeLabel(row.C)).filter(Boolean)).size;
      const uniqueFamilyCount = new Set(rows.map((row) => normalizeLabel(row.D)).filter(Boolean)).size;
      const filledDraftCount = (["A", "B", "C", "D", "E", "F", "G"] as SheetColumnKey[]).filter(
        (column) => String(draftCells[column] ?? "").trim().length > 0
      ).length;

      return (
        <section ref={sheetEditorRef} className="panel editor-panel editor-panel--sheet">
          <h2>{editorTitle}</h2>
          <p className="field-help">
            Renseignez ici les dégradations et leurs causes probables. Les trois premiers champs servent à identifier
            clairement la dégradation dans le catalogue.
          </p>
          {draftFormError ? <p className="modal-feedback modal-feedback--error">{draftFormError}</p> : null}

          <div className="card card--spaced feuil7-summary-card">
            <h3>Vue rapide du catalogue</h3>
            <div className="feuil7-summary-grid">
              <div className="feuil7-summary-item">
                <span>Dégradations recensées</span>
                <strong>{uniqueDegradationCount}</strong>
              </div>
              <div className="feuil7-summary-item">
                <span>Familles</span>
                <strong>{uniqueFamilyCount}</strong>
              </div>
              <div className="feuil7-summary-item">
                <span>Causes détaillées</span>
                <strong>{causeRowsCount}</strong>
              </div>
              <div className="feuil7-summary-item">
                <span>Champs remplis</span>
                <strong>{filledDraftCount} / 7</strong>
              </div>
            </div>
          </div>

          {(draftCells.C || draftCells.B || draftCells.D || draftCells.E) && (
            <div className="card card--spaced feuil7-preview-card">
              <h3>Fiche en cours</h3>
              <div className="feuil7-preview-grid">
                <div className="feuil7-preview-item">
                  <span>Dégradation</span>
                  <strong>{toDisplay(draftCells.C)}</strong>
                </div>
                <div className="feuil7-preview-item">
                  <span>Référence</span>
                  <strong>{toDisplay(draftCells.B)}</strong>
                </div>
                <div className="feuil7-preview-item">
                  <span>Famille</span>
                  <strong>{toDisplay(draftCells.D)}</strong>
                </div>
                <div className="feuil7-preview-item">
                  <span>Sous-famille</span>
                  <strong>{toDisplay(draftCells.E)}</strong>
                </div>
              </div>
            </div>
          )}

          <div className="card card--spaced">
            <h3>Identification de la dégradation</h3>
            <p className="muted">
              Identifiez d'abord clairement la dégradation dans le catalogue. Cette base sert ensuite à alimenter
              l'aide à la décision, les causes probables et les traitements de référence.
            </p>
            <div className="cells-grid">
              {(["A", "B", "C"] as SheetColumnKey[]).map((column) => renderEditorField(column))}
            </div>
          </div>

          <div className="card card--spaced">
            <h3>Classement</h3>
            <p className="muted">
              Utilisez la famille et la sous-famille pour regrouper visuellement les dégradations proches et faciliter
              les recherches dans le catalogue.
            </p>
            <div className="cells-grid">
              {(["D", "E"] as SheetColumnKey[]).map((column) => renderEditorField(column))}
            </div>
          </div>

          <div className="card card--spaced">
            <h3>Observation et cause</h3>
            <p className="muted">
              La note décrit le contexte utile. La cause probable doit rester exploitable dans l'aide à la décision et
              dans le catalogue enrichi des dégradations.
            </p>
            <div className="cells-grid">
              {(["F", "G"] as SheetColumnKey[]).map((column) => renderEditorField(column))}
            </div>
          </div>

          <div className="editor-actions">
            <button className="primary" type="button" onClick={handleSaveRow} disabled={isBusy || !activeSheet}>
              {editingRowId ? "Enregistrer" : "Ajouter"}
            </button>
            <button className="row-action" type="button" onClick={handleStartNewRow} disabled={isBusy}>
              Nouvelle ligne
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
      );
    }

    return (
      <section ref={sheetEditorRef} className="panel editor-panel editor-panel--sheet">
        <h2>{editorTitle}</h2>
        <p className="field-help">{editorHelp}</p>
        {draftFormError ? <p className="modal-feedback modal-feedback--error">{draftFormError}</p> : null}

        <div className="cells-grid">
          {editableColumns.map((column) => renderEditorField(column))}
        </div>

        <div className="editor-actions">
          <button className="primary" type="button" onClick={handleSaveRow} disabled={isBusy || !activeSheet}>
            {editingRowId ? "Enregistrer" : "Ajouter"}
          </button>
          <button className="row-action" type="button" onClick={handleStartNewRow} disabled={isBusy}>
            Nouvelle ligne
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
    );
  }

  function renderFeuil2View() {
    const totalSections = feuil2Groups.reduce((sum, group) => sum + group.rows.length, 0);
    const totalLength = feuil2Groups.reduce((sum, group) => sum + group.totalLengthM, 0);

    return (
      <main className="workspace">
        {renderSheetEditorPanel()}
        <section className="panel table-panel table-panel--full sheet-print-view">
          {renderStandardSheetPrintHeader(
            activeSheet ? getSheetDisplayName(activeSheet.name) : "Sections",
            activeSheet ? getSheetPrintSubtitle(activeSheet.name) : getSheetPrintSubtitle("Feuil2")
          )}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet ? getSheetDisplayName(activeSheet.name) : "Sections"}</h2>
              <p className="muted">
                Référentiel des sections du réseau, groupé par SAP, utilisé pour structurer les voies et leurs bornes.
              </p>
            </div>
            <div className="sheet-header-actions">
              <div className="measurement-toolbar__meta">
                <span className="pill">SAP: {feuil2Groups.length}</span>
                <span className="pill">Sections: {totalSections}</span>
                <span className="pill">Linéaire: {formatMeasurementNumber(totalLength)} m</span>
              </div>
              {renderSheetPrintButton(activeSheet ? getSheetDisplayName(activeSheet.name) : "Sections")}
            </div>
          </div>

          <div className="feuil2-main">
            <div className="measurement-toolbar">
              <div className="measurement-toolbar__field">
                <label htmlFor="feuil2-search">Recherche section</label>
                <input
                  id="feuil2-search"
                  value={search}
                  onChange={(event) => setSearch(event.target.value)}
                  placeholder="Voie, désignation, début, fin"
                />
              </div>
              <div className="feuil2-filter">
                <label htmlFor="feuil2-sap">Filtre SAP</label>
                <select id="feuil2-sap" value={feuil2SapFilter} onChange={(event) => setFeuil2SapFilter(event.target.value)}>
                  <option value="">Tous les SAP</option>
                  {feuil2SapOptions.map((sapCode) => (
                    <option key={sapCode} value={sapCode}>
                      {sapCode}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {feuil2Groups.map((group) => (
              <div className="card feuil2-group" key={group.sapCode}>
                <div className="dashboard-card__header">
                  <h3>{group.sapCode}</h3>
                  <span className="status-pill">{formatMeasurementNumber(group.totalLengthM)} m</span>
                </div>
                <div className="table-wrap feuil2-table-wrap">
                  <table className="table--feuil2">
                    <thead>
                      <tr>
                        <th className="col-actions">Actions</th>
                        <th>N°</th>
                        <th>N° tronçon</th>
                        <th>N° section</th>
                        <th>Voie</th>
                        <th>Désignation</th>
                        <th>Début</th>
                        <th>Fin</th>
                        <th>Longueur (m)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {group.rows.map((item, index) => (
                        <tr key={item.section.id}>
                          <td className="col-actions">
                            <div className="row-buttons row-buttons--compact row-buttons--wrap">
                              <button
                                className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                                type="button"
                                onClick={() =>
                                  handleUseCentralRoad(
                                    item.linkedRoad,
                                    "Aucune voie centralisée correspondante n'a été trouvée pour cette section.",
                                    `Section ${toDisplay(item.roadLabel)} chargée dans l'aide à la décision.`,
                                    item.sapCode
                                  )
                                }
                                disabled={!item.linkedRoad}
                              >
                                <Gauge size={14} aria-hidden="true" />
                                <span>Évaluer</span>
                              </button>
                              <button
                                className="row-action row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() =>
                                  item.sourceRow
                                    ? handleEditSourceRow(
                                        item.sourceRow,
                                        "Cette section centralisée n'a pas de ligne source Feuil2 modifiable."
                                      )
                                    : void handleCreateSourceRowAndEdit(
                                        "Feuil2",
                                        buildFeuil2SourcePayload(item),
                                        `Ligne maître Feuil2 créée pour ${toDisplay(item.roadLabel)} puis ouverte en édition.`
                                      )
                                }
                                title={item.sourceRow ? "Éditer" : "Créer la ligne maître puis éditer"}
                                aria-label={item.sourceRow ? "Éditer" : "Créer la ligne maître puis éditer"}
                              >
                                <Pencil size={15} aria-hidden="true" />
                              </button>
                              <button
                                className="row-action row-action--danger row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() =>
                                  handleDeleteSourceRow(
                                    item.sourceRow,
                                    "Cette section centralisée n'a pas de ligne source Feuil2 à supprimer."
                                  )
                                }
                                title="Supprimer"
                                aria-label="Supprimer"
                                disabled={!item.sourceRow}
                              >
                                <Trash2 size={15} aria-hidden="true" />
                              </button>
                            </div>
                          </td>
                          <td>{index + 1}</td>
                          <td>{toDisplay(item.tronconNo)}</td>
                          <td>{toDisplay(item.sectionNo)}</td>
                          <td>{toDisplay(item.roadLabel)}</td>
                          <td>{toDisplay(item.designation)}</td>
                          <td>{toDisplay(item.startLabel)}</td>
                          <td>{toDisplay(item.endLabel)}</td>
                          <td>{formatMeasurementNumber(item.lengthM)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}

            {feuil2Groups.length === 0 ? (
              <div className="card">
                <p className="muted">{isLoadingRows ? "Chargement..." : "Aucune section Feuil2 disponible."}</p>
              </div>
            ) : null}
          </div>
          {renderStandardSheetPrintFooter()}
        </section>
      </main>
    );
  }

  function renderFeuil6View() {
    const totalRoads = feuil6Groups.reduce((sum, group) => sum + group.rows.length, 0);
    const totalLinear = feuil6Groups.reduce((sum, group) => sum + group.totalLinearM, 0);

    return (
      <main className="workspace">
        {renderSheetEditorPanel()}
        <section className="panel table-panel table-panel--full sheet-print-view">
          {renderStandardSheetPrintHeader(
            activeSheet ? getSheetDisplayName(activeSheet.name) : "Voies",
            activeSheet ? getSheetPrintSubtitle(activeSheet.name) : getSheetPrintSubtitle("Feuil6")
          )}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet ? getSheetDisplayName(activeSheet.name) : "Voies"}</h2>
              <p className="muted">
                Répertoire codifié central des voies: type, code, début, fin et justification, réutilisés dans tout le système.
              </p>
            </div>
            <div className="sheet-header-actions">
              <div className="measurement-toolbar__meta">
                <span className="pill">SAP: {feuil6Groups.length}</span>
                <span className="pill">Voies: {totalRoads}</span>
                <span className="pill">Linéaire: {formatMeasurementNumber(totalLinear)} m</span>
              </div>
              {renderSheetPrintButton(activeSheet ? getSheetDisplayName(activeSheet.name) : "Voies")}
            </div>
          </div>

          <div className="feuil2-layout feuil2-layout--single">
            <div className="feuil2-main">
              <div className="measurement-toolbar">
                <div className="measurement-toolbar__field">
                  <label htmlFor="feuil6-search">Recherche voie codifiée</label>
                  <input
                    id="feuil6-search"
                    value={search}
                    onChange={(event) => setSearch(event.target.value)}
                    placeholder="Code, nom proposé, début, fin, justification"
                  />
                </div>
                <div className="feuil2-filter">
                  <label htmlFor="feuil6-sap">Filtre SAP</label>
                  <select id="feuil6-sap" value={feuil6SapFilter} onChange={(event) => setFeuil6SapFilter(event.target.value)}>
                    <option value="">Tous les SAP</option>
                    {feuil6SapOptions.map((sapCode) => (
                      <option key={sapCode} value={sapCode}>
                        {sapCode}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {feuil6Groups.map((group) => (
                <div className="card feuil2-group" key={group.sapCode}>
                  <div className="dashboard-card__header">
                    <h3>{group.sapCode}</h3>
                    <span className="status-pill">{formatMeasurementNumber(group.totalLinearM)} m</span>
                  </div>
                  <div className="table-wrap feuil2-table-wrap">
                  <table className="table--feuil6">
                    <thead>
                      <tr>
                        <th className="col-actions">Action</th>
                        <th>N°</th>
                        <th>Type de voie</th>
                        <th>Code</th>
                          <th>Linéaire (ml)</th>
                          <th>Noms proposés</th>
                          <th>Début</th>
                          <th>Fin</th>
                          <th>Justification</th>
                        </tr>
                      </thead>
                      <tbody>
                        {group.rows.map((item, index) => (
                          <tr key={item.linkedRoad.id}>
                            <td className="col-actions">
                              <div className="row-buttons row-buttons--compact row-buttons--wrap">
                                <button
                                  className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                                  type="button"
                                  onClick={() =>
                                    handleUseCentralRoad(
                                      item.linkedRoad,
                                      "Aucune voie centralisée correspondante n'a été trouvée pour cette entrée du répertoire.",
                                      `Voie ${item.roadCode} chargée depuis le répertoire codifié.`,
                                      item.sapCode
                                    )
                                  }
                                >
                                  <Gauge size={14} aria-hidden="true" />
                                  <span>Évaluer</span>
                                </button>
                                <button
                                  className="row-action row-action--icon row-action--icon-sm"
                                  type="button"
                                  onClick={() =>
                                    item.sourceRow
                                      ? handleEditSourceRow(
                                          item.sourceRow,
                                          "Cette voie centrale n'a pas encore de ligne source Feuil6 modifiable."
                                        )
                                      : void handleCreateSourceRowAndEdit(
                                          "Feuil6",
                                          buildFeuil6SourcePayload(item),
                                          `Fiche Voies créée pour ${toDisplay(item.roadCode)} puis ouverte en édition.`
                                        )
                                  }
                                  title="Éditer"
                                  aria-label="Éditer"
                                >
                                  <Pencil size={15} aria-hidden="true" />
                                </button>
                                <button
                                  className="row-action row-action--danger row-action--icon row-action--icon-sm"
                                  type="button"
                                  onClick={() =>
                                    handleDeleteSourceRow(
                                      item.sourceRow,
                                      "Cette voie centrale n'a pas encore de ligne source Feuil6 à supprimer."
                                    )
                                  }
                                  title="Supprimer"
                                  aria-label="Supprimer"
                                  disabled={!item.sourceRow}
                                >
                                  <Trash2 size={15} aria-hidden="true" />
                              </button>
                            </div>
                          </td>
                          <td>{index + 1}</td>
                          <td>{toDisplay(item.roadType)}</td>
                          <td>{toDisplay(item.roadCode)}</td>
                          <td>{formatMeasurementNumber(item.linearM)}</td>
                            <td>{toDisplay(item.proposedName)}</td>
                            <td>{toDisplay(item.startLabel)}</td>
                            <td>{toDisplay(item.endLabel)}</td>
                            <td>{toDisplay(item.justification)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ))}

              {feuil6Groups.length === 0 ? (
                <div className="card">
                  <p className="muted">{isLoadingRows ? "Chargement..." : "Aucune voie Feuil6 disponible."}</p>
                </div>
              ) : null}
            </div>
          </div>
          {renderStandardSheetPrintFooter()}
        </section>
      </main>
    );
  }

  function renderFeuil3View() {
    const totalProfiles = feuil3Groups.reduce((sum, group) => sum + group.rows.length, 0);

    return (
      <main className="workspace">
        <div className="workspace-side-stack">
          {renderSheetEditorPanel()}
          <section className="panel">
            <div className="card">
              <h3>Lecture métier</h3>
              <p className="muted">
                Diagnostic porte l'état technique courant de la section. C'est la vue de référence pour la chaussée,
                l'assainissement observé et l'intervention à prévoir avant décision.
              </p>
            </div>

            <div className="card">
              <h3>Interventions à préciser</h3>
              <ul className="count-list">
                {feuil3Groups.map((group) => {
                  const pending = group.rows.filter((item) => isToDetermineIntervention(item.interventionHint)).length;
                  return (
                    <li key={`${group.sapCode}-pending`}>
                      <span>{group.sapCode}</span>
                      <strong>{pending}</strong>
                    </li>
                  );
                })}
              </ul>
            </div>

            <div className="card">
              <h3>Assainissement sous surveillance</h3>
              <ul className="count-list">
                {feuil3Groups.map((group) => {
                  const flagged = group.rows.filter((item) => {
                    const normalized = normalizeLabel(item.drainageState);
                    return Boolean(normalized && !["-", "BON"].includes(normalized));
                  }).length;
                  return (
                    <li key={`${group.sapCode}-drainage`}>
                      <span>{group.sapCode}</span>
                      <strong>{flagged}</strong>
                    </li>
                  );
                })}
              </ul>
            </div>

            <div className="card">
              <h3>États spéciaux réseau</h3>
              <ul className="count-list">
                <li>
                  <span>Travaux en cours</span>
                  <strong>{feuil3SpecialStateCounts.inProgress}</strong>
                </li>
                <li>
                  <span>Non aménagées</span>
                  <strong>{feuil3SpecialStateCounts.notBuilt}</strong>
                </li>
                <li>
                  <span>RAS</span>
                  <strong>{feuil3SpecialStateCounts.ras}</strong>
                </li>
              </ul>
            </div>
          </section>
        </div>
        <section className="panel table-panel table-panel--full sheet-print-view">
          {renderStandardSheetPrintHeader(
            activeSheet ? getSheetDisplayName(activeSheet.name) : "Diagnostic",
            activeSheet ? getSheetPrintSubtitle(activeSheet.name) : getSheetPrintSubtitle("Feuil3")
          )}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet ? getSheetDisplayName(activeSheet.name) : "Diagnostic"}</h2>
              <p className="muted">
                Diagnostic technique des sections: état de la chaussée, assainissement et intervention à prévoir.
              </p>
            </div>
            <div className="sheet-header-actions">
              <div className="measurement-toolbar__meta">
                <span className="pill">SAP: {feuil3Groups.length}</span>
                <span className="pill">Profils: {totalProfiles}</span>
                <span className="pill">À déterminer: {feuil3PendingInterventions}</span>
                <span className="pill">Assainissement à surveiller: {feuil3DrainageAlerts}</span>
              </div>
              {renderSheetPrintButton(activeSheet ? getSheetDisplayName(activeSheet.name) : "Diagnostic")}
            </div>
          </div>

          <div className="feuil2-main">
            <div className="measurement-toolbar">
              <div className="measurement-toolbar__field">
                <label htmlFor="feuil3-search">Recherche profil technique</label>
                <input
                  id="feuil3-search"
                  value={search}
                  onChange={(event) => setSearch(event.target.value)}
                  placeholder="Voie, désignation, état chaussée, assainissement"
                />
              </div>
              <div className="feuil2-filter">
                <label htmlFor="feuil3-sap">Filtre SAP</label>
                <select id="feuil3-sap" value={feuil3SapFilter} onChange={(event) => setFeuil3SapFilter(event.target.value)}>
                  <option value="">Tous les SAP</option>
                  {feuil3SapOptions.map((sapCode) => (
                    <option key={sapCode} value={sapCode}>
                      {sapCode}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {feuil3Groups.map((group) => (
              <div className="card feuil2-group" key={group.sapCode}>
                <div className="dashboard-card__header">
                  <h3>{group.sapCode}</h3>
                  <span className="status-pill">{group.rows.length} profil(s)</span>
                </div>
                <div className="table-wrap feuil2-table-wrap">
                  <table className="table--feuil3">
                    <thead>
                      <tr>
                        <th className="col-actions" rowSpan={3}>Action</th>
                        <th rowSpan={3}>N°</th>
                        <th rowSpan={3}>Voies</th>
                        <th rowSpan={3}>Désignation</th>
                        <th rowSpan={3}>Début</th>
                        <th rowSpan={3}>Fin</th>
                        <th rowSpan={3}>Longueur (m)</th>
                        <th rowSpan={3}>Largeur min. façade (m)</th>
                        <th colSpan={2}>CHAUSSÉE (en %)</th>
                        <th colSpan={2}>ASSAINISSEMENT</th>
                        <th rowSpan={3}>Largeur min. trottoirs (m)</th>
                        <th rowSpan={3}>Nature de l'intervention</th>
                      </tr>
                      <tr>
                        <th>Nature du revêtement</th>
                        <th>État de la chaussée</th>
                        <th>Type</th>
                        <th>État</th>
                      </tr>
                      <tr>
                        <th>&nbsp;</th>
                        <th>&nbsp;</th>
                        <th>caniveaux</th>
                        <th>description</th>
                      </tr>
                    </thead>
                    <tbody>
                      {group.rows.map((item, index) => (
                        <tr key={item.section.id}>
                          <td className="col-actions">
                            <div className="row-buttons row-buttons--compact row-buttons--wrap">
                              <button
                                className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                                type="button"
                                onClick={() =>
                                  handleUseCentralRoad(
                                    item.linkedRoad,
                                    "Aucune voie centralisée correspondante n'a été trouvée pour ce profil technique.",
                                    `Profil technique ${toDisplay(item.roadLabel)} chargé dans l'aide à la décision.`,
                                    item.sapCode
                                  )
                                }
                                disabled={!item.linkedRoad}
                              >
                                <Gauge size={14} aria-hidden="true" />
                                <span>Évaluer</span>
                              </button>
                              <button
                                className="row-action row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() =>
                                  item.sourceRow
                                    ? handleEditSourceRow(
                                        item.sourceRow,
                                        "Ce profil centralisé n'a pas de ligne source Feuil3 modifiable."
                                      )
                                    : void handleCreateSourceRowAndEdit(
                                        "Feuil3",
                                        buildFeuil3SourcePayload(item),
                                        `Ligne source Feuil3 créée pour ${toDisplay(item.roadLabel)} puis ouverte en édition.`
                                      )
                                }
                                title={item.sourceRow ? "Éditer" : "Créer la ligne source puis éditer"}
                                aria-label={item.sourceRow ? "Éditer" : "Créer la ligne source puis éditer"}
                              >
                                <Pencil size={15} aria-hidden="true" />
                              </button>
                              <button
                                className="row-action row-action--danger row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() =>
                                  handleDeleteSourceRow(
                                    item.sourceRow,
                                    "Ce profil centralisé n'a pas de ligne source Feuil3 à supprimer."
                                  )
                                }
                                title="Supprimer"
                                aria-label="Supprimer"
                                disabled={!item.sourceRow}
                              >
                                <Trash2 size={15} aria-hidden="true" />
                              </button>
                            </div>
                          </td>
                          <td>{index + 1}</td>
                          <td>{toDisplay(item.roadLabel)}</td>
                          <td>{toDisplay(item.designation)}</td>
                          <td>{toDisplay(item.startLabel)}</td>
                          <td>{toDisplay(item.endLabel)}</td>
                          <td>{formatMeasurementNumber(item.lengthM)}</td>
                          <td>{formatMeasurementNumber(item.facadeWidthM)}</td>
                          <td>{toDisplay(item.surfaceType)}</td>
                          <td className={normalizeLabel(item.pavementState) === "BON" ? "" : "cell-warning"}>
                            {toDisplay(item.pavementState)}
                          </td>
                          <td>{toDisplay(item.drainageType)}</td>
                          <td>{toDisplay(item.drainageState)}</td>
                          <td>{formatMeasurementNumber(item.sidewalkMinM)}</td>
                          <td className={isToDetermineIntervention(item.interventionHint) ? "cell-warning" : ""}>
                            {item.interventionHint || "à déterminer (A D)"}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}

            {feuil3Groups.length === 0 ? (
              <div className="card">
                <p className="muted">{isLoadingRows ? "Chargement..." : "Aucun profil Feuil3 disponible."}</p>
              </div>
            ) : null}
          </div>
          {renderStandardSheetPrintFooter()}
        </section>
      </main>
    );
  }

  function renderFeuil5View() {
    const totalProfiles = feuil5Groups.reduce((sum, group) => sum + group.rows.length, 0);

    return (
      <main className="workspace">
        <div className="workspace-side-stack">
          {renderSheetEditorPanel()}
          <section className="panel">
            <div className="card">
              <h3>Lecture métier</h3>
              <p className="muted">
                Compléments sert à documenter ce qui gravite autour de la section: largeur utile, trottoirs,
                stationnement et contexte d'assainissement. Le revêtement et l'état de chaussée restent dans Diagnostic.
              </p>
            </div>

            <div className="card">
              <h3>Assainissement sous surveillance</h3>
              <ul className="count-list">
                {feuil5Groups.map((group) => {
                  const count = group.rows.filter((item) => {
                    const normalized = normalizeLabel(item.drainageState);
                    return Boolean(normalized && !["-", "BON"].includes(normalized));
                  }).length;
                  return (
                    <li key={`${group.sapCode}-drainage`}>
                      <span>{group.sapCode}</span>
                      <strong>{count}</strong>
                    </li>
                  );
                })}
              </ul>
            </div>
          </section>
        </div>
        <section className="panel table-panel table-panel--full sheet-print-view">
          {renderStandardSheetPrintHeader(
            activeSheet ? getSheetDisplayName(activeSheet.name) : "Compléments",
            activeSheet ? getSheetPrintSubtitle(activeSheet.name) : getSheetPrintSubtitle("Feuil5")
          )}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet ? getSheetDisplayName(activeSheet.name) : "Compléments"}</h2>
              <p className="muted">
                Compléments latéraux des sections: largeurs utiles, trottoirs, stationnement et contexte d'assainissement.
              </p>
            </div>
            <div className="sheet-header-actions">
              <div className="measurement-toolbar__meta">
                <span className="pill">SAP: {feuil5Groups.length}</span>
                <span className="pill">Profils: {totalProfiles}</span>
                <span className="pill">Stationnement recensé: {feuil5ParkingCount}</span>
                <span className="pill">Assainissement à surveiller: {feuil5DrainageWatchCount}</span>
              </div>
              {renderSheetPrintButton(activeSheet ? getSheetDisplayName(activeSheet.name) : "Compléments")}
            </div>
          </div>

          <div className="feuil2-main">
            <div className="measurement-toolbar">
              <div className="measurement-toolbar__field">
                <label htmlFor="feuil5-search">Recherche profil complémentaire</label>
                <input
                  id="feuil5-search"
                  value={search}
                  onChange={(event) => setSearch(event.target.value)}
                  placeholder="Voie, désignation, assainissement, trottoirs, stationnement"
                />
              </div>
              <div className="feuil2-filter">
                <label htmlFor="feuil5-sap">Filtre SAP</label>
                <select id="feuil5-sap" value={feuil5SapFilter} onChange={(event) => setFeuil5SapFilter(event.target.value)}>
                  <option value="">Tous les SAP</option>
                  {feuil5SapOptions.map((sapCode) => (
                    <option key={sapCode} value={sapCode}>
                      {sapCode}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {feuil5Groups.map((group) => (
              <div className="card feuil2-group" key={group.sapCode}>
                <div className="dashboard-card__header">
                  <h3>{group.sapCode}</h3>
                  <span className="status-pill">{group.rows.length} complément(s)</span>
                </div>
                <div className="table-wrap feuil2-table-wrap">
                  <table className="table--feuil5">
                    <thead>
                      <tr>
                        <th className="col-actions" rowSpan={2}>Action</th>
                        <th rowSpan={2}>N°</th>
                        <th rowSpan={2}>N° tronçon</th>
                        <th rowSpan={2}>N° section</th>
                        <th rowSpan={2}>Voies</th>
                        <th rowSpan={2}>Désignation</th>
                        <th rowSpan={2}>Début</th>
                        <th rowSpan={2}>Fin</th>
                        <th rowSpan={2}>Longueur (m)</th>
                        <th rowSpan={2}>Largeur min. façade (m)</th>
                        <th colSpan={2}>ASSAINISSEMENT</th>
                        <th rowSpan={2}>Largeur min. trottoirs (m)</th>
                        <th colSpan={3}>STATIONNEMENT</th>
                      </tr>
                      <tr>
                        <th>Type</th>
                        <th>État</th>
                        <th>Gauche</th>
                        <th>Droit</th>
                        <th>Autres</th>
                      </tr>
                    </thead>
                    <tbody>
                      {group.rows.map((item, index) => (
                        <tr key={item.section.id}>
                          <td className="col-actions">
                            <div className="row-buttons row-buttons--compact row-buttons--wrap">
                              <button
                                className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                                type="button"
                                onClick={() =>
                                  handleUseCentralRoad(
                                    item.linkedRoad,
                                    "Aucune voie centralisée correspondante n'a été trouvée pour ce profil complémentaire.",
                                    `Profil complémentaire ${toDisplay(item.roadLabel)} chargé dans l'aide à la décision.`,
                                    item.sapCode
                                  )
                                }
                                disabled={!item.linkedRoad}
                              >
                                <Gauge size={14} aria-hidden="true" />
                                <span>Évaluer</span>
                              </button>
                              <button
                                className="row-action row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() =>
                                  item.sourceRow
                                    ? handleEditSourceRow(
                                        item.sourceRow,
                                        "Ce profil centralisé n'a pas de ligne source Feuil5 modifiable."
                                      )
                                    : handleStartPrefilledFeuil5Draft(
                                        buildFeuil5SourcePayload(item),
                                        `Complément ${toDisplay(item.roadLabel)} chargé dans l'éditeur. Complétez les champs manquants puis enregistrez.`
                                      )
                                }
                                title={item.sourceRow ? "Éditer" : "Ouvrir dans l'éditeur"}
                                aria-label={item.sourceRow ? "Éditer" : "Ouvrir dans l'éditeur"}
                              >
                                <Pencil size={15} aria-hidden="true" />
                              </button>
                              <button
                                className="row-action row-action--danger row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() =>
                                  handleDeleteSourceRow(
                                    item.sourceRow,
                                    "Ce profil centralisé n'a pas de ligne source Feuil5 à supprimer."
                                  )
                                }
                                title="Supprimer"
                                aria-label="Supprimer"
                                disabled={!item.sourceRow}
                              >
                                <Trash2 size={15} aria-hidden="true" />
                              </button>
                            </div>
                          </td>
                          <td>{index + 1}</td>
                          <td>{toDisplay(item.tronconNo)}</td>
                          <td>{toDisplay(item.sectionNo)}</td>
                          <td>{toDisplay(item.roadLabel)}</td>
                          <td>{toDisplay(item.designation)}</td>
                          <td>{toDisplay(item.startLabel)}</td>
                          <td>{toDisplay(item.endLabel)}</td>
                          <td>{formatMeasurementNumber(item.lengthM)}</td>
                          <td>{formatMeasurementNumber(item.facadeWidthM)}</td>
                          <td>{toDisplay(item.drainageType)}</td>
                          <td className={normalizeLabel(item.drainageState) === "BON" ? "" : "cell-warning"}>
                            {toDisplay(item.drainageState)}
                          </td>
                          <td>{formatMeasurementNumber(item.sidewalkMinM)}</td>
                          <td>{toDisplay(item.parkingLeft)}</td>
                          <td>{toDisplay(item.parkingRight)}</td>
                          <td>{toDisplay(item.parkingOther)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}

            {feuil5Groups.length === 0 ? (
              <div className="card">
                <p className="muted">{isLoadingRows ? "Chargement..." : "Aucun complément Feuil5 disponible."}</p>
              </div>
            ) : null}
          </div>
          {renderStandardSheetPrintFooter()}
        </section>
      </main>
    );
  }

  function renderFeuil7View() {
    const totalRows = rows.length;
    const uniqueDegradationCount = new Set(
      rows.map((row) => normalizeLabel(row.C)).filter(Boolean)
    ).size;
    const uniqueCategoryCount = new Set(
      rows.map((row) => normalizeLabel(row.A)).filter(Boolean)
    ).size;
    const detailedCauseCount = rows.filter((row) => String(row.G ?? "").trim()).length;

    return (
      <main className="workspace">
        {renderSheetEditorPanel()}

        <section className="panel table-panel sheet-print-view">
          {renderStandardSheetPrintHeader(
            activeSheet ? getSheetDisplayName(activeSheet.name) : "Causes",
            activeSheet ? getSheetPrintSubtitle(activeSheet.name) : getSheetPrintSubtitle("Feuil7")
          )}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet ? getSheetDisplayName(activeSheet.name) : "Causes"}</h2>
              <p className="muted">
                Catalogue structuré des dégradations, de leurs familles et des causes probables utiles à la décision.
              </p>
            </div>
            <div className="sheet-header-actions">
              <div className="measurement-toolbar__meta">
                <span className="pill">Lignes: {totalRows}</span>
                <span className="pill">Dégradations: {uniqueDegradationCount}</span>
                <span className="pill">Catégories: {uniqueCategoryCount}</span>
                <span className="pill">Causes détaillées: {detailedCauseCount}</span>
              </div>
              {renderSheetPrintButton(activeSheet ? getSheetDisplayName(activeSheet.name) : "Causes")}
            </div>
          </div>

          <div className="table-toolbar">
            <input
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Rechercher référence, dégradation, famille, cause"
            />
            <span className="muted">{status?.sheetCounts?.Feuil7 ?? 0} ligne(s) en base</span>
          </div>

          <div className="table-wrap">
            <table className="table--feuil7">
              <thead>
                <tr>
                  <th className="col-actions">Actions</th>
                  <th>N°</th>
                  <th>Catégorie</th>
                  <th>Référence</th>
                  <th>Dégradation</th>
                  <th>Famille / Sous-famille</th>
                  <th>Cause probable</th>
                  <th>Notes</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => (
                  <tr key={row.id} className={editingRowId === row.id ? "is-selected" : ""}>
                    <td className="col-actions">
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
                    <td>{index + 1}</td>
                    <td>{toDisplay(row.A)}</td>
                    <td>{toDisplay(row.B)}</td>
                    <td>
                      <div className="table-cell-stack">
                        <strong>{toDisplay(row.C)}</strong>
                      </div>
                    </td>
                    <td>
                      <div className="table-cell-stack">
                        <strong>{toDisplay(row.D)}</strong>
                        <span>{toDisplay(row.E)}</span>
                      </div>
                    </td>
                    <td title={String(row.G ?? "")}>{trimPreviewText(row.G, 220) || "-"}</td>
                    <td title={String(row.F ?? "")}>{trimPreviewText(row.F, 160) || "-"}</td>
                  </tr>
                ))}
                {rows.length === 0 ? (
                  <tr>
                    <td colSpan={8}>{isLoadingRows ? "Chargement..." : "Aucune ligne."}</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>
          {renderStandardSheetPrintFooter()}
        </section>
      </main>
    );
  }

  function renderSheetView() {
    if (activeSheet?.name === "Feuil1") {
      return renderFeuil1View();
    }
    if (activeSheet?.name === "Feuil2") {
      return renderFeuil2View();
    }
    if (activeSheet?.name === "Feuil3") {
      return renderFeuil3View();
    }
    if (activeSheet?.name === "Feuil5") {
      return renderFeuil5View();
    }
    if (activeSheet?.name === "Feuil6") {
      return renderFeuil6View();
    }
    if (activeSheet?.name === "Feuil7") {
      return renderFeuil7View();
    }

    return (
      <main className="workspace">
        {renderSheetEditorPanel()}

        <section className="panel table-panel sheet-print-view">
          {renderStandardSheetPrintHeader(
            activeSheet ? getSheetDisplayName(activeSheet.name) : "Feuille",
            activeSheet ? getSheetPrintSubtitle(activeSheet.name) : undefined
          )}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet ? getSheetDisplayName(activeSheet.name) : "Feuille"}</h2>
              <p className="muted">
                {activeSheet ? getSheetPrintSubtitle(activeSheet.name) : "Impression du tableau de la feuille active."}
              </p>
            </div>
            {renderSheetPrintButton(activeSheet ? getSheetDisplayName(activeSheet.name) : "Feuille")}
          </div>

          <div className="table-toolbar">
            <input
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder={
                activeSheet?.name === "Feuil4"
                  ? "Rechercher libellé, valeur, intervention, zone"
                  : activeSheet?.name === "Feuil7"
                    ? "Rechercher référence, dégradation, cause"
                    : "Rechercher dans la feuille active"
              }
            />
            <span className="muted">
              {activeSheet ? `${status?.sheetCounts?.[activeSheet.name] ?? 0} ligne(s) en base` : "Aucune feuille"}
            </span>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th className="col-actions">Actions</th>
                  <th>N°</th>
                  {activeColumns.map((column) => (
                    <th key={`head-${column}`}>
                      {getColumnLabel(activeSheet, column)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => (
                  <tr key={row.id} className={editingRowId === row.id ? "is-selected" : ""}>
                    <td className="col-actions">
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
                    <td>{index + 1}</td>
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
          {renderStandardSheetPrintFooter()}
        </section>
      </main>
    );
  }

  if (!hasElectronBridge) {
    return (
      <div className="bridge-error-shell">
        <div className="bridge-error-card">
          <img className="hero-logo" src={appLogoSrc} alt="Logo Port Autonome de Douala" />
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
          <img className="hero-logo" src={appLogoSrc} alt="Logo Port Autonome de Douala" />
          <div>
            <h1>PAD Outil de Maintenance Routière</h1>
            <p>Pilotez la maintenance routière du PAD avec des décisions rapides et fiables.</p>
          </div>
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
          <span className="pill">Dernier import: {status?.lastImportAt ? fmtDate(status.lastImportAt) : "-"}</span>
        </div>

        {error ? <p className="hero__error">{error}</p> : null}
        {notice ? <p className="hero__notice">{notice}</p> : null}
        {dashboardSummary?.integrity && !(dashboardSummary.integrity.status === "OK" && isIntegrityAlertDismissed) ? (
          <div className="integrity-alert">
            <span className="integrity-alert__content">
              {dashboardSummary.integrity.status === "OK" ? (
                <ShieldCheck size={16} aria-hidden="true" />
              ) : (
                <TriangleAlert size={16} aria-hidden="true" />
              )}
              <span>
                Cohérence: {dashboardSummary.integrity.status === "OK" ? "OK" : "à vérifier"} ·{" "}
                {dashboardSummary.integrity.issues.length} point(s) détecté(s)
              </span>
            </span>
            {dashboardSummary.integrity.status === "OK" ? (
              <button
                className="integrity-alert__action integrity-alert__action--icon"
                type="button"
                onClick={() => setIsIntegrityAlertDismissed(true)}
                title="Fermer"
                aria-label="Fermer"
              >
                <X size={14} aria-hidden="true" />
              </button>
            ) : (
              <button className="integrity-alert__action" type="button" onClick={handleManageIntegrityFromAlert}>
                Gérer
              </button>
            )}
          </div>
        ) : null}
        <nav className="hero__nav">
          <button className={activeView === "dashboard" ? "active" : ""} type="button" onClick={() => setActiveView("dashboard")}>
            <BarChart3 size={16} aria-hidden="true" />
            <span>Tableau de bord</span>
          </button>
          <button className={activeView === "decision" ? "active" : ""} type="button" onClick={() => setActiveView("decision")}>
            <Gauge size={16} aria-hidden="true" />
            <span>Aide à la décision</span>
          </button>
          <button className={activeView === "catalogue" ? "active" : ""} type="button" onClick={() => setActiveView("catalogue")}>
            <BookOpen size={16} aria-hidden="true" />
            <span>Catalogue</span>
          </button>
          <button
            className={activeView === "degradations" ? "active" : ""}
            type="button"
            onClick={() => setActiveView("degradations")}
          >
            <TriangleAlert size={16} aria-hidden="true" />
            <span>Dégradations</span>
          </button>
          <button className={activeView === "maintenance" ? "active" : ""} type="button" onClick={() => setActiveView("maintenance")}>
            <ClipboardPlus size={16} aria-hidden="true" />
            <span>Suivi</span>
          </button>
          <button className={activeView === "history" ? "active" : ""} type="button" onClick={() => setActiveView("history")}>
            <History size={16} aria-hidden="true" />
            <span>Historique ({status?.decisionHistoryCount ?? 0})</span>
          </button>
          {definitions.filter((sheet) => sheet.name !== "Feuil7").map((sheet) => (
            <button
              key={sheet.name}
              className={activeView === `sheet:${sheet.name}` ? "active" : ""}
              type="button"
              onClick={() => setActiveView(`sheet:${sheet.name}`)}
            >
              {renderSheetNavIcon(sheet.name)}
              <span>{getSheetDisplayName(sheet.name)}</span>
            </button>
          ))}
        </nav>
      </header>

      {activeView === "dashboard" ? renderDashboardView() : null}
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




