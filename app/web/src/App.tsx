import { useCallback, useEffect, useMemo, useRef, useState, type ChangeEvent } from "react";
import { padApi } from "./lib/pad-api";
import {
  BarChart3,
  ChevronDown,
  ChevronUp,
  DatabaseBackup,
  Eye,
  ExternalLink,
  FileSpreadsheet,
  FolderOpen,
  Gauge,
  Paperclip,
  Pencil,
  Plus,
  Printer,
  RefreshCw,
  Settings2,
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
  DrainageRule,
  DegradationItem,
  ImportPreview,
  MeasurementCampaignItem,
  MaintenanceInterventionItem,
  MaintenanceInterventionPayload,
  MaintenanceInterventionStatus,
  MaintenanceSolutionTemplate,
  RoadCatalogItem,
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

function getEditableColumns(sheet: SheetDefinition | null): SheetColumnKey[] {
  if (!sheet) {
    return [];
  }
  if (sheet.name === "Feuil2") {
    return ["A", "B", "C", "D", "E", "F", "G"];
  }
  if (sheet.name === "Feuil6") {
    return ["B", "C", "D", "E", "F", "G"];
  }
  return sheet.columns;
}

function getSheetFieldPlaceholder(sheetName: string | undefined, column: SheetColumnKey) {
  if (sheetName === "Feuil2") {
    if (column === "A") return "Ex: 1";
    if (column === "B") return "Ex: 1_1";
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
    if (column === "B") return "Ex: 1_1";
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
    if (column === "P") return "Ex: 0";
  }
  if (sheetName === "Feuil6") {
    if (column === "A") return "Ex: 1";
    if (column === "B") return "Ex: Rue, Boulevard ou Avenue";
    if (column === "C") return "Ex: Rue.01";
    if (column === "D") return "Ex: 700";
    if (column === "E") return "Ex: Rue des Archives";
    if (column === "F") return "Ex: de Dangote à Quai 60";
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
      return "Écrivez le numéro du tronçon. Exemple : 1, 2, 3.";
    }
    if (column === "B") {
      return "Écrivez le numéro de section sous la forme 1_1, 2_3 ou 6_1. Le nombre avant le tiret bas (_) crée le groupe SAP automatiquement.";
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
    if (column === "A") return "Numéro du tronçon.";
    if (column === "B") return "Numéro de section, par exemple 1_1 ou 2_3.";
    if (column === "C") return "Code court de la voie.";
    if (column === "D") return "Nom complet de la voie.";
    if (column === "E") return "Nom du lieu où la section commence. Ce n'est pas une date.";
    if (column === "F") return "Nom du lieu où la section se termine. Ce n'est pas une date.";
    if (column === "G") return "Longueur de la section en mètres.";
    if (column === "H") return "Largeur minimale en mètres.";
    if (column === "I") return "Type de revêtement observé sur la chaussée.";
    if (column === "J") return "État général de la chaussée.";
    if (column === "K") return "Type de caniveaux ou d'assainissement.";
    if (column === "L") return "État de l'assainissement.";
    if (column === "M") return "Largeur minimale des trottoirs en mètres.";
    if (column === "N") return "Valeur de stationnement du côté gauche.";
    if (column === "O") return "Valeur de stationnement du côté droit.";
    if (column === "P") return "Autre information utile sur le stationnement.";
  }
  if (sheetName === "Feuil6") {
    if (column === "A") return "Numéro d'ordre de la ligne dans le répertoire.";
    if (column === "B") return "Choisissez le type de voie : Rue, Boulevard ou Avenue.";
    if (column === "C") return "Code court de la voie.";
    if (column === "D") return "Longueur totale en mètres.";
    if (column === "E") return "Nom proposé pour cette voie.";
    if (column === "F") return "Écrivez l'itinéraire sous la forme de ... à ...";
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
    return ["A", "B", "C", "D", "E", "F", "G"];
  }
  if (sheetName === "Feuil3") {
    return ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "L"];
  }
  if (sheetName === "Feuil5") {
    return ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"];
  }
  if (sheetName === "Feuil6") {
    return ["B", "C", "D", "E", "F", "G"];
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
    if (column === "A") return "Veuillez renseigner le numéro du tronçon.";
    if (column === "B") return "Veuillez renseigner le numéro de section.";
    if (column === "C") return "Veuillez renseigner le code de la voie.";
    if (column === "D") return "Veuillez renseigner le nom de la voie.";
    if (column === "E") return "Veuillez renseigner le lieu de départ.";
    if (column === "F") return "Veuillez renseigner le lieu d'arrivée.";
    if (column === "G") return "Veuillez renseigner la longueur en mètres.";
  }
  if (sheetName === "Feuil3") {
    if (column === "A") return "Veuillez renseigner le code de la voie.";
    if (column === "B") return "Veuillez renseigner le nom de la voie.";
    if (column === "C") return "Veuillez renseigner le lieu de départ.";
    if (column === "D") return "Veuillez renseigner le lieu d'arrivée.";
    if (column === "E") return "Veuillez renseigner la longueur en mètres.";
    if (column === "F") return "Veuillez renseigner la largeur minimale côté façade.";
    if (column === "G") return "Veuillez renseigner le type de revêtement.";
    if (column === "H") return "Veuillez renseigner l'état de la chaussée.";
    if (column === "I") return "Veuillez renseigner le type de caniveaux.";
    if (column === "J") return "Veuillez renseigner l'état de l'assainissement.";
    if (column === "L") return "Veuillez renseigner l'intervention à prévoir.";
  }
  if (sheetName === "Feuil5") {
    if (column === "A") return "Veuillez renseigner le numéro du tronçon.";
    if (column === "B") return "Veuillez renseigner le numéro de section.";
    if (column === "C") return "Veuillez renseigner le code de la voie.";
    if (column === "D") return "Veuillez renseigner le nom de la voie.";
    if (column === "E") return "Veuillez renseigner le lieu de départ.";
    if (column === "F") return "Veuillez renseigner le lieu d'arrivée.";
    if (column === "G") return "Veuillez renseigner la longueur en mètres.";
    if (column === "H") return "Veuillez renseigner la largeur minimale côté façade.";
    if (column === "I") return "Veuillez renseigner le type de revêtement.";
    if (column === "J") return "Veuillez renseigner l'état de la chaussée.";
    if (column === "K") return "Veuillez renseigner le type d'assainissement.";
    if (column === "L") return "Veuillez renseigner l'état de l'assainissement.";
    if (column === "M") return "Veuillez renseigner la largeur minimale des trottoirs.";
  }
  if (sheetName === "Feuil6") {
    if (column === "B") return "Veuillez renseigner le type de voie.";
    if (column === "C") return "Veuillez renseigner le code de la voie.";
    if (column === "D") return "Veuillez renseigner le linéaire en mètres.";
    if (column === "E") return "Veuillez renseigner le nom proposé.";
    if (column === "F") return "Veuillez renseigner l'itinéraire.";
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
    const tronconNo = String(cells.A ?? "").trim();
    const sectionNo = String(cells.B ?? "").trim();
    const lengthM = parseNumberValue(cells.G);
    if (tronconNo && !/^[1-9][0-9]*$/.test(tronconNo)) {
      setSheetFieldError(fieldErrors, "A", "Le numéro du tronçon doit être un nombre entier positif.");
    }
    if (sectionNo && !/^[1-9][0-9]*_[1-9][0-9]*$/.test(sectionNo)) {
      setSheetFieldError(fieldErrors, "B", "Écrivez par exemple 1_1, 2_3 ou 6_1. Le nombre avant _ crée le groupe SAP.");
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
    const tronconNo = String(cells.A ?? "").trim();
    const sectionNo = String(cells.B ?? "").trim();
    const numericColumns: Array<[SheetColumnKey, string, boolean]> = [
      ["G", "La longueur doit être un nombre supérieur à 0.", true],
      ["H", "La largeur côté façade doit être un nombre supérieur à 0.", true],
      ["M", "La largeur des trottoirs doit être un nombre positif ou nul.", false],
      ["N", "La valeur du stationnement à gauche doit être un nombre positif ou nul.", false],
      ["O", "La valeur du stationnement à droite doit être un nombre positif ou nul.", false]
    ];

    if (tronconNo && !/^[1-9][0-9]*$/.test(tronconNo)) {
      setSheetFieldError(fieldErrors, "A", "Le numéro du tronçon doit être un nombre entier positif.");
    }
    if (sectionNo && !/^[1-9][0-9]*_[1-9][0-9]*$/.test(sectionNo)) {
      setSheetFieldError(fieldErrors, "B", "Écrivez par exemple 1_1, 2_3 ou 6_1.");
    }

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
    const linearM = parseNumberValue(cells.D);
    if (String(cells.D ?? "").trim() && (!Number.isFinite(linearM) || Number(linearM) <= 0)) {
      setSheetFieldError(fieldErrors, "D", "Le linéaire doit être un nombre supérieur à 0.");
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

function parseFeuil6SapMarker(value: unknown) {
  const text = String(value ?? "").trim();
  const explicit = text.match(/SAP\s*([1-9][0-9]?)/i);
  if (explicit) {
    return `SAP${explicit[1]}`;
  }
  return "";
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
  const itineraryKey = normalizeRoadCompareKey(row.F);

  return (
    roads.find((road) => {
      const roadCode = normalizeRoadCompareKey(road.roadCode);
      const roadDesignation = normalizeRoadCompareKey(road.designation);
      const roadItinerary = normalizeRoadCompareKey(
        road.startLabel && road.endLabel ? `${road.startLabel} à ${road.endLabel}` : road.itinerary
      );

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

export default function App() {
  const hasElectronBridge = Boolean(window.padApp);
  const appName = window.padApp?.appName || "PAD Maintenance Routière";
  const appVersion = window.padApp?.appVersion || "0.0.0";
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
  const [maintenanceResponsibleName, setMaintenanceResponsibleName] = useState("");
  const [maintenanceAttachmentPath, setMaintenanceAttachmentPath] = useState("");
  const [maintenanceObservation, setMaintenanceObservation] = useState("");
  const [maintenanceCostAmount, setMaintenanceCostAmount] = useState("");
  const [feuil4Snapshot, setFeuil4Snapshot] = useState<Feuil4Snapshot | null>(null);
  const [solutionTemplates, setSolutionTemplates] = useState<MaintenanceSolutionTemplate[]>([]);
  const [selectedTemplateKey, setSelectedTemplateKey] = useState("");
  const [solutionDraft, setSolutionDraft] = useState("");
  const [shouldScrollToDegradationEditor, setShouldScrollToDegradationEditor] = useState(false);
  const [isDegradationEditorHighlighted, setIsDegradationEditorHighlighted] = useState(false);
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
  const [isMeasurementLoading, setIsMeasurementLoading] = useState(false);
  const [isMeasurementBusy, setIsMeasurementBusy] = useState(false);
  const [isHistoryLoading, setIsHistoryLoading] = useState(false);
  const [isMaintenanceLoading, setIsMaintenanceLoading] = useState(false);
  const [isMaintenanceBusy, setIsMaintenanceBusy] = useState(false);
  const [isPreviewingImport, setIsPreviewingImport] = useState(false);
  const [isBackupBusy, setIsBackupBusy] = useState(false);
  const [isReportingBusy, setIsReportingBusy] = useState(false);
  const [isSolutionBusy, setIsSolutionBusy] = useState(false);
  const [isDrainageRuleBusy, setIsDrainageRuleBusy] = useState(false);
  const [isAttachmentBusy, setIsAttachmentBusy] = useState(false);
  const supportsMaintenanceAttachments =
    typeof window.padApp?.maintenance?.pickAttachment === "function" &&
    typeof window.padApp?.maintenance?.openAttachment === "function";
  const [isImportAssistantCollapsed, setIsImportAssistantCollapsed] = useState(false);
  const [isIntegrityAlertDismissed, setIsIntegrityAlertDismissed] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const degradationEditorRef = useRef<HTMLDivElement | null>(null);
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
  const feuil2Sections = useMemo(() => {
    if (activeSheetName !== "Feuil2") {
      return [];
    }

    return rows
      .map((row) => {
        const sapCode = parseFeuil2SapCode(row);
        const lengthM = parseNumberValue(row.G);
        return {
          row,
          sapCode,
          tronconNo: String(row.A ?? "").trim(),
          sectionNo: String(row.B ?? "").trim(),
          roadLabel: String(row.C ?? "").trim(),
          designation: String(row.D ?? "").trim(),
          startLabel: String(row.E ?? "").trim(),
          endLabel: String(row.F ?? "").trim(),
          lengthM
        };
      })
      .filter((item) => item.roadLabel || item.designation || item.sapCode);
  }, [activeSheetName, rows]);
  const feuil2SapOptions = useMemo(
    () => [...new Set(feuil2Sections.map((item) => item.sapCode).filter(Boolean))].sort((a, b) => a.localeCompare(b)),
    [feuil2Sections]
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
      .sort((a, b) => a.sapCode.localeCompare(b.sapCode))
      .map((group) => ({
        ...group,
        rows: group.rows.sort((left, right) => getDisplayRowNumber(left.row) - getDisplayRowNumber(right.row))
      }));
  }, [feuil2SapFilter, feuil2Sections]);
  const feuil6DirectoryRows = useMemo(() => {
    if (activeSheetName !== "Feuil6") {
      return [];
    }

    const items: Array<{
      row: SheetRow;
      sapCode: string;
      roadType: string;
      roadCode: string;
      linearM: number | null;
      proposedName: string;
      itinerary: string;
      justification: string;
      linkedRoad: RoadCatalogItem | null;
    }> = [];

    let currentSap = "";
    for (const row of rows) {
      const sapMarker =
        parseFeuil6SapMarker(row.A) || parseFeuil6SapMarker(row.F) || parseFeuil6SapMarker(row.E) || currentSap;
      if (sapMarker && sapMarker !== currentSap) {
        currentSap = sapMarker;
      }

      const roadCode = String(row.C ?? "").trim();
      const proposedName = String(row.E ?? "").trim();
      if (!roadCode && !proposedName) {
        continue;
      }

      items.push({
        row,
        sapCode: currentSap || "",
        roadType: String(row.B ?? "").trim(),
        roadCode,
        linearM: parseNumberValue(row.D),
        proposedName,
        itinerary: String(row.F ?? "").trim(),
        justification: String(row.G ?? "").trim(),
        linkedRoad: resolveRoadFromFeuil6Row(row, allRoads)
      });
    }

    return items;
  }, [activeSheetName, allRoads, rows]);
  const feuil6SapOptions = useMemo(
    () => [...new Set(feuil6DirectoryRows.map((item) => item.sapCode).filter(Boolean))].sort((a, b) => a.localeCompare(b)),
    [feuil6DirectoryRows]
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

    return [...grouped.values()].sort((a, b) => a.sapCode.localeCompare(b.sapCode));
  }, [feuil6DirectoryRows, feuil6SapFilter]);
  const feuil6LinkedCount = useMemo(
    () => feuil6DirectoryRows.filter((item) => item.linkedRoad).length,
    [feuil6DirectoryRows]
  );
  const feuil3Profiles = useMemo(() => {
    if (activeSheetName !== "Feuil3") {
      return [];
    }

    return rows
      .map((row) => {
        const linkedRoad = resolveRoadFromFeuil3Row(row, allRoads);
        return {
          row,
          linkedRoad,
          sapCode: linkedRoad?.sapCode || "",
          roadLabel: String(row.A ?? "").trim(),
          designation: String(row.B ?? "").trim(),
          startLabel: String(row.C ?? "").trim(),
          endLabel: String(row.D ?? "").trim(),
          lengthM: parseNumberValue(row.E),
          facadeWidthM: parseNumberValue(row.F),
          surfaceType: String(row.G ?? "").trim(),
          pavementState: String(row.H ?? "").trim(),
          drainageType: String(row.I ?? "").trim(),
          drainageState: String(row.J ?? "").trim(),
          sidewalkMinM: parseNumberValue(row.K),
          interventionHint: String(row.L ?? "").trim()
        };
      })
      .filter((item) => item.roadLabel || item.designation || item.surfaceType || item.interventionHint);
  }, [activeSheetName, allRoads, rows]);
  const feuil3SapOptions = useMemo(
    () => [...new Set(feuil3Profiles.map((item) => item.sapCode).filter(Boolean))].sort((a, b) => a.localeCompare(b)),
    [feuil3Profiles]
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

    return [...grouped.values()].sort((a, b) => a.sapCode.localeCompare(b.sapCode));
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
  const feuil5Profiles = useMemo(() => {
    if (activeSheetName !== "Feuil5") {
      return [];
    }

    return rows
      .map((row) => {
        const linkedRoad = resolveRoadFromFeuil2Row(row, allRoads);
        return {
          row,
          linkedRoad,
          sapCode: linkedRoad?.sapCode || parseFeuil2SapCode(row) || "",
          tronconNo: String(row.A ?? "").trim(),
          sectionNo: String(row.B ?? "").trim(),
          roadLabel: String(row.C ?? "").trim(),
          designation: String(row.D ?? "").trim(),
          startLabel: String(row.E ?? "").trim(),
          endLabel: String(row.F ?? "").trim(),
          lengthM: parseNumberValue(row.G),
          facadeWidthM: parseNumberValue(row.H),
          surfaceType: String(row.I ?? "").trim(),
          pavementState: String(row.J ?? "").trim(),
          drainageType: String(row.K ?? "").trim(),
          drainageState: String(row.L ?? "").trim(),
          sidewalkMinM: parseNumberValue(row.M),
          parkingLeft: String(row.N ?? "").trim(),
          parkingRight: String(row.O ?? "").trim(),
          parkingOther: String(row.P ?? "").trim()
        };
      })
      .filter((item) => item.roadLabel || item.designation || item.surfaceType || item.drainageState);
  }, [activeSheetName, allRoads, rows]);
  const feuil5SapOptions = useMemo(
    () => [...new Set(feuil5Profiles.map((item) => item.sapCode).filter(Boolean))].sort((a, b) => a.localeCompare(b)),
    [feuil5Profiles]
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

    return [...grouped.values()].sort((a, b) => a.sapCode.localeCompare(b.sapCode));
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

  const inferRoadType = useCallback((roadCode: string) => {
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
  }, []);

  const suggestSectionForSap = useCallback(
    (sapCode: string) => {
      const sapNumber = String(sapCode ?? "").match(/([0-9]+)/)?.[1] || "";
      if (!sapNumber) {
        return { tronconNo: "", sectionNo: "" };
      }
      const prefix = `${sapNumber}_`;
      const currentIndexes = rows
        .map((row) => String(row.B ?? "").trim())
        .filter((value) => value.startsWith(prefix))
        .map((value) => Number(value.split("_")[1] || 0))
        .filter((value) => Number.isFinite(value) && value > 0);
      const nextIndex = currentIndexes.length > 0 ? Math.max(...currentIndexes) + 1 : 1;
      return {
        tronconNo: sapNumber,
        sectionNo: `${sapNumber}_${nextIndex}`
      };
    },
    [rows]
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
          next.A = String(existingRow.A ?? "");
          next.B = String(existingRow.B ?? "");
        } else {
          const suggestion = suggestSectionForSap(road.sapCode);
          next.A = next.A || suggestion.tronconNo;
          next.B = next.B || suggestion.sectionNo;
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
          next.A = String(existingRow.A ?? "");
          next.B = String(existingRow.B ?? "");
        } else {
          const suggestion = suggestSectionForSap(road.sapCode);
          next.A = next.A || suggestion.tronconNo;
          next.B = next.B || suggestion.sectionNo;
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
        next.B = inferRoadType(road.roadCode) || next.B || "";
        next.C = road.roadCode;
        next.D = road.lengthM !== null && road.lengthM !== undefined ? String(road.lengthM) : next.D || "";
        next.E = road.designation;
        next.F = road.itinerary || (road.startLabel && road.endLabel ? `${road.startLabel} à ${road.endLabel}` : next.F || "");
        next.G = road.justification || next.G || "";
      }

      return next;
    },
    [allRoads, editingRowId, inferRoadType, rows, suggestSectionForSap]
  );

  const validateDraftDuplicates = useCallback(
    (
      sheetName: string | undefined,
      cells: Partial<Record<SheetColumnKey, string>>
    ): { fieldErrors: Partial<Record<SheetColumnKey, string>>; formError: string } => {
      const fieldErrors: Partial<Record<SheetColumnKey, string>> = {};
      const comparableRows = rows.filter((row) => !editingRowId || row.id !== editingRowId);
      const draftRoad =
        sheetName === "Feuil2" || sheetName === "Feuil5" || sheetName === "Feuil3" || sheetName === "Feuil6"
          ? resolveDraftRoadMatch(sheetName, cells)
          : null;

      if (sheetName === "Feuil2" || sheetName === "Feuil5") {
        const sectionNo = normalizeLabel(cells.B);
        const startKey = normalizeLabel(cells.E);
        const endKey = normalizeLabel(cells.F);
        const duplicate = comparableRows.find((row) => {
          const rowRoad = resolveRoadFromFeuil2Row(row, allRoads);
          const sameRoad = draftRoad ? rowRoad?.id === draftRoad.id : normalizeRoadCompareKey(row.C) === normalizeRoadCompareKey(cells.C);
          if (!sameRoad) {
            return false;
          }
          if (sectionNo && normalizeLabel(row.B) === sectionNo) {
            return true;
          }
          return Boolean(startKey && endKey && normalizeLabel(row.E) === startKey && normalizeLabel(row.F) === endKey);
        });

        if (duplicate) {
          fieldErrors.B = `Cette section existe déjà pour cette voie (${toDisplay(duplicate.B)}).`;
          return { fieldErrors, formError: fieldErrors.B };
        }
      }

      if (sheetName === "Feuil3") {
        const duplicate = comparableRows.find((row) => {
          const rowRoad = resolveRoadFromFeuil3Row(row, allRoads);
          if (draftRoad) {
            return rowRoad?.id === draftRoad.id;
          }
          return normalizeRoadCompareKey(row.A) === normalizeRoadCompareKey(cells.A) || normalizeRoadCompareKey(row.B) === normalizeRoadCompareKey(cells.B);
        });

        if (duplicate) {
          fieldErrors.A = `Cette voie possède déjà un profil dans cette feuille (${toDisplay(duplicate.A)}).`;
          return { fieldErrors, formError: fieldErrors.A };
        }
      }

      if (sheetName === "Feuil6") {
        const codeKey = normalizeRoadCompareKey(cells.C);
        const duplicate = comparableRows.find((row) => codeKey && normalizeRoadCompareKey(row.C) === codeKey);
        if (duplicate) {
          fieldErrors.C = `Ce code de voie existe déjà dans le répertoire (${toDisplay(duplicate.C)}).`;
          return { fieldErrors, formError: fieldErrors.C };
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
            fieldErrors.B = `Cette référence existe déjà (${toDisplay(duplicate.B)}).`;
            return { fieldErrors, formError: fieldErrors.B };
          }
          fieldErrors.C = `Cette dégradation existe déjà (${toDisplay(duplicate.C)}).`;
          return { fieldErrors, formError: fieldErrors.C };
        }
      }

      if (sheetName === "Feuil4") {
        const labelKey = normalizeRoadCompareKey(cells.A);
        const duplicate = comparableRows.find((row) => labelKey && normalizeRoadCompareKey(row.A) === labelKey);
        if (duplicate) {
          fieldErrors.A = `Cette ligne existe déjà dans le programme d'évaluation (${toDisplay(duplicate.A)}).`;
          return { fieldErrors, formError: fieldErrors.A };
        }
      }

      return { fieldErrors, formError: "" };
    },
    [allRoads, editingRowId, resolveDraftRoadMatch, rows]
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

        const isRoadSelectorField =
          (activeSheetName === "Feuil2" && (column === "C" || column === "D")) ||
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
          loadRoads(),
          loadMeasurementCampaigns(),
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
    loadDashboardSummary,
    loadAllRoads,
    loadDrainageRules,
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
    if (!selectedDegradation) {
      setSelectedTemplateKey("");
      setSolutionDraft("");
      return;
    }
    setSelectedTemplateKey(selectedDegradation.templateKey ?? "");
    setSolutionDraft(selectedDegradation.solution || "");
  }, [selectedDegradation]);

  useEffect(() => {
    if (!shouldScrollToDegradationEditor || !selectedDegradation || !degradationEditorRef.current) {
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
  }, [selectedDegradation, shouldScrollToDegradationEditor]);

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
        loadRoads(),
        loadMeasurementCampaigns(),
        loadHistory(),
        loadMaintenanceRows(),
        loadFeuil4Snapshot(),
        loadSolutionTemplates(),
        loadDrainageRules()
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
          loadRoads(),
          loadMeasurementCampaigns(),
          loadHistory(),
          loadMaintenanceRows(),
          loadFeuil4Snapshot(),
          loadSolutionTemplates(),
          loadDrainageRules()
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

  function renderStandardSheetPrintHeader(_sheetTitle?: string) {
    return (
      <div className="print-sheet-header">
        <div className="print-sheet-header__brand">
          <img className="print-sheet-header__logo" src="/logo-pad.png" alt="Logo Port Autonome de Douala" />
          <div>
            <strong>{appName}</strong>
            <div>Pilotez la maintenance routière du PAD avec des décisions rapides et fiables.</div>
          </div>
        </div>
      </div>
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
        loadRoads(),
        loadMeasurementCampaigns(),
        loadHistory(),
        loadMaintenanceRows(),
        loadFeuil4Snapshot(),
        loadSolutionTemplates(),
        loadDrainageRules()
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
    if (!selectedRoadId) {
      setError("Sélectionne une voie.");
      scrollPageTop();
      return;
    }
    if (!selectedDegradationId) {
      setError("Sélectionne une dégradation.");
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
  }

  function handleUseMeasurementCampaign(campaign: MeasurementCampaignItem) {
    if (campaign.roadId) {
      setSelectedRoadId(campaign.roadId);
    }
    if (campaign.sapCode) {
      setSelectedSap(campaign.sapCode);
    }
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
    }
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

  function handleUseFeuil2Section(row: SheetRow) {
    const matchedRoad = resolveRoadFromFeuil2Row(row, allRoads);
    if (!matchedRoad) {
      setError("Aucune voie normalisée correspondante n'a été trouvée pour cette section.");
      return;
    }

    setSelectedRoadId(matchedRoad.id);
    if (matchedRoad.sapCode) {
      setSelectedSap(matchedRoad.sapCode);
    }
    setDecisionResult(null);
    setActiveView("decision");
    setNotice(`Section ${toDisplay(row.C)} chargée dans l'aide à la décision.`);
    setError("");
  }

  function handleUseFeuil6Road(row: SheetRow) {
    const matchedRoad = resolveRoadFromFeuil6Row(row, allRoads);
    if (!matchedRoad) {
      setError("Aucune voie normalisée correspondante n'a été trouvée pour cette entrée du répertoire.");
      return;
    }

    setSelectedRoadId(matchedRoad.id);
    if (matchedRoad.sapCode) {
      setSelectedSap(matchedRoad.sapCode);
    }
    setDecisionResult(null);
    setActiveView("decision");
    setNotice(`Voie ${matchedRoad.roadCode} chargée depuis le répertoire codifié.`);
    setError("");
  }

  function handleUseFeuil3Profile(row: SheetRow) {
    const matchedRoad = resolveRoadFromFeuil3Row(row, allRoads);
    if (!matchedRoad) {
      setError("Aucune voie normalisée correspondante n'a été trouvée pour ce profil technique.");
      return;
    }

    setSelectedRoadId(matchedRoad.id);
    if (matchedRoad.sapCode) {
      setSelectedSap(matchedRoad.sapCode);
    }
    setDecisionResult(null);
    setActiveView("decision");
    setNotice(`Profil technique ${matchedRoad.roadCode} chargé dans l'aide à la décision.`);
    setError("");
  }

  function handleUseFeuil5Profile(row: SheetRow) {
    const matchedRoad = resolveRoadFromFeuil2Row(row, allRoads);
    if (!matchedRoad) {
      setError("Aucune voie normalisée correspondante n'a été trouvée pour ce profil complémentaire.");
      return;
    }

    setSelectedRoadId(matchedRoad.id);
    if (matchedRoad.sapCode) {
      setSelectedSap(matchedRoad.sapCode);
    }
    setDecisionResult(null);
    setActiveView("decision");
    setNotice(`Profil complémentaire ${matchedRoad.roadCode} chargé dans l'aide à la décision.`);
    setError("");
  }

  function handleEdit(row: SheetRow) {
    if (!activeSheet) {
      return;
    }

    const nextCells = createEmptyCells(editableColumns);
    for (const column of editableColumns) {
      nextCells[column] = String(row[column] ?? "");
    }

    setEditingRowId(row.id);
    setDraftCells(nextCells);
    setDraftFieldErrors({});
    setDraftFormError("");
    setNotice(`Édition de la ligne ${getDisplayRowNumber(row)}.`);
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
        await padApi.createSheetRow(activeSheet.name, payload);
        setNotice(`Nouvelle ligne ajoutée.${isNewSap ? ` ${inferredSapCode} sera ajouté automatiquement à la liste des SAP.` : ""}`);
      }

      await Promise.all([
        refreshStatus(),
        loadRows(),
        loadDashboardSummary(),
        loadIntegrityReport(),
        refreshDecisionCatalogs(),
        loadAllRoads(),
        loadRoads(),
        loadMeasurementCampaigns()
      ]);
      if (activeSheet.name === "Feuil4") {
        await loadFeuil4Snapshot();
      }
      resetDraft();
      setError("");
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
      setError("Le pattern est obligatoire sauf pour l'opérateur ALWAYS.");
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
    setMaintenanceResponsibleName(intervention.responsibleName || "");
    setMaintenanceAttachmentPath(intervention.attachmentPath || "");
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
    const pavementStates = uniqueValues([...allRoads.map((item) => item.pavementState), "Bon", "Moy", "Mau", "Mauvais"]);
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
              <span className="kpi-card__label">Budget déclaré</span>
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
                              <strong>{sheet.name}</strong>
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

            <div className="card card--spaced">
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
            placeholder="Code ou désignation"
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

          <label htmlFor="measurement-campaign">Campagne Feuil1</label>
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

          <label className="field-label--spaced" htmlFor="degradation">Dégradation</label>
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
                  <img className="print-sheet-header__logo" src="/logo-pad.png" alt="Logo Port Autonome de Douala" />
                  <div>
                    <strong>{appName}</strong>
                    <div>Fiche d'aide à la décision de maintenance routière</div>
                  </div>
                </div>
              </div>

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
                  <th className="col-actions">Action</th>
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
                    <td className="col-actions">
                      <div className="row-buttons">
                        <button
                          className="row-action row-action--configure row-action--with-icon"
                          type="button"
                          onClick={() => {
                            setSelectedDegradationId(item.id);
                            setShouldScrollToDegradationEditor(true);
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
            <div
              ref={degradationEditorRef}
              className={`card card--spaced${isDegradationEditorHighlighted ? " card--highlighted" : ""}`}
            >
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
                    <th className="col-actions">Actions</th>
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
                      <td className="col-actions">
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
    const hasMaintenanceAttachment = maintenanceAttachmentPath.trim().length > 0;

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
              <label htmlFor="maintenance-date">Date prévue</label>
              <input
                id="maintenance-date"
                type="date"
                value={maintenanceDate}
                onChange={(event) => setMaintenanceDate(event.target.value)}
                disabled={isMaintenanceBusy}
              />
            </div>

            <div className="cell-field">
              <label htmlFor="maintenance-completion-date">Date réelle / clôture</label>
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
              <label htmlFor="maintenance-responsible">Responsable PAD</label>
              <input
                id="maintenance-responsible"
                value={maintenanceResponsibleName}
                onChange={(event) => setMaintenanceResponsibleName(event.target.value)}
                placeholder="Ex: Chef section voirie"
                disabled={isMaintenanceBusy}
              />
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
                    <td colSpan={14}>{isMaintenanceLoading ? "Chargement..." : "Aucun entretien enregistré."}</td>
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

  function renderFeuil1View() {
    const campaign = selectedMeasurementCampaign;
    const campaignLabel = [campaign?.roadCode, campaign?.designation].filter(Boolean).join(" - ") || "Campagne sélectionnée";
    const canCreateMeasurementRow = Boolean(selectedMeasurementCampaignKey);

    return (
      <main className="workspace workspace--full">
        <section className="panel table-panel table-panel--full sheet-print-view">
          {renderStandardSheetPrintHeader(activeSheet?.title ?? "Feuil1")}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet?.title ?? "Feuil1"}</h2>
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
                activeSheet?.title ?? "Feuil1",
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

    return (
      <section ref={sheetEditorRef} className="panel editor-panel editor-panel--sheet">
        <h2>{activeSheet.title}</h2>
        <p className="field-help">Les champs marqués d'un * sont obligatoires.</p>
        {draftFormError ? <p className="modal-feedback modal-feedback--error">{draftFormError}</p> : null}

        <div className="cells-grid">
          {editableColumns.map((column) => (
            <div className={`cell-field${draftFieldErrors[column] ? " cell-field--error" : ""}`} key={column}>
              <label htmlFor={`cell-${column}`}>
                {getColumnLabel(activeSheet, column)}
                {isSheetFieldRequired(activeSheet.name, column) ? <span className="field-label__required"> *</span> : null}
              </label>
              {(() => {
                const suggestions = getSheetFieldSuggestions(activeSheet?.name, column);
                const inputId = `cell-${column}`;
                const datalistId = suggestions.length > 0 ? `cell-suggestions-${activeSheet?.name ?? "sheet"}-${column}` : undefined;
                const useTextarea =
                  (activeSheet?.name === "Feuil3" && ["H", "J", "L"].includes(column)) ||
                  (activeSheet?.name === "Feuil5" && ["M", "N", "O", "P"].includes(column));
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

                if (useTextarea) {
                  return (
                    <>
                      <textarea
                        {...commonProps}
                        className={`input-textarea${draftFieldErrors[column] ? " cell-field--error" : ""}`}
                        rows={column === "L" ? 4 : 3}
                        placeholder={fieldPlaceholder}
                      />
                      {draftFieldErrors[column] ? <p className="field-error">{draftFieldErrors[column]}</p> : null}
                      {fieldHelp ? <p className="field-help">{fieldHelp}</p> : null}
                    </>
                  );
                }

                return (
                  <>
                    <input
                      {...commonProps}
                      className={draftFieldErrors[column] ? "cell-field--error" : undefined}
                      list={datalistId}
                      placeholder={fieldPlaceholder}
                    />
                    {draftFieldErrors[column] ? <p className="field-error">{draftFieldErrors[column]}</p> : null}
                    {datalistId ? (
                      <datalist id={datalistId}>
                        {suggestions.map((item) => (
                          <option key={`${column}-${item}`} value={item} />
                        ))}
                      </datalist>
                    ) : null}
                    {fieldHelp ? <p className="field-help">{fieldHelp}</p> : null}
                  </>
                );
              })()}
            </div>
          ))}
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
          {renderStandardSheetPrintHeader(activeSheet?.title ?? "Feuil2")}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet?.title ?? "Feuil2"}</h2>
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
              {renderSheetPrintButton(activeSheet?.title ?? "Feuil2")}
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
                        <tr key={item.row.id}>
                          <td className="col-actions">
                            <div className="row-buttons row-buttons--compact row-buttons--wrap">
                              <button
                                className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                                type="button"
                                onClick={() => handleUseFeuil2Section(item.row)}
                              >
                                <Gauge size={14} aria-hidden="true" />
                                <span>Évaluer</span>
                              </button>
                              <button
                                className="row-action row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() => handleEdit(item.row)}
                                title="Éditer"
                                aria-label="Éditer"
                              >
                                <Pencil size={15} aria-hidden="true" />
                              </button>
                              <button
                                className="row-action row-action--danger row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() => handleDeleteRow(item.row.id)}
                                title="Supprimer"
                                aria-label="Supprimer"
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
          {renderStandardSheetPrintHeader(activeSheet?.title ?? "Feuil6")}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet?.title ?? "Feuil6"}</h2>
              <p className="muted">
                Répertoire codifié central des voies: type, code, nom proposé, itinéraire et justification, réutilisés dans tout le système.
              </p>
            </div>
            <div className="sheet-header-actions">
              <div className="measurement-toolbar__meta">
                <span className="pill">SAP: {feuil6Groups.length}</span>
                <span className="pill">Voies: {totalRoads}</span>
                <span className="pill">Liées: {feuil6LinkedCount}</span>
                <span className="pill">Linéaire: {formatMeasurementNumber(totalLinear)} m</span>
              </div>
              {renderSheetPrintButton(activeSheet?.title ?? "Feuil6")}
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
                    placeholder="Code, nom proposé, itinéraire, justification"
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
                          <th>Itinéraires</th>
                          <th>Justification</th>
                          <th>Raccordement</th>
                        </tr>
                      </thead>
                      <tbody>
                        {group.rows.map((item, index) => (
                          <tr key={item.row.id}>
                            <td className="col-actions">
                              <div className="row-buttons row-buttons--compact row-buttons--wrap">
                                <button
                                  className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                                  type="button"
                                  onClick={() => handleUseFeuil6Road(item.row)}
                                  disabled={!item.linkedRoad}
                                >
                                  <Gauge size={14} aria-hidden="true" />
                                  <span>Évaluer</span>
                                </button>
                                <button
                                  className="row-action row-action--icon row-action--icon-sm"
                                  type="button"
                                  onClick={() => handleEdit(item.row)}
                                  title="Éditer"
                                  aria-label="Éditer"
                                >
                                  <Pencil size={15} aria-hidden="true" />
                                </button>
                                <button
                                  className="row-action row-action--danger row-action--icon row-action--icon-sm"
                                  type="button"
                                  onClick={() => handleDeleteRow(item.row.id)}
                                  title="Supprimer"
                                  aria-label="Supprimer"
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
                            <td>{toDisplay(item.itinerary)}</td>
                            <td>{toDisplay(item.justification)}</td>
                            <td>
                              <span className={`status-pill ${item.linkedRoad ? "status-pill--ok" : "status-pill--warning"}`}>
                                {item.linkedRoad ? "Liée" : "À vérifier"}
                              </span>
                            </td>
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
                Feuil3 porte l'état technique courant de la voie. C'est la feuille la plus proche du diagnostic terrain avant décision.
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
          </section>
        </div>
        <section className="panel table-panel table-panel--full sheet-print-view">
          {renderStandardSheetPrintHeader(activeSheet?.title ?? "Feuil3")}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet?.title ?? "Feuil3"}</h2>
              <p className="muted">
                Profil technique détaillé des voies: chaussée, assainissement, trottoirs et nature d'intervention recommandée.
              </p>
            </div>
            <div className="sheet-header-actions">
              <div className="measurement-toolbar__meta">
                <span className="pill">SAP: {feuil3Groups.length}</span>
                <span className="pill">Profils: {totalProfiles}</span>
                <span className="pill">À déterminer: {feuil3PendingInterventions}</span>
                <span className="pill">Assainissement à surveiller: {feuil3DrainageAlerts}</span>
              </div>
              {renderSheetPrintButton(activeSheet?.title ?? "Feuil3")}
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
                        <tr key={item.row.id}>
                          <td className="col-actions">
                            <div className="row-buttons row-buttons--compact row-buttons--wrap">
                              <button
                                className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                                type="button"
                                onClick={() => handleUseFeuil3Profile(item.row)}
                                disabled={!item.linkedRoad}
                              >
                                <Gauge size={14} aria-hidden="true" />
                                <span>Évaluer</span>
                              </button>
                              <button
                                className="row-action row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() => handleEdit(item.row)}
                                title="Éditer"
                                aria-label="Éditer"
                              >
                                <Pencil size={15} aria-hidden="true" />
                              </button>
                              <button
                                className="row-action row-action--danger row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() => handleDeleteRow(item.row.id)}
                                title="Supprimer"
                                aria-label="Supprimer"
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
                          <td>{toDisplay(item.pavementState)}</td>
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
                Feuil5 enrichit Feuil3 avec les usages latéraux de la voie: stationnement, trottoirs et niveau d'assainissement.
              </p>
            </div>

            <div className="card">
              <h3>Stationnement recensé</h3>
              <ul className="count-list">
                {feuil5Groups.map((group) => {
                  const count = group.rows.filter(
                    (item) =>
                      hasFeuil5ParkingValue(item.parkingLeft) ||
                      hasFeuil5ParkingValue(item.parkingRight) ||
                      hasFeuil5ParkingValue(item.parkingOther)
                  ).length;
                  return (
                    <li key={`${group.sapCode}-parking`}>
                      <span>{group.sapCode}</span>
                      <strong>{count}</strong>
                    </li>
                  );
                })}
              </ul>
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
          {renderStandardSheetPrintHeader(activeSheet?.title ?? "Feuil5")}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet?.title ?? "Feuil5"}</h2>
              <p className="muted">
                Compléments techniques des sections: assainissement, largeur minimale, trottoirs et stationnement, rattachés à la même voie centrale.
              </p>
            </div>
            <div className="sheet-header-actions">
              <div className="measurement-toolbar__meta">
                <span className="pill">SAP: {feuil5Groups.length}</span>
                <span className="pill">Profils: {totalProfiles}</span>
                <span className="pill">Stationnement recensé: {feuil5ParkingCount}</span>
                <span className="pill">Assainissement à surveiller: {feuil5DrainageWatchCount}</span>
              </div>
              {renderSheetPrintButton(activeSheet?.title ?? "Feuil5")}
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
                  placeholder="Voie, désignation, assainissement, stationnement"
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
                  <span className="status-pill">{group.rows.length} section(s)</span>
                </div>
                <div className="table-wrap feuil2-table-wrap">
                  <table className="table--feuil5">
                    <thead>
                      <tr>
                        <th className="col-actions" rowSpan={3}>Action</th>
                        <th rowSpan={3}>N°</th>
                        <th rowSpan={3}>N° tronçon</th>
                        <th rowSpan={3}>N° sections</th>
                        <th rowSpan={3}>Voies</th>
                        <th rowSpan={3}>Désignation</th>
                        <th rowSpan={3}>Début</th>
                        <th rowSpan={3}>Fin</th>
                        <th rowSpan={3}>Longueur (m)</th>
                        <th rowSpan={3}>Largeur min. façade (m)</th>
                        <th colSpan={2}>CHAUSSÉE (en %)</th>
                        <th colSpan={2}>ASSAINISSEMENT</th>
                        <th rowSpan={3}>Largeur min. trottoirs (m)</th>
                        <th colSpan={3}>STATIONNEMENT</th>
                      </tr>
                      <tr>
                        <th>Nature du revêtement</th>
                        <th>État de la chaussée</th>
                        <th>Type</th>
                        <th>État</th>
                        <th>Gauche</th>
                        <th>Droit</th>
                        <th>Autres</th>
                      </tr>
                      <tr>
                        <th>&nbsp;</th>
                        <th>&nbsp;</th>
                        <th>caniveaux</th>
                        <th>description</th>
                        <th>&nbsp;</th>
                        <th>&nbsp;</th>
                        <th>&nbsp;</th>
                      </tr>
                    </thead>
                    <tbody>
                      {group.rows.map((item, index) => (
                        <tr key={item.row.id}>
                          <td className="col-actions">
                            <div className="row-buttons row-buttons--compact row-buttons--wrap">
                              <button
                                className="row-action row-action--evaluate row-action--with-icon row-action--compact"
                                type="button"
                                onClick={() => handleUseFeuil5Profile(item.row)}
                                disabled={!item.linkedRoad}
                              >
                                <Gauge size={14} aria-hidden="true" />
                                <span>Évaluer</span>
                              </button>
                              <button
                                className="row-action row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() => handleEdit(item.row)}
                                title="Éditer"
                                aria-label="Éditer"
                              >
                                <Pencil size={15} aria-hidden="true" />
                              </button>
                              <button
                                className="row-action row-action--danger row-action--icon row-action--icon-sm"
                                type="button"
                                onClick={() => handleDeleteRow(item.row.id)}
                                title="Supprimer"
                                aria-label="Supprimer"
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
                          <td>{toDisplay(item.surfaceType)}</td>
                          <td>{toDisplay(item.pavementState)}</td>
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
                <p className="muted">{isLoadingRows ? "Chargement..." : "Aucun profil Feuil5 disponible."}</p>
              </div>
            ) : null}
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

    return (
      <main className="workspace">
        {renderSheetEditorPanel()}

        <section className="panel table-panel sheet-print-view">
          {renderStandardSheetPrintHeader(activeSheet?.title ?? "Feuille")}
          <div className="dashboard-card__header">
            <div>
              <h2>{activeSheet?.title ?? "Feuille"}</h2>
              <p className="muted">Impression du tableau de la feuille active, avec les colonnes et lignes actuellement affichées.</p>
            </div>
            {renderSheetPrintButton(activeSheet?.title ?? "Feuille")}
          </div>

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
              <button className="integrity-alert__action" type="button" onClick={() => setActiveView("dashboard")}>
                Gérer
              </button>
            )}
          </div>
        ) : null}
        <nav className="hero__nav">
          <button className={activeView === "dashboard" ? "active" : ""} type="button" onClick={() => setActiveView("dashboard")}>
            Tableau de bord
          </button>
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
            Historique ({status?.decisionHistoryCount ?? 0})
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




