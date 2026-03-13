export const SHEET_COLUMN_KEYS = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P"
] as const;

export type SheetColumnKey = (typeof SHEET_COLUMN_KEYS)[number];

export type SheetDefinition = {
  name: string;
  title: string;
  description: string;
  columns: SheetColumnKey[];
  columnLabels?: Partial<Record<SheetColumnKey, string>>;
};

export type SheetRow = {
  id: number;
  rowNo: number;
} & Partial<Record<SheetColumnKey, string>>;

export type SheetRowPayload = {
  rowNo?: number;
} & Partial<Record<SheetColumnKey, string>>;

export type SheetRowFilters = {
  search?: string;
  limit?: number;
};

export type DataStatus = {
  sheetCounts: Record<string, number>;
  totalRows: number;
  decisionHistoryCount: number;
  lastImportPath: string | null;
  lastImportAt: string | null;
};

export type DataIntegrityIssue = {
  code: string;
  level: "WARNING" | "ERROR";
  count: number;
  message: string;
};

export type DataIntegrityReport = {
  generatedAt: string;
  status: "OK" | "WARNING";
  totals: {
    roads: number;
    degradations: number;
    roadSections: number;
    roadMeasurements: number;
    profileInputs: number;
    decisionHistory: number;
  };
  issues: DataIntegrityIssue[];
};

export type ImportSheetPreview = {
  name: string;
  title: string;
  present: boolean;
  rowCount: number;
  expectedColumns: number;
};

export type ImportPreview = {
  filePath: string;
  workbookSheetNames: string[];
  missingSheets: string[];
  warnings: string[];
  ready: boolean;
  totals: {
    rows: number;
    roads: number;
    degradations: number;
    sections: number;
  };
  sheetPreviews: ImportSheetPreview[];
};

export type CountByLabel = {
  label: string;
  count: number;
};

export type BackupResult = {
  filePath: string;
  size?: number;
  exportedAt?: string;
};

export type ReportExportResult = {
  filePath: string;
  reportType: "history" | "maintenance";
  rowCount: number;
};

export type AttachmentUploadResult = {
  storedPath: string;
  fileName: string;
  size: number;
};

export type DashboardSummary = {
  generatedAt: string;
  totals: {
    roads: number;
    degradations: number;
    decisionHistory: number;
    maintenance: number;
    pendingMaintenance: number;
    completedMaintenance: number;
    estimatedBudget: number;
    urgentDrainage: number;
  };
  roadsBySap: CountByLabel[];
  roadsByState: CountByLabel[];
  maintenanceByStatus: CountByLabel[];
  topDegradations: CountByLabel[];
  integrity: DataIntegrityReport;
  recentMaintenance: MaintenanceInterventionItem[];
};

export type SapSector = {
  code: string;
  name: string;
};

export type RoadCatalogItem = {
  id: number;
  roadKey: string;
  roadCode: string;
  designation: string;
  sapCode: string;
  startLabel: string;
  endLabel: string;
  lengthM: number | null;
  widthM: number | null;
  surfaceType: string;
  pavementState: string;
  drainageType: string;
  drainageState: string;
  sidewalkMinM: number | null;
  parkingLeft: string;
  parkingRight: string;
  parkingOther: string;
  itinerary: string;
  justification: string;
  interventionHint?: string;
};

export type DegradationItem = {
  id: number;
  code: string;
  name: string;
  causes: string[];
  solution: string;
  solutionSource: "TEMPLATE" | "OVERRIDE" | "MISSING";
  templateKey: string | null;
};

export type MaintenanceSolutionTemplate = {
  templateKey: string;
  title: string;
  description: string;
};

export type DrainageRule = {
  id: number;
  ruleOrder: number;
  matchOperator: "CONTAINS" | "EQUALS" | "REGEX" | "ALWAYS";
  pattern: string;
  askRequired: boolean;
  needsAttention: boolean;
  recommendation: string;
  isActive: boolean;
};

export type DecisionDeflection = {
  value: number | null;
  severity: string;
  recommendation: string;
};

export type DecisionDrainage = {
  needsAttention: boolean;
  recommendation: string;
  ruleId?: number | null;
};

export type DecisionResult = {
  road: RoadCatalogItem;
  degradation: DegradationItem;
  probableCause: string;
  maintenanceSolution: string;
  contextualIntervention: string | null;
  deflection: DecisionDeflection;
  drainage: DecisionDrainage;
};

export type DecisionHistoryItem = {
  id: number;
  createdAt: string;
  roadId: number | null;
  roadCode: string;
  roadDesignation: string;
  sapCode: string;
  startLabel: string;
  endLabel: string;
  degradationName: string;
  probableCause: string;
  maintenanceSolution: string;
  contextualIntervention: string;
  deflectionValue: number | null;
  deflectionSeverity: string;
  deflectionRecommendation: string;
  drainageNeedsAttention: boolean;
  drainageRecommendation: string;
};

export type MaintenanceInterventionStatus = "PREVU" | "EN_COURS" | "TERMINE";

export type MaintenanceInterventionItem = {
  id: number;
  createdAt: string;
  updatedAt: string;
  roadId: number | null;
  roadKey: string;
  roadCode: string;
  roadDesignation: string;
  sapCode: string;
  degradationCode: string;
  degradationName: string;
  interventionType: string;
  status: MaintenanceInterventionStatus;
  interventionDate: string;
  completionDate: string;
  stateBefore: string;
  stateAfter: string;
  deflectionBefore: number | null;
  deflectionAfter: number | null;
  solutionApplied: string;
  contractorName: string;
  responsibleName: string;
  attachmentPath: string;
  observation: string;
  costAmount: number | null;
};

export type MaintenanceInterventionPayload = {
  id?: number;
  roadId: number;
  degradationCode?: string;
  interventionType: string;
  status: MaintenanceInterventionStatus;
  interventionDate: string;
  completionDate?: string;
  stateBefore?: string;
  stateAfter?: string;
  deflectionBefore?: number;
  deflectionAfter?: number;
  solutionApplied?: string;
  contractorName?: string;
  responsibleName?: string;
  attachmentPath?: string;
  observation?: string;
  costAmount?: number;
};

