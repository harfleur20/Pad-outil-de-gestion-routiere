import type {
  DataIntegrityReport,
  DataStatus,
  DrainageRule,
  DecisionHistoryItem,
  DecisionResult,
  DegradationItem,
  MaintenanceSolutionTemplate,
  RoadCatalogItem,
  SapSector,
  SheetDefinition,
  SheetRow,
  SheetRowFilters,
  SheetRowPayload
} from "../types/pad-api";

const FALLBACK_STATUS: DataStatus = {
  sheetCounts: {},
  totalRows: 0,
  decisionHistoryCount: 0,
  lastImportPath: null,
  lastImportAt: null
};

function requireBridge() {
  if (!window.padApp) {
    throw new Error("Bridge Electron indisponible. Lance l'application avec 'npm run dev' (pas 'npm run dev:web').");
  }
  return window.padApp;
}

export const padApi = {
  async getDataStatus(): Promise<DataStatus> {
    if (!window.padApp?.data?.getStatus) {
      return FALLBACK_STATUS;
    }
    return window.padApp.data.getStatus();
  },

  async getDataIntegrityReport(): Promise<DataIntegrityReport> {
    const bridge = requireBridge();
    return bridge.audit.integrity();
  },

  async importFromExcel(excelPath?: string): Promise<DataStatus> {
    const bridge = requireBridge();
    return bridge.data.importFromExcel(excelPath);
  },

  async pickExcelFile(): Promise<string | null> {
    const bridge = requireBridge();
    return bridge.data.pickExcelFile();
  },

  async listSheetDefinitions(): Promise<SheetDefinition[]> {
    const bridge = requireBridge();
    return bridge.sheet.definitions();
  },

  async listSheetRows(sheetName: string, filters?: SheetRowFilters): Promise<SheetRow[]> {
    const bridge = requireBridge();
    return bridge.sheet.list(sheetName, filters);
  },

  async createSheetRow(sheetName: string, payload: SheetRowPayload): Promise<SheetRow> {
    const bridge = requireBridge();
    return bridge.sheet.create(sheetName, payload);
  },

  async updateSheetRow(sheetName: string, rowId: number, payload: SheetRowPayload): Promise<SheetRow> {
    const bridge = requireBridge();
    return bridge.sheet.update(sheetName, rowId, payload);
  },

  async deleteSheetRow(sheetName: string, rowId: number): Promise<boolean> {
    const bridge = requireBridge();
    return bridge.sheet.delete(sheetName, rowId);
  },

  async listSapSectors(): Promise<SapSector[]> {
    const bridge = requireBridge();
    return bridge.sap.list();
  },

  async listRoads(filters?: { sapCode?: string; search?: string }): Promise<RoadCatalogItem[]> {
    const bridge = requireBridge();
    return bridge.roads.list(filters);
  },

  async listDegradations(): Promise<DegradationItem[]> {
    const bridge = requireBridge();
    return bridge.degradations.list();
  },

  async listDrainageRules(): Promise<DrainageRule[]> {
    const bridge = requireBridge();
    return bridge.drainageRules.list();
  },

  async upsertDrainageRule(payload: {
    id?: number;
    ruleOrder: number;
    matchOperator: "CONTAINS" | "EQUALS" | "REGEX" | "ALWAYS";
    pattern?: string;
    askRequired?: boolean;
    needsAttention?: boolean;
    recommendation: string;
    isActive?: boolean;
  }): Promise<DrainageRule | null> {
    const bridge = requireBridge();
    return bridge.drainageRules.upsert(payload);
  },

  async deleteDrainageRule(ruleId: number): Promise<{ deleted: boolean }> {
    const bridge = requireBridge();
    return bridge.drainageRules.delete(ruleId);
  },

  async listSolutionTemplates(): Promise<MaintenanceSolutionTemplate[]> {
    const bridge = requireBridge();
    return bridge.solutions.listTemplates();
  },

  async upsertSolutionTemplate(payload: {
    templateKey: string;
    title: string;
    description: string;
  }): Promise<MaintenanceSolutionTemplate | null> {
    const bridge = requireBridge();
    return bridge.solutions.upsertTemplate(payload);
  },

  async assignTemplateToDegradation(degradationCode: string, templateKey: string): Promise<{ degradationCode: string; templateKey: string }> {
    const bridge = requireBridge();
    return bridge.solutions.assignTemplate(degradationCode, templateKey);
  },

  async setDegradationSolutionOverride(degradationCode: string, solutionText: string): Promise<{ degradationCode: string; solutionText: string }> {
    const bridge = requireBridge();
    return bridge.solutions.setOverride(degradationCode, solutionText);
  },

  async clearDegradationSolutionOverride(degradationCode: string): Promise<{ degradationCode: string; cleared: boolean }> {
    const bridge = requireBridge();
    return bridge.solutions.clearOverride(degradationCode);
  },

  async evaluateDecision(payload: {
    roadId?: number;
    roadKey?: string;
    roadCode?: string;
    designation?: string;
    degradationId?: number;
    degradationName?: string;
    deflectionValue?: number;
    askDrainage?: boolean;
  }): Promise<DecisionResult> {
    const bridge = requireBridge();
    return bridge.decision.evaluate(payload);
  },

  async listDecisionHistory(filters?: { sapCode?: string; search?: string; limit?: number }): Promise<DecisionHistoryItem[]> {
    const bridge = requireBridge();
    return bridge.reporting.listHistory(filters);
  },

  async clearDecisionHistory(): Promise<{ deleted: boolean }> {
    const bridge = requireBridge();
    return bridge.reporting.clearHistory();
  }
};
