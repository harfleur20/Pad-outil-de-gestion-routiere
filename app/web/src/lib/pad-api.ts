import type {
  AttachmentUploadResult,
  BackupResult,
  DataIntegrityReport,
  DataStatus,
  DashboardSummary,
  DrainageRule,
  DecisionHistoryItem,
  DecisionResult,
  DegradationItem,
  ImportPreview,
  MeasurementCampaignItem,
  MeasurementCampaignPayload,
  MaintenanceInterventionItem,
  MaintenanceInterventionPayload,
  MaintenanceSolutionTemplate,
  PdfExportResult,
  ReportExportResult,
  RoadCatalogItem,
  RoadSectionItem,
  RoadMeasurementItem,
  RoadMeasurementPayload,
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

function requireFunctionBridgeMethod<T extends (...args: never[]) => unknown>(
  method: T | undefined,
  message: string
): T {
  if (typeof method !== "function") {
    throw new Error(message);
  }
  return method;
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

  async getDashboardSummary(): Promise<DashboardSummary> {
    const bridge = requireBridge();
    return bridge.dashboard.summary();
  },

  async importFromExcel(excelPath?: string): Promise<DataStatus> {
    const bridge = requireBridge();
    return bridge.data.importFromExcel(excelPath);
  },

  async previewExcelImport(excelPath?: string): Promise<ImportPreview> {
    const bridge = requireBridge();
    return bridge.data.previewExcelImport(excelPath);
  },

  async pickExcelFile(): Promise<string | null> {
    const bridge = requireBridge();
    return bridge.data.pickExcelFile();
  },

  async exportBackup(): Promise<BackupResult | null> {
    const bridge = requireBridge();
    return bridge.backup.export();
  },

  async restoreBackup(): Promise<DataStatus | null> {
    const bridge = requireBridge();
    return bridge.backup.restore();
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

  async listRoadSections(filters?: { sapCode?: string; search?: string }): Promise<RoadSectionItem[]> {
    const bridge = requireBridge();
    return bridge.roadSections.list(filters);
  },

  async listMeasurementCampaigns(filters?: {
    roadId?: number;
    search?: string;
    limit?: number;
  }): Promise<MeasurementCampaignItem[]> {
    const bridge = requireBridge();
    return bridge.measurement.listCampaigns(filters);
  },

  async listRoadMeasurements(filters?: {
    campaignKey?: string;
    roadId?: number;
    limit?: number;
  }): Promise<RoadMeasurementItem[]> {
    const bridge = requireBridge();
    return bridge.measurement.listRows(filters);
  },

  async upsertMeasurementCampaign(payload: MeasurementCampaignPayload): Promise<MeasurementCampaignItem | null> {
    const bridge = requireBridge();
    return bridge.measurement.upsertCampaign(payload);
  },

  async deleteMeasurementCampaign(campaignId: number): Promise<{ deleted: boolean }> {
    const bridge = requireBridge();
    return bridge.measurement.deleteCampaign(campaignId);
  },

  async upsertRoadMeasurement(payload: RoadMeasurementPayload): Promise<RoadMeasurementItem | null> {
    const bridge = requireBridge();
    return bridge.measurement.upsertRow(payload);
  },

  async deleteRoadMeasurement(measurementId: number): Promise<{ deleted: boolean }> {
    const bridge = requireBridge();
    return bridge.measurement.deleteRow(measurementId);
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

  async listMaintenanceInterventions(filters?: {
    sapCode?: string;
    roadId?: number;
    status?: string;
    search?: string;
    limit?: number;
  }): Promise<MaintenanceInterventionItem[]> {
    const bridge = requireBridge();
    return bridge.maintenance.list(filters);
  },

  async upsertMaintenanceIntervention(payload: MaintenanceInterventionPayload): Promise<MaintenanceInterventionItem | null> {
    const bridge = requireBridge();
    return bridge.maintenance.upsert(payload);
  },

  async deleteMaintenanceIntervention(interventionId: number): Promise<{ deleted: boolean }> {
    const bridge = requireBridge();
    return bridge.maintenance.delete(interventionId);
  },

  async pickMaintenanceAttachment(): Promise<AttachmentUploadResult | null> {
    const bridge = requireBridge();
    const pickAttachment = requireFunctionBridgeMethod(
      bridge.maintenance?.pickAttachment,
      "La fonction de pièce jointe n'est pas disponible dans cette session Electron. Ferme puis relance l'application."
    );
    return pickAttachment();
  },

  async openMaintenanceAttachment(attachmentPath: string): Promise<{ opened: boolean }> {
    const bridge = requireBridge();
    const openAttachment = requireFunctionBridgeMethod(
      bridge.maintenance?.openAttachment,
      "L'ouverture des pièces jointes n'est pas disponible dans cette session Electron. Ferme puis relance l'application."
    );
    return openAttachment(attachmentPath);
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
  },

  async exportDecisionHistoryXlsx(): Promise<ReportExportResult | null> {
    const bridge = requireBridge();
    return bridge.reporting.exportHistoryXlsx();
  },

  async exportMaintenanceHistoryXlsx(): Promise<ReportExportResult | null> {
    const bridge = requireBridge();
    return bridge.reporting.exportMaintenanceXlsx();
  },

  async exportCurrentViewPdf(suggestedName?: string): Promise<PdfExportResult | null> {
    const bridge = requireBridge();
    return bridge.printing.exportCurrentViewPdf(suggestedName);
  }
};




