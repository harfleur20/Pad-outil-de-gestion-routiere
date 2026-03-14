/// <reference types="vite/client" />

import type {
  AttachmentUploadResult,
  BackupResult,
  DataIntegrityReport,
  DataStatus,
  DashboardSummary,
  DecisionHistoryItem,
  DecisionResult,
  DrainageRule,
  DegradationItem,
  ImportPreview,
  MeasurementCampaignItem,
  MeasurementCampaignPayload,
  MaintenanceInterventionItem,
  MaintenanceInterventionPayload,
  MaintenanceSolutionTemplate,
  ReportExportResult,
  RoadCatalogItem,
  RoadMeasurementItem,
  RoadMeasurementPayload,
  SapSector,
  SheetDefinition,
  SheetRow,
  SheetRowFilters,
  SheetRowPayload
} from "./types/pad-api";

declare global {
  interface Window {
    padApp?: {
      appName: string;
      appVersion: string;
      data: {
        getStatus: () => Promise<DataStatus>;
        importFromExcel: (excelPath?: string) => Promise<DataStatus>;
        previewExcelImport: (excelPath?: string) => Promise<ImportPreview>;
        pickExcelFile: () => Promise<string | null>;
      };
      audit: {
        integrity: () => Promise<DataIntegrityReport>;
      };
      dashboard: {
        summary: () => Promise<DashboardSummary>;
      };
      backup: {
        export: () => Promise<BackupResult | null>;
        restore: () => Promise<DataStatus | null>;
      };
      sheet: {
        definitions: () => Promise<SheetDefinition[]>;
        list: (sheetName: string, filters?: SheetRowFilters) => Promise<SheetRow[]>;
        create: (sheetName: string, payload: SheetRowPayload) => Promise<SheetRow>;
        update: (sheetName: string, rowId: number, payload: SheetRowPayload) => Promise<SheetRow>;
        delete: (sheetName: string, rowId: number) => Promise<boolean>;
      };
      sap: {
        list: () => Promise<SapSector[]>;
      };
      roads: {
        list: (filters?: { sapCode?: string; search?: string }) => Promise<RoadCatalogItem[]>;
      };
      measurement: {
        listCampaigns: (filters?: {
          roadId?: number;
          search?: string;
          limit?: number;
        }) => Promise<MeasurementCampaignItem[]>;
        listRows: (filters?: {
          campaignKey?: string;
          roadId?: number;
          limit?: number;
        }) => Promise<RoadMeasurementItem[]>;
        upsertCampaign: (payload: MeasurementCampaignPayload) => Promise<MeasurementCampaignItem | null>;
        deleteCampaign: (campaignId: number) => Promise<{ deleted: boolean }>;
        upsertRow: (payload: RoadMeasurementPayload) => Promise<RoadMeasurementItem | null>;
        deleteRow: (measurementId: number) => Promise<{ deleted: boolean }>;
      };
      degradations: {
        list: () => Promise<DegradationItem[]>;
      };
      drainageRules: {
        list: () => Promise<DrainageRule[]>;
        upsert: (payload: {
          id?: number;
          ruleOrder: number;
          matchOperator: "CONTAINS" | "EQUALS" | "REGEX" | "ALWAYS";
          pattern?: string;
          askRequired?: boolean;
          needsAttention?: boolean;
          recommendation: string;
          isActive?: boolean;
        }) => Promise<DrainageRule | null>;
        delete: (ruleId: number) => Promise<{ deleted: boolean }>;
      };
      solutions: {
        listTemplates: () => Promise<MaintenanceSolutionTemplate[]>;
        upsertTemplate: (payload: {
          templateKey: string;
          title: string;
          description: string;
        }) => Promise<MaintenanceSolutionTemplate | null>;
        assignTemplate: (degradationCode: string, templateKey: string) => Promise<{ degradationCode: string; templateKey: string }>;
        setOverride: (degradationCode: string, solutionText: string) => Promise<{ degradationCode: string; solutionText: string }>;
        clearOverride: (degradationCode: string) => Promise<{ degradationCode: string; cleared: boolean }>;
      };
      maintenance: {
        list: (filters?: {
          sapCode?: string;
          roadId?: number;
          status?: string;
          search?: string;
          limit?: number;
        }) => Promise<MaintenanceInterventionItem[]>;
        upsert: (payload: MaintenanceInterventionPayload) => Promise<MaintenanceInterventionItem | null>;
        delete: (interventionId: number) => Promise<{ deleted: boolean }>;
        pickAttachment: () => Promise<AttachmentUploadResult | null>;
        openAttachment: (attachmentPath: string) => Promise<{ opened: boolean }>;
      };
      decision: {
        evaluate: (payload: {
          roadId?: number;
          roadKey?: string;
          roadCode?: string;
          designation?: string;
          degradationId?: number;
          degradationName?: string;
          deflectionValue?: number;
          askDrainage?: boolean;
        }) => Promise<DecisionResult>;
      };
      reporting: {
        listHistory: (filters?: { sapCode?: string; search?: string; limit?: number }) => Promise<DecisionHistoryItem[]>;
        clearHistory: () => Promise<{ deleted: boolean }>;
        exportHistoryXlsx: () => Promise<ReportExportResult | null>;
        exportMaintenanceXlsx: () => Promise<ReportExportResult | null>;
      };
      ping: () => boolean;
    };
  }
}

export {};



