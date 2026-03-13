/// <reference types="vite/client" />

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
        pickExcelFile: () => Promise<string | null>;
      };
      audit: {
        integrity: () => Promise<DataIntegrityReport>;
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
      };
      ping: () => boolean;
    };
  }
}

export {};
