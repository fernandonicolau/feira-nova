export interface BatchWarning {
  code: string;
  message: string;
  entryId?: string;
}

export interface BatchArtifact {
  id: string;
  fileName: string;
  mediaType: string;
  size: number;
  downloadUrl: string;
}

export interface BatchSummary {
  entries: number;
  files: number;
  texts: number;
  items: number;
  artifacts: number;
}

export interface ProcessBatchData {
  batchId: string;
  summary: BatchSummary;
  artifacts: BatchArtifact[];
}

export interface ProcessBatchResponse {
  requestId: string;
  data: ProcessBatchData;
  warnings: BatchWarning[];
}

export interface ApiErrorResponse {
  requestId: string;
  error: {
    code: string;
    message: string;
    details: unknown[];
  };
}
