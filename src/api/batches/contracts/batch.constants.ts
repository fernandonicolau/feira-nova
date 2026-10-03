export const BATCH_LIMITS = {
  maxEntries: 20,
  maxFiles: 10,
  maxFileBytes: 10 * 1024 * 1024,
  maxTotalFileBytes: 50 * 1024 * 1024,
  maxBatchFieldBytes: 2 * 1024 * 1024,
  maxTextCharacters: 50_000,
} as const;

export const SPREADSHEET_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
