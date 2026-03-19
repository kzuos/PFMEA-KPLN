var LogService = (function() {
  function createEntry(params) {
    return {
      TIMESTAMP: params.timestamp || SyncUtils.nowIso(),
      USER: params.user || SyncUtils.getUserEmail(),
      MODE: params.mode || 'SYNC',
      SOURCE_SHEET: params.sourceSheet || '',
      SOURCE_ROW: params.sourceRow || '',
      PFMEA_ROW_ID: params.pfmeaRowId || '',
      STEP_ID: params.stepId || '',
      TARGET_TYPE: params.targetType || '',
      TARGET_ID: params.targetId || '',
      ACTION: params.action || '',
      STATUS: params.status || '',
      BEFORE_SUMMARY: SyncUtils.serializeSummary(params.beforeSummary || ''),
      AFTER_SUMMARY: SyncUtils.serializeSummary(params.afterSummary || ''),
      MESSAGE: SyncUtils.truncate(params.message || '', 50000)
    };
  }

  function logEntries(entries, config) {
    if (!entries || !entries.length) {
      return;
    }
    SheetsService.appendRows(config.CHANGE_LOG_SHEET, config.CHANGE_LOG_HEADER_ROW, entries);
  }

  return {
    createEntry: createEntry,
    logEntries: logEntries
  };
})();
