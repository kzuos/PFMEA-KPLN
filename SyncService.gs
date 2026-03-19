var SyncService = (function() {
  function handleSimpleEdit(e) {
    if (!e || !e.range) {
      return;
    }
    try {
      var properties = PropertiesService.getDocumentProperties();
      properties.setProperty('PFMEA_SYNC_LAST_EDIT', JSON.stringify({
        sheetName: e.range.getSheet().getName(),
        rowNumber: e.range.getRow(),
        editedAt: SyncUtils.nowIso()
      }));
    } catch (ignored) {
      // Simple trigger support is intentionally best-effort only.
    }
  }

  function handleEditTrigger(e) {
    if (!e || !e.range) {
      return;
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return;
    }

    try {
      ConfigService.ensureRequiredSheets();
      var config = ConfigService.getConfig();
      if (config.SYNC_MODE !== APP_CONSTANTS.MODES.AUTO) {
        return;
      }

      var sheet = e.range.getSheet();
      if (sheet.getName() !== config.PFMEA_SHEET || e.range.getRow() <= config.PFMEA_HEADER_ROW) {
        return;
      }

      syncPfmeaRowInternal_(e.range.getRow(), buildContext_({
        dryRun: false,
        initiatedBy: 'INSTALLABLE_EDIT'
      }));
    } catch (error) {
      safeLogTriggerError_(error, e);
    } finally {
      lock.releaseLock();
    }
  }

  function runFullSync(options) {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var context = buildContext_(options || {});
      if (!context.dryRun && context.config.BACKUP_BEFORE_WRITE) {
        context.backupSpreadsheetId = SheetsService.backupActiveSpreadsheet(context.config, 'FULL_SYNC');
      }

      var pfmeaRecords = SheetsService.getPfmeaRecords(context.config);
      var summary = createSummary_(context.dryRun);
      pfmeaRecords.forEach(function(record) {
        if (!isProcessRow_(record.values)) {
          return;
        }
        mergeSummary_(summary, syncPfmeaRecord_(record, context));
      });

      if (summary.logEntries.length) {
        LogService.logEntries(summary.logEntries, context.config);
      }
      summary.backupSpreadsheetId = context.backupSpreadsheetId || '';
      return summary;
    } finally {
      lock.releaseLock();
    }
  }

  function syncPfmeaRow(rowNumber, options) {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      return syncPfmeaRowInternal_(rowNumber, buildContext_(options || {}));
    } finally {
      lock.releaseLock();
    }
  }

  function syncPfmeaRowInternal_(rowNumber, context) {
    var result = syncPfmeaRecord_(
      SheetsService.getRecordByRowNumber(context.config.PFMEA_SHEET, context.config.PFMEA_HEADER_ROW, rowNumber),
      context
    );
    if (result.logEntries.length) {
      LogService.logEntries(result.logEntries, context.config);
    }
    return result;
  }

  function syncPfmeaRecord_(record, context) {
    var summary = createSummary_(context.dryRun);
    if (!isProcessRow_(record.values)) {
      return summary;
    }

    var ensuredValues = SheetsService.ensurePfmeaIdentifiers(record.rowNumber, context.config, context.dryRun);
    record.values = SyncUtils.mergeObjects(record.values, ensuredValues);

    var pfmeaRowId = SyncUtils.asString(record.values.PFMEA_ROW_ID);
    if (!pfmeaRowId) {
      throw new Error('PFMEA_ROW_ID could not be resolved for row ' + record.rowNumber);
    }

    var currentSnapshot = buildSnapshot_(record.values);
    var currentHash = SyncUtils.stableHashObject(currentSnapshot);
    var previousState = loadRowState_(pfmeaRowId);
    if (previousState && previousState.hash === currentHash) {
      summary.skipped += 1;
      summary.logEntries.push(
        LogService.createEntry({
          mode: context.initiatedBy,
          sourceSheet: context.config.PFMEA_SHEET,
          sourceRow: record.rowNumber,
          pfmeaRowId: pfmeaRowId,
          stepId: record.values.STEP_ID,
          targetType: 'SYSTEM',
          targetId: pfmeaRowId,
          action: 'NO_SOURCE_CHANGE',
          status: context.dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED,
          beforeSummary: previousState.snapshot,
          afterSummary: currentSnapshot,
          message: 'PFMEA row hash unchanged; downstream sync skipped.'
        })
      );
      return summary;
    }

    var controlPlanResult = syncControlPlan_(record, context);
    var workInstructionResult = syncWorkInstructions_(record, context);

    mergeSummary_(summary, controlPlanResult);
    mergeSummary_(summary, workInstructionResult);
    summary.processed += 1;

    if (!context.dryRun) {
      saveRowState_(pfmeaRowId, {
        hash: currentHash,
        snapshot: currentSnapshot,
        stepId: record.values.STEP_ID,
        updatedAt: SyncUtils.nowIso()
      });
    }

    summary.logEntries.push(
      LogService.createEntry({
        mode: context.initiatedBy,
        sourceSheet: context.config.PFMEA_SHEET,
        sourceRow: record.rowNumber,
        pfmeaRowId: pfmeaRowId,
        stepId: record.values.STEP_ID,
        targetType: 'SYSTEM',
        targetId: pfmeaRowId,
        action: 'ROW_SYNC',
        status: summary.errors ? APP_CONSTANTS.STATUS.ERROR : (context.dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SUCCESS),
        beforeSummary: previousState ? previousState.snapshot : '',
        afterSummary: currentSnapshot,
        message: 'Control Plan actions: ' + controlPlanResult.changed + ', Work Instruction actions: ' + workInstructionResult.changed
      })
    );
    return summary;
  }

  function syncControlPlan_(record, context) {
    var summary = createSummary_(context.dryRun);
    var pfmeaRowId = record.values.PFMEA_ROW_ID;
    var payload = MappingService.buildControlPlanPayload(record.values, context.mappings);
    var existingRows = SheetsService.findRecordsByField(
      context.config.CONTROL_PLAN_SHEET,
      context.config.CONTROL_PLAN_HEADER_ROW,
      'PFMEA_ROW_ID',
      pfmeaRowId
    );
    var targetDocId = resolveWorkInstructionDocId_(record.values, context.config);

    payload.PFMEA_ROW_ID = pfmeaRowId;
    payload.STEP_ID = record.values.STEP_ID;
    payload.CHARACTERISTIC_ID = record.values.CHARACTERISTIC_ID;
    payload.OPERATION_NO = record.values.OPERATION_NO;
    payload.WORK_INSTRUCTION_DOC_ID = targetDocId;
    payload.WORK_INSTRUCTION_STEP_TAG = record.values.STEP_ID;
    payload.STATUS = payload.STATUS || (isActiveRecord_(record.values) ? APP_CONSTANTS.STATUS.ACTIVE : APP_CONSTANTS.STATUS.FLAGGED_INACTIVE);
    payload.LAST_SYNC_AT = SyncUtils.nowIso();
    payload.LAST_SYNC_BY = SyncUtils.getUserEmail();

    if (!existingRows.length) {
      if (!context.config.CREATE_MISSING_CP_ROWS) {
        summary.skipped += 1;
        summary.logEntries.push(
          buildTargetLogEntry_(context, record, APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, pfmeaRowId, 'MISSING_CONTROL_PLAN_ROW', APP_CONSTANTS.STATUS.SKIPPED, '', payload, 'No linked Control Plan row found and auto-create is disabled.')
        );
        return summary;
      }

      var newRow = SyncUtils.deepClone(payload);
      newRow.CONTROL_PLAN_ROW_ID = SyncUtils.generateId('CPR');
      if (context.dryRun) {
        summary.changed += 1;
        summary.controlPlanWrites += 1;
        summary.logEntries.push(
          buildTargetLogEntry_(context, record, APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, newRow.CONTROL_PLAN_ROW_ID, 'CREATE_CONTROL_PLAN_ROW', APP_CONSTANTS.STATUS.PREVIEW, '', newRow, 'Control Plan row would be created.')
        );
        return summary;
      }

      SheetsService.appendRecord(context.config.CONTROL_PLAN_SHEET, context.config.CONTROL_PLAN_HEADER_ROW, newRow);
      summary.changed += 1;
      summary.controlPlanWrites += 1;
      summary.logEntries.push(
        buildTargetLogEntry_(context, record, APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, newRow.CONTROL_PLAN_ROW_ID, 'CREATE_CONTROL_PLAN_ROW', APP_CONSTANTS.STATUS.SUCCESS, '', newRow, 'Control Plan row created.')
      );
      return summary;
    }

    existingRows.forEach(function(existingRow) {
      if (SyncUtils.asString(existingRow.values.STATUS).toUpperCase() === APP_CONSTANTS.STATUS.LOCKED) {
        summary.skipped += 1;
        summary.logEntries.push(
          buildTargetLogEntry_(context, record, APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, existingRow.values.CONTROL_PLAN_ROW_ID, 'LOCKED_CONTROL_PLAN_ROW', APP_CONSTANTS.STATUS.SKIPPED, existingRow.values, payload, 'Control Plan row status is LOCKED.')
        );
        return;
      }

      var changeSet = computeSheetChangeSet_(existingRow.values, payload, context.config.ALLOW_OVERWRITE);
      if (!changeSet.hasChanges) {
        summary.skipped += 1;
        summary.logEntries.push(
          buildTargetLogEntry_(
            context,
            record,
            APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN,
            existingRow.values.CONTROL_PLAN_ROW_ID,
            changeSet.conflicts.length ? 'OVERWRITE_BLOCKED_FIELDS' : 'NO_CONTROL_PLAN_CHANGE',
            context.dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED,
            existingRow.values,
            payload,
            changeSet.conflicts.length ? buildConflictMessage_(changeSet, 'Control Plan row has blocked field changes.') : 'Control Plan row already aligned.'
          )
        );
        return;
      }

      if (context.dryRun) {
        summary.changed += 1;
        summary.controlPlanWrites += 1;
        summary.logEntries.push(
          buildTargetLogEntry_(context, record, APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, existingRow.values.CONTROL_PLAN_ROW_ID, 'UPDATE_CONTROL_PLAN_ROW', APP_CONSTANTS.STATUS.PREVIEW, changeSet.before, changeSet.after, buildConflictMessage_(changeSet, 'Control Plan row would be updated.'))
        );
        return;
      }

      SheetsService.updateRecordByRow(context.config.CONTROL_PLAN_SHEET, context.config.CONTROL_PLAN_HEADER_ROW, existingRow.rowNumber, changeSet.updates);
      summary.changed += 1;
      summary.controlPlanWrites += 1;
      summary.logEntries.push(
        buildTargetLogEntry_(context, record, APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, existingRow.values.CONTROL_PLAN_ROW_ID, 'UPDATE_CONTROL_PLAN_ROW', APP_CONSTANTS.STATUS.SUCCESS, changeSet.before, changeSet.after, buildConflictMessage_(changeSet, 'Control Plan row updated.'))
      );
    });

    return summary;
  }

  function syncWorkInstructions_(record, context) {
    var summary = createSummary_(context.dryRun);
    var stepId = record.values.STEP_ID;
    var allStepRecords = SheetsService.findRecordsByField(context.config.PFMEA_SHEET, context.config.PFMEA_HEADER_ROW, 'STEP_ID', stepId);
    if (!allStepRecords.length) {
      summary.skipped += 1;
      return summary;
    }

    mergeSummary_(summary, prepareWorkInstructionAssignments_(record, allStepRecords, context));

    var groupedByDoc = groupRecordsByWorkInstruction_(allStepRecords, context.config);
    Object.keys(groupedByDoc).forEach(function(docId) {
      if (docId.indexOf('PREVIEW-') === 0) {
        return;
      }
      var recordsForDoc = groupedByDoc[docId];
      var activeRecords = recordsForDoc.filter(function(stepRecord) {
        return isActiveRecord_(stepRecord.values);
      });
      var sourceRecords = activeRecords.length ? activeRecords : recordsForDoc;
      var payload = MappingService.buildWorkInstructionPayload(
        sourceRecords.map(function(stepRecord) {
          return stepRecord.values;
        }),
        context.mappings
      );

      payload.PFMEA_ROW_IDS = SyncUtils.unique(
        sourceRecords.map(function(stepRecord) {
          return stepRecord.values.PFMEA_ROW_ID;
        })
      ).join(', ');
      payload.SECTION_STATUS = activeRecords.length ? APP_CONSTANTS.STATUS.ACTIVE : APP_CONSTANTS.STATUS.FLAGGED_INACTIVE;
      payload.LAST_SYNC_AT = SyncUtils.nowIso();

      var docResult = DocsService.syncStepSection(docId, stepId, payload, {
        dryRun: context.dryRun,
        allowOverwrite: context.config.ALLOW_OVERWRITE,
        createMissingSection: context.config.CREATE_MISSING_WI_SECTION,
        backupBeforeWrite: context.config.BACKUP_BEFORE_WRITE,
        backupFolderId: context.config.BACKUP_FOLDER_ID,
        backupDocIds: context.backupDocIds
      });

      if (docResult.status === APP_CONSTANTS.STATUS.ERROR) {
        summary.errors += 1;
      } else if (docResult.status === APP_CONSTANTS.STATUS.SUCCESS || docResult.status === APP_CONSTANTS.STATUS.PREVIEW) {
        summary.changed += 1;
        summary.docWrites += 1;
      } else {
        summary.skipped += 1;
      }

      summary.logEntries.push(
        buildTargetLogEntry_(context, record, APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, docId, docResult.action, docResult.status, docResult.beforeSummary || '', docResult.afterSummary || payload, docResult.message)
      );
    });

    return summary;
  }

  function prepareWorkInstructionAssignments_(record, stepRecords, context) {
    var summary = createSummary_(context.dryRun);
    var stepId = stepRecords[0].values.STEP_ID || record.values.STEP_ID;
    var blankRecords = stepRecords.filter(function(stepRecord) {
      return SyncUtils.isBlank(stepRecord.values.WI_DOC_ID);
    });

    if (!blankRecords.length) {
      return summary;
    }

    var accessibleDocIds = SyncUtils.unique(stepRecords.map(function(stepRecord) {
      var docId = SyncUtils.asString(stepRecord.values.WI_DOC_ID);
      return DocsService.isDocumentAccessible(docId) ? docId : '';
    }));

    if (accessibleDocIds.length === 1) {
      applyWorkInstructionDocId_(blankRecords, accessibleDocIds[0], context, summary, record, 'ASSIGN_EXISTING_WI_DOC_ID', 'Blank PFMEA WI_DOC_ID values were linked to the existing step document.');
      return summary;
    }

    if (accessibleDocIds.length > 1) {
      summary.skipped += 1;
      summary.logEntries.push(
        buildTargetLogEntry_(
          context,
          record,
          APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION,
          stepRecords[0].values.STEP_ID,
          'AMBIGUOUS_WI_DOC_ASSIGNMENT',
          APP_CONSTANTS.STATUS.SKIPPED,
          '',
          '',
          'Multiple Work Instruction documents already exist for step ' + stepRecords[0].values.STEP_ID + '. Blank WI_DOC_ID values were not auto-assigned.'
        )
      );
      return summary;
    }

    if (!context.config.CREATE_MISSING_WI_DOCS) {
      return summary;
    }

    var seedRecord = chooseSeedRecordForStep_(stepRecords);
    var creationResult = DocsService.ensureWorkInstructionDocument(stepId, seedRecord.values, {
      dryRun: context.dryRun,
      createMissingDocument: true,
      defaultDocId: context.config.DEFAULT_WI_DOC_ID,
      templateDocId: context.config.WI_TEMPLATE_DOC_ID,
      folderId: context.config.WI_FOLDER_ID
    });

    if (creationResult.status === APP_CONSTANTS.STATUS.ERROR) {
      summary.errors += 1;
      summary.logEntries.push(
        buildTargetLogEntry_(context, record, APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, creationResult.docId || stepId, creationResult.action, creationResult.status, '', '', creationResult.message)
      );
      return summary;
    }

    summary.changed += 1;
    summary.docWrites += 1;
    summary.logEntries.push(
      buildTargetLogEntry_(context, record, APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, creationResult.docId || stepId, creationResult.action, creationResult.status, '', {
        STEP_ID: stepId,
        DOCUMENT_NAME: creationResult.documentName || ''
      }, creationResult.message)
    );

    if ((creationResult.status === APP_CONSTANTS.STATUS.SUCCESS || creationResult.status === APP_CONSTANTS.STATUS.PREVIEW) && creationResult.docId) {
      applyWorkInstructionDocId_(blankRecords, creationResult.docId, context, summary, record, 'ASSIGN_NEW_WI_DOC_ID', 'Created and assigned a new Work Instruction document for the step.');
    }

    return summary;
  }

  function groupRecordsByWorkInstruction_(records, config) {
    var groups = {};
    records.forEach(function(record) {
      var docId = resolveWorkInstructionDocId_(record.values, config);
      if (!groups[docId]) {
        groups[docId] = [];
      }
      groups[docId].push(record);
    });
    return groups;
  }

  function resolveWorkInstructionDocId_(recordValues, config) {
    return SyncUtils.asString(recordValues.WI_DOC_ID) || config.DEFAULT_WI_DOC_ID;
  }

  function applyWorkInstructionDocId_(records, docId, context, summary, sourceRecord, action, message) {
    if (!records.length || !docId) {
      return;
    }

    records.forEach(function(record) {
      record.values.WI_DOC_ID = docId;
      if (!context.dryRun) {
        SheetsService.updateRecordByRow(context.config.PFMEA_SHEET, context.config.PFMEA_HEADER_ROW, record.rowNumber, {
          WI_DOC_ID: docId
        });
      }
    });

    summary.changed += 1;
    summary.logEntries.push(
      buildTargetLogEntry_(
        context,
        sourceRecord,
        APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION,
        docId,
        action,
        context.dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SUCCESS,
        '',
        {
          STEP_ID: sourceRecord.values.STEP_ID,
          WI_DOC_ID: docId,
          PFMEA_ROWS: records.map(function(record) {
            return record.rowNumber;
          }).join(', ')
        },
        message
      )
    );
  }

  function chooseSeedRecordForStep_(records) {
    var activeRecords = records.filter(function(record) {
      return isActiveRecord_(record.values);
    });
    return activeRecords.length ? activeRecords[0] : records[0];
  }

  function computeSheetChangeSet_(existingValues, payload, allowOverwrite) {
    var updates = {};
    var before = {};
    var after = {};
    var conflicts = [];

    Object.keys(payload).forEach(function(field) {
      var currentValue = SyncUtils.asString(existingValues[field]);
      var nextValue = SyncUtils.asString(payload[field]);
      if (currentValue === nextValue) {
        return;
      }

      var overwriteAllowed = allowOverwrite || APP_CONSTANTS.SYSTEM_FIELDS_NO_OVERWRITE.indexOf(field) > -1;
      if (!overwriteAllowed && currentValue !== '') {
        conflicts.push(field);
        return;
      }

      updates[field] = payload[field];
      before[field] = existingValues[field];
      after[field] = payload[field];
    });

    return {
      hasChanges: Object.keys(updates).length > 0,
      updates: updates,
      before: before,
      after: after,
      conflicts: conflicts
    };
  }

  function buildConflictMessage_(changeSet, baseMessage) {
    if (!changeSet.conflicts.length) {
      return baseMessage;
    }
    return baseMessage + ' Skipped fields because ALLOW_OVERWRITE is FALSE: ' + changeSet.conflicts.join(', ');
  }

  function buildSnapshot_(recordValues) {
    var snapshot = {};
    APP_CONSTANTS.SNAPSHOT_FIELDS.forEach(function(field) {
      snapshot[field] = recordValues[field] || '';
    });
    return snapshot;
  }

  function loadRowState_(pfmeaRowId) {
    var rawValue = PropertiesService.getDocumentProperties().getProperty(APP_CONSTANTS.PROPERTY_PREFIXES.ROW_STATE + pfmeaRowId);
    return rawValue ? JSON.parse(rawValue) : null;
  }

  function saveRowState_(pfmeaRowId, state) {
    PropertiesService.getDocumentProperties().setProperty(
      APP_CONSTANTS.PROPERTY_PREFIXES.ROW_STATE + pfmeaRowId,
      JSON.stringify(state)
    );
  }

  function buildTargetLogEntry_(context, record, targetType, targetId, action, status, beforeSummary, afterSummary, message) {
    return LogService.createEntry({
      mode: context.initiatedBy,
      sourceSheet: context.config.PFMEA_SHEET,
      sourceRow: record.rowNumber,
      pfmeaRowId: record.values.PFMEA_ROW_ID,
      stepId: record.values.STEP_ID,
      targetType: targetType,
      targetId: targetId,
      action: action,
      status: status,
      beforeSummary: beforeSummary,
      afterSummary: afterSummary,
      message: message
    });
  }

  function buildContext_(options) {
    ConfigService.ensureRequiredSheets();
    var config = ConfigService.getConfig();
    if (!config.DEFAULT_WI_DOC_ID || !config.WI_FOLDER_ID || !config.BACKUP_FOLDER_ID) {
      config = ConfigService.ensureDriveArtifacts();
    }
    return {
      config: config,
      dryRun: !!options.dryRun || config.DRY_RUN_MODE,
      initiatedBy: options.initiatedBy || 'MENU',
      mappings: MappingService.loadMappings(config, false),
      backupDocIds: {}
    };
  }

  function createSummary_(dryRun) {
    return {
      dryRun: dryRun,
      processed: 0,
      changed: 0,
      skipped: 0,
      errors: 0,
      controlPlanWrites: 0,
      docWrites: 0,
      logEntries: []
    };
  }

  function mergeSummary_(summary, partial) {
    summary.processed += partial.processed || 0;
    summary.changed += partial.changed || 0;
    summary.skipped += partial.skipped || 0;
    summary.errors += partial.errors || 0;
    summary.controlPlanWrites += partial.controlPlanWrites || 0;
    summary.docWrites += partial.docWrites || 0;
    summary.logEntries = summary.logEntries.concat(partial.logEntries || []);
  }

  function isProcessRow_(recordValues) {
    return !SyncUtils.isBlank(recordValues.OPERATION_NO) ||
      !SyncUtils.isBlank(recordValues.PROCESS_STEP) ||
      !SyncUtils.isBlank(recordValues.FAILURE_MODE);
  }

  function isActiveRecord_(recordValues) {
    var activeValue = SyncUtils.asString(recordValues.ACTIVE).toUpperCase();
    return activeValue === '' || activeValue === 'TRUE' || activeValue === 'YES' || activeValue === 'Y' || activeValue === '1';
  }

  function safeLogTriggerError_(error, event) {
    try {
      ConfigService.ensureRequiredSheets();
      var config = ConfigService.getConfig();
      LogService.logEntries([
        LogService.createEntry({
          mode: 'INSTALLABLE_EDIT',
          sourceSheet: event && event.range ? event.range.getSheet().getName() : '',
          sourceRow: event && event.range ? event.range.getRow() : '',
          targetType: 'SYSTEM',
          targetId: 'TRIGGER',
          action: 'TRIGGER_ERROR',
          status: APP_CONSTANTS.STATUS.ERROR,
          message: error.message || String(error)
        })
      ], config);
    } catch (ignored) {
      // Intentionally ignored so the trigger fails gracefully.
    }
  }

  return {
    handleSimpleEdit: handleSimpleEdit,
    handleEditTrigger: handleEditTrigger,
    runFullSync: runFullSync,
    syncPfmeaRow: syncPfmeaRow
  };
})();
