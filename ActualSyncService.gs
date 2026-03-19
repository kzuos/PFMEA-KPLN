var ActualSyncService = (function() {
  var CONSTANTS = {
    CONFIG_SHEET: 'SYNC_CONFIG',
    LINKS_SHEET: 'SYNC_LINKS',
    WI_REGISTRY_SHEET: 'WI_REGISTRY',
    WI_TEMPLATES_SHEET: 'WI_TEMPLATES',
    PFMEA_VIEW_SHEET: 'PFMEA_SYNC_VIEW',
    LOG_SHEET: 'CHANGE_LOG',
    DEFAULT_PFMEA_SPREADSHEET_ID: '1h8Xa_qbSM9r9fCu6OcitXYanY-WWMDlKDqcUzOZHwGQ',
    DEFAULT_KPLN_SHEET_NAME: 'FR.000189',
    DEFAULT_WI_TEMPLATE_FOLDER_ID: '1FUzjKn9EE-CTZZydxfotjREvyr9lZDei',
    CONFIG_HEADERS: ['KEY', 'VALUE', 'DESCRIPTION'],
    LINKS_HEADERS: [
      'ACTIVE',
      'LINK_STATUS',
      'LINK_KEY',
      'PFMEA_SHEET_NAME',
      'PFMEA_PROCESS_NAME',
      'PFMEA_STEP_FILTER',
      'KPLN_PROCESS_NO',
      'KPLN_STEP_TITLE',
      'KPLN_ROW_START',
      'KPLN_ROW_END',
      'UPDATE_STEP_TITLE',
      'UPDATE_CONTROL_METHOD',
      'UPDATE_REACTION_PLAN',
      'UPDATE_WI',
      'WI_TEMPLATE_KEY',
      'WI_DOC_ID',
      'WI_TITLE',
      'NOTES'
    ],
    WI_REGISTRY_HEADERS: [
      'ACTIVE',
      'LINK_KEY',
      'KPLN_PROCESS_NO',
      'WI_TEMPLATE_KEY',
      'DOC_ID',
      'DOC_NAME',
      'LAST_SYNC_AT',
      'NOTES'
    ],
    WI_TEMPLATE_HEADERS: [
      'ACTIVE',
      'TEMPLATE_KEY',
      'DISPLAY_NAME',
      'MATCH_KEYWORDS',
      'SOURCE_FILE_ID',
      'SOURCE_FILE_NAME',
      'SOURCE_MIME_TYPE',
      'GOOGLE_TEMPLATE_DOC_ID',
      'GOOGLE_TEMPLATE_NAME',
      'TARGET_FOLDER_ID',
      'NOTES'
    ],
    PFMEA_VIEW_HEADERS: [
      'PFMEA_SHEET_NAME',
      'SOURCE_ROW',
      'ISSUE_NO',
      'PROCESS_ITEM',
      'PROCESS_STEP',
      'WORK_ELEMENT_4M',
      'FAILURE_MODE',
      'FAILURE_EFFECT',
      'FAILURE_CAUSE',
      'PREVENTION_CONTROLS',
      'DETECTION_CONTROLS',
      'SPECIAL_CHARACTERISTIC',
      'PFMEA_AP'
    ],
    LINK_STATUS: {
      APPROVED: 'APPROVED',
      SUGGESTED: 'SUGGESTED',
      UNMAPPED: 'UNMAPPED',
      IGNORE: 'IGNORE'
    },
    UPDATE_COLUMNS: {
      STEP_TITLE: 2,
      CONTROL_METHOD: 13,
      REACTION_PLAN: 15
    }
  };

  function setup() {
    ensureHelperSheets_();
    writeConfigDefaults_();
    var config = getConfig_();
    ensureWorkInstructionFolder_(config);
    var templates = refreshTemplatesFromFolder_(config);
    var refreshResult = refreshLinks();
    return {
      config: config,
      templates: templates,
      refresh: refreshResult
    };
  }

  function refreshLinks() {
    ensureHelperSheets_();
    writeConfigDefaults_();

    var config = getConfig_();
    refreshTemplatesFromFolder_(config);
    var pfmeaSpreadsheet = openPfmeaSpreadsheet_(config);
    var pfmeaSheets = collectPfmeaSheetSummaries_(pfmeaSpreadsheet);
    var pfmeaViewRows = collectPfmeaViewRows_(pfmeaSpreadsheet, pfmeaSheets);
    var kplnBlocks = scanKplnBlocks_(config);
    var existingLinks = loadExistingLinkRows_();
    var suggestedLinks = buildSuggestedLinks_(kplnBlocks, pfmeaSheets, existingLinks);

    writeSheetRows_(CONSTANTS.LINKS_SHEET, CONSTANTS.LINKS_HEADERS, suggestedLinks);
    writeSheetRows_(CONSTANTS.PFMEA_VIEW_SHEET, CONSTANTS.PFMEA_VIEW_HEADERS, pfmeaViewRows);

    return {
      pfmeaSheets: pfmeaSheets.length,
      pfmeaRows: pfmeaViewRows.length,
      kplnBlocks: kplnBlocks.length,
      linkRows: suggestedLinks.length
    };
  }

  function refreshTemplates() {
    ensureHelperSheets_();
    writeConfigDefaults_();
    return {
      templateRows: refreshTemplatesFromFolder_(getConfig_()).length
    };
  }

  function previewSync() {
    return runSync_(true, 'ACTUAL_PREVIEW');
  }

  function runSync() {
    return runSync_(false, 'ACTUAL_SYNC');
  }

  function openConfig() {
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(getSheet_(CONSTANTS.CONFIG_SHEET));
  }

  function openLinks() {
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(getSheet_(CONSTANTS.LINKS_SHEET));
  }

  function openTemplates() {
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(getSheet_(CONSTANTS.WI_TEMPLATES_SHEET));
  }

  function runSync_(dryRun, mode) {
    ensureHelperSheets_();
    var config = getConfig_();
    var pfmeaSpreadsheet = openPfmeaSpreadsheet_(config);
    var links = loadSyncLinks_(config);
    var summary = {
      dryRun: dryRun,
      processed: 0,
      changed: 0,
      skipped: 0,
      errors: 0,
      wiWrites: 0,
      kplnWrites: 0,
      logEntries: []
    };
    var pfmeaCache = {};

    links.forEach(function(link) {
      var result = syncLink_(link, config, pfmeaSpreadsheet, pfmeaCache, dryRun, mode);
      summary.processed += 1;
      summary.changed += result.changed;
      summary.skipped += result.skipped;
      summary.errors += result.errors;
      summary.wiWrites += result.wiWrites;
      summary.kplnWrites += result.kplnWrites;
      summary.logEntries = summary.logEntries.concat(result.logEntries);
    });

    if (summary.logEntries.length) {
      LogService.logEntries(summary.logEntries, {
        CHANGE_LOG_SHEET: CONSTANTS.LOG_SHEET,
        CHANGE_LOG_HEADER_ROW: 1
      });
    }

    return summary;
  }

  function syncLink_(link, config, pfmeaSpreadsheet, pfmeaCache, dryRun, mode) {
    var result = {
      changed: 0,
      skipped: 0,
      errors: 0,
      wiWrites: 0,
      kplnWrites: 0,
      logEntries: []
    };
    var pfmeaRows = getPfmeaRowsForLink_(link, pfmeaSpreadsheet, pfmeaCache);
    if (!pfmeaRows.length) {
      result.skipped += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, 'NO_PFMEA_MATCH', APP_CONSTANTS.STATUS.SKIPPED, '', '', 'No PFMEA rows matched this approved link.'));
      return result;
    }

    var payload = buildActualPayload_(pfmeaRows, link);
    var kplnSheet = getSheet_(config.KPLN_SHEET_NAME);
    var beforeSnapshot = readKplnSnapshot_(kplnSheet, link);
    var afterSnapshot = SyncUtils.deepClone(beforeSnapshot);

    if (SyncUtils.toBoolean(link.UPDATE_STEP_TITLE) && payload.stepTitle) {
      afterSnapshot.stepTitle = payload.stepTitle;
    }
    if (SyncUtils.toBoolean(link.UPDATE_CONTROL_METHOD) && payload.controlMethod) {
      afterSnapshot.controlMethod = payload.controlMethod;
    }
    if (SyncUtils.toBoolean(link.UPDATE_REACTION_PLAN) && payload.reactionPlan) {
      afterSnapshot.reactionPlan = payload.reactionPlan;
    }

    var kplnChanged = beforeSnapshot.stepTitle !== afterSnapshot.stepTitle ||
      beforeSnapshot.controlMethod !== afterSnapshot.controlMethod ||
      beforeSnapshot.reactionPlan !== afterSnapshot.reactionPlan;

    if (kplnChanged) {
      result.changed += 1;
      result.kplnWrites += 1;
      if (!dryRun) {
        writeKplnSnapshot_(kplnSheet, link, afterSnapshot);
      }
      result.logEntries.push(
        createActualLogEntry_(
          mode,
          link,
          'SYNC_KPLN_BLOCK',
          dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SUCCESS,
          beforeSnapshot,
          afterSnapshot,
          dryRun ? 'KPLN block would be updated.' : 'KPLN block updated.'
        )
      );
    } else {
      result.skipped += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, 'NO_KPLN_CHANGE', dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED, beforeSnapshot, afterSnapshot, 'KPLN block already aligned or all updates are disabled.'));
    }

    if (SyncUtils.toBoolean(link.UPDATE_WI)) {
      var wiResult = syncWorkInstructionForLink_(link, payload, config, dryRun, mode);
      result.changed += wiResult.changed;
      result.skipped += wiResult.skipped;
      result.errors += wiResult.errors;
      result.wiWrites += wiResult.wiWrites;
      result.logEntries = result.logEntries.concat(wiResult.logEntries);
    }

    return result;
  }

  function syncWorkInstructionForLink_(link, payload, config, dryRun, mode) {
    var result = {
      changed: 0,
      skipped: 0,
      errors: 0,
      wiWrites: 0,
      logEntries: []
    };
    var templateResolution = resolveTemplateForLink_(link, config, dryRun, mode);
    result.changed += templateResolution.changed;
    result.skipped += templateResolution.skipped;
    result.errors += templateResolution.errors;
    result.wiWrites += templateResolution.wiWrites;
    result.logEntries = result.logEntries.concat(templateResolution.logEntries);

    if (templateResolution.errors) {
      return result;
    }

    var registryEntry = getRegistryEntryByLinkKey_(link.LINK_KEY);
    var docId = SyncUtils.asString(link.WI_DOC_ID) || (registryEntry ? registryEntry.DOC_ID : '');
    var docResolution = DocsService.ensureWorkInstructionDocument(link.LINK_KEY, {
      WI_DOC_ID: docId,
      STEP_ID: link.LINK_KEY,
      OPERATION_NO: link.KPLN_PROCESS_NO,
      PROCESS_STEP: payload.stepTitle || link.KPLN_STEP_TITLE
    }, {
      dryRun: dryRun,
      createMissingDocument: SyncUtils.toBoolean(config.CREATE_MISSING_WI_DOCS),
      defaultDocId: '',
      templateDocId: templateResolution.templateDocId || config.WI_TEMPLATE_DOC_ID,
      templateSourceFileId: templateResolution.sourceFileId,
      templateName: templateResolution.templateName,
      folderId: config.WI_FOLDER_ID
    });

    if (docResolution.status === APP_CONSTANTS.STATUS.ERROR) {
      result.errors += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, docResolution.action, docResolution.status, '', '', docResolution.message));
      return result;
    }

    if ((docResolution.status === APP_CONSTANTS.STATUS.SUCCESS || docResolution.status === APP_CONSTANTS.STATUS.PREVIEW) && docResolution.docId) {
      if (!dryRun && docResolution.docId !== docId) {
        updateLinkDocumentId_(link.LINK_KEY, docResolution.docId, docResolution.documentName || payload.stepTitle || link.KPLN_STEP_TITLE);
        upsertRegistryEntry_(link, docResolution.docId, docResolution.documentName || payload.stepTitle || link.KPLN_STEP_TITLE);
        link.WI_DOC_ID = docResolution.docId;
      } else if (!dryRun && !link.WI_DOC_ID) {
        updateLinkDocumentId_(link.LINK_KEY, docResolution.docId, docResolution.documentName || payload.stepTitle || link.KPLN_STEP_TITLE);
        link.WI_DOC_ID = docResolution.docId;
      } else if (dryRun && docResolution.docId.indexOf('PREVIEW-') !== 0) {
        link.WI_DOC_ID = docResolution.docId;
      }
      result.changed += 1;
      result.wiWrites += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, docResolution.action, docResolution.status, '', {
        docId: docResolution.docId,
        name: docResolution.documentName || ''
      }, docResolution.message));
    }

    if (dryRun && docResolution.docId && docResolution.docId.indexOf('PREVIEW-') === 0) {
      return result;
    }

    if (!link.WI_DOC_ID) {
      result.skipped += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, 'WI_DOC_UNRESOLVED', dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED, '', '', 'No Work Instruction document is available for this link.'));
      return result;
    }

    var payloadForDoc = {
      STEP_TITLE: payload.stepTitle || link.KPLN_STEP_TITLE,
      OPERATION_NO: link.KPLN_PROCESS_NO,
      PROCESS_DESCRIPTION: payload.processDescription,
      FAILURE_SUMMARY: payload.failureSummary,
      PRODUCT_CHARACTERISTICS: payload.productCharacteristics,
      PROCESS_CHARACTERISTICS: payload.processCharacteristics,
      SPECIAL_CHARACTERISTICS: payload.specialCharacteristics,
      SPECIFICATION_TOLERANCE: payload.specificationTolerance,
      CONTROL_METHOD: payload.controlMethod,
      REACTION_PLAN: payload.reactionPlan,
      PREVENTION_CONTROLS: payload.preventionControls,
      DETECTION_CONTROLS: payload.detectionControls,
      PFMEA_ROW_IDS: payload.issueNos,
      SECTION_STATUS: APP_CONSTANTS.STATUS.ACTIVE,
      LAST_SYNC_AT: SyncUtils.nowIso()
    };
    var docSync = DocsService.syncStepSection(link.WI_DOC_ID, link.LINK_KEY, payloadForDoc, {
      dryRun: dryRun,
      allowOverwrite: true,
      createMissingSection: true,
      backupBeforeWrite: false,
      backupFolderId: '',
      backupDocIds: {}
    });

    if (docSync.status === APP_CONSTANTS.STATUS.ERROR) {
      result.errors += 1;
    } else if (docSync.status === APP_CONSTANTS.STATUS.SUCCESS || docSync.status === APP_CONSTANTS.STATUS.PREVIEW) {
      result.changed += 1;
      result.wiWrites += 1;
    } else {
      result.skipped += 1;
    }

    result.logEntries.push(createActualLogEntry_(mode, link, docSync.action, docSync.status, docSync.beforeSummary || '', docSync.afterSummary || payloadForDoc, docSync.message));
    return result;
  }

  function buildActualPayload_(pfmeaRows, link) {
    return {
      stepTitle: chooseDominantValue_(pfmeaRows, 'PROCESS_STEP') || link.KPLN_STEP_TITLE,
      processDescription: chooseDominantValue_(pfmeaRows, 'PROCESS_ITEM') || link.PFMEA_PROCESS_NAME,
      failureSummary: buildFailureSummary_(pfmeaRows),
      productCharacteristics: chooseDominantValue_(pfmeaRows, 'PROCESS_STEP'),
      processCharacteristics: chooseDominantValue_(pfmeaRows, 'WORK_ELEMENT_4M'),
      specialCharacteristics: aggregateUniqueField_(pfmeaRows, 'SPECIAL_CHARACTERISTIC').join(', '),
      specificationTolerance: '',
      controlMethod: buildControlMethodSummary_(pfmeaRows),
      reactionPlan: '',
      preventionControls: aggregateUniqueField_(pfmeaRows, 'PREVENTION_CONTROLS').join('; '),
      detectionControls: aggregateUniqueField_(pfmeaRows, 'DETECTION_CONTROLS').join('; '),
      issueNos: aggregateUniqueField_(pfmeaRows, 'ISSUE_NO').join(', ')
    };
  }

  function buildFailureSummary_(pfmeaRows) {
    return SyncUtils.unique(pfmeaRows.map(function(row) {
      var parts = [];
      if (!SyncUtils.isBlank(row.FAILURE_MODE)) {
        parts.push('Mode: ' + row.FAILURE_MODE);
      }
      if (!SyncUtils.isBlank(row.FAILURE_EFFECT)) {
        parts.push('Effect: ' + row.FAILURE_EFFECT);
      }
      if (!SyncUtils.isBlank(row.FAILURE_CAUSE)) {
        parts.push('Cause: ' + row.FAILURE_CAUSE);
      }
      return parts.join(' | ');
    }).filter(function(line) {
      return !!line;
    })).join('\n');
  }

  function buildControlMethodSummary_(pfmeaRows) {
    var prevention = aggregateUniqueField_(pfmeaRows, 'PREVENTION_CONTROLS');
    var detection = aggregateUniqueField_(pfmeaRows, 'DETECTION_CONTROLS');
    var parts = [];
    if (prevention.length) {
      parts.push('Prevention: ' + prevention.join('; '));
    }
    if (detection.length) {
      parts.push('Detection: ' + detection.join('; '));
    }
    return parts.join(' | ');
  }

  function aggregateUniqueField_(rows, fieldName) {
    return SyncUtils.unique(rows.map(function(row) {
      return SyncUtils.asString(row[fieldName]);
    }).filter(function(value) {
      return !!value;
    }));
  }

  function chooseDominantValue_(rows, fieldName) {
    var counts = {};
    var bestValue = '';
    var bestCount = 0;
    rows.forEach(function(row) {
      var value = SyncUtils.asString(row[fieldName]);
      if (!value) {
        return;
      }
      counts[value] = (counts[value] || 0) + 1;
      if (counts[value] > bestCount) {
        bestValue = value;
        bestCount = counts[value];
      }
    });
    return bestValue;
  }

  function readKplnSnapshot_(sheet, link) {
    return {
      stepTitle: SyncUtils.asString(sheet.getRange(toNumber_(link.KPLN_ROW_START), CONSTANTS.UPDATE_COLUMNS.STEP_TITLE).getDisplayValue()),
      controlMethod: SyncUtils.asString(sheet.getRange(toNumber_(link.KPLN_ROW_START), CONSTANTS.UPDATE_COLUMNS.CONTROL_METHOD).getDisplayValue()),
      reactionPlan: SyncUtils.asString(sheet.getRange(toNumber_(link.KPLN_ROW_START), CONSTANTS.UPDATE_COLUMNS.REACTION_PLAN).getDisplayValue())
    };
  }

  function writeKplnSnapshot_(sheet, link, snapshot) {
    var row = toNumber_(link.KPLN_ROW_START);
    sheet.getRange(row, CONSTANTS.UPDATE_COLUMNS.STEP_TITLE).setValue(snapshot.stepTitle);
    sheet.getRange(row, CONSTANTS.UPDATE_COLUMNS.CONTROL_METHOD).setValue(snapshot.controlMethod);
    sheet.getRange(row, CONSTANTS.UPDATE_COLUMNS.REACTION_PLAN).setValue(snapshot.reactionPlan);
  }

  function collectPfmeaSheetSummaries_(pfmeaSpreadsheet) {
    return getPfmeaProcessSheets_(pfmeaSpreadsheet).map(function(sheet) {
      var rows = parsePfmeaSheet_(sheet);
      return {
        sheetName: sheet.getName(),
        processName: chooseDominantValue_(rows, 'PROCESS_ITEM') || sheet.getName(),
        stepName: chooseDominantValue_(rows, 'PROCESS_STEP') || '',
        issueCount: rows.length
      };
    }).filter(function(summary) {
      return summary.issueCount > 0;
    });
  }

  function collectPfmeaViewRows_(pfmeaSpreadsheet, summaries) {
    var summaryMap = {};
    summaries.forEach(function(summary) {
      summaryMap[summary.sheetName] = true;
    });
    var output = [];
    getPfmeaProcessSheets_(pfmeaSpreadsheet).forEach(function(sheet) {
      if (!summaryMap[sheet.getName()]) {
        return;
      }
      parsePfmeaSheet_(sheet).forEach(function(row) {
        output.push(CONSTANTS.PFMEA_VIEW_HEADERS.map(function(header) {
          return row[header] || '';
        }));
      });
    });
    return output;
  }

  function getPfmeaRowsForLink_(link, pfmeaSpreadsheet, cache) {
    var sheetName = SyncUtils.asString(link.PFMEA_SHEET_NAME);
    if (!sheetName) {
      return [];
    }
    if (!cache[sheetName]) {
      cache[sheetName] = parsePfmeaSheet_(pfmeaSpreadsheet.getSheetByName(sheetName));
    }
    var rows = cache[sheetName] || [];
    var filterText = normalizeText_(link.PFMEA_STEP_FILTER);
    if (!filterText) {
      return rows;
    }
    return rows.filter(function(row) {
      var haystack = normalizeText_([
        row.PROCESS_ITEM,
        row.PROCESS_STEP,
        row.FAILURE_MODE,
        row.FAILURE_CAUSE
      ].join(' '));
      return haystack.indexOf(filterText) > -1;
    });
  }

  function parsePfmeaSheet_(sheet) {
    if (!sheet) {
      return [];
    }
    var values = sheet.getDataRange().getDisplayValues();
    if (!values.length) {
      return [];
    }
    var issueHeaderRow = -1;
    for (var rowIndex = 0; rowIndex < Math.min(values.length, 12); rowIndex += 1) {
      if (SyncUtils.asString(values[rowIndex][0]) === 'Issue #') {
        issueHeaderRow = rowIndex;
        break;
      }
    }
    if (issueHeaderRow === -1 || issueHeaderRow + 1 >= values.length) {
      return [];
    }

    var detailHeaders = values[issueHeaderRow + 1];
    var indexMap = {
      ISSUE_NO: 0,
      PROCESS_ITEM: findHeaderIndex_(detailHeaders, '1. PROCESS ITEM'),
      PROCESS_STEP: findHeaderIndex_(detailHeaders, '2. PROCESS STEP'),
      WORK_ELEMENT_4M: findHeaderIndex_(detailHeaders, '3. PROCESS WORK ELEMENT'),
      FAILURE_EFFECT: findHeaderIndex_(detailHeaders, '1. FAILURE EFFECTS'),
      FAILURE_MODE: findHeaderIndex_(detailHeaders, '2. FAILURE MODE'),
      FAILURE_CAUSE: findHeaderIndex_(detailHeaders, '3. FAILURE CAUSE'),
      PREVENTION_CONTROLS: findHeaderIndex_(detailHeaders, 'CURRENT PREVENTION CONTROL'),
      DETECTION_CONTROLS: findHeaderIndex_(detailHeaders, 'CURRENT DETECTION CONTROL'),
      SPECIAL_CHARACTERISTIC: findHeaderIndex_(detailHeaders, 'SPECIAL CHARACTERISTIC'),
      PFMEA_AP: findHeaderIndex_(detailHeaders, 'PFMEA AP')
    };

    var records = [];
    for (var dataRowIndex = issueHeaderRow + 2; dataRowIndex < values.length; dataRowIndex += 1) {
      var row = values[dataRowIndex];
      var issueNo = SyncUtils.asString(row[indexMap.ISSUE_NO]);
      var processItem = getCellByIndex_(row, indexMap.PROCESS_ITEM);
      var processStep = getCellByIndex_(row, indexMap.PROCESS_STEP);
      var failureMode = getCellByIndex_(row, indexMap.FAILURE_MODE);
      var prevention = getCellByIndex_(row, indexMap.PREVENTION_CONTROLS);
      var detection = getCellByIndex_(row, indexMap.DETECTION_CONTROLS);

      if (!issueNo && !processItem && !processStep && !failureMode && !prevention && !detection) {
        continue;
      }

      records.push({
        PFMEA_SHEET_NAME: sheet.getName(),
        SOURCE_ROW: dataRowIndex + 1,
        ISSUE_NO: issueNo,
        PROCESS_ITEM: processItem,
        PROCESS_STEP: processStep,
        WORK_ELEMENT_4M: getCellByIndex_(row, indexMap.WORK_ELEMENT_4M),
        FAILURE_MODE: failureMode,
        FAILURE_EFFECT: getCellByIndex_(row, indexMap.FAILURE_EFFECT),
        FAILURE_CAUSE: getCellByIndex_(row, indexMap.FAILURE_CAUSE),
        PREVENTION_CONTROLS: prevention,
        DETECTION_CONTROLS: detection,
        SPECIAL_CHARACTERISTIC: getCellByIndex_(row, indexMap.SPECIAL_CHARACTERISTIC),
        PFMEA_AP: getCellByIndex_(row, indexMap.PFMEA_AP)
      });
    }
    return records;
  }

  function getPfmeaProcessSheets_(pfmeaSpreadsheet) {
    return pfmeaSpreadsheet.getSheets().filter(function(sheet) {
      return /^\d+$/.test(SyncUtils.asString(sheet.getName()));
    });
  }

  function findHeaderIndex_(headers, containsText) {
    var normalizedNeedle = normalizeText_(containsText);
    for (var index = 0; index < headers.length; index += 1) {
      if (normalizeText_(headers[index]).indexOf(normalizedNeedle) > -1) {
        return index;
      }
    }
    return -1;
  }

  function getCellByIndex_(row, index) {
    return index > -1 ? SyncUtils.asString(row[index]) : '';
  }

  function scanKplnBlocks_(config) {
    var sheet = getSheet_(config.KPLN_SHEET_NAME);
    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(toNumber_(config.KPLN_DATA_START_ROW), 1, Math.max(lastRow - toNumber_(config.KPLN_DATA_START_ROW) + 1, 1), 15).getDisplayValues();
    var blocks = [];
    var currentBlock = null;
    var currentMajor = {number: '', title: ''};

    values.forEach(function(row, index) {
      var absoluteRow = toNumber_(config.KPLN_DATA_START_ROW) + index;
      var processNo = SyncUtils.asString(row[0]);
      var stepTitle = SyncUtils.asString(row[1]);

      if (isFooterRow_(processNo, stepTitle)) {
        if (currentBlock) {
          currentBlock.rowEnd = absoluteRow - 1;
          blocks.push(currentBlock);
          currentBlock = null;
        }
        return;
      }

      if (isMajorProcessRow_(processNo, stepTitle)) {
        if (currentBlock) {
          currentBlock.rowEnd = absoluteRow - 1;
          blocks.push(currentBlock);
          currentBlock = null;
        }
        currentMajor = {
          number: processNo,
          title: stepTitle
        };
        return;
      }

      if (isStepProcessRow_(processNo)) {
        if (currentBlock) {
          currentBlock.rowEnd = absoluteRow - 1;
          blocks.push(currentBlock);
        }
        currentBlock = {
          linkKey: 'LINK-' + SyncUtils.sanitizeDriveName(processNo, processNo).replace(/\s+/g, '_'),
          processNo: processNo,
          stepTitle: stepTitle,
          rowStart: absoluteRow,
          rowEnd: absoluteRow,
          majorProcessNo: currentMajor.number,
          majorProcessTitle: currentMajor.title
        };
      } else if (currentBlock) {
        currentBlock.rowEnd = absoluteRow;
      }
    });

    if (currentBlock) {
      blocks.push(currentBlock);
    }
    return blocks;
  }

  function isMajorProcessRow_(processNo, stepTitle) {
    return /^\d+$/.test(processNo) && !!stepTitle;
  }

  function isStepProcessRow_(processNo) {
    return /^\d+\.\d+$/.test(processNo);
  }

  function isFooterRow_(processNo, stepTitle) {
    var joined = normalizeText_([processNo, stepTitle].join(' '));
    return joined.indexOf('REVIZYON') > -1 || joined.indexOf('REVISION') > -1 || joined.indexOf('FR 000189') > -1;
  }

  function buildSuggestedLinks_(kplnBlocks, pfmeaSheets, existingLinks) {
    return kplnBlocks.map(function(block) {
      var existing = existingLinks[block.linkKey] || {};
      var suggestion = suggestPfmeaSheet_(block, pfmeaSheets);
      var suggestedTemplateKey = suggestTemplateKeyFromBlock_(block);
      var linkStatus = existing.LINK_STATUS || suggestion.status;
      var linkRow = {
        ACTIVE: existing.ACTIVE || (linkStatus === CONSTANTS.LINK_STATUS.UNMAPPED ? 'FALSE' : 'TRUE'),
        LINK_STATUS: linkStatus,
        LINK_KEY: block.linkKey,
        PFMEA_SHEET_NAME: existing.PFMEA_SHEET_NAME || suggestion.sheetName,
        PFMEA_PROCESS_NAME: existing.PFMEA_PROCESS_NAME || suggestion.processName,
        PFMEA_STEP_FILTER: existing.PFMEA_STEP_FILTER || '',
        KPLN_PROCESS_NO: block.processNo,
        KPLN_STEP_TITLE: block.stepTitle,
        KPLN_ROW_START: String(block.rowStart),
        KPLN_ROW_END: String(block.rowEnd),
        UPDATE_STEP_TITLE: existing.UPDATE_STEP_TITLE || 'FALSE',
        UPDATE_CONTROL_METHOD: existing.UPDATE_CONTROL_METHOD || 'FALSE',
        UPDATE_REACTION_PLAN: existing.UPDATE_REACTION_PLAN || 'FALSE',
        UPDATE_WI: existing.UPDATE_WI || 'TRUE',
        WI_TEMPLATE_KEY: existing.WI_TEMPLATE_KEY || suggestedTemplateKey,
        WI_DOC_ID: existing.WI_DOC_ID || '',
        WI_TITLE: existing.WI_TITLE || ('WI - ' + block.processNo + ' - ' + block.stepTitle),
        NOTES: existing.NOTES || (suggestion.note + ' | Suggested WI template: ' + suggestedTemplateKey)
      };
      return CONSTANTS.LINKS_HEADERS.map(function(header) {
        return linkRow[header] || '';
      });
    });
  }

  function resolveTemplateForLink_(link, config, dryRun, mode) {
    var result = {
      templateKey: SyncUtils.asString(link.WI_TEMPLATE_KEY) || suggestTemplateKeyFromStepTitle_(link.KPLN_STEP_TITLE),
      templateDocId: '',
      templateName: '',
      sourceFileId: '',
      changed: 0,
      skipped: 0,
      errors: 0,
      wiWrites: 0,
      logEntries: []
    };

    if (!result.templateKey || result.templateKey === 'GENERIC_MANAGED') {
      return result;
    }

    var templateRow = getTemplateRowByKey_(result.templateKey);
    if (!templateRow) {
      result.skipped += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, 'WI_TEMPLATE_FALLBACK', dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED, '', '', 'No WI template row is configured for template key ' + result.templateKey + '. Falling back to the generic managed WI.'));
      return result;
    }

    result.templateName = templateRow.DISPLAY_NAME || templateRow.GOOGLE_TEMPLATE_NAME || templateRow.SOURCE_FILE_NAME;
    result.sourceFileId = SyncUtils.asString(templateRow.SOURCE_FILE_ID);

    if (DocsService.isDocumentAccessible(templateRow.GOOGLE_TEMPLATE_DOC_ID)) {
      result.templateDocId = SyncUtils.asString(templateRow.GOOGLE_TEMPLATE_DOC_ID);
      return result;
    }

    if (!result.sourceFileId) {
      result.skipped += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, 'WI_TEMPLATE_FALLBACK', dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED, '', '', 'Template key ' + result.templateKey + ' has no source file yet. Falling back to the generic managed WI.'));
      return result;
    }

    if (dryRun) {
      result.changed += 1;
      result.wiWrites += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, 'IMPORT_WI_TEMPLATE', APP_CONSTANTS.STATUS.PREVIEW, '', {
        templateKey: result.templateKey,
        sourceFileId: result.sourceFileId
      }, 'The source template file would be imported to Google Docs for template key ' + result.templateKey + '.'));
      return result;
    }

    try {
      var templateDoc = DocsService.ensureGoogleDocTemplate(result.sourceFileId, result.templateName, templateRow.TARGET_FOLDER_ID || config.WI_FOLDER_ID);
      result.templateDocId = templateDoc.docId;
      result.templateName = templateDoc.name;
      updateTemplateGoogleDocId_(result.templateKey, templateDoc.docId, templateDoc.name);
      result.changed += 1;
      result.wiWrites += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, 'IMPORT_WI_TEMPLATE', APP_CONSTANTS.STATUS.SUCCESS, '', {
        templateKey: result.templateKey,
        templateDocId: templateDoc.docId,
        templateName: templateDoc.name
      }, 'Imported WI template ' + result.templateKey + ' to Google Docs.'));
    } catch (error) {
      result.errors += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, 'WI_TEMPLATE_IMPORT_ERROR', APP_CONSTANTS.STATUS.ERROR, '', '', 'Unable to import WI template ' + result.templateKey + ': ' + error.message));
    }

    return result;
  }

  function suggestPfmeaSheet_(block, pfmeaSheets) {
    var searchText = normalizeText_([block.majorProcessTitle, block.stepTitle].join(' '));
    var best = {
      score: 0,
      sheetName: '',
      processName: '',
      note: 'No PFMEA suggestion found.',
      status: CONSTANTS.LINK_STATUS.UNMAPPED
    };

    pfmeaSheets.forEach(function(summary) {
      var candidateText = normalizeText_([summary.processName, summary.stepName, summary.sheetName].join(' '));
      var score = scoreCandidate_(searchText, candidateText);
      if (score > best.score) {
        best = {
          score: score,
          sheetName: summary.sheetName,
          processName: summary.processName,
          note: 'Suggested from PFMEA sheet ' + summary.sheetName + ' (' + summary.processName + ') with score ' + score.toFixed(2),
          status: score >= 0.65 ? CONSTANTS.LINK_STATUS.APPROVED : (score >= 0.35 ? CONSTANTS.LINK_STATUS.SUGGESTED : CONSTANTS.LINK_STATUS.UNMAPPED)
        };
      }
    });

    return best;
  }

  function scoreCandidate_(sourceText, candidateText) {
    var sourceTokens = tokenize_(sourceText);
    var candidateTokens = tokenize_(candidateText);
    if (!sourceTokens.length || !candidateTokens.length) {
      return 0;
    }
    var overlap = 0;
    sourceTokens.forEach(function(token) {
      if (candidateTokens.indexOf(token) > -1) {
        overlap += 1;
      }
    });
    return overlap / Math.max(sourceTokens.length, candidateTokens.length);
  }

  function tokenize_(value) {
    return SyncUtils.unique(normalizeText_(value).split(' ').filter(function(token) {
      return token && token.length > 2;
    }));
  }

  function normalizeText_(value) {
    var text = SyncUtils.asString(value).toUpperCase();
    return text
      .replace(/[Ç]/g, 'C')
      .replace(/[Ğ]/g, 'G')
      .replace(/[İI]/g, 'I')
      .replace(/[Ö]/g, 'O')
      .replace(/[Ş]/g, 'S')
      .replace(/[Ü]/g, 'U')
      .replace(/[^A-Z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function ensureHelperSheets_() {
    ensureSheetWithHeaders_(CONSTANTS.CONFIG_SHEET, CONSTANTS.CONFIG_HEADERS);
    ensureSheetWithHeaders_(CONSTANTS.LINKS_SHEET, CONSTANTS.LINKS_HEADERS);
    ensureSheetWithHeaders_(CONSTANTS.WI_REGISTRY_SHEET, CONSTANTS.WI_REGISTRY_HEADERS);
    ensureSheetWithHeaders_(CONSTANTS.WI_TEMPLATES_SHEET, CONSTANTS.WI_TEMPLATE_HEADERS);
    ensureSheetWithHeaders_(CONSTANTS.PFMEA_VIEW_SHEET, CONSTANTS.PFMEA_VIEW_HEADERS);
    ensureSheetWithHeaders_(CONSTANTS.LOG_SHEET, APP_CONSTANTS.HEADERS.CHANGE_LOG);
  }

  function ensureSheetWithHeaders_(sheetName, headers) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getDisplayValues()[0];
    var hasHeaders = headers.every(function(header) {
      return currentHeaders.indexOf(header) > -1;
    });
    if (!hasHeaders) {
      sheet.clear();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }

  function writeConfigDefaults_() {
    var defaults = {
      PFMEA_SPREADSHEET_ID: CONSTANTS.DEFAULT_PFMEA_SPREADSHEET_ID,
      KPLN_SHEET_NAME: CONSTANTS.DEFAULT_KPLN_SHEET_NAME,
      KPLN_DATA_START_ROW: '10',
      ONLY_APPROVED_LINKS: 'TRUE',
      CREATE_MISSING_WI_DOCS: 'TRUE',
      WI_FOLDER_ID: '',
      WI_TEMPLATE_DOC_ID: '',
      WI_TEMPLATE_FOLDER_ID: CONSTANTS.DEFAULT_WI_TEMPLATE_FOLDER_ID
    };
    var descriptions = {
      PFMEA_SPREADSHEET_ID: 'Source PFMEA spreadsheet ID.',
      KPLN_SHEET_NAME: 'Formatted KPLN sheet name in this spreadsheet.',
      KPLN_DATA_START_ROW: 'First KPLN data row after the header band.',
      ONLY_APPROVED_LINKS: 'TRUE to sync only APPROVED rows from SYNC_LINKS.',
      CREATE_MISSING_WI_DOCS: 'TRUE to create a Google Doc when no WI_DOC_ID exists.',
      WI_FOLDER_ID: 'Drive folder ID used for created Work Instructions.',
      WI_TEMPLATE_DOC_ID: 'Optional Google Doc template for created Work Instructions.',
      WI_TEMPLATE_FOLDER_ID: 'Drive folder ID that stores the company Work Instruction template files (.docx or Google Docs).'
    };
    var sheet = getSheet_(CONSTANTS.CONFIG_SHEET);
    var existing = {};
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues().forEach(function(row) {
        existing[SyncUtils.asString(row[0])] = true;
      });
    }
    Object.keys(defaults).forEach(function(key) {
      if (!existing[key]) {
        sheet.appendRow([key, defaults[key], descriptions[key]]);
      }
    });
  }

  function getConfig_() {
    var sheet = getSheet_(CONSTANTS.CONFIG_SHEET);
    var config = {};
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues().forEach(function(row) {
        config[SyncUtils.asString(row[0])] = SyncUtils.asString(row[1]);
      });
    }
    return {
      PFMEA_SPREADSHEET_ID: config.PFMEA_SPREADSHEET_ID || CONSTANTS.DEFAULT_PFMEA_SPREADSHEET_ID,
      KPLN_SHEET_NAME: config.KPLN_SHEET_NAME || CONSTANTS.DEFAULT_KPLN_SHEET_NAME,
      KPLN_DATA_START_ROW: config.KPLN_DATA_START_ROW || '10',
      ONLY_APPROVED_LINKS: config.ONLY_APPROVED_LINKS || 'TRUE',
      CREATE_MISSING_WI_DOCS: config.CREATE_MISSING_WI_DOCS || 'TRUE',
      WI_FOLDER_ID: config.WI_FOLDER_ID || '',
      WI_TEMPLATE_DOC_ID: config.WI_TEMPLATE_DOC_ID || '',
      WI_TEMPLATE_FOLDER_ID: config.WI_TEMPLATE_FOLDER_ID || CONSTANTS.DEFAULT_WI_TEMPLATE_FOLDER_ID
    };
  }

  function ensureWorkInstructionFolder_(config) {
    if (config.WI_FOLDER_ID) {
      try {
        DriveApp.getFolderById(config.WI_FOLDER_ID);
        return;
      } catch (ignored) {
        // Continue and create a new folder.
      }
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var file = DriveApp.getFileById(spreadsheet.getId());
    var parentIterator = file.getParents();
    var parent = parentIterator.hasNext() ? parentIterator.next() : DriveApp.getRootFolder();
    var folderName = spreadsheet.getName() + ' - Work Instructions';
    var folderIterator = parent.getFoldersByName(folderName);
    var folder = folderIterator.hasNext() ? folderIterator.next() : parent.createFolder(folderName);
    updateConfigValue_('WI_FOLDER_ID', folder.getId());
  }

  function updateConfigValue_(key, value) {
    var sheet = getSheet_(CONSTANTS.CONFIG_SHEET);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      for (var index = 0; index < rows.length; index += 1) {
        if (SyncUtils.asString(rows[index][0]) === key) {
          sheet.getRange(index + 2, 2).setValue(value);
          return;
        }
      }
    }
    sheet.appendRow([key, value, '']);
  }

  function loadExistingLinkRows_() {
    var sheet = getSheet_(CONSTANTS.LINKS_SHEET);
    var lastRow = sheet.getLastRow();
    var map = {};
    if (lastRow <= 1) {
      return map;
    }
    var rows = sheet.getRange(2, 1, lastRow - 1, CONSTANTS.LINKS_HEADERS.length).getDisplayValues();
    rows.forEach(function(row) {
      var record = {};
      CONSTANTS.LINKS_HEADERS.forEach(function(header, index) {
        record[header] = SyncUtils.asString(row[index]);
      });
      if (record.LINK_KEY) {
        map[record.LINK_KEY] = record;
      }
    });
    return map;
  }

  function loadSyncLinks_(config) {
    var sheet = getSheet_(CONSTANTS.LINKS_SHEET);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    var rows = sheet.getRange(2, 1, lastRow - 1, CONSTANTS.LINKS_HEADERS.length).getDisplayValues();
    return rows.map(function(row) {
      var record = {};
      CONSTANTS.LINKS_HEADERS.forEach(function(header, index) {
        record[header] = SyncUtils.asString(row[index]);
      });
      return record;
    }).filter(function(link) {
      if (!SyncUtils.toBoolean(link.ACTIVE)) {
        return false;
      }
      if (SyncUtils.asString(link.LINK_STATUS) === CONSTANTS.LINK_STATUS.IGNORE) {
        return false;
      }
      return !SyncUtils.toBoolean(config.ONLY_APPROVED_LINKS) || link.LINK_STATUS === CONSTANTS.LINK_STATUS.APPROVED;
    });
  }

  function getRegistryEntryByLinkKey_(linkKey) {
    var sheet = getSheet_(CONSTANTS.WI_REGISTRY_SHEET);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return null;
    }
    var rows = sheet.getRange(2, 1, lastRow - 1, CONSTANTS.WI_REGISTRY_HEADERS.length).getDisplayValues();
    for (var index = 0; index < rows.length; index += 1) {
      if (SyncUtils.asString(rows[index][1]) === linkKey) {
        return {
          ACTIVE: SyncUtils.asString(rows[index][0]),
          LINK_KEY: SyncUtils.asString(rows[index][1]),
          KPLN_PROCESS_NO: SyncUtils.asString(rows[index][2]),
          WI_TEMPLATE_KEY: SyncUtils.asString(rows[index][3]),
          DOC_ID: SyncUtils.asString(rows[index][4]),
          DOC_NAME: SyncUtils.asString(rows[index][5]),
          LAST_SYNC_AT: SyncUtils.asString(rows[index][6]),
          NOTES: SyncUtils.asString(rows[index][7])
        };
      }
    }
    return null;
  }

  function upsertRegistryEntry_(link, docId, docName) {
    var sheet = getSheet_(CONSTANTS.WI_REGISTRY_SHEET);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var rows = sheet.getRange(2, 1, lastRow - 1, CONSTANTS.WI_REGISTRY_HEADERS.length).getDisplayValues();
      for (var index = 0; index < rows.length; index += 1) {
        if (SyncUtils.asString(rows[index][1]) === link.LINK_KEY) {
          sheet.getRange(index + 2, 1, 1, CONSTANTS.WI_REGISTRY_HEADERS.length).setValues([[
            'TRUE',
            link.LINK_KEY,
            link.KPLN_PROCESS_NO,
            SyncUtils.asString(link.WI_TEMPLATE_KEY),
            docId,
            docName,
            SyncUtils.nowIso(),
            ''
          ]]);
          return;
        }
      }
    }
    sheet.appendRow(['TRUE', link.LINK_KEY, link.KPLN_PROCESS_NO, SyncUtils.asString(link.WI_TEMPLATE_KEY), docId, docName, SyncUtils.nowIso(), '']);
  }

  function updateLinkDocumentId_(linkKey, docId, wiTitle) {
    var sheet = getSheet_(CONSTANTS.LINKS_SHEET);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return;
    }
    var rows = sheet.getRange(2, 1, lastRow - 1, CONSTANTS.LINKS_HEADERS.length).getDisplayValues();
    var docIdColumn = CONSTANTS.LINKS_HEADERS.indexOf('WI_DOC_ID') + 1;
    var wiTitleColumn = CONSTANTS.LINKS_HEADERS.indexOf('WI_TITLE') + 1;
    for (var index = 0; index < rows.length; index += 1) {
      if (SyncUtils.asString(rows[index][2]) === linkKey) {
        sheet.getRange(index + 2, docIdColumn).setValue(docId);
        sheet.getRange(index + 2, wiTitleColumn).setValue(wiTitle);
        return;
      }
    }
  }

  function refreshTemplatesFromFolder_(config) {
    var existing = loadTemplateRows_();
    var discovered = {};
    var folderId = SyncUtils.asString(config.WI_TEMPLATE_FOLDER_ID);

    if (folderId) {
      try {
        var folder = DriveApp.getFolderById(folderId);
        var files = folder.getFiles();
        while (files.hasNext()) {
          var file = files.next();
          var templateKey = classifyTemplateKeyFromName_(file.getName());
          if (!discovered[templateKey]) {
            discovered[templateKey] = {
              SOURCE_FILE_ID: file.getId(),
              SOURCE_FILE_NAME: file.getName(),
              SOURCE_MIME_TYPE: file.getMimeType(),
              TARGET_FOLDER_ID: folderId
            };
          }
        }
      } catch (error) {
        Logger.log('Unable to scan WI template folder: ' + error.message);
      }
    }

    var rows = getTemplateCatalog_().map(function(catalogItem) {
      var current = existing[catalogItem.TEMPLATE_KEY] || {};
      var autoFile = discovered[catalogItem.TEMPLATE_KEY] || {};
      var row = {
        ACTIVE: current.ACTIVE || 'TRUE',
        TEMPLATE_KEY: catalogItem.TEMPLATE_KEY,
        DISPLAY_NAME: current.DISPLAY_NAME || catalogItem.DISPLAY_NAME,
        MATCH_KEYWORDS: current.MATCH_KEYWORDS || catalogItem.MATCH_KEYWORDS,
        SOURCE_FILE_ID: current.SOURCE_FILE_ID || autoFile.SOURCE_FILE_ID || '',
        SOURCE_FILE_NAME: current.SOURCE_FILE_NAME || autoFile.SOURCE_FILE_NAME || '',
        SOURCE_MIME_TYPE: current.SOURCE_MIME_TYPE || autoFile.SOURCE_MIME_TYPE || '',
        GOOGLE_TEMPLATE_DOC_ID: current.GOOGLE_TEMPLATE_DOC_ID || '',
        GOOGLE_TEMPLATE_NAME: current.GOOGLE_TEMPLATE_NAME || '',
        TARGET_FOLDER_ID: current.TARGET_FOLDER_ID || autoFile.TARGET_FOLDER_ID || config.WI_FOLDER_ID || '',
        NOTES: current.NOTES || (autoFile.SOURCE_FILE_ID ? 'Discovered from Drive template folder.' : catalogItem.NOTES)
      };
      return CONSTANTS.WI_TEMPLATE_HEADERS.map(function(header) {
        return row[header] || '';
      });
    });

    writeSheetRows_(CONSTANTS.WI_TEMPLATES_SHEET, CONSTANTS.WI_TEMPLATE_HEADERS, rows);
    return rows;
  }

  function loadTemplateRows_() {
    var sheet = getSheet_(CONSTANTS.WI_TEMPLATES_SHEET);
    var lastRow = sheet.getLastRow();
    var map = {};
    if (lastRow <= 1) {
      return map;
    }
    var rows = sheet.getRange(2, 1, lastRow - 1, CONSTANTS.WI_TEMPLATE_HEADERS.length).getDisplayValues();
    rows.forEach(function(row) {
      var record = {};
      CONSTANTS.WI_TEMPLATE_HEADERS.forEach(function(header, index) {
        record[header] = SyncUtils.asString(row[index]);
      });
      if (record.TEMPLATE_KEY) {
        map[record.TEMPLATE_KEY] = record;
      }
    });
    return map;
  }

  function getTemplateRowByKey_(templateKey) {
    return loadTemplateRows_()[templateKey] || null;
  }

  function updateTemplateGoogleDocId_(templateKey, docId, docName) {
    var sheet = getSheet_(CONSTANTS.WI_TEMPLATES_SHEET);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return;
    }
    var rows = sheet.getRange(2, 1, lastRow - 1, CONSTANTS.WI_TEMPLATE_HEADERS.length).getDisplayValues();
    var docIdColumn = CONSTANTS.WI_TEMPLATE_HEADERS.indexOf('GOOGLE_TEMPLATE_DOC_ID') + 1;
    var docNameColumn = CONSTANTS.WI_TEMPLATE_HEADERS.indexOf('GOOGLE_TEMPLATE_NAME') + 1;
    for (var index = 0; index < rows.length; index += 1) {
      if (SyncUtils.asString(rows[index][1]) === templateKey) {
        sheet.getRange(index + 2, docIdColumn).setValue(docId);
        sheet.getRange(index + 2, docNameColumn).setValue(docName);
        return;
      }
    }
  }

  function getTemplateCatalog_() {
    return [
      {
        TEMPLATE_KEY: 'FOUNDRY_PROCESS_CONTROL',
        DISPLAY_NAME: 'Dökümhane Proses Kontrol Talimatı',
        MATCH_KEYWORDS: 'dökümhane;proses kontrol;operator control;process control',
        NOTES: 'Use for foundry process control instructions with control-point tables.'
      },
      {
        TEMPLATE_KEY: 'CAST_CONTROL',
        DISPLAY_NAME: 'Dökülmüş Kontrol Talimatı',
        MATCH_KEYWORDS: 'dökülmüş;casting quality;kalite kontrol;quality control',
        NOTES: 'Use for visual or post-casting quality control instructions.'
      },
      {
        TEMPLATE_KEY: 'TRIMMING',
        DISPLAY_NAME: 'Trimleme Talimatı',
        MATCH_KEYWORDS: 'trimleme;trimming',
        NOTES: 'Use for trimming operation instructions.'
      },
      {
        TEMPLATE_KEY: 'MILLING',
        DISPLAY_NAME: 'Freze Talimatı',
        MATCH_KEYWORDS: 'freze;milling;mekanik işlem',
        NOTES: 'Use for milling or machining operation instructions.'
      },
      {
        TEMPLATE_KEY: 'FINAL_CONTROL',
        DISPLAY_NAME: 'Final Kontrol Talimatı',
        MATCH_KEYWORDS: 'final kontrol;final control',
        NOTES: 'Use for final inspection instructions.'
      },
      {
        TEMPLATE_KEY: 'PACKAGING',
        DISPLAY_NAME: 'Paketleme Talimatı',
        MATCH_KEYWORDS: 'paketleme;packaging;ambalaj',
        NOTES: 'Use for packing and shipping instructions.'
      },
      {
        TEMPLATE_KEY: 'GENERIC_MANAGED',
        DISPLAY_NAME: 'Generic Managed Appendix',
        MATCH_KEYWORDS: 'generic',
        NOTES: 'Fallback when no company template is assigned.'
      }
    ];
  }

  function classifyTemplateKeyFromName_(fileName) {
    var normalized = normalizeRoutingText_(fileName);
    if (containsAny_(normalized, ['PAKETLEME', 'PACKAGING'])) {
      return 'PACKAGING';
    }
    if (containsAny_(normalized, ['FINAL KONTROL', 'FINAL CONTROL'])) {
      return 'FINAL_CONTROL';
    }
    if (containsAny_(normalized, ['FREZE', 'MILLING'])) {
      return 'MILLING';
    }
    if (containsAny_(normalized, ['TRIMLEME', 'TRIMMING'])) {
      return 'TRIMMING';
    }
    if (containsAny_(normalized, ['DOKULMUS', 'CASTING QUALITY', 'KALITE KONTROL'])) {
      return 'CAST_CONTROL';
    }
    if (containsAny_(normalized, ['DOKUMHANE', 'PROSES KONTROL', 'PROCESS CONTROL'])) {
      return 'FOUNDRY_PROCESS_CONTROL';
    }
    return 'GENERIC_MANAGED';
  }

  function suggestTemplateKeyFromBlock_(block) {
    return suggestTemplateKeyFromStepTitle_([block.majorProcessTitle, block.stepTitle].join(' '));
  }

  function suggestTemplateKeyFromStepTitle_(value) {
    var normalized = normalizeRoutingText_(value);
    if (containsAny_(normalized, ['PAKETLEME', 'PACKAGING', 'AMBALAJ'])) {
      return 'PACKAGING';
    }
    if (containsAny_(normalized, ['FINAL KONTROL', 'FINAL CONTROL'])) {
      return 'FINAL_CONTROL';
    }
    if (containsAny_(normalized, ['FREZE', 'MILLING', 'MEKANIK ISLEM'])) {
      return 'MILLING';
    }
    if (containsAny_(normalized, ['TRIM', 'TRIMLEME'])) {
      return 'TRIMMING';
    }
    if (containsAny_(normalized, ['DOKUM KALITE', 'CASTING QUALITY', 'DOKULMUS'])) {
      return 'CAST_CONTROL';
    }
    if (containsAny_(normalized, ['DOKUM', 'DOKUMHANE', 'ERGITME', 'KUMLAMA', 'EMPRENYELEME', 'SANDBLASTING', 'BLASTING'])) {
      return 'FOUNDRY_PROCESS_CONTROL';
    }
    return 'GENERIC_MANAGED';
  }

  function containsAny_(value, candidates) {
    return candidates.some(function(candidate) {
      return value.indexOf(candidate) > -1;
    });
  }

  function normalizeRoutingText_(value) {
    var text = SyncUtils.asString(value).toUpperCase();
    return text
      .replace(/[Ç]/g, 'C')
      .replace(/[Ğ]/g, 'G')
      .replace(/[İI]/g, 'I')
      .replace(/[Ö]/g, 'O')
      .replace(/[Ş]/g, 'S')
      .replace(/[Ü]/g, 'U')
      .replace(/[^A-Z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function writeSheetRows_(sheetName, headers, rows) {
    var sheet = getSheet_(sheetName);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
  }

  function openPfmeaSpreadsheet_(config) {
    return SpreadsheetApp.openById(config.PFMEA_SPREADSHEET_ID);
  }

  function getSheet_(sheetName) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('Missing sheet: ' + sheetName);
    }
    return sheet;
  }

  function toNumber_(value) {
    return parseInt(SyncUtils.asString(value), 10) || 0;
  }

  function createActualLogEntry_(mode, link, action, status, beforeSummary, afterSummary, message) {
    return LogService.createEntry({
      mode: mode,
      sourceSheet: SyncUtils.asString(link.PFMEA_SHEET_NAME),
      sourceRow: '',
      pfmeaRowId: SyncUtils.asString(link.LINK_KEY),
      stepId: SyncUtils.asString(link.KPLN_PROCESS_NO),
      targetType: 'ACTUAL_SYNC',
      targetId: SyncUtils.asString(link.KPLN_PROCESS_NO),
      action: action,
      status: status,
      beforeSummary: beforeSummary,
      afterSummary: afterSummary,
      message: message
    });
  }

  return {
    setup: setup,
    refreshLinks: refreshLinks,
    refreshTemplates: refreshTemplates,
    previewSync: previewSync,
    runSync: runSync,
    openConfig: openConfig,
    openLinks: openLinks,
    openTemplates: openTemplates
  };
})();
