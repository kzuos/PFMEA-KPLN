var ActualSyncService = (function() {
  var CONSTANTS = {
    CONFIG_SHEET: 'SYNC_CONFIG',
    LINKS_SHEET: 'SYNC_LINKS',
    WI_REGISTRY_SHEET: 'WI_REGISTRY',
    WI_TEMPLATES_SHEET: 'WI_TEMPLATES',
    PFMEA_VIEW_SHEET: 'PFMEA_SYNC_VIEW',
    LOG_SHEET: 'CHANGE_LOG',
    DEFAULT_PFMEA_SPREADSHEET_ID: '',
    DEFAULT_KPLN_SHEET_NAME: '',
    DEFAULT_WI_TEMPLATE_FOLDER_ID: '',
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
      'PRODUCT_CHARACTERISTIC',
      'PROCESS_CHARACTERISTIC',
      'SPECIFICATION_TOLERANCE',
      'REACTION_PLAN',
      'PFMEA_AP'
    ],
    LINK_STATUS: {
      APPROVED: 'APPROVED',
      SUGGESTED: 'SUGGESTED',
      UNMAPPED: 'UNMAPPED',
      IGNORE: 'IGNORE'
    },
    PFMEA_SELECTION: {
      WARNING_ROW_COUNT: 25,
      ERROR_ROW_COUNT: 100
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
    config = getConfig_();
    var templates = refreshTemplatesFromFolder_(config);
    var validation = validateActualConfiguration_(config);
    var refreshResult = createEmptyRefreshSummary_();
    if (validation.ok) {
      refreshResult = refreshLinks();
    }
    return {
      config: config,
      templates: templates,
      refresh: refreshResult,
      validation: validation
    };
  }

  function refreshLinks() {
    ensureHelperSheets_();
    writeConfigDefaults_();

    var config = getConfig_();
    var validation = assertActualConfigurationReady_(config);
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
      linkRows: suggestedLinks.length,
      validation: validation
    };
  }

  function refreshTemplates() {
    ensureHelperSheets_();
    writeConfigDefaults_();
    return {
      templateRows: refreshTemplatesFromFolder_(getConfig_()).length
    };
  }

  function validateSetup() {
    ensureHelperSheets_();
    writeConfigDefaults_();

    var config = getConfig_();
    var validation = validateActualConfiguration_(config);
    var sourceSpreadsheet = null;
    var kplnSheet = null;

    if (validation.ok) {
      sourceSpreadsheet = openPfmeaSpreadsheet_(config);
      kplnSheet = getSheet_(config.KPLN_SHEET_NAME);
    }

    var templates = loadTemplateRows_();
    var links = loadAllLinkRows_();
    var linkSummary = buildLinkValidationSummary_(links, config, sourceSpreadsheet, kplnSheet, templates);
    var templateSummary = buildTemplateValidationSummary_(templates);

    return {
      ok: validation.errors.concat(linkSummary.errors).length === 0,
      errors: validation.errors.concat(linkSummary.errors),
      warnings: validation.warnings.concat(linkSummary.warnings, templateSummary.warnings),
      linkSummary: linkSummary,
      templateSummary: templateSummary
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
    assertActualConfigurationReady_(config);
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
    var selectionInspection = inspectPfmeaSelection_(link, pfmeaRows);
    if (selectionInspection.error) {
      result.errors += 1;
      result.logEntries.push(createActualLogEntry_(mode, link, 'PFMEA_SELECTION_TOO_BROAD', APP_CONSTANTS.STATUS.ERROR, '', {
        rowCount: selectionInspection.rowCount,
        distinctSteps: selectionInspection.distinctSteps,
        distinctProcesses: selectionInspection.distinctProcesses,
        pfmeaStepFilter: SyncUtils.asString(link.PFMEA_STEP_FILTER)
      }, selectionInspection.error));
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
      var resolvedDocName = docResolution.documentName || SyncUtils.asString(link.WI_TITLE) || payload.stepTitle || link.KPLN_STEP_TITLE;
      if (!dryRun) {
        if (docResolution.docId !== docId || !link.WI_DOC_ID) {
          updateLinkDocumentId_(link.LINK_KEY, docResolution.docId, resolvedDocName);
        }
        upsertRegistryEntry_(link, docResolution.docId, resolvedDocName);
        link.WI_DOC_ID = docResolution.docId;
      } else if (docResolution.docId.indexOf('PREVIEW-') !== 0) {
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
      MANAGED_SECTION_TITLE: getManagedSectionTitle_(templateResolution.templateKey || link.WI_TEMPLATE_KEY),
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
    var sectionSyncOptions = getWorkInstructionSectionOptions_(templateResolution.templateKey || link.WI_TEMPLATE_KEY);
    var docSync = DocsService.syncStepSection(link.WI_DOC_ID, link.LINK_KEY, payloadForDoc, {
      dryRun: dryRun,
      allowOverwrite: true,
      createMissingSection: true,
      backupBeforeWrite: false,
      backupFolderId: '',
      backupDocIds: {},
      relocateManagedSection: true,
      anchorPatterns: sectionSyncOptions.anchorPatterns,
      anchorMatch: sectionSyncOptions.anchorMatch,
      nextSectionPatterns: sectionSyncOptions.nextSectionPatterns
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
      productCharacteristics: joinUniqueField_(pfmeaRows, 'PRODUCT_CHARACTERISTIC', ', '),
      processCharacteristics: joinUniqueField_(pfmeaRows, 'PROCESS_CHARACTERISTIC', ', ') || chooseDominantValue_(pfmeaRows, 'WORK_ELEMENT_4M'),
      specialCharacteristics: aggregateUniqueField_(pfmeaRows, 'SPECIAL_CHARACTERISTIC').join(', '),
      specificationTolerance: joinUniqueField_(pfmeaRows, 'SPECIFICATION_TOLERANCE', ', '),
      controlMethod: buildControlMethodSummary_(pfmeaRows),
      reactionPlan: joinUniqueField_(pfmeaRows, 'REACTION_PLAN', '; '),
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

  function joinUniqueField_(rows, fieldName, separator) {
    return aggregateUniqueField_(rows, fieldName).join(separator || ', ');
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
    var filterTerms = splitFilterTerms_(link.PFMEA_STEP_FILTER);
    if (!filterTerms.length) {
      return rows;
    }
    return rows.filter(function(row) {
      var haystack = normalizeText_([
        row.PROCESS_ITEM,
        row.PROCESS_STEP,
        row.FAILURE_MODE,
        row.FAILURE_CAUSE
      ].join(' '));
      return filterTerms.some(function(filterText) {
        return haystack.indexOf(filterText) > -1;
      });
    });
  }

  function debugLinkSelection(linkKey, sampleLimit) {
    ensureHelperSheets_();
    writeConfigDefaults_();

    var config = getConfig_();
    assertActualConfigurationReady_(config);

    var links = loadAllLinkRows_();
    var link = null;
    for (var index = 0; index < links.length; index += 1) {
      if (SyncUtils.asString(links[index].LINK_KEY) === SyncUtils.asString(linkKey)) {
        link = links[index];
        break;
      }
    }
    if (!link) {
      throw new Error('Link not found: ' + linkKey);
    }

    var pfmeaSpreadsheet = openPfmeaSpreadsheet_(config);
    var pfmeaRows = getPfmeaRowsForLink_(link, pfmeaSpreadsheet, {});
    var selectionInspection = inspectPfmeaSelection_(link, pfmeaRows);
    var limit = Math.max(1, Math.min(toNumber_(sampleLimit) || 10, 25));

    return {
      linkKey: link.LINK_KEY,
      pfmeaSheetName: link.PFMEA_SHEET_NAME,
      pfmeaStepFilter: link.PFMEA_STEP_FILTER,
      filterTerms: splitFilterTerms_(link.PFMEA_STEP_FILTER),
      rowCount: pfmeaRows.length,
      selectionInspection: selectionInspection,
      distinctIssueNos: aggregateUniqueField_(pfmeaRows, 'ISSUE_NO').length,
      distinctProcessItems: aggregateUniqueField_(pfmeaRows, 'PROCESS_ITEM'),
      distinctProcessSteps: aggregateUniqueField_(pfmeaRows, 'PROCESS_STEP'),
      sampleRows: pfmeaRows.slice(0, limit).map(function(row) {
        return {
          SOURCE_ROW: row.SOURCE_ROW,
          ISSUE_NO: row.ISSUE_NO,
          PROCESS_ITEM: row.PROCESS_ITEM,
          PROCESS_STEP: row.PROCESS_STEP,
          FAILURE_MODE: row.FAILURE_MODE,
          FAILURE_CAUSE: row.FAILURE_CAUSE
        };
      })
    };
  }

  function debugPfmeaSheet(sheetName, sampleLimit) {
    ensureHelperSheets_();
    writeConfigDefaults_();

    var config = getConfig_();
    assertActualConfigurationReady_(config);

    var pfmeaSpreadsheet = openPfmeaSpreadsheet_(config);
    var sheet = pfmeaSpreadsheet.getSheetByName(SyncUtils.asString(sheetName));
    if (!sheet) {
      throw new Error('PFMEA sheet not found: ' + sheetName);
    }

    var values = sheet.getDataRange().getDisplayValues();
    var issueHeaderRow = -1;
    for (var rowIndex = 0; rowIndex < Math.min(values.length, 12); rowIndex += 1) {
      if (SyncUtils.asString(values[rowIndex][0]) === 'Issue #') {
        issueHeaderRow = rowIndex;
        break;
      }
    }
    var detailHeaders = issueHeaderRow > -1 && issueHeaderRow + 1 < values.length ? values[issueHeaderRow + 1] : [];
    var indexMap = buildPfmeaIndexMap_(detailHeaders);
    var rows = parsePfmeaSheet_(sheet);
    var limit = Math.max(1, Math.min(toNumber_(sampleLimit) || 20, 50));

    return {
      sheetName: sheet.getName(),
      issueHeaderRow: issueHeaderRow + 1,
      mappedHeaders: buildMappedHeaderDebug_(detailHeaders, indexMap),
      rowCount: rows.length,
      distinctIssueNos: aggregateUniqueField_(rows, 'ISSUE_NO').length,
      distinctProcessItems: aggregateUniqueField_(rows, 'PROCESS_ITEM').slice(0, 20),
      distinctProcessSteps: aggregateUniqueField_(rows, 'PROCESS_STEP').slice(0, 20),
      sampleRows: rows.slice(0, limit).map(function(row) {
        return {
          SOURCE_ROW: row.SOURCE_ROW,
          ISSUE_NO: row.ISSUE_NO,
          PROCESS_ITEM: row.PROCESS_ITEM,
          PROCESS_STEP: row.PROCESS_STEP,
          FAILURE_MODE: row.FAILURE_MODE,
          FAILURE_CAUSE: row.FAILURE_CAUSE
        };
      })
    };
  }

  function buildMappedHeaderDebug_(headers, indexMap) {
    var debug = {};
    Object.keys(indexMap).forEach(function(key) {
      var index = indexMap[key];
      debug[key] = {
        index: index,
        header: index > -1 ? SyncUtils.asString(headers[index]) : ''
      };
    });
    return debug;
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
    var indexMap = buildPfmeaIndexMap_(detailHeaders);

    var records = [];
    var carried = {
      PROCESS_ITEM: '',
      PROCESS_STEP: '',
      WORK_ELEMENT_4M: '',
      SPECIAL_CHARACTERISTIC: '',
      PRODUCT_CHARACTERISTIC: '',
      PROCESS_CHARACTERISTIC: '',
      SPECIFICATION_TOLERANCE: '',
      REACTION_PLAN: ''
    };
    for (var dataRowIndex = issueHeaderRow + 2; dataRowIndex < values.length; dataRowIndex += 1) {
      var row = values[dataRowIndex];
      var issueNo = SyncUtils.asString(row[indexMap.ISSUE_NO]);
      var processItem = fillDownValue_(getCellByIndex_(row, indexMap.PROCESS_ITEM), carried, 'PROCESS_ITEM');
      var processStep = fillDownValue_(getCellByIndex_(row, indexMap.PROCESS_STEP), carried, 'PROCESS_STEP');
      var workElement = fillDownValue_(getCellByIndex_(row, indexMap.WORK_ELEMENT_4M), carried, 'WORK_ELEMENT_4M');
      var failureMode = getCellByIndex_(row, indexMap.FAILURE_MODE);
      var prevention = getCellByIndex_(row, indexMap.PREVENTION_CONTROLS);
      var detection = getCellByIndex_(row, indexMap.DETECTION_CONTROLS);
      var specialCharacteristic = fillDownValue_(getCellByIndex_(row, indexMap.SPECIAL_CHARACTERISTIC), carried, 'SPECIAL_CHARACTERISTIC');
      var productCharacteristic = fillDownValue_(getCellByIndex_(row, indexMap.PRODUCT_CHARACTERISTIC), carried, 'PRODUCT_CHARACTERISTIC');
      var processCharacteristic = fillDownValue_(getCellByIndex_(row, indexMap.PROCESS_CHARACTERISTIC), carried, 'PROCESS_CHARACTERISTIC');
      var specificationTolerance = fillDownValue_(getCellByIndex_(row, indexMap.SPECIFICATION_TOLERANCE), carried, 'SPECIFICATION_TOLERANCE');
      var reactionPlan = fillDownValue_(getCellByIndex_(row, indexMap.REACTION_PLAN), carried, 'REACTION_PLAN');

      if (!issueNo && !processItem && !processStep && !failureMode && !prevention && !detection) {
        continue;
      }

      records.push({
        PFMEA_SHEET_NAME: sheet.getName(),
        SOURCE_ROW: dataRowIndex + 1,
        ISSUE_NO: issueNo,
        PROCESS_ITEM: processItem,
        PROCESS_STEP: processStep,
        WORK_ELEMENT_4M: workElement,
        FAILURE_MODE: failureMode,
        FAILURE_EFFECT: getCellByIndex_(row, indexMap.FAILURE_EFFECT),
        FAILURE_CAUSE: getCellByIndex_(row, indexMap.FAILURE_CAUSE),
        PREVENTION_CONTROLS: prevention,
        DETECTION_CONTROLS: detection,
        SPECIAL_CHARACTERISTIC: specialCharacteristic,
        PRODUCT_CHARACTERISTIC: productCharacteristic,
        PROCESS_CHARACTERISTIC: processCharacteristic,
        SPECIFICATION_TOLERANCE: specificationTolerance,
        REACTION_PLAN: reactionPlan,
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

  function findFirstHeaderIndex_(headers, containsTexts) {
    for (var index = 0; index < containsTexts.length; index += 1) {
      var headerIndex = findHeaderIndex_(headers, containsTexts[index]);
      if (headerIndex > -1) {
        return headerIndex;
      }
    }
    return -1;
  }

  function getCellByIndex_(row, index) {
    return index > -1 ? SyncUtils.asString(row[index]) : '';
  }

  function buildPfmeaIndexMap_(detailHeaders) {
    return {
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
      PRODUCT_CHARACTERISTIC: findFirstHeaderIndex_(detailHeaders, ['PRODUCT CHARACTERISTIC', 'PRODUCT CHAR']),
      PROCESS_CHARACTERISTIC: findFirstHeaderIndex_(detailHeaders, ['PROCESS CHARACTERISTIC', 'PROCESS CHAR']),
      SPECIFICATION_TOLERANCE: findFirstHeaderIndex_(detailHeaders, ['SPECIFICATION / TOLERANCE', 'SPECIFICATION', 'TOLERANCE']),
      REACTION_PLAN: findFirstHeaderIndex_(detailHeaders, ['REACTION PLAN', 'ACTION PLAN', 'CONTINGENCY ACTION']),
      PFMEA_AP: findHeaderIndex_(detailHeaders, 'PFMEA AP')
    };
  }

  function fillDownValue_(value, carried, key) {
    var text = SyncUtils.asString(value);
    if (text) {
      carried[key] = text;
      return text;
    }
    return carried[key] || '';
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
        PFMEA_STEP_FILTER: existing.PFMEA_STEP_FILTER || buildSuggestedStepFilter_(block, suggestion),
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

  function buildSuggestedStepFilter_(block, suggestion) {
    var candidates = [];
    collectFilterCandidates_(candidates, block.stepTitle);
    collectFilterCandidates_(candidates, suggestion ? suggestion.processName : '');
    collectFilterCandidates_(candidates, block.majorProcessTitle);
    return SyncUtils.unique(candidates).slice(0, 3).join(' | ');
  }

  function collectFilterCandidates_(candidates, value) {
    SyncUtils.asString(value).split(/(?:\r?\n|\s\/\s|[|;])/).forEach(function(part) {
      var candidate = SyncUtils.asString(part).trim();
      if (normalizeText_(candidate).length >= 4) {
        candidates.push(candidate);
      }
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

  function splitFilterTerms_(value) {
    return SyncUtils.unique(SyncUtils.asString(value).split(/(?:\r?\n|[|;])/).map(function(part) {
      return normalizeText_(part);
    }).filter(function(part) {
      return !!part;
    }));
  }

  function countDistinctFieldValues_(rows, fieldName) {
    var values = {};
    rows.forEach(function(row) {
      var normalized = normalizeText_(row[fieldName]);
      if (normalized) {
        values[normalized] = true;
      }
    });
    return Object.keys(values).length;
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
      PFMEA_SPREADSHEET_ID: 'Source PFMEA spreadsheet ID. Required before refreshing links or running actual sync.',
      KPLN_SHEET_NAME: 'Formatted KPLN sheet name in this spreadsheet. Required before refreshing links or running actual sync.',
      KPLN_DATA_START_ROW: 'First KPLN data row after the header band.',
      ONLY_APPROVED_LINKS: 'TRUE to sync only APPROVED rows from SYNC_LINKS.',
      CREATE_MISSING_WI_DOCS: 'TRUE to create a Google Doc when no WI_DOC_ID exists.',
      WI_FOLDER_ID: 'Drive folder ID used for created Work Instructions.',
      WI_TEMPLATE_DOC_ID: 'Optional Google Doc template for created Work Instructions.',
      WI_TEMPLATE_FOLDER_ID: 'Optional Drive folder ID that stores the company Work Instruction template files (.docx or Google Docs).'
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

  function validateActualConfiguration_(config) {
    var validation = {
      ok: true,
      errors: [],
      warnings: []
    };
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var startRow = toNumber_(config.KPLN_DATA_START_ROW);

    if (!SyncUtils.asString(config.PFMEA_SPREADSHEET_ID)) {
      validation.errors.push('Set PFMEA_SPREADSHEET_ID in SYNC_CONFIG before refreshing links or running actual sync.');
    } else {
      try {
        SpreadsheetApp.openById(config.PFMEA_SPREADSHEET_ID);
      } catch (error) {
        validation.errors.push('PFMEA_SPREADSHEET_ID is not accessible: ' + error.message);
      }
    }

    if (!SyncUtils.asString(config.KPLN_SHEET_NAME)) {
      validation.errors.push('Set KPLN_SHEET_NAME in SYNC_CONFIG before refreshing links or running actual sync.');
    } else if (!spreadsheet.getSheetByName(config.KPLN_SHEET_NAME)) {
      validation.errors.push('KPLN sheet "' + config.KPLN_SHEET_NAME + '" was not found in this spreadsheet.');
    }

    if (startRow < 1) {
      validation.errors.push('KPLN_DATA_START_ROW must be a positive integer.');
    }

    if (!SyncUtils.asString(config.WI_TEMPLATE_FOLDER_ID)) {
      validation.warnings.push('WI_TEMPLATE_FOLDER_ID is blank. Template discovery will rely on manual WI_TEMPLATES rows.');
    }

    validation.ok = validation.errors.length === 0;
    return validation;
  }

  function assertActualConfigurationReady_(config) {
    var validation = validateActualConfiguration_(config);
    if (!validation.ok) {
      throw new Error('Actual sync configuration is incomplete: ' + validation.errors.join(' | '));
    }
    return validation;
  }

  function createEmptyRefreshSummary_() {
    return {
      pfmeaSheets: 0,
      pfmeaRows: 0,
      kplnBlocks: 0,
      linkRows: 0
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
    var map = {};
    loadAllLinkRows_().forEach(function(record) {
      if (record.LINK_KEY) {
        map[record.LINK_KEY] = record;
      }
    });
    return map;
  }

  function loadSyncLinks_(config) {
    return loadAllLinkRows_().filter(function(link) {
      if (!SyncUtils.toBoolean(link.ACTIVE)) {
        return false;
      }
      if (SyncUtils.asString(link.LINK_STATUS) === CONSTANTS.LINK_STATUS.IGNORE) {
        return false;
      }
      return !SyncUtils.toBoolean(config.ONLY_APPROVED_LINKS) || link.LINK_STATUS === CONSTANTS.LINK_STATUS.APPROVED;
    });
  }

  function loadAllLinkRows_() {
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
    }).filter(function(record) {
      return !!record.LINK_KEY;
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

  function buildLinkValidationSummary_(links, config, sourceSpreadsheet, kplnSheet, templates) {
    var summary = {
      total: links.length,
      active: 0,
      approved: 0,
      suggested: 0,
      unmapped: 0,
      ignored: 0,
      inactive: 0,
      approvedMissingPfmeaSheet: 0,
      approvedInvalidKplnRow: 0,
      approvedMissingTemplateConfig: 0,
      approvedMissingWiDocAssignment: 0,
      approvedSelectionWarnings: 0,
      approvedSelectionErrors: 0,
      errors: [],
      warnings: []
    };
    var pfmeaSheetMap = {};
    var pfmeaCache = {};
    var kplnLastRow = kplnSheet ? kplnSheet.getLastRow() : 0;
    var kplnDataStartRow = toNumber_(config.KPLN_DATA_START_ROW);

    if (sourceSpreadsheet) {
      sourceSpreadsheet.getSheets().forEach(function(sheet) {
        pfmeaSheetMap[sheet.getName()] = true;
      });
    }

    links.forEach(function(link) {
      var status = SyncUtils.asString(link.LINK_STATUS) || CONSTANTS.LINK_STATUS.UNMAPPED;
      var isActive = SyncUtils.toBoolean(link.ACTIVE);

      if (!isActive) {
        summary.inactive += 1;
      } else {
        summary.active += 1;
      }

      if (status === CONSTANTS.LINK_STATUS.APPROVED) {
        summary.approved += 1;
      } else if (status === CONSTANTS.LINK_STATUS.SUGGESTED) {
        summary.suggested += 1;
      } else if (status === CONSTANTS.LINK_STATUS.IGNORE) {
        summary.ignored += 1;
      } else {
        summary.unmapped += 1;
      }

      if (!isActive || status !== CONSTANTS.LINK_STATUS.APPROVED) {
        return;
      }

      if (!SyncUtils.asString(link.PFMEA_SHEET_NAME)) {
        summary.approvedMissingPfmeaSheet += 1;
      } else if (sourceSpreadsheet && !pfmeaSheetMap[link.PFMEA_SHEET_NAME]) {
        summary.approvedMissingPfmeaSheet += 1;
      } else if (sourceSpreadsheet) {
        var pfmeaRows = getPfmeaRowsForLink_(link, sourceSpreadsheet, pfmeaCache);
        var selectionInspection = inspectPfmeaSelection_(link, pfmeaRows);
        if (selectionInspection.error) {
          summary.approvedSelectionErrors += 1;
          summary.errors.push(selectionInspection.error);
        } else if (selectionInspection.warning) {
          summary.approvedSelectionWarnings += 1;
          summary.warnings.push(selectionInspection.warning);
        }
      }

      if (kplnSheet) {
        var rowStart = toNumber_(link.KPLN_ROW_START);
        if (rowStart < kplnDataStartRow || rowStart > kplnLastRow) {
          summary.approvedInvalidKplnRow += 1;
        }
      }

      if (SyncUtils.toBoolean(link.UPDATE_WI)) {
        var templateKey = SyncUtils.asString(link.WI_TEMPLATE_KEY) || 'GENERIC_MANAGED';
        var templateRow = templates[templateKey];
        if (templateKey !== 'GENERIC_MANAGED' && (!templateRow || !hasUsableTemplateConfig_(templateRow))) {
          summary.approvedMissingTemplateConfig += 1;
        }
        if (!SyncUtils.asString(link.WI_DOC_ID) && !SyncUtils.toBoolean(config.CREATE_MISSING_WI_DOCS)) {
          summary.approvedMissingWiDocAssignment += 1;
        }
      }
    });

    if (!summary.total) {
      summary.warnings.push('SYNC_LINKS is empty. Run Refresh Link Matrix before previewing or syncing.');
    }
    if (summary.active && !summary.approved) {
      summary.warnings.push('No active APPROVED links are ready for preview or sync yet.');
    }
    if (summary.suggested) {
      summary.warnings.push('SUGGESTED links still need review: ' + summary.suggested);
    }
    if (summary.unmapped) {
      summary.warnings.push('UNMAPPED links still need manual mapping: ' + summary.unmapped);
    }
    if (summary.approvedMissingPfmeaSheet) {
      summary.errors.push('Approved links with missing or invalid PFMEA_SHEET_NAME: ' + summary.approvedMissingPfmeaSheet);
    }
    if (summary.approvedInvalidKplnRow) {
      summary.errors.push('Approved links with invalid KPLN_ROW_START: ' + summary.approvedInvalidKplnRow);
    }
    if (summary.approvedMissingTemplateConfig) {
      summary.warnings.push('Approved WI links missing a usable template configuration: ' + summary.approvedMissingTemplateConfig);
    }
    if (summary.approvedMissingWiDocAssignment) {
      summary.warnings.push('Approved WI links have no WI_DOC_ID and CREATE_MISSING_WI_DOCS is FALSE: ' + summary.approvedMissingWiDocAssignment);
    }

    return summary;
  }

  function inspectPfmeaSelection_(link, pfmeaRows) {
    var filterTerms = splitFilterTerms_(link.PFMEA_STEP_FILTER);
    var distinctSteps = countDistinctFieldValues_(pfmeaRows, 'PROCESS_STEP');
    var distinctProcesses = countDistinctFieldValues_(pfmeaRows, 'PROCESS_ITEM');
    var linkLabel = SyncUtils.asString(link.LINK_KEY) || SyncUtils.asString(link.KPLN_PROCESS_NO) || 'UNKNOWN_LINK';
    var sheetLabel = SyncUtils.asString(link.PFMEA_SHEET_NAME) || 'UNKNOWN_SHEET';
    var result = {
      rowCount: pfmeaRows.length,
      distinctSteps: distinctSteps,
      distinctProcesses: distinctProcesses,
      usesFilter: filterTerms.length > 0,
      warning: '',
      error: ''
    };

    if (!pfmeaRows.length) {
      result.error = 'Approved link ' + linkLabel + ' currently matches 0 PFMEA rows on sheet ' + sheetLabel + '. Review PFMEA_SHEET_NAME and PFMEA_STEP_FILTER before preview or sync.';
      return result;
    }

    if (!result.usesFilter && pfmeaRows.length >= CONSTANTS.PFMEA_SELECTION.ERROR_ROW_COUNT) {
      result.error = 'Approved link ' + linkLabel + ' currently matches ' + pfmeaRows.length + ' PFMEA rows on sheet ' + sheetLabel + '. Narrow PFMEA_STEP_FILTER before preview or sync.';
      return result;
    }

    if (!result.usesFilter && (pfmeaRows.length >= CONSTANTS.PFMEA_SELECTION.WARNING_ROW_COUNT || distinctSteps > 1 || distinctProcesses > 1)) {
      result.warning = 'Approved link ' + linkLabel + ' uses a blank PFMEA_STEP_FILTER and currently matches ' + pfmeaRows.length + ' PFMEA rows on sheet ' + sheetLabel + '. Add a step filter before live sync.';
      return result;
    }

    if (result.usesFilter && pfmeaRows.length >= CONSTANTS.PFMEA_SELECTION.ERROR_ROW_COUNT) {
      result.warning = 'Approved link ' + linkLabel + ' matches ' + pfmeaRows.length + ' PFMEA rows on sheet ' + sheetLabel + ' even with PFMEA_STEP_FILTER set. Review the mapping, but preview can continue.';
      return result;
    }

    if (result.usesFilter && pfmeaRows.length >= CONSTANTS.PFMEA_SELECTION.WARNING_ROW_COUNT && (distinctSteps > 1 || distinctProcesses > 1)) {
      result.warning = 'Approved link ' + linkLabel + ' still matches a broad PFMEA selection (' + pfmeaRows.length + ' rows on sheet ' + sheetLabel + '). Review PFMEA_STEP_FILTER before live sync.';
    }

    return result;
  }

  function buildTemplateValidationSummary_(templates) {
    var templateKeys = Object.keys(templates);
    var summary = {
      total: templateKeys.length,
      active: 0,
      ready: 0,
      missingConfig: 0,
      warnings: []
    };

    templateKeys.forEach(function(templateKey) {
      var templateRow = templates[templateKey];
      if (!SyncUtils.toBoolean(templateRow.ACTIVE)) {
        return;
      }
      summary.active += 1;
      if (templateKey === 'GENERIC_MANAGED' || hasUsableTemplateConfig_(templateRow)) {
        summary.ready += 1;
      } else {
        summary.missingConfig += 1;
      }
    });

    if (!summary.total) {
      summary.warnings.push('WI_TEMPLATES has no rows yet. Run Refresh WI Templates.');
    } else if (summary.missingConfig) {
      summary.warnings.push('Active WI template rows missing a source file or Google template doc: ' + summary.missingConfig);
    }

    return summary;
  }

  function hasUsableTemplateConfig_(templateRow) {
    return !!SyncUtils.asString(templateRow.GOOGLE_TEMPLATE_DOC_ID) || !!SyncUtils.asString(templateRow.SOURCE_FILE_ID);
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
        DISPLAY_NAME: 'Dokumhane Proses Kontrol Talimati',
        MATCH_KEYWORDS: 'dokumhane;proses kontrol;operator control;process control',
        NOTES: 'Use for foundry process control instructions with control-point tables.'
      },
      {
        TEMPLATE_KEY: 'CAST_CONTROL',
        DISPLAY_NAME: 'Dokulmus Kontrol Talimati',
        MATCH_KEYWORDS: 'dokulmus;casting quality;kalite kontrol;quality control',
        NOTES: 'Use for visual or post-casting quality control instructions.'
      },
      {
        TEMPLATE_KEY: 'TRIMMING',
        DISPLAY_NAME: 'Trimleme Talimati',
        MATCH_KEYWORDS: 'trimleme;trimming',
        NOTES: 'Use for trimming operation instructions.'
      },
      {
        TEMPLATE_KEY: 'MILLING',
        DISPLAY_NAME: 'Freze Talimati',
        MATCH_KEYWORDS: 'freze;milling;mekanik islem',
        NOTES: 'Use for milling or machining operation instructions.'
      },
      {
        TEMPLATE_KEY: 'FINAL_CONTROL',
        DISPLAY_NAME: 'Final Kontrol Talimati',
        MATCH_KEYWORDS: 'final kontrol;final control',
        NOTES: 'Use for final inspection instructions.'
      },
      {
        TEMPLATE_KEY: 'PACKAGING',
        DISPLAY_NAME: 'Paketleme Talimati',
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

  function getManagedSectionTitle_(templateKey) {
    switch (templateKey) {
      case 'FINAL_CONTROL':
        return 'PFMEA / KPLN Final Kontrol Ozet';
      case 'PACKAGING':
        return 'PFMEA / KPLN Paketleme Ozet';
      case 'MILLING':
        return 'PFMEA / KPLN Freze Operasyon Ozet';
      case 'TRIMMING':
        return 'PFMEA / KPLN Trimleme Ozet';
      case 'CAST_CONTROL':
        return 'PFMEA / KPLN Dokum Kalite Ozet';
      case 'FOUNDRY_PROCESS_CONTROL':
        return 'PFMEA / KPLN Proses Kontrol Ozet';
      default:
        return 'PFMEA / KPLN Sync Ozet';
    }
  }

  function getWorkInstructionSectionOptions_(templateKey) {
    switch (templateKey) {
      case 'PACKAGING':
        return {
          anchorPatterns: ['UYGULAMA'],
          anchorMatch: 'last',
          nextSectionPatterns: ['ILGILI DOKUMANLAR', 'REVIZYON GECMISI']
        };
      case 'FINAL_CONTROL':
      case 'MILLING':
      case 'TRIMMING':
      case 'CAST_CONTROL':
        return {
          anchorPatterns: ['UYGULAMA'],
          anchorMatch: 'first',
          nextSectionPatterns: ['ILGILI DOKUMANLAR', 'REVIZYON GECMISI']
        };
      case 'FOUNDRY_PROCESS_CONTROL':
        return {
          anchorPatterns: ['OLASI HATALARA KARSI GORSEL KONTROLLER', 'KONTROL NOKTALARI'],
          anchorMatch: 'last',
          nextSectionPatterns: ['REVIZYON GECMISI', 'ACIKLAMA']
        };
      default:
        return {
          anchorPatterns: [],
          anchorMatch: 'first',
          nextSectionPatterns: []
        };
    }
  }

  function containsAny_(value, candidates) {
    return candidates.some(function(candidate) {
      return value.indexOf(candidate) > -1;
    });
  }

  function normalizeRoutingText_(value) {
    var text = SyncUtils.asString(value);
    [
      ['\u00E7', 'c'],
      ['\u00C7', 'C'],
      ['\u011F', 'g'],
      ['\u011E', 'G'],
      ['\u0131', 'i'],
      ['\u0130', 'I'],
      ['\u00F6', 'o'],
      ['\u00D6', 'O'],
      ['\u015F', 's'],
      ['\u015E', 'S'],
      ['\u00FC', 'u'],
      ['\u00DC', 'U']
    ].forEach(function(replacement) {
      text = text.split(replacement[0]).join(replacement[1]);
    });
    return text
      .toUpperCase()
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
    if (!SyncUtils.asString(config.PFMEA_SPREADSHEET_ID)) {
      throw new Error('PFMEA_SPREADSHEET_ID is blank. Set it in SYNC_CONFIG first.');
    }
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
    validateSetup: validateSetup,
    previewSync: previewSync,
    runSync: runSync,
    debugLinkSelection: debugLinkSelection,
    debugPfmeaSheet: debugPfmeaSheet,
    openConfig: openConfig,
    openLinks: openLinks,
    openTemplates: openTemplates
  };
})();
