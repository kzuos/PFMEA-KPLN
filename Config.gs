var APP_CONSTANTS = {
  PROJECT_NAME: 'PFMEA Sync System',
  VERSION: '1.0.0',
  SHEETS: {
    PFMEA: 'PFMEA',
    CONTROL_PLAN: 'CONTROL_PLAN',
    MAPPING: 'MAPPING',
    CHANGE_LOG: 'CHANGE_LOG',
    CONFIG: 'CONFIG'
  },
  TARGET_TYPES: {
    CONTROL_PLAN: 'CONTROL_PLAN',
    WORK_INSTRUCTION: 'WORK_INSTRUCTION'
  },
  MODES: {
    AUTO: 'AUTO',
    MANUAL: 'MANUAL'
  },
  STATUS: {
    SUCCESS: 'SUCCESS',
    SKIPPED: 'SKIPPED',
    ERROR: 'ERROR',
    PREVIEW: 'PREVIEW',
    ACTIVE: 'ACTIVE',
    FLAGGED_INACTIVE: 'FLAGGED_INACTIVE',
    LOCKED: 'LOCKED'
  },
  TRANSFORMS: {
    DIRECT: 'DIRECT',
    CONTROL_METHOD: 'CONTROL_METHOD',
    STATUS_FROM_ACTIVE: 'STATUS_FROM_ACTIVE',
    STEP_TITLE: 'STEP_TITLE',
    AGGREGATE_UNIQUE: 'AGGREGATE_UNIQUE',
    FAILURE_SUMMARY: 'FAILURE_SUMMARY',
    STEP_TAG: 'STEP_TAG'
  },
  PROPERTY_PREFIXES: {
    ROW_STATE: 'PFMEA_SYNC_ROW_STATE_'
  },
  DOC_MARKERS: {
    START_PREFIX: '[[STEP_START:',
    END_PREFIX: '[[STEP_END:',
    LOCK_PREFIX: '[[LOCKED:'
  },
  SYSTEM_FIELDS_NO_OVERWRITE: ['LAST_SYNC_AT', 'LAST_SYNC_BY', 'STATUS'],
  WI_SUPPORTED_FIELDS: [
    'STEP_TITLE',
    'OPERATION_NO',
    'PROCESS_DESCRIPTION',
    'FAILURE_SUMMARY',
    'PRODUCT_CHARACTERISTICS',
    'PROCESS_CHARACTERISTICS',
    'SPECIAL_CHARACTERISTICS',
    'SPECIFICATION_TOLERANCE',
    'CONTROL_METHOD',
    'REACTION_PLAN',
    'PREVENTION_CONTROLS',
    'DETECTION_CONTROLS',
    'PFMEA_ROW_IDS',
    'SECTION_STATUS',
    'LAST_SYNC_AT'
  ],
  SNAPSHOT_FIELDS: [
    'PFMEA_ROW_ID',
    'STEP_ID',
    'CHARACTERISTIC_ID',
    'OPERATION_NO',
    'PROCESS_STEP',
    'PROCESS_FUNCTION_REQUIREMENT',
    'FAILURE_MODE',
    'FAILURE_EFFECT',
    'CAUSE_MECHANISM',
    'PREVENTION_CONTROLS',
    'DETECTION_CONTROLS',
    'SPECIAL_CHARACTERISTICS',
    'PRODUCT_CHARACTERISTICS',
    'PROCESS_CHARACTERISTICS',
    'SPECIFICATION_TOLERANCE',
    'EVALUATION_MEASUREMENT_TECHNIQUE',
    'SAMPLE_SIZE',
    'SAMPLING_FREQUENCY',
    'CONTROL_METHOD_OVERRIDE',
    'REACTION_PLAN',
    'WI_DOC_ID',
    'ACTIVE'
  ],
  HEADERS: {
    PFMEA: [
      'PFMEA_ROW_ID',
      'STEP_ID',
      'CHARACTERISTIC_ID',
      'OPERATION_NO',
      'PROCESS_STEP',
      'PROCESS_FUNCTION_REQUIREMENT',
      'FAILURE_MODE',
      'FAILURE_EFFECT',
      'CAUSE_MECHANISM',
      'PREVENTION_CONTROLS',
      'DETECTION_CONTROLS',
      'SPECIAL_CHARACTERISTICS',
      'PRODUCT_CHARACTERISTICS',
      'PROCESS_CHARACTERISTICS',
      'SPECIFICATION_TOLERANCE',
      'EVALUATION_MEASUREMENT_TECHNIQUE',
      'SAMPLE_SIZE',
      'SAMPLING_FREQUENCY',
      'CONTROL_METHOD_OVERRIDE',
      'REACTION_PLAN',
      'WI_DOC_ID',
      'ACTIVE',
      'OWNER',
      'LAST_REVIEW_DATE',
      'NOTES'
    ],
    CONTROL_PLAN: [
      'CONTROL_PLAN_ROW_ID',
      'PFMEA_ROW_ID',
      'STEP_ID',
      'CHARACTERISTIC_ID',
      'OPERATION_NO',
      'PROCESS_STEP',
      'PRODUCT_CHARACTERISTICS',
      'PROCESS_CHARACTERISTICS',
      'SPECIAL_CHARACTERISTICS',
      'SPECIFICATION_TOLERANCE',
      'EVALUATION_TECHNIQUE',
      'SAMPLE_SIZE',
      'SAMPLING_FREQUENCY',
      'CONTROL_METHOD',
      'REACTION_PLAN',
      'WORK_INSTRUCTION_DOC_ID',
      'WORK_INSTRUCTION_STEP_TAG',
      'STATUS',
      'LAST_SYNC_AT',
      'LAST_SYNC_BY',
      'NOTES'
    ],
    MAPPING: [
      'ACTIVE',
      'TARGET_TYPE',
      'SOURCE_COLUMNS',
      'TARGET_FIELD',
      'TRANSFORM',
      'ON_MISSING',
      'NOTES'
    ],
    CHANGE_LOG: [
      'TIMESTAMP',
      'USER',
      'MODE',
      'SOURCE_SHEET',
      'SOURCE_ROW',
      'PFMEA_ROW_ID',
      'STEP_ID',
      'TARGET_TYPE',
      'TARGET_ID',
      'ACTION',
      'STATUS',
      'BEFORE_SUMMARY',
      'AFTER_SUMMARY',
      'MESSAGE'
    ],
    CONFIG: ['KEY', 'VALUE', 'DESCRIPTION']
  }
};

var SyncUtils = (function() {
  function asString(value) {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  }

  function isBlank(value) {
    return asString(value) === '';
  }

  function toBoolean(value) {
    var normalized = asString(value).toUpperCase();
    return normalized === 'TRUE' || normalized === 'YES' || normalized === 'Y' || normalized === '1';
  }

  function toNumber(value, fallback) {
    var parsed = parseInt(asString(value), 10);
    return isNaN(parsed) ? fallback : parsed;
  }

  function deepClone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function nowIso() {
    return new Date().toISOString();
  }

  function generateId(prefix) {
    return prefix + '-' + Utilities.getUuid().split('-')[0].toUpperCase();
  }

  function normalizeList(values) {
    var raw = Array.isArray(values) ? values : [values];
    var output = [];
    raw.forEach(function(item) {
      asString(item)
        .split(/\r?\n|;|,/)
        .map(function(part) {
          return part.trim();
        })
        .forEach(function(part) {
          if (part) {
            output.push(part);
          }
        });
    });
    return unique(output);
  }

  function unique(values) {
    var map = {};
    var output = [];
    values.forEach(function(item) {
      var key = asString(item);
      if (key && !map[key]) {
        map[key] = true;
        output.push(key);
      }
    });
    return output;
  }

  function truncate(value, maxLength) {
    var text = asString(value);
    if (text.length <= maxLength) {
      return text;
    }
    return text.substring(0, maxLength - 3) + '...';
  }

  function mergeObjects(base, extra) {
    var result = {};
    Object.keys(base || {}).forEach(function(key) {
      result[key] = base[key];
    });
    Object.keys(extra || {}).forEach(function(key) {
      result[key] = extra[key];
    });
    return result;
  }

  function stableHashObject(objectValue) {
    var raw = JSON.stringify(objectValue);
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
    var encoded = Utilities.base64EncodeWebSafe(digest);
    return encoded.substring(0, 20);
  }

  function formatStepId(operationNo) {
    var normalized = asString(operationNo).toUpperCase().replace(/[^A-Z0-9]+/g, '_');
    if (!normalized) {
      normalized = 'UNSPECIFIED';
    }
    return 'STEP-' + normalized;
  }

  function getUserEmail() {
    var email = '';
    try {
      email = Session.getActiveUser().getEmail();
    } catch (error) {
      email = '';
    }
    if (!email) {
      try {
        email = Session.getEffectiveUser().getEmail();
      } catch (ignored) {
        email = '';
      }
    }
    return email || 'unknown@local';
  }

  function serializeSummary(value) {
    if (typeof value === 'string') {
      return truncate(value, 50000);
    }
    return truncate(JSON.stringify(value), 50000);
  }

  return {
    asString: asString,
    isBlank: isBlank,
    toBoolean: toBoolean,
    toNumber: toNumber,
    deepClone: deepClone,
    nowIso: nowIso,
    generateId: generateId,
    normalizeList: normalizeList,
    unique: unique,
    truncate: truncate,
    mergeObjects: mergeObjects,
    stableHashObject: stableHashObject,
    formatStepId: formatStepId,
    getUserEmail: getUserEmail,
    serializeSummary: serializeSummary
  };
})();

var ConfigService = (function() {
  function getDefaultConfig() {
    return {
      PFMEA_SHEET: APP_CONSTANTS.SHEETS.PFMEA,
      CONTROL_PLAN_SHEET: APP_CONSTANTS.SHEETS.CONTROL_PLAN,
      MAPPING_SHEET: APP_CONSTANTS.SHEETS.MAPPING,
      CHANGE_LOG_SHEET: APP_CONSTANTS.SHEETS.CHANGE_LOG,
      CONFIG_SHEET: APP_CONSTANTS.SHEETS.CONFIG,
      PFMEA_HEADER_ROW: '1',
      CONTROL_PLAN_HEADER_ROW: '1',
      MAPPING_HEADER_ROW: '1',
      CHANGE_LOG_HEADER_ROW: '1',
      SYNC_MODE: APP_CONSTANTS.MODES.AUTO,
      DRY_RUN_MODE: 'FALSE',
      ALLOW_OVERWRITE: 'TRUE',
      CONFIRM_BULK_SYNC: 'TRUE',
      BACKUP_BEFORE_WRITE: 'TRUE',
      CREATE_MISSING_CP_ROWS: 'TRUE',
      CREATE_MISSING_WI_SECTION: 'TRUE',
      AUTO_CREATE_IDS: 'TRUE',
      DEFAULT_WI_DOC_ID: '',
      WI_FOLDER_ID: '',
      BACKUP_FOLDER_ID: ''
    };
  }

  function getConfig() {
    var defaults = getDefaultConfig();
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = spreadsheet.getSheetByName(defaults.CONFIG_SHEET);
    var rawConfig = SyncUtils.deepClone(defaults);

    if (configSheet && configSheet.getLastRow() > 1) {
      var values = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 3).getDisplayValues();
      values.forEach(function(row) {
        var key = SyncUtils.asString(row[0]);
        if (key) {
          rawConfig[key] = row[1];
        }
      });
    }

    return coerceConfig_(rawConfig);
  }

  function coerceConfig_(rawConfig) {
    return {
      PFMEA_SHEET: rawConfig.PFMEA_SHEET,
      CONTROL_PLAN_SHEET: rawConfig.CONTROL_PLAN_SHEET,
      MAPPING_SHEET: rawConfig.MAPPING_SHEET,
      CHANGE_LOG_SHEET: rawConfig.CHANGE_LOG_SHEET,
      CONFIG_SHEET: rawConfig.CONFIG_SHEET,
      PFMEA_HEADER_ROW: SyncUtils.toNumber(rawConfig.PFMEA_HEADER_ROW, 1),
      CONTROL_PLAN_HEADER_ROW: SyncUtils.toNumber(rawConfig.CONTROL_PLAN_HEADER_ROW, 1),
      MAPPING_HEADER_ROW: SyncUtils.toNumber(rawConfig.MAPPING_HEADER_ROW, 1),
      CHANGE_LOG_HEADER_ROW: SyncUtils.toNumber(rawConfig.CHANGE_LOG_HEADER_ROW, 1),
      SYNC_MODE: SyncUtils.asString(rawConfig.SYNC_MODE) || APP_CONSTANTS.MODES.AUTO,
      DRY_RUN_MODE: SyncUtils.toBoolean(rawConfig.DRY_RUN_MODE),
      ALLOW_OVERWRITE: SyncUtils.toBoolean(rawConfig.ALLOW_OVERWRITE),
      CONFIRM_BULK_SYNC: SyncUtils.toBoolean(rawConfig.CONFIRM_BULK_SYNC),
      BACKUP_BEFORE_WRITE: SyncUtils.toBoolean(rawConfig.BACKUP_BEFORE_WRITE),
      CREATE_MISSING_CP_ROWS: SyncUtils.toBoolean(rawConfig.CREATE_MISSING_CP_ROWS),
      CREATE_MISSING_WI_SECTION: SyncUtils.toBoolean(rawConfig.CREATE_MISSING_WI_SECTION),
      AUTO_CREATE_IDS: SyncUtils.toBoolean(rawConfig.AUTO_CREATE_IDS),
      DEFAULT_WI_DOC_ID: SyncUtils.asString(rawConfig.DEFAULT_WI_DOC_ID),
      WI_FOLDER_ID: SyncUtils.asString(rawConfig.WI_FOLDER_ID),
      BACKUP_FOLDER_ID: SyncUtils.asString(rawConfig.BACKUP_FOLDER_ID)
    };
  }

  function initializeSystem() {
    ensureRequiredSheets();
    writeMissingConfigDefaults_();
    seedDefaultMappings_();
    ensureDriveArtifacts();
    installEditTrigger();
    return getConfig();
  }

  function ensureRequiredSheets() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheetWithHeaders_(spreadsheet, APP_CONSTANTS.SHEETS.PFMEA, APP_CONSTANTS.HEADERS.PFMEA, 1);
    ensureSheetWithHeaders_(spreadsheet, APP_CONSTANTS.SHEETS.CONTROL_PLAN, APP_CONSTANTS.HEADERS.CONTROL_PLAN, 1);
    ensureSheetWithHeaders_(spreadsheet, APP_CONSTANTS.SHEETS.MAPPING, APP_CONSTANTS.HEADERS.MAPPING, 1);
    ensureSheetWithHeaders_(spreadsheet, APP_CONSTANTS.SHEETS.CHANGE_LOG, APP_CONSTANTS.HEADERS.CHANGE_LOG, 1);
    ensureSheetWithHeaders_(spreadsheet, APP_CONSTANTS.SHEETS.CONFIG, APP_CONSTANTS.HEADERS.CONFIG, 1);
  }

  function ensureSheetWithHeaders_(spreadsheet, sheetName, requiredHeaders, headerRow) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    var lastColumn = Math.max(sheet.getLastColumn(), requiredHeaders.length);
    if (lastColumn === 0) {
      lastColumn = requiredHeaders.length;
    }
    var headerValues = sheet.getRange(headerRow, 1, 1, lastColumn).getDisplayValues()[0];
    var existingHeaders = {};
    headerValues.forEach(function(header) {
      var key = SyncUtils.asString(header);
      if (key) {
        existingHeaders[key] = true;
      }
    });

    if (sheet.getLastRow() === 0 || SyncUtils.isBlank(headerValues[0])) {
      sheet.getRange(headerRow, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
      sheet.setFrozenRows(headerRow);
      return;
    }

    var missingHeaders = requiredHeaders.filter(function(header) {
      return !existingHeaders[header];
    });

    if (missingHeaders.length) {
      var startColumn = sheet.getLastColumn() + 1;
      sheet.getRange(headerRow, startColumn, 1, missingHeaders.length).setValues([missingHeaders]);
    }
    sheet.setFrozenRows(headerRow);
  }

  function writeMissingConfigDefaults_() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONSTANTS.SHEETS.CONFIG);
    var existing = {};
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues().forEach(function(row) {
        var key = SyncUtils.asString(row[0]);
        if (key) {
          existing[key] = true;
        }
      });
    }

    var rowsToAppend = [];
    Object.keys(getDefaultConfig()).forEach(function(key) {
      if (!existing[key]) {
        rowsToAppend.push([key, getDefaultConfig()[key], buildConfigDescription_(key)]);
      }
    });

    if (rowsToAppend.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 3).setValues(rowsToAppend);
    }
  }

  function buildConfigDescription_(key) {
    var descriptions = {
      PFMEA_SHEET: 'Source PFMEA sheet name.',
      CONTROL_PLAN_SHEET: 'Downstream control plan sheet name.',
      MAPPING_SHEET: 'Mapping rules sheet name.',
      CHANGE_LOG_SHEET: 'Audit log sheet name.',
      CONFIG_SHEET: 'Configuration sheet name.',
      PFMEA_HEADER_ROW: 'Header row number for PFMEA.',
      CONTROL_PLAN_HEADER_ROW: 'Header row number for Control Plan.',
      MAPPING_HEADER_ROW: 'Header row number for Mapping.',
      CHANGE_LOG_HEADER_ROW: 'Header row number for Change Log.',
      SYNC_MODE: 'AUTO for edit-trigger sync, MANUAL for menu-only sync.',
      DRY_RUN_MODE: 'TRUE to preview changes without writing.',
      ALLOW_OVERWRITE: 'TRUE to overwrite mapped downstream fields.',
      CONFIRM_BULK_SYNC: 'TRUE to prompt before full sync.',
      BACKUP_BEFORE_WRITE: 'TRUE to create Drive backups before bulk writes.',
      CREATE_MISSING_CP_ROWS: 'TRUE to append missing Control Plan rows.',
      CREATE_MISSING_WI_SECTION: 'TRUE to append missing Work Instruction sections.',
      AUTO_CREATE_IDS: 'TRUE to auto-generate PFMEA_ROW_ID and STEP_ID when blank.',
      DEFAULT_WI_DOC_ID: 'Fallback Google Doc ID for Work Instructions.',
      WI_FOLDER_ID: 'Folder that stores Work Instruction docs.',
      BACKUP_FOLDER_ID: 'Folder that stores backup copies.'
    };
    return descriptions[key] || '';
  }

  function seedDefaultMappings_() {
    var config = getConfig();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.MAPPING_SHEET);
    if (sheet.getLastRow() > 1) {
      return;
    }
    var headers = APP_CONSTANTS.HEADERS.MAPPING;
    var rows = MappingService.getDefaultMappings().map(function(mapping) {
      return headers.map(function(header) {
        return mapping[header] || '';
      });
    });
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
  }

  function upsertConfigValues(updates) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONSTANTS.SHEETS.CONFIG);
    var data = sheet.getDataRange().getDisplayValues();
    var rowMap = {};
    for (var index = 1; index < data.length; index += 1) {
      var key = SyncUtils.asString(data[index][0]);
      if (key) {
        rowMap[key] = index + 1;
      }
    }

    Object.keys(updates).forEach(function(key) {
      var rowNumber = rowMap[key];
      if (rowNumber) {
        sheet.getRange(rowNumber, 2).setValue(updates[key]);
      } else {
        sheet.appendRow([key, updates[key], buildConfigDescription_(key)]);
      }
    });
  }

  function ensureDriveArtifacts() {
    var config = getConfig();
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var updates = {};
    var projectFolder = getOrCreateProjectFolder_(spreadsheet, config.WI_FOLDER_ID);

    if (!config.WI_FOLDER_ID || config.WI_FOLDER_ID !== projectFolder.getId()) {
      updates.WI_FOLDER_ID = projectFolder.getId();
    }

    var backupFolder = getOrCreateBackupFolder_(projectFolder, config.BACKUP_FOLDER_ID);
    if (!config.BACKUP_FOLDER_ID || config.BACKUP_FOLDER_ID !== backupFolder.getId()) {
      updates.BACKUP_FOLDER_ID = backupFolder.getId();
    }

    var defaultDocId = config.DEFAULT_WI_DOC_ID;
    if (!defaultDocId || !isAccessibleDocument_(defaultDocId)) {
      var doc = createDefaultWorkInstructionDoc_(projectFolder);
      updates.DEFAULT_WI_DOC_ID = doc.getId();
    }

    if (Object.keys(updates).length) {
      upsertConfigValues(updates);
    }
    return getConfig();
  }

  function getOrCreateProjectFolder_(spreadsheet, folderId) {
    if (folderId) {
      try {
        return DriveApp.getFolderById(folderId);
      } catch (error) {
        // Fall through to create a new folder.
      }
    }

    var spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
    var parentFolderIterator = spreadsheetFile.getParents();
    var parentFolder = parentFolderIterator.hasNext() ? parentFolderIterator.next() : DriveApp.getRootFolder();
    var folderIterator = parentFolder.getFoldersByName(APP_CONSTANTS.PROJECT_NAME + ' Assets');
    if (folderIterator.hasNext()) {
      return folderIterator.next();
    }
    return parentFolder.createFolder(APP_CONSTANTS.PROJECT_NAME + ' Assets');
  }

  function getOrCreateBackupFolder_(projectFolder, backupFolderId) {
    if (backupFolderId) {
      try {
        return DriveApp.getFolderById(backupFolderId);
      } catch (error) {
        // Fall through to create a new folder.
      }
    }
    var folderIterator = projectFolder.getFoldersByName('Backups');
    if (folderIterator.hasNext()) {
      return folderIterator.next();
    }
    return projectFolder.createFolder('Backups');
  }

  function createDefaultWorkInstructionDoc_(projectFolder) {
    var doc = DocumentApp.create(APP_CONSTANTS.PROJECT_NAME + ' - Work Instructions');
    var file = DriveApp.getFileById(doc.getId());
    file.moveTo(projectFolder);
    var body = doc.getBody();
    body.clear();
    body.appendParagraph(APP_CONSTANTS.PROJECT_NAME + ' Work Instructions').setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph('Managed step sections are inserted between [[STEP_START:STEP_ID]] and [[STEP_END:STEP_ID]] markers.');
    body.appendParagraph('Manual text outside managed markers is preserved during sync.');
    doc.saveAndClose();
    return doc;
  }

  function isAccessibleDocument_(documentId) {
    try {
      DocumentApp.openById(documentId);
      return true;
    } catch (error) {
      return false;
    }
  }

  function installEditTrigger() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var triggers = ScriptApp.getProjectTriggers();
    var hasTrigger = triggers.some(function(trigger) {
      return trigger.getHandlerFunction() === 'handleSpreadsheetEdit';
    });
    if (!hasTrigger) {
      ScriptApp.newTrigger('handleSpreadsheetEdit').forSpreadsheet(spreadsheet).onEdit().create();
    }
  }

  return {
    getDefaultConfig: getDefaultConfig,
    getConfig: getConfig,
    initializeSystem: initializeSystem,
    ensureRequiredSheets: ensureRequiredSheets,
    ensureDriveArtifacts: ensureDriveArtifacts,
    upsertConfigValues: upsertConfigValues,
    installEditTrigger: installEditTrigger
  };
})();
