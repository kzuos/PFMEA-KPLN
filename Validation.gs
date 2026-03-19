var ValidationService = (function() {
  function validateSystem() {
    var errors = [];
    var warnings = [];
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var config = ConfigService.getConfig();

    validateSheet_(spreadsheet, config.PFMEA_SHEET, APP_CONSTANTS.HEADERS.PFMEA, errors);
    validateSheet_(spreadsheet, config.CONTROL_PLAN_SHEET, APP_CONSTANTS.HEADERS.CONTROL_PLAN, errors);
    validateSheet_(spreadsheet, config.MAPPING_SHEET, APP_CONSTANTS.HEADERS.MAPPING, errors);
    validateSheet_(spreadsheet, config.CHANGE_LOG_SHEET, APP_CONSTANTS.HEADERS.CHANGE_LOG, errors);
    validateSheet_(spreadsheet, config.CONFIG_SHEET, APP_CONSTANTS.HEADERS.CONFIG, errors);
    validateMappings_(config, errors, warnings);
    validateDriveArtifacts_(config, warnings);
    validateTrigger_(warnings);

    return {
      ok: errors.length === 0,
      errors: errors,
      warnings: warnings
    };
  }

  function validateMappingsOnly() {
    var validation = validateSystem();
    return {
      ok: validation.ok,
      errors: validation.errors.filter(function(error) {
        return error.indexOf('Mapping') > -1 || error.indexOf('header') > -1;
      }),
      warnings: validation.warnings
    };
  }

  function validateSheet_(spreadsheet, sheetName, requiredHeaders, errors) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      errors.push('Missing sheet: ' + sheetName);
      return;
    }
    var lastColumn = Math.max(sheet.getLastColumn(), requiredHeaders.length);
    var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(header) {
      return SyncUtils.asString(header);
    });
    requiredHeaders.forEach(function(requiredHeader) {
      if (headers.indexOf(requiredHeader) === -1) {
        errors.push('Missing header "' + requiredHeader + '" in sheet ' + sheetName);
      }
    });
  }

  function validateMappings_(config, errors, warnings) {
    var mappings = MappingService.loadMappings(config, true);
    if (!mappings.length) {
      errors.push('Mapping sheet is empty.');
      return;
    }

    var pfmeaHeaders = SheetsService.getHeaders(config.PFMEA_SHEET, config.PFMEA_HEADER_ROW);
    var controlPlanHeaders = SheetsService.getHeaders(config.CONTROL_PLAN_SHEET, config.CONTROL_PLAN_HEADER_ROW);
    var activeCount = 0;

    mappings.forEach(function(mapping) {
      if (SyncUtils.toBoolean(mapping.ACTIVE)) {
        activeCount += 1;
      }

      SyncUtils.asString(mapping.SOURCE_COLUMNS)
        .split(',')
        .map(function(column) {
          return column.trim();
        })
        .filter(function(column) {
          return !!column;
        })
        .forEach(function(column) {
          if (pfmeaHeaders.indexOf(column) === -1) {
            errors.push('Mapping source column "' + column + '" does not exist in PFMEA.');
          }
        });

      if (mapping.TARGET_TYPE === APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN && controlPlanHeaders.indexOf(mapping.TARGET_FIELD) === -1) {
        errors.push('Mapping target field "' + mapping.TARGET_FIELD + '" does not exist in CONTROL_PLAN.');
      }

      if (mapping.TARGET_TYPE === APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION &&
          APP_CONSTANTS.WI_SUPPORTED_FIELDS.indexOf(mapping.TARGET_FIELD) === -1) {
        errors.push('Mapping target field "' + mapping.TARGET_FIELD + '" is not supported for Work Instructions.');
      }
    });

    if (!activeCount) {
      warnings.push('No active mapping rows found.');
    }
  }

  function validateDriveArtifacts_(config, warnings) {
    if (!config.DEFAULT_WI_DOC_ID) {
      warnings.push('DEFAULT_WI_DOC_ID is not set yet. Setup can create one.');
    } else {
      try {
        DocumentApp.openById(config.DEFAULT_WI_DOC_ID);
      } catch (error) {
        warnings.push('DEFAULT_WI_DOC_ID is not accessible.');
      }
    }

    if (config.WI_FOLDER_ID) {
      try {
        DriveApp.getFolderById(config.WI_FOLDER_ID);
      } catch (error) {
        warnings.push('WI_FOLDER_ID is not accessible.');
      }
    }

    if (config.BACKUP_FOLDER_ID) {
      try {
        DriveApp.getFolderById(config.BACKUP_FOLDER_ID);
      } catch (error) {
        warnings.push('BACKUP_FOLDER_ID is not accessible.');
      }
    }
  }

  function validateTrigger_(warnings) {
    var hasTrigger = ScriptApp.getProjectTriggers().some(function(trigger) {
      return trigger.getHandlerFunction() === 'handleSpreadsheetEdit';
    });
    if (!hasTrigger) {
      warnings.push('Installable edit trigger is not installed.');
    }
  }

  function assertReadyOrThrow() {
    var validation = validateSystem();
    if (!validation.ok) {
      throw new Error('System validation failed: ' + validation.errors.join(' | '));
    }
  }

  return {
    validateSystem: validateSystem,
    validateMappingsOnly: validateMappingsOnly,
    assertReadyOrThrow: assertReadyOrThrow
  };
})();
