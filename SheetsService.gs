var SheetsService = (function() {
  function getSheet(sheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('Missing sheet: ' + sheetName);
    }
    return sheet;
  }

  function getHeaderMap(sheetName, headerRow) {
    var headers = getHeaders(sheetName, headerRow);
    var map = {};
    headers.forEach(function(header, index) {
      if (header) {
        map[header] = index + 1;
      }
    });
    return map;
  }

  function getHeaders(sheetName, headerRow) {
    var sheet = getSheet(sheetName);
    var lastColumn = Math.max(sheet.getLastColumn(), 1);
    return sheet.getRange(headerRow, 1, 1, lastColumn).getDisplayValues()[0].map(function(header) {
      return SyncUtils.asString(header);
    });
  }

  function getRecords(sheetName, headerRow) {
    var sheet = getSheet(sheetName);
    var lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) {
      return [];
    }
    var headers = getHeaders(sheetName, headerRow);
    var values = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, headers.length).getDisplayValues();
    var records = [];
    values.forEach(function(row, index) {
      var data = {};
      headers.forEach(function(header, headerIndex) {
        if (header) {
          data[header] = row[headerIndex];
        }
      });
      if (isMeaningfulRecord_(data)) {
        records.push({
          rowNumber: headerRow + index + 1,
          values: data
        });
      }
    });
    return records;
  }

  function getRecordByRowNumber(sheetName, headerRow, rowNumber) {
    var headers = getHeaders(sheetName, headerRow);
    var rowValues = getSheet(sheetName).getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
    var data = {};
    headers.forEach(function(header, index) {
      if (header) {
        data[header] = rowValues[index];
      }
    });
    return {
      rowNumber: rowNumber,
      values: data
    };
  }

  function updateRecordByRow(sheetName, headerRow, rowNumber, updates) {
    var sheet = getSheet(sheetName);
    var headers = getHeaders(sheetName, headerRow);
    var headerMap = getHeaderMap(sheetName, headerRow);
    var range = sheet.getRange(rowNumber, 1, 1, headers.length);
    var values = range.getValues()[0];

    Object.keys(updates).forEach(function(field) {
      if (headerMap[field]) {
        values[headerMap[field] - 1] = updates[field];
      }
    });

    range.setValues([values]);
  }

  function appendRecord(sheetName, headerRow, record) {
    var sheet = getSheet(sheetName);
    var headers = getHeaders(sheetName, headerRow);
    var row = headers.map(function(header) {
      return record.hasOwnProperty(header) ? record[header] : '';
    });
    sheet.appendRow(row);
    return sheet.getLastRow();
  }

  function appendRows(sheetName, headerRow, rows) {
    if (!rows.length) {
      return;
    }
    var sheet = getSheet(sheetName);
    var headers = getHeaders(sheetName, headerRow);
    var values = rows.map(function(record) {
      return headers.map(function(header) {
        return record.hasOwnProperty(header) ? record[header] : '';
      });
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
  }

  function findRecordsByField(sheetName, headerRow, fieldName, fieldValue) {
    var expected = SyncUtils.asString(fieldValue);
    return getRecords(sheetName, headerRow).filter(function(record) {
      return SyncUtils.asString(record.values[fieldName]) === expected;
    });
  }

  function ensurePfmeaIdentifiers(rowNumber, config, dryRun) {
    var record = getRecordByRowNumber(config.PFMEA_SHEET, config.PFMEA_HEADER_ROW, rowNumber);
    var updates = {};
    if (config.AUTO_CREATE_IDS && SyncUtils.isBlank(record.values.PFMEA_ROW_ID)) {
      updates.PFMEA_ROW_ID = SyncUtils.generateId('PFR');
    }
    if (config.AUTO_CREATE_IDS && SyncUtils.isBlank(record.values.STEP_ID)) {
      updates.STEP_ID = SyncUtils.formatStepId(record.values.OPERATION_NO || record.values.PROCESS_STEP);
    }
    if (SyncUtils.isBlank(record.values.ACTIVE)) {
      updates.ACTIVE = 'TRUE';
    }
    if (Object.keys(updates).length && !dryRun) {
      updateRecordByRow(config.PFMEA_SHEET, config.PFMEA_HEADER_ROW, rowNumber, updates);
    }
    return updates;
  }

  function getPfmeaRecords(config) {
    return getRecords(config.PFMEA_SHEET, config.PFMEA_HEADER_ROW);
  }

  function getControlPlanRecords(config) {
    return getRecords(config.CONTROL_PLAN_SHEET, config.CONTROL_PLAN_HEADER_ROW);
  }

  function backupActiveSpreadsheet(config, reason) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var backupFolder = DriveApp.getFolderById(config.BACKUP_FOLDER_ID);
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    var backupName = spreadsheet.getName() + ' - ' + reason + ' - ' + timestamp;
    var fileCopy = DriveApp.getFileById(spreadsheet.getId()).makeCopy(backupName, backupFolder);
    return fileCopy.getId();
  }

  function isMeaningfulRecord_(recordValues) {
    return Object.keys(recordValues).some(function(key) {
      return !SyncUtils.isBlank(recordValues[key]);
    });
  }

  return {
    getSheet: getSheet,
    getHeaders: getHeaders,
    getHeaderMap: getHeaderMap,
    getRecords: getRecords,
    getRecordByRowNumber: getRecordByRowNumber,
    updateRecordByRow: updateRecordByRow,
    appendRecord: appendRecord,
    appendRows: appendRows,
    findRecordsByField: findRecordsByField,
    ensurePfmeaIdentifiers: ensurePfmeaIdentifiers,
    getPfmeaRecords: getPfmeaRecords,
    getControlPlanRecords: getControlPlanRecords,
    backupActiveSpreadsheet: backupActiveSpreadsheet
  };
})();
