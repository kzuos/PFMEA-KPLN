function onOpen(e) {
  UIService.buildMenu();
}

function doGet(e) {
  return handleRemoteWebRequest_(e);
}

function doPost(e) {
  return handleRemoteWebRequest_(e);
}

function onEdit(e) {
  return SyncService.handleSimpleEdit(e);
}

function setupSystem() {
  return UIService.setupSystemAction();
}

function setupActualSync() {
  return UIService.setupActualSyncAction();
}

function refreshActualLinks() {
  return UIService.refreshActualLinksAction();
}

function refreshActualTemplates() {
  return UIService.refreshActualTemplatesAction();
}

function validateActualSync() {
  return UIService.validateActualSyncAction();
}

function previewActualSync() {
  return UIService.previewActualSyncAction();
}

function runActualSync() {
  return UIService.runActualSyncAction();
}

function openActualConfig() {
  return UIService.openActualConfigAction();
}

function openActualLinks() {
  return UIService.openActualLinksAction();
}

function openActualTemplates() {
  return UIService.openActualTemplatesAction();
}

function remoteSetupActualSync() {
  return runRemoteActualAction_(function() {
    return ActualSyncService.setup();
  });
}

function remoteRefreshActualLinks() {
  return runRemoteActualAction_(function() {
    return ActualSyncService.refreshLinks();
  });
}

function remoteRefreshActualTemplates() {
  return runRemoteActualAction_(function() {
    return ActualSyncService.refreshTemplates();
  });
}

function remoteValidateActualSync() {
  return runRemoteActualAction_(function() {
    return ActualSyncService.validateSetup();
  });
}

function remotePreviewActualSync() {
  return runRemoteActualAction_(function() {
    return ActualSyncService.previewSync();
  });
}

function remoteRunActualSync() {
  return runRemoteActualAction_(function() {
    return ActualSyncService.runSync();
  });
}

function remoteDebugActualSelection() {
  return runRemoteActualAction_(function() {
    return ActualSyncService.debugLinkSelection('LINK-10.1', 10);
  });
}

function runFullSync() {
  return UIService.runFullSyncAction(false);
}

function previewChanges() {
  return UIService.previewChangesAction();
}

function syncSelectedPfmeaRow() {
  return UIService.syncSelectedPfmeaRowAction(false);
}

function validateMapping() {
  return UIService.validateMappingAction();
}

function openConfig() {
  return UIService.openConfigAction();
}

function installPfmeaSyncTrigger() {
  return UIService.installTriggerAction();
}

function handleSpreadsheetEdit(e) {
  return SyncService.handleEditTrigger(e);
}

function runRemoteActualAction_(action) {
  try {
    return action();
  } catch (error) {
    return {
      ok: false,
      error: error && error.message ? error.message : String(error)
    };
  }
}

function handleRemoteWebRequest_(e) {
  var request = parseRemoteWebRequest_(e);
  var action = request.action;
  var result;

  if (action === 'refresh') {
    result = runRemoteActualAction_(function() {
      return ActualSyncService.refreshLinks();
    });
  } else if (action === 'inspectLink') {
    result = runRemoteActualAction_(function() {
      return inspectActualLink_(request.linkKey || 'LINK-10.1');
    });
  } else if (action === 'debugSelection') {
    result = runRemoteActualAction_(function() {
      return ActualSyncService.debugLinkSelection(request.linkKey || 'LINK-10.1', request.limit || 10);
    });
  } else if (action === 'debugSheet') {
    result = runRemoteActualAction_(function() {
      return ActualSyncService.debugPfmeaSheet(request.sheetName || '13', request.limit || 20);
    });
  } else if (action === 'refreshTemplates') {
    result = runRemoteActualAction_(function() {
      return ActualSyncService.refreshTemplates();
    });
  } else if (action === 'validate') {
    result = runRemoteActualAction_(function() {
      return ActualSyncService.validateSetup();
    });
  } else if (action === 'preview') {
    result = runRemoteActualAction_(function() {
      return ActualSyncService.previewSync();
    });
  } else {
    result = {
      ok: false,
      error: 'Unsupported action. Use refresh, inspectLink, debugSelection, debugSheet, refreshTemplates, validate, or preview.'
    };
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      ok: !result || result.ok !== false,
      action: action || '',
      result: result
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseRemoteWebRequest_(e) {
  var request = {};
  if (e && e.parameter) {
    request.action = e.parameter.action || '';
    request.linkKey = e.parameter.linkKey || '';
    request.sheetName = e.parameter.sheetName || '';
    request.limit = e.parameter.limit || '';
  }
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      if (body && body.action) {
        request.action = body.action;
      }
      if (body && body.linkKey) {
        request.linkKey = body.linkKey;
      }
      if (body && body.sheetName) {
        request.sheetName = body.sheetName;
      }
      if (body && body.limit) {
        request.limit = body.limit;
      }
    } catch (ignored) {
      // Fall back to query parameters.
    }
  }
  return request;
}

function inspectActualLink_(linkKey) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SYNC_LINKS');
  if (!sheet) {
    throw new Error('SYNC_LINKS sheet was not found.');
  }
  var values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) {
    throw new Error('SYNC_LINKS has no data rows.');
  }
  var headers = values[0];
  for (var index = 1; index < values.length; index += 1) {
    if (values[index][2] === linkKey) {
      var row = {};
      headers.forEach(function(header, headerIndex) {
        row[header] = values[index][headerIndex];
      });
      return row;
    }
  }
  throw new Error('Link not found: ' + linkKey);
}
