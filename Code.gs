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
      error: 'Unsupported action. Use refresh, refreshTemplates, validate, or preview.'
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
  }
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      if (body && body.action) {
        request.action = body.action;
      }
    } catch (ignored) {
      // Fall back to query parameters.
    }
  }
  return request;
}
