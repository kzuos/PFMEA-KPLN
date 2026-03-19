function onOpen(e) {
  UIService.buildMenu();
}

function onEdit(e) {
  return SyncService.handleSimpleEdit(e);
}

function setupSystem() {
  return UIService.setupSystemAction();
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
