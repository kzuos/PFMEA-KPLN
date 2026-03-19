function onOpen(e) {
  UIService.buildMenu();
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
