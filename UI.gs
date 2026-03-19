var UIService = (function() {
  function buildMenu() {
    SpreadsheetApp.getUi()
      .createMenu(APP_CONSTANTS.PROJECT_NAME)
      .addItem('Setup Actual Sync', 'setupActualSync')
      .addItem('Refresh Link Matrix', 'refreshActualLinks')
      .addItem('Refresh WI Templates', 'refreshActualTemplates')
      .addItem('Preview Actual Sync', 'previewActualSync')
      .addItem('Run Actual Sync', 'runActualSync')
      .addSeparator()
      .addItem('Open Actual Config', 'openActualConfig')
      .addItem('Open Actual Links', 'openActualLinks')
      .addItem('Open WI Templates', 'openActualTemplates')
      .addSeparator()
      .addItem('Setup Generic MVP', 'setupSystem')
      .addItem('Run Full Sync', 'runFullSync')
      .addItem('Sync Selected PFMEA Row', 'syncSelectedPfmeaRow')
      .addItem('Preview Changes', 'previewChanges')
      .addSeparator()
      .addItem('Validate Mapping', 'validateMapping')
      .addItem('Open Config', 'openConfig')
      .addItem('Install Edit Trigger', 'installPfmeaSyncTrigger')
      .addToUi();
  }

  function setupSystemAction() {
    var config = ConfigService.initializeSystem();
    var validation = ValidationService.validateSystem();
    var ui = SpreadsheetApp.getUi();
    var message = [
      'Setup completed.',
      'Default Work Instruction Doc ID: ' + config.DEFAULT_WI_DOC_ID,
      'Work Instruction Folder ID: ' + config.WI_FOLDER_ID,
      validation.warnings.length ? 'Warnings: ' + validation.warnings.join(' | ') : 'Validation passed.'
    ].join('\n');
    ui.alert(APP_CONSTANTS.PROJECT_NAME, message, ui.ButtonSet.OK);
    return config;
  }

  function runFullSyncAction(forceDryRun) {
    ConfigService.initializeSystem();
    ValidationService.assertReadyOrThrow();
    var config = ConfigService.getConfig();
    var ui = SpreadsheetApp.getUi();

    if (!forceDryRun && config.CONFIRM_BULK_SYNC) {
      var response = ui.alert(
        APP_CONSTANTS.PROJECT_NAME,
        'Run full PFMEA to Control Plan / Work Instruction sync now?',
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) {
        return null;
      }
    }

    var result = SyncService.runFullSync({
      dryRun: !!forceDryRun,
      initiatedBy: forceDryRun ? 'MENU_PREVIEW' : 'MENU_FULL_SYNC'
    });
    showSummary_(result, forceDryRun ? 'Preview completed.' : 'Full sync completed.');
    return result;
  }

  function syncSelectedPfmeaRowAction(forceDryRun) {
    ConfigService.initializeSystem();
    ValidationService.assertReadyOrThrow();
    var config = ConfigService.getConfig();
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var activeRange = activeSheet.getActiveRange();
    if (!activeRange || activeSheet.getName() !== config.PFMEA_SHEET || activeRange.getRow() <= config.PFMEA_HEADER_ROW) {
      SpreadsheetApp.getUi().alert(APP_CONSTANTS.PROJECT_NAME, 'Select a PFMEA data row first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return null;
    }

    var result = SyncService.syncPfmeaRow(activeRange.getRow(), {
      dryRun: !!forceDryRun,
      initiatedBy: forceDryRun ? 'MENU_PREVIEW' : 'MENU_ROW_SYNC'
    });
    showSummary_(result, forceDryRun ? 'Preview completed for selected PFMEA row.' : 'Selected PFMEA row synced.');
    return result;
  }

  function previewChangesAction() {
    var config = ConfigService.getConfig();
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var activeRange = activeSheet.getActiveRange();

    if (activeSheet.getName() === config.PFMEA_SHEET && activeRange && activeRange.getRow() > config.PFMEA_HEADER_ROW) {
      return syncSelectedPfmeaRowAction(true);
    }
    return runFullSyncAction(true);
  }

  function validateMappingAction() {
    ConfigService.ensureRequiredSheets();
    var validation = ValidationService.validateMappingsOnly();
    var ui = SpreadsheetApp.getUi();
    var message = [
      validation.ok ? 'Mapping validation passed.' : 'Mapping validation found issues.',
      validation.errors.length ? 'Errors: ' + validation.errors.join(' | ') : 'Errors: none',
      validation.warnings.length ? 'Warnings: ' + validation.warnings.join(' | ') : 'Warnings: none'
    ].join('\n');
    ui.alert(APP_CONSTANTS.PROJECT_NAME, message, ui.ButtonSet.OK);
    return validation;
  }

  function openConfigAction() {
    var config = ConfigService.getConfig();
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SheetsService.getSheet(config.CONFIG_SHEET));
  }

  function installTriggerAction() {
    ConfigService.installEditTrigger();
    SpreadsheetApp.getUi().alert(APP_CONSTANTS.PROJECT_NAME, 'Installable edit trigger ensured.', SpreadsheetApp.getUi().ButtonSet.OK);
  }

  function showSummary_(result, intro) {
    if (!result) {
      return;
    }
    var ui = SpreadsheetApp.getUi();
    var message = [
      intro,
      'Processed PFMEA rows: ' + result.processed,
      'Changed targets: ' + result.changed,
      'Skipped actions: ' + result.skipped,
      'Errors: ' + result.errors,
      'Control Plan writes: ' + result.controlPlanWrites,
      'Work Instruction writes: ' + result.docWrites,
      result.backupSpreadsheetId ? 'Spreadsheet backup ID: ' + result.backupSpreadsheetId : ''
    ].filter(function(line) {
      return !!line;
    }).join('\n');
    ui.alert(APP_CONSTANTS.PROJECT_NAME, message, ui.ButtonSet.OK);
  }

  function setupActualSyncAction() {
    var result = ActualSyncService.setup();
    var ui = SpreadsheetApp.getUi();
    var message = [
      'Actual sync helper sheets are ready.',
      'WI templates indexed: ' + result.templates.length,
      'PFMEA sheets scanned: ' + result.refresh.pfmeaSheets,
      'PFMEA rows indexed: ' + result.refresh.pfmeaRows,
      'KPLN blocks found: ' + result.refresh.kplnBlocks,
      'SYNC_LINKS rows created: ' + result.refresh.linkRows,
      'Next: review WI_TEMPLATES, then open SYNC_LINKS and approve mappings.'
    ].join('\n');
    ui.alert(APP_CONSTANTS.PROJECT_NAME, message, ui.ButtonSet.OK);
    return result;
  }

  function refreshActualLinksAction() {
    var result = ActualSyncService.refreshLinks();
    var ui = SpreadsheetApp.getUi();
    var message = [
      'Actual link matrix refreshed.',
      'PFMEA sheets scanned: ' + result.pfmeaSheets,
      'PFMEA rows indexed: ' + result.pfmeaRows,
      'KPLN blocks found: ' + result.kplnBlocks,
      'SYNC_LINKS rows created: ' + result.linkRows
    ].join('\n');
    ui.alert(APP_CONSTANTS.PROJECT_NAME, message, ui.ButtonSet.OK);
    return result;
  }

  function previewActualSyncAction() {
    var result = ActualSyncService.previewSync();
    showActualSummary_(result, 'Actual sync preview completed.');
    return result;
  }

  function refreshActualTemplatesAction() {
    var result = ActualSyncService.refreshTemplates();
    SpreadsheetApp.getUi().alert(
      APP_CONSTANTS.PROJECT_NAME,
      'WI templates refreshed.\nTemplate rows: ' + result.templateRows,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return result;
  }

  function runActualSyncAction() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      APP_CONSTANTS.PROJECT_NAME,
      'Run actual sync against APPROVED rows in SYNC_LINKS?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return null;
    }
    var result = ActualSyncService.runSync();
    showActualSummary_(result, 'Actual sync completed.');
    return result;
  }

  function openActualConfigAction() {
    ActualSyncService.openConfig();
  }

  function openActualLinksAction() {
    ActualSyncService.openLinks();
  }

  function openActualTemplatesAction() {
    ActualSyncService.openTemplates();
  }

  function showActualSummary_(result, intro) {
    var ui = SpreadsheetApp.getUi();
    var message = [
      intro,
      'Links processed: ' + result.processed,
      'Changed actions: ' + result.changed,
      'Skipped actions: ' + result.skipped,
      'Errors: ' + result.errors,
      'KPLN writes: ' + result.kplnWrites,
      'WI writes: ' + result.wiWrites
    ].join('\n');
    ui.alert(APP_CONSTANTS.PROJECT_NAME, message, ui.ButtonSet.OK);
  }

  return {
    buildMenu: buildMenu,
    setupSystemAction: setupSystemAction,
    runFullSyncAction: runFullSyncAction,
    syncSelectedPfmeaRowAction: syncSelectedPfmeaRowAction,
    previewChangesAction: previewChangesAction,
    validateMappingAction: validateMappingAction,
    openConfigAction: openConfigAction,
    installTriggerAction: installTriggerAction,
    setupActualSyncAction: setupActualSyncAction,
    refreshActualLinksAction: refreshActualLinksAction,
    refreshActualTemplatesAction: refreshActualTemplatesAction,
    previewActualSyncAction: previewActualSyncAction,
    runActualSyncAction: runActualSyncAction,
    openActualConfigAction: openActualConfigAction,
    openActualLinksAction: openActualLinksAction,
    openActualTemplatesAction: openActualTemplatesAction
  };
})();
