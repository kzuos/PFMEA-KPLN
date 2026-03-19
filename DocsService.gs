var DocsService = (function() {
  function syncStepSection(docId, stepId, payload, options) {
    if (!docId) {
      return {
        status: APP_CONSTANTS.STATUS.ERROR,
        action: 'DOC_NOT_FOUND',
        message: 'No Work Instruction document ID resolved for step ' + stepId
      };
    }

    var startMarker = APP_CONSTANTS.DOC_MARKERS.START_PREFIX + stepId + ']]';
    var endMarker = APP_CONSTANTS.DOC_MARKERS.END_PREFIX + stepId + ']]';
    var document = DocumentApp.openById(docId);
    var body = document.getBody();
    var section = findSection_(body, startMarker, endMarker, stepId);
    var sectionModel = buildSectionModel_(payload, stepId);
    var newPlainText = buildPlainText_(sectionModel);

    if (!section.found) {
      if (!options.createMissingSection) {
        return {
          status: options.dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED,
          action: 'MISSING_SECTION',
          message: 'Missing markers for step ' + stepId + ' in document ' + docId,
          beforeSummary: '',
          afterSummary: newPlainText
        };
      }

      if (options.dryRun) {
        return {
          status: APP_CONSTANTS.STATUS.PREVIEW,
          action: 'CREATE_SECTION',
          beforeSummary: '',
          afterSummary: newPlainText,
          message: 'Section would be appended to document ' + docId
        };
      }

      backupDocumentIfNeeded_(docId, options);
      appendSection_(body, startMarker, endMarker, sectionModel);
      document.saveAndClose();
      return {
        status: APP_CONSTANTS.STATUS.SUCCESS,
        action: 'CREATE_SECTION',
        beforeSummary: '',
        afterSummary: newPlainText,
        message: 'Section created in document ' + docId
      };
    }

    if (section.locked) {
      return {
        status: APP_CONSTANTS.STATUS.SKIPPED,
        action: 'LOCKED_SECTION',
        beforeSummary: section.currentText,
        afterSummary: newPlainText,
        message: 'Section is locked with a [[LOCKED:' + stepId + ']] marker'
      };
    }

    if (!options.allowOverwrite) {
      return {
        status: options.dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED,
        action: 'OVERWRITE_BLOCKED',
        beforeSummary: section.currentText,
        afterSummary: newPlainText,
        message: 'ALLOW_OVERWRITE is FALSE; existing section left unchanged'
      };
    }

    if (section.currentText === newPlainText) {
      return {
        status: options.dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED,
        action: 'NO_DOC_CHANGE',
        beforeSummary: section.currentText,
        afterSummary: newPlainText,
        message: 'Work Instruction section already aligned'
      };
    }

    if (options.dryRun) {
      return {
        status: APP_CONSTANTS.STATUS.PREVIEW,
        action: 'UPDATE_SECTION',
        beforeSummary: section.currentText,
        afterSummary: newPlainText,
        message: 'Section would be updated in document ' + docId
      };
    }

    backupDocumentIfNeeded_(docId, options);
    replaceSectionContent_(body, section.startIndex, section.endIndex, sectionModel);
    document.saveAndClose();
    return {
      status: APP_CONSTANTS.STATUS.SUCCESS,
      action: 'UPDATE_SECTION',
      beforeSummary: section.currentText,
      afterSummary: newPlainText,
      message: 'Section updated in document ' + docId
    };
  }

  function findSection_(body, startMarker, endMarker, stepId) {
    var result = {
      found: false,
      startIndex: -1,
      endIndex: -1,
      currentText: '',
      locked: false
    };
    var lockMarker = APP_CONSTANTS.DOC_MARKERS.LOCK_PREFIX + stepId + ']]';

    for (var index = 0; index < body.getNumChildren(); index += 1) {
      var child = body.getChild(index);
      var text = getElementText_(child);
      if (text === startMarker) {
        result.found = true;
        result.startIndex = index;
      } else if (text === endMarker && result.startIndex > -1) {
        result.endIndex = index;
        break;
      }
    }

    if (!result.found || result.endIndex === -1) {
      result.found = false;
      return result;
    }

    var lines = [];
    for (var contentIndex = result.startIndex + 1; contentIndex < result.endIndex; contentIndex += 1) {
      var sectionText = getElementText_(body.getChild(contentIndex));
      if (sectionText === lockMarker) {
        result.locked = true;
      }
      if (sectionText) {
        lines.push(sectionText);
      }
    }
    result.currentText = lines.join('\n');
    return result;
  }

  function getElementText_(element) {
    var type = element.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      return element.asParagraph().getText();
    }
    if (type === DocumentApp.ElementType.LIST_ITEM) {
      return element.asListItem().getText();
    }
    if (type === DocumentApp.ElementType.TABLE) {
      return element.asTable().getText();
    }
    return '';
  }

  function buildSectionModel_(payload, stepId) {
    var lines = [];
    lines.push({type: 'heading', text: payload.STEP_TITLE || stepId});
    lines.push({type: 'paragraph', text: 'Step ID: ' + stepId});
    if (!SyncUtils.isBlank(payload.OPERATION_NO)) {
      lines.push({type: 'paragraph', text: 'Operation No: ' + payload.OPERATION_NO});
    }
    if (!SyncUtils.isBlank(payload.PROCESS_DESCRIPTION)) {
      lines.push({type: 'paragraph', text: 'Process Requirement: ' + payload.PROCESS_DESCRIPTION});
    }
    appendMultiLineField_(lines, 'Failure Risks', payload.FAILURE_SUMMARY);
    appendSingleField_(lines, 'Product Characteristics', payload.PRODUCT_CHARACTERISTICS);
    appendSingleField_(lines, 'Process Characteristics', payload.PROCESS_CHARACTERISTICS);
    appendSingleField_(lines, 'Special Characteristics', payload.SPECIAL_CHARACTERISTICS);
    appendSingleField_(lines, 'Specification / Tolerance', payload.SPECIFICATION_TOLERANCE);
    appendSingleField_(lines, 'Prevention Controls', payload.PREVENTION_CONTROLS);
    appendSingleField_(lines, 'Detection Controls', payload.DETECTION_CONTROLS);
    appendSingleField_(lines, 'Control Method', payload.CONTROL_METHOD);
    appendSingleField_(lines, 'Reaction Plan', payload.REACTION_PLAN);
    appendSingleField_(lines, 'PFMEA Row IDs', payload.PFMEA_ROW_IDS);
    appendSingleField_(lines, 'Section Status', payload.SECTION_STATUS);
    appendSingleField_(lines, 'Last Sync At', payload.LAST_SYNC_AT);
    return lines;
  }

  function appendSingleField_(lines, label, value) {
    if (!SyncUtils.isBlank(value)) {
      lines.push({type: 'paragraph', text: label + ': ' + value});
    }
  }

  function appendMultiLineField_(lines, label, value) {
    if (SyncUtils.isBlank(value)) {
      return;
    }
    lines.push({type: 'paragraph', text: label + ':'});
    SyncUtils.asString(value).split(/\r?\n/).forEach(function(line) {
      if (line.trim()) {
        lines.push({type: 'paragraph', text: '- ' + line.trim()});
      }
    });
  }

  function buildPlainText_(sectionModel) {
    return sectionModel.map(function(item) {
      return item.text;
    }).join('\n');
  }

  function appendSection_(body, startMarker, endMarker, sectionModel) {
    body.appendParagraph(startMarker);
    sectionModel.forEach(function(item) {
      var paragraph = body.appendParagraph(item.text);
      if (item.type === 'heading') {
        paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      }
    });
    body.appendParagraph(endMarker);
  }

  function replaceSectionContent_(body, startIndex, endIndex, sectionModel) {
    for (var index = endIndex - 1; index > startIndex; index -= 1) {
      body.removeChild(body.getChild(index));
    }
    insertSectionModel_(body, startIndex + 1, sectionModel);
  }

  function insertSectionModel_(body, insertIndex, sectionModel) {
    var currentIndex = insertIndex;
    sectionModel.forEach(function(item) {
      var paragraph = body.insertParagraph(currentIndex, item.text);
      if (item.type === 'heading') {
        paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      }
      currentIndex += 1;
    });
  }

  function backupDocumentIfNeeded_(docId, options) {
    if (!options.backupBeforeWrite || !options.backupFolderId) {
      return;
    }
    if (!options.backupDocIds) {
      options.backupDocIds = {};
    }
    if (options.backupDocIds[docId]) {
      return;
    }
    var folder = DriveApp.getFolderById(options.backupFolderId);
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    var copy = DriveApp.getFileById(docId).makeCopy('WI Backup - ' + docId + ' - ' + timestamp, folder);
    options.backupDocIds[docId] = copy.getId();
  }

  return {
    syncStepSection: syncStepSection
  };
})();
