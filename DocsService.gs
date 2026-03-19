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
    var document;
    try {
      document = DocumentApp.openById(docId);
    } catch (error) {
      return {
        status: APP_CONSTANTS.STATUS.ERROR,
        action: 'DOC_OPEN_ERROR',
        message: 'Unable to open Work Instruction document ' + docId + ': ' + error.message
      };
    }
    var body = document.getBody();
    var section = findSection_(body, startMarker, endMarker, stepId);
    var sectionModel = buildSectionModel_(payload, stepId);
    var newPlainText = buildPlainText_(sectionModel);
    var placement = resolvePlacement_(body, options);

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
          message: 'Section would be inserted ' + placement.messageSuffix + ' in document ' + docId
        };
      }

      backupDocumentIfNeeded_(docId, options);
      insertManagedSection_(body, placement.insertIndex, startMarker, endMarker, sectionModel);
      document.saveAndClose();
      return {
        status: APP_CONSTANTS.STATUS.SUCCESS,
        action: 'CREATE_SECTION',
        beforeSummary: '',
        afterSummary: newPlainText,
        message: 'Section created ' + placement.messageSuffix + ' in document ' + docId
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
      if (options.relocateManagedSection && shouldRelocateSection_(section, placement)) {
        if (options.dryRun) {
          return {
            status: APP_CONSTANTS.STATUS.PREVIEW,
            action: 'RELOCATE_SECTION',
            beforeSummary: section.currentText,
            afterSummary: newPlainText,
            message: 'Section would be moved ' + placement.messageSuffix + ' in document ' + docId
          };
        }

        backupDocumentIfNeeded_(docId, options);
        relocateManagedSection_(body, section, placement, startMarker, endMarker, sectionModel);
        document.saveAndClose();
        return {
          status: APP_CONSTANTS.STATUS.SUCCESS,
          action: 'RELOCATE_SECTION',
          beforeSummary: section.currentText,
          afterSummary: newPlainText,
          message: 'Section moved ' + placement.messageSuffix + ' in document ' + docId
        };
      }

      return {
        status: options.dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED,
        action: 'NO_DOC_CHANGE',
        beforeSummary: section.currentText,
        afterSummary: newPlainText,
        message: 'Work Instruction section already aligned'
      };
    }

    if (options.dryRun) {
      if (options.relocateManagedSection && shouldRelocateSection_(section, placement)) {
        return {
          status: APP_CONSTANTS.STATUS.PREVIEW,
          action: 'RELOCATE_SECTION',
          beforeSummary: section.currentText,
          afterSummary: newPlainText,
          message: 'Section would be moved and updated ' + placement.messageSuffix + ' in document ' + docId
        };
      }

      return {
        status: APP_CONSTANTS.STATUS.PREVIEW,
        action: 'UPDATE_SECTION',
        beforeSummary: section.currentText,
        afterSummary: newPlainText,
        message: 'Section would be updated in document ' + docId
      };
    }

    backupDocumentIfNeeded_(docId, options);
    if (options.relocateManagedSection && shouldRelocateSection_(section, placement)) {
      relocateManagedSection_(body, section, placement, startMarker, endMarker, sectionModel);
    } else {
      replaceSectionContent_(body, section.startIndex, section.endIndex, sectionModel);
    }
    document.saveAndClose();
    return {
      status: APP_CONSTANTS.STATUS.SUCCESS,
      action: options.relocateManagedSection && shouldRelocateSection_(section, placement) ? 'RELOCATE_SECTION' : 'UPDATE_SECTION',
      beforeSummary: section.currentText,
      afterSummary: newPlainText,
      message: (options.relocateManagedSection && shouldRelocateSection_(section, placement) ? 'Section moved and updated ' + placement.messageSuffix : 'Section updated') + ' in document ' + docId
    };
  }

  function ensureWorkInstructionDocument(stepId, recordValues, options) {
    var explicitDocId = SyncUtils.asString(recordValues.WI_DOC_ID);
    if (explicitDocId) {
      if (isDocumentAccessible_(explicitDocId)) {
        return {
          status: APP_CONSTANTS.STATUS.SUCCESS,
          action: 'USE_EXISTING_DOC',
          docId: explicitDocId,
          created: false,
          message: 'Using existing Work Instruction document ' + explicitDocId
        };
      }
      return {
        status: APP_CONSTANTS.STATUS.ERROR,
        action: 'INVALID_WI_DOC_ID',
        docId: explicitDocId,
        created: false,
        message: 'WI_DOC_ID ' + explicitDocId + ' is set but not accessible for step ' + stepId
      };
    }

    if (!options.createMissingDocument) {
      return {
        status: options.dryRun ? APP_CONSTANTS.STATUS.PREVIEW : APP_CONSTANTS.STATUS.SKIPPED,
        action: 'USE_DEFAULT_DOC',
        docId: SyncUtils.asString(options.defaultDocId),
        created: false,
        message: 'No WI_DOC_ID set for step ' + stepId + '; using default Work Instruction document.'
      };
    }

    var documentName = buildWorkInstructionName_(stepId, recordValues);
    if (options.dryRun) {
      return {
        status: APP_CONSTANTS.STATUS.PREVIEW,
        action: 'CREATE_WI_DOC',
        docId: 'PREVIEW-' + stepId,
        created: true,
        documentName: documentName,
        message: 'A new Work Instruction document would be created: ' + documentName
      };
    }

    var document = createWorkInstructionDocument_(documentName, recordValues, options);
    return {
      status: APP_CONSTANTS.STATUS.SUCCESS,
      action: 'CREATE_WI_DOC',
      docId: document.getId(),
      created: true,
      documentName: documentName,
      message: 'Created Work Instruction document ' + documentName + ' (' + document.getId() + ')'
    };
  }

  function isDocumentAccessible(documentId) {
    return isDocumentAccessible_(documentId);
  }

  function ensureGoogleDocTemplate(sourceFileId, templateName, folderId) {
    return ensureGoogleDocTemplate_(sourceFileId, templateName, folderId);
  }

  function findSection_(body, startMarker, endMarker, stepId) {
    var result = {
      found: false,
      startIndex: -1,
      endIndex: -1,
      currentText: '',
      locked: false,
      inclusiveEndIndex: -1
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
        result.inclusiveEndIndex = index;
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
    if (!SyncUtils.isBlank(payload.MANAGED_SECTION_TITLE)) {
      lines.push({type: 'heading', text: payload.MANAGED_SECTION_TITLE});
      lines.push({type: 'subheading', text: payload.STEP_TITLE || stepId});
    } else {
      lines.push({type: 'heading', text: payload.STEP_TITLE || stepId});
    }
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

  function insertManagedSection_(body, insertIndex, startMarker, endMarker, sectionModel) {
    if (insertIndex >= body.getNumChildren()) {
      body.appendParagraph(startMarker);
      sectionModel.forEach(function(item) {
        var paragraph = body.appendParagraph(item.text);
        applySectionFormatting_(paragraph, item);
      });
      body.appendParagraph(endMarker);
      return;
    }

    var currentIndex = insertIndex;
    body.insertParagraph(currentIndex, startMarker);
    currentIndex += 1;
    sectionModel.forEach(function(item) {
      var paragraph = body.insertParagraph(currentIndex, item.text);
      applySectionFormatting_(paragraph, item);
      currentIndex += 1;
    });
    body.insertParagraph(currentIndex, endMarker);
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
      applySectionFormatting_(paragraph, item);
      currentIndex += 1;
    });
  }

  function applySectionFormatting_(paragraph, item) {
    if (item.type === 'heading') {
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      return;
    }
    if (item.type === 'subheading') {
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      return;
    }
    paragraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  }

  function relocateManagedSection_(body, section, placement, startMarker, endMarker, sectionModel) {
    var insertedLength = sectionModel.length + 2;
    var insertIndex = placement.insertIndex;
    insertManagedSection_(body, insertIndex, startMarker, endMarker, sectionModel);

    var removeStart = section.startIndex;
    var removeEnd = section.inclusiveEndIndex;
    if (insertIndex <= section.startIndex) {
      removeStart += insertedLength;
      removeEnd += insertedLength;
    }

    for (var removeIndex = removeEnd; removeIndex >= removeStart; removeIndex -= 1) {
      body.removeChild(body.getChild(removeIndex));
    }
  }

  function shouldRelocateSection_(section, placement) {
    if (!placement || !placement.isAnchored) {
      return false;
    }
    if (section.startIndex < placement.sectionStartIndex) {
      return true;
    }
    return section.inclusiveEndIndex > placement.sectionEndIndex;
  }

  function resolvePlacement_(body, options) {
    var normalizedOptions = options || {};
    var placement = {
      insertIndex: body.getNumChildren(),
      sectionStartIndex: 0,
      sectionEndIndex: Math.max(body.getNumChildren() - 1, 0),
      anchorText: '',
      isAnchored: false,
      messageSuffix: 'at the end',
      options: normalizedOptions
    };

    var anchorPatterns = normalizedOptions.anchorPatterns || [];
    if (!anchorPatterns.length) {
      return placement;
    }

    var anchorIndex = findAnchorIndex_(body, anchorPatterns, normalizedOptions.anchorMatch || 'first');
    if (anchorIndex === -1) {
      return placement;
    }

    var boundaryIndex = findSectionBoundaryIndex_(body, anchorIndex + 1, normalizedOptions.nextSectionPatterns || []);
    placement.insertIndex = boundaryIndex;
    placement.sectionStartIndex = anchorIndex;
    placement.sectionEndIndex = Math.max(boundaryIndex - 1, anchorIndex);
    placement.anchorText = getElementText_(body.getChild(anchorIndex));
    placement.isAnchored = true;
    placement.messageSuffix = 'under "' + placement.anchorText + '"';
    return placement;
  }

  function findAnchorIndex_(body, anchorPatterns, anchorMatch) {
    var matches = [];
    for (var index = 0; index < body.getNumChildren(); index += 1) {
      var element = body.getChild(index);
      var normalized = normalizeMarkerText_(getElementText_(element));
      if (!normalized) {
        continue;
      }
      if (anchorPatterns.some(function(pattern) {
        return normalized.indexOf(normalizeMarkerText_(pattern)) > -1;
      }) && isHeadingLikeElement_(element, normalized)) {
        matches.push(index);
      }
    }
    if (!matches.length) {
      return -1;
    }
    return anchorMatch === 'last' ? matches[matches.length - 1] : matches[0];
  }

  function findSectionBoundaryIndex_(body, startIndex, nextSectionPatterns) {
    for (var index = startIndex; index < body.getNumChildren(); index += 1) {
      var element = body.getChild(index);
      var text = getElementText_(element);
      var normalized = normalizeMarkerText_(text);
      if (!normalized) {
        continue;
      }
      if (matchesSectionBoundary_(element, normalized, nextSectionPatterns)) {
        return index;
      }
    }
    return body.getNumChildren();
  }

  function matchesSectionBoundary_(element, normalizedText, nextSectionPatterns) {
    if (nextSectionPatterns.some(function(pattern) {
      return normalizedText.indexOf(normalizeMarkerText_(pattern)) > -1;
    })) {
      return true;
    }
    return isHeadingLikeElement_(element, normalizedText) && /^\d+\s/.test(normalizedText);
  }

  function isHeadingLikeElement_(element, normalizedText) {
    var type = element.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      var heading = element.asParagraph().getHeading();
      if (heading && heading !== DocumentApp.ParagraphHeading.NORMAL) {
        return true;
      }
    }

    if (normalizedText.length <= 80 && /^\d+\s+[A-Z]/.test(normalizedText)) {
      return true;
    }

    return normalizedText.length <= 60 && normalizedText === normalizedText.toUpperCase();
  }

  function normalizeMarkerText_(value) {
    return SyncUtils.asString(value)
      .toUpperCase()
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

  function createWorkInstructionDocument_(documentName, recordValues, options) {
    var document;
    var usedTemplate = false;

    if (options.templateDocId && isDocumentAccessible_(options.templateDocId)) {
      var templateCopy = DriveApp.getFileById(options.templateDocId).makeCopy(documentName);
      document = DocumentApp.openById(templateCopy.getId());
      usedTemplate = true;
    } else if (options.templateSourceFileId) {
      var importedTemplate = ensureGoogleDocTemplate_(options.templateSourceFileId, options.templateName, options.folderId);
      var importedTemplateCopy = DriveApp.getFileById(importedTemplate.docId).makeCopy(documentName);
      document = DocumentApp.openById(importedTemplateCopy.getId());
      usedTemplate = true;
    } else {
      document = DocumentApp.create(documentName);
    }

    var file = DriveApp.getFileById(document.getId());
    if (options.folderId) {
      try {
        file.moveTo(DriveApp.getFolderById(options.folderId));
      } catch (ignored) {
        // If move fails, keep the document where it was created.
      }
    }

    initializeWorkInstructionDocument_(document, recordValues, usedTemplate);
    document.saveAndClose();
    return DocumentApp.openById(document.getId());
  }

  function initializeWorkInstructionDocument_(document, recordValues, usedTemplate) {
    var body = document.getBody();
    var templateHasContent = !SyncUtils.isBlank(body.getText());
    replacePlaceholders_(body, {
      DOC_TITLE: document.getName(),
      STEP_ID: recordValues.STEP_ID,
      OPERATION_NO: recordValues.OPERATION_NO,
      PROCESS_STEP: recordValues.PROCESS_STEP
    });

    if (!usedTemplate || !templateHasContent) {
      body.clear();
      body.appendParagraph(document.getName()).setHeading(DocumentApp.ParagraphHeading.TITLE);
      body.appendParagraph('System-managed Work Instruction generated from PFMEA.');
      body.appendParagraph('Manual text outside [[STEP_START:STEP_ID]] and [[STEP_END:STEP_ID]] markers is preserved.');
      if (!SyncUtils.isBlank(recordValues.STEP_ID)) {
        body.appendParagraph('Step ID: ' + recordValues.STEP_ID);
      }
      if (!SyncUtils.isBlank(recordValues.OPERATION_NO)) {
        body.appendParagraph('Operation No: ' + recordValues.OPERATION_NO);
      }
      if (!SyncUtils.isBlank(recordValues.PROCESS_STEP)) {
        body.appendParagraph('Process Step: ' + recordValues.PROCESS_STEP);
      }
      body.appendParagraph('');
    }
  }

  function buildWorkInstructionName_(stepId, recordValues) {
    var segments = ['WI'];
    if (!SyncUtils.isBlank(recordValues.OPERATION_NO)) {
      segments.push(SyncUtils.asString(recordValues.OPERATION_NO));
    }
    if (!SyncUtils.isBlank(recordValues.PROCESS_STEP)) {
      segments.push(SyncUtils.asString(recordValues.PROCESS_STEP));
    }
    segments.push(stepId);
    return SyncUtils.sanitizeDriveName(segments.join(' - '), 'WI - ' + stepId);
  }

  function isDocumentAccessible_(documentId) {
    if (!documentId) {
      return false;
    }
    try {
      DocumentApp.openById(documentId);
      return true;
    } catch (error) {
      return false;
    }
  }

  function ensureGoogleDocTemplate_(sourceFileId, templateName, folderId) {
    if (!sourceFileId) {
      throw new Error('No source template file ID was provided.');
    }

    var sourceFile = DriveApp.getFileById(sourceFileId);
    var mimeType = sourceFile.getMimeType();
    if (mimeType === MimeType.GOOGLE_DOCS || mimeType === 'application/vnd.google-apps.document') {
      return {
        docId: sourceFileId,
        name: sourceFile.getName(),
        imported: false
      };
    }

    var metadata = {
      name: SyncUtils.sanitizeDriveName(templateName || ('Template - ' + sourceFile.getName()), 'WI Template'),
      mimeType: 'application/vnd.google-apps.document'
    };
    if (folderId) {
      metadata.parents = [folderId];
    }

    var converted = Drive.Files.create(metadata, sourceFile.getBlob(), {
      fields: 'id,name,mimeType'
    });

    return {
      docId: converted.id,
      name: converted.name || metadata.name,
      imported: true
    };
  }

  function replacePlaceholders_(body, replacements) {
    Object.keys(replacements || {}).forEach(function(key) {
      var value = SyncUtils.asString(replacements[key]);
      body.replaceText('\\[\\[' + escapeRegex_(key) + '\\]\\]', value);
    });
  }

  function escapeRegex_(value) {
    return SyncUtils.asString(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  return {
    syncStepSection: syncStepSection,
    ensureWorkInstructionDocument: ensureWorkInstructionDocument,
    ensureGoogleDocTemplate: ensureGoogleDocTemplate,
    isDocumentAccessible: isDocumentAccessible
  };
})();
