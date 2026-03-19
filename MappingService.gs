var MappingService = (function() {
  function getDefaultMappings() {
    return [
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'OPERATION_NO', 'OPERATION_NO', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Sync operation number'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'PROCESS_STEP', 'PROCESS_STEP', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Sync process step wording'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'PRODUCT_CHARACTERISTICS', 'PRODUCT_CHARACTERISTICS', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Product characteristic'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'PROCESS_CHARACTERISTICS', 'PROCESS_CHARACTERISTICS', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Process characteristic'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'SPECIAL_CHARACTERISTICS', 'SPECIAL_CHARACTERISTICS', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Special characteristics'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'SPECIFICATION_TOLERANCE', 'SPECIFICATION_TOLERANCE', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Specification or tolerance'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'EVALUATION_MEASUREMENT_TECHNIQUE', 'EVALUATION_TECHNIQUE', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Measurement method'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'SAMPLE_SIZE', 'SAMPLE_SIZE', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Sample size'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'SAMPLING_FREQUENCY', 'SAMPLING_FREQUENCY', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Sampling frequency'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'PREVENTION_CONTROLS,DETECTION_CONTROLS,CONTROL_METHOD_OVERRIDE', 'CONTROL_METHOD', APP_CONSTANTS.TRANSFORMS.CONTROL_METHOD, 'LOG', 'Derived from PFMEA controls'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'REACTION_PLAN', 'REACTION_PLAN', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Reaction plan'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'ACTIVE', 'STATUS', APP_CONSTANTS.TRANSFORMS.STATUS_FROM_ACTIVE, 'LOG', 'Flag inactive rows instead of deleting'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'WI_DOC_ID', 'WORK_INSTRUCTION_DOC_ID', APP_CONSTANTS.TRANSFORMS.DIRECT, 'SKIP', 'Optional document override'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN, 'STEP_ID', 'WORK_INSTRUCTION_STEP_TAG', APP_CONSTANTS.TRANSFORMS.STEP_TAG, 'LOG', 'Deterministic Work Instruction section tag'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'PROCESS_STEP,OPERATION_NO', 'STEP_TITLE', APP_CONSTANTS.TRANSFORMS.STEP_TITLE, 'LOG', 'Section heading'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'OPERATION_NO', 'OPERATION_NO', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Operation number'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'PROCESS_FUNCTION_REQUIREMENT', 'PROCESS_DESCRIPTION', APP_CONSTANTS.TRANSFORMS.DIRECT, 'LOG', 'Process requirement'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'FAILURE_MODE,FAILURE_EFFECT,CAUSE_MECHANISM', 'FAILURE_SUMMARY', APP_CONSTANTS.TRANSFORMS.FAILURE_SUMMARY, 'LOG', 'Step risk summary'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'PRODUCT_CHARACTERISTICS', 'PRODUCT_CHARACTERISTICS', APP_CONSTANTS.TRANSFORMS.AGGREGATE_UNIQUE, 'LOG', 'Aggregate product characteristics'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'PROCESS_CHARACTERISTICS', 'PROCESS_CHARACTERISTICS', APP_CONSTANTS.TRANSFORMS.AGGREGATE_UNIQUE, 'LOG', 'Aggregate process characteristics'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'SPECIAL_CHARACTERISTICS', 'SPECIAL_CHARACTERISTICS', APP_CONSTANTS.TRANSFORMS.AGGREGATE_UNIQUE, 'LOG', 'Aggregate special characteristics'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'SPECIFICATION_TOLERANCE', 'SPECIFICATION_TOLERANCE', APP_CONSTANTS.TRANSFORMS.AGGREGATE_UNIQUE, 'LOG', 'Aggregate specification limits'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'PREVENTION_CONTROLS', 'PREVENTION_CONTROLS', APP_CONSTANTS.TRANSFORMS.AGGREGATE_UNIQUE, 'LOG', 'Aggregate prevention controls'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'DETECTION_CONTROLS', 'DETECTION_CONTROLS', APP_CONSTANTS.TRANSFORMS.AGGREGATE_UNIQUE, 'LOG', 'Aggregate detection controls'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'PREVENTION_CONTROLS,DETECTION_CONTROLS,CONTROL_METHOD_OVERRIDE', 'CONTROL_METHOD', APP_CONSTANTS.TRANSFORMS.CONTROL_METHOD, 'LOG', 'Derived control method'),
      createMapping_('TRUE', APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION, 'REACTION_PLAN', 'REACTION_PLAN', APP_CONSTANTS.TRANSFORMS.AGGREGATE_UNIQUE, 'LOG', 'Aggregate reaction plan')
    ];
  }

  function createMapping_(active, targetType, sourceColumns, targetField, transform, onMissing, notes) {
    return {
      ACTIVE: active,
      TARGET_TYPE: targetType,
      SOURCE_COLUMNS: sourceColumns,
      TARGET_FIELD: targetField,
      TRANSFORM: transform,
      ON_MISSING: onMissing,
      NOTES: notes
    };
  }

  function loadMappings(config, includeInactive) {
    var records = SheetsService.getRecords(config.MAPPING_SHEET, config.MAPPING_HEADER_ROW);
    return records
      .map(function(record) {
        return record.values;
      })
      .filter(function(mapping) {
        return includeInactive || SyncUtils.toBoolean(mapping.ACTIVE);
      });
  }

  function buildControlPlanPayload(pfmeaRecord, mappings) {
    return buildPayload_([pfmeaRecord], filterMappings_(mappings, APP_CONSTANTS.TARGET_TYPES.CONTROL_PLAN));
  }

  function buildWorkInstructionPayload(stepRecords, mappings) {
    return buildPayload_(stepRecords, filterMappings_(mappings, APP_CONSTANTS.TARGET_TYPES.WORK_INSTRUCTION));
  }

  function filterMappings_(mappings, targetType) {
    return mappings.filter(function(mapping) {
      return SyncUtils.asString(mapping.TARGET_TYPE) === targetType;
    });
  }

  function buildPayload_(records, mappings) {
    var payload = {};
    mappings.forEach(function(mapping) {
      payload[mapping.TARGET_FIELD] = resolveMappingValue_(mapping, records);
    });
    return payload;
  }

  function resolveMappingValue_(mapping, records) {
    var transform = SyncUtils.asString(mapping.TRANSFORM) || APP_CONSTANTS.TRANSFORMS.DIRECT;
    var sourceColumns = parseSourceColumns_(mapping.SOURCE_COLUMNS);

    switch (transform) {
      case APP_CONSTANTS.TRANSFORMS.CONTROL_METHOD:
        return buildControlMethod_(records);
      case APP_CONSTANTS.TRANSFORMS.STATUS_FROM_ACTIVE:
        return hasActiveRecord_(records) ? APP_CONSTANTS.STATUS.ACTIVE : APP_CONSTANTS.STATUS.FLAGGED_INACTIVE;
      case APP_CONSTANTS.TRANSFORMS.STEP_TITLE:
        return buildStepTitle_(records);
      case APP_CONSTANTS.TRANSFORMS.AGGREGATE_UNIQUE:
        return aggregateUnique_(records, sourceColumns).join(', ');
      case APP_CONSTANTS.TRANSFORMS.FAILURE_SUMMARY:
        return buildFailureSummary_(records);
      case APP_CONSTANTS.TRANSFORMS.STEP_TAG:
        return getPrimaryRecord_(records).STEP_ID || '';
      case APP_CONSTANTS.TRANSFORMS.DIRECT:
      default:
        return resolveDirectValue_(records, sourceColumns);
    }
  }

  function parseSourceColumns_(sourceColumns) {
    return SyncUtils.asString(sourceColumns)
      .split(',')
      .map(function(column) {
        return column.trim();
      })
      .filter(function(column) {
        return !!column;
      });
  }

  function resolveDirectValue_(records, sourceColumns) {
    var primaryRecord = getPrimaryRecord_(records);
    for (var sourceIndex = 0; sourceIndex < sourceColumns.length; sourceIndex += 1) {
      var value = primaryRecord[sourceColumns[sourceIndex]];
      if (!SyncUtils.isBlank(value)) {
        return value;
      }
    }
    return sourceColumns.length ? primaryRecord[sourceColumns[0]] || '' : '';
  }

  function aggregateUnique_(records, sourceColumns) {
    var values = [];
    records.forEach(function(record) {
      sourceColumns.forEach(function(sourceColumn) {
        values = values.concat(SyncUtils.normalizeList(record[sourceColumn]));
      });
    });
    return SyncUtils.unique(values);
  }

  function buildControlMethod_(records) {
    var primaryRecord = getPrimaryRecord_(records);
    var overrideValue = SyncUtils.asString(primaryRecord.CONTROL_METHOD_OVERRIDE);
    if (overrideValue) {
      return overrideValue;
    }

    var prevention = aggregateUnique_(records, ['PREVENTION_CONTROLS']);
    var detection = aggregateUnique_(records, ['DETECTION_CONTROLS']);
    var parts = [];
    if (prevention.length) {
      parts.push('Prevention: ' + prevention.join('; '));
    }
    if (detection.length) {
      parts.push('Detection: ' + detection.join('; '));
    }
    return parts.join(' | ');
  }

  function buildStepTitle_(records) {
    var primaryRecord = getPrimaryRecord_(records);
    var operationNo = SyncUtils.asString(primaryRecord.OPERATION_NO);
    var processStep = SyncUtils.asString(primaryRecord.PROCESS_STEP);
    if (operationNo && processStep) {
      return operationNo + ' - ' + processStep;
    }
    return operationNo || processStep || primaryRecord.STEP_ID || 'Unnamed Step';
  }

  function buildFailureSummary_(records) {
    var lines = [];
    records.forEach(function(record) {
      var pieces = [];
      if (!SyncUtils.isBlank(record.FAILURE_MODE)) {
        pieces.push('Mode: ' + record.FAILURE_MODE);
      }
      if (!SyncUtils.isBlank(record.FAILURE_EFFECT)) {
        pieces.push('Effect: ' + record.FAILURE_EFFECT);
      }
      if (!SyncUtils.isBlank(record.CAUSE_MECHANISM)) {
        pieces.push('Cause: ' + record.CAUSE_MECHANISM);
      }
      if (pieces.length) {
        lines.push(pieces.join(' | '));
      }
    });
    return SyncUtils.unique(lines).join('\n');
  }

  function hasActiveRecord_(records) {
    return records.some(function(record) {
      return isActiveRecord_(record);
    });
  }

  function getPrimaryRecord_(records) {
    for (var index = 0; index < records.length; index += 1) {
      if (isActiveRecord_(records[index])) {
        return records[index];
      }
    }
    return records[0] || {};
  }

  function isActiveRecord_(record) {
    var activeValue = SyncUtils.asString(record.ACTIVE).toUpperCase();
    return activeValue === '' || activeValue === 'TRUE' || activeValue === 'YES' || activeValue === 'Y' || activeValue === '1';
  }

  return {
    getDefaultMappings: getDefaultMappings,
    loadMappings: loadMappings,
    buildControlPlanPayload: buildControlPlanPayload,
    buildWorkInstructionPayload: buildWorkInstructionPayload
  };
})();
