# PFMEA Sync System

## 1. Short system architecture explanation
PFMEA Sync System is a Google Workspace-native MVP built as a bound Google Apps Script project inside the master Google Sheet. The `PFMEA` tab is the source of truth. An installable edit trigger or a manual menu action reads PFMEA rows, applies the mapping rules stored in the `MAPPING` tab, updates linked rows in `CONTROL_PLAN`, updates deterministic step sections in Google Docs Work Instructions, and records every action in `CHANGE_LOG`.

This is deliberately a practical manufacturing-first design:

- Google Sheets stores PFMEA, Control Plan, mapping rules, logs, and settings.
- Google Docs stores Work Instructions.
- Google Drive stores the Work Instruction docs and backup copies.
- Google Apps Script orchestrates validation, sync, logging, and trigger handling.

## 2. Data model explanation
The solution uses a two-key strategy:

- `PFMEA_ROW_ID`: stable immutable row key. This is the primary link between `PFMEA` and `CONTROL_PLAN`.
- `STEP_ID`: stable process-step key. This is the primary link for Work Instruction sections inside Google Docs.

Recommended business meaning:

- One PFMEA row = one process risk or control record.
- One Control Plan row = one downstream row linked back to exactly one `PFMEA_ROW_ID`.
- One Work Instruction section = one process step (`STEP_ID`), which can aggregate multiple PFMEA rows for the same step.

## 3. Assumed Google Sheets column structure for PFMEA and CONTROL_PLAN
### PFMEA
- `PFMEA_ROW_ID`
- `STEP_ID`
- `CHARACTERISTIC_ID`
- `OPERATION_NO`
- `PROCESS_STEP`
- `PROCESS_FUNCTION_REQUIREMENT`
- `FAILURE_MODE`
- `FAILURE_EFFECT`
- `CAUSE_MECHANISM`
- `PREVENTION_CONTROLS`
- `DETECTION_CONTROLS`
- `SPECIAL_CHARACTERISTICS`
- `PRODUCT_CHARACTERISTICS`
- `PROCESS_CHARACTERISTICS`
- `SPECIFICATION_TOLERANCE`
- `EVALUATION_MEASUREMENT_TECHNIQUE`
- `SAMPLE_SIZE`
- `SAMPLING_FREQUENCY`
- `CONTROL_METHOD_OVERRIDE`
- `REACTION_PLAN`
- `WI_DOC_ID`
- `ACTIVE`
- `OWNER`
- `LAST_REVIEW_DATE`
- `NOTES`

### CONTROL_PLAN
- `CONTROL_PLAN_ROW_ID`
- `PFMEA_ROW_ID`
- `STEP_ID`
- `CHARACTERISTIC_ID`
- `OPERATION_NO`
- `PROCESS_STEP`
- `PRODUCT_CHARACTERISTICS`
- `PROCESS_CHARACTERISTICS`
- `SPECIAL_CHARACTERISTICS`
- `SPECIFICATION_TOLERANCE`
- `EVALUATION_TECHNIQUE`
- `SAMPLE_SIZE`
- `SAMPLING_FREQUENCY`
- `CONTROL_METHOD`
- `REACTION_PLAN`
- `WORK_INSTRUCTION_DOC_ID`
- `WORK_INSTRUCTION_STEP_TAG`
- `STATUS`
- `LAST_SYNC_AT`
- `LAST_SYNC_BY`
- `NOTES`

## 4. Mapping logic explanation
The `MAPPING` tab drives downstream behavior with these columns:

- `ACTIVE`
- `TARGET_TYPE`
- `SOURCE_COLUMNS`
- `TARGET_FIELD`
- `TRANSFORM`
- `ON_MISSING`
- `NOTES`

Supported target types:

- `CONTROL_PLAN`
- `WORK_INSTRUCTION`

Supported transform rules in this MVP:

- `DIRECT`
- `CONTROL_METHOD`
- `STATUS_FROM_ACTIVE`
- `STEP_TITLE`
- `AGGREGATE_UNIQUE`
- `FAILURE_SUMMARY`
- `STEP_TAG`

Examples already seeded by setup:

- PFMEA `PREVENTION_CONTROLS + DETECTION_CONTROLS` -> Control Plan `CONTROL_METHOD`
- PFMEA `PROCESS_STEP` -> Control Plan `PROCESS_STEP`
- PFMEA `REACTION_PLAN` -> Control Plan `REACTION_PLAN`
- PFMEA step group -> Work Instruction `FAILURE_SUMMARY`
- PFMEA inactive row -> downstream `FLAGGED_INACTIVE`

## 5. Work Instruction tagging strategy
The Google Docs strategy is deterministic and marker-based:

```text
[[STEP_START:STEP-OP10]]
... system-managed section content ...
[[STEP_END:STEP-OP10]]
```

Rules:

- All manual content outside the markers is preserved.
- Only the content between the matching markers is rewritten.
- If a section is missing and `CREATE_MISSING_WI_SECTION=TRUE`, it is appended automatically.
- If a PFMEA step has no `WI_DOC_ID` and `CREATE_MISSING_WI_DOCS=TRUE`, the system creates a dedicated Google Doc for that step and writes the new `WI_DOC_ID` back to PFMEA.
- If `WI_TEMPLATE_DOC_ID` is configured, new Work Instructions are copied from that template before the managed section is inserted.
- If a section contains `[[LOCKED:STEP-OP10]]`, it is skipped and logged.
- Work Instruction content is aggregated at `STEP_ID` level, which is more realistic for automotive process instructions than one PFMEA row per document section.

## 6. Full project file list
- `Code.gs`
- `Config.gs`
- `SyncService.gs`
- `MappingService.gs`
- `SheetsService.gs`
- `DocsService.gs`
- `LogService.gs`
- `UI.gs`
- `Validation.gs`
- `appsscript.json`
- `README.md`

## 7. Full code for every Apps Script file
The full runnable code is in the repository files listed above. Paste each `.gs` file and the manifest into a bound Apps Script project attached to the master spreadsheet.

## 8. appsscript.json
The manifest is included in `appsscript.json` and already contains the scopes needed for:

- Spreadsheet access
- Google Docs access
- Google Drive backup operations
- installable trigger creation
- user email capture for audit logs

## 9. Setup instructions
1. Open the master Google Sheet that will own the PFMEA sync process.
2. Open `Extensions -> Apps Script`.
3. Create Apps Script files matching the repository names and paste in the code.
4. Replace the manifest with `appsscript.json`.
5. Save the script project and reload the spreadsheet.
6. From the custom menu `PFMEA Sync System`, run `Setup System`.
7. Authorize the required Google permissions.
8. Review the seeded `PFMEA`, `CONTROL_PLAN`, `MAPPING`, `CHANGE_LOG`, and `CONFIG` tabs.
9. Adjust config values as needed:
   - `DEFAULT_WI_DOC_ID`
   - `CREATE_MISSING_WI_DOCS`
   - `WI_TEMPLATE_DOC_ID`
   - `WI_FOLDER_ID`
   - `BACKUP_FOLDER_ID`
   - `SYNC_MODE`
   - `DRY_RUN_MODE`
   - `ALLOW_OVERWRITE`
10. Replace or extend the default mapping rules to match your plant naming and downstream logic.

## 10. How to deploy/use inside Google Sheets
### Menu actions
- `Setup System`: creates missing tabs, seeds defaults, creates a default Work Instruction doc, and ensures the installable edit trigger exists.
- `Run Full Sync`: processes all PFMEA rows.
- `Sync Selected PFMEA Row`: syncs only the selected PFMEA row.
- `Preview Changes`: dry-run for the selected row or the whole PFMEA list.
- `Validate Mapping`: validates required headers, mapping rows, docs, and trigger state.
- `Open Config`: jumps to the `CONFIG` tab.

### Trigger behavior
- The installable trigger runs `handleSpreadsheetEdit`.
- If `SYNC_MODE=AUTO`, PFMEA edits trigger downstream updates automatically.
- If `SYNC_MODE=MANUAL`, users sync via menu actions only.

## 11. Example workflow
1. A quality engineer edits PFMEA row `PFR-1234ABCD` for `STEP-OP10`.
2. The installable edit trigger detects the row change.
3. The system recalculates mapped values and updates the linked `CONTROL_PLAN` row with the same `PFMEA_ROW_ID`.
4. If the PFMEA step has no assigned `WI_DOC_ID` and auto-create is enabled, the system creates a dedicated Google Doc, stores that `WI_DOC_ID` back in PFMEA, and then updates the managed section.
5. The system opens the target Google Doc and updates the section between `[[STEP_START:STEP-OP10]]` and `[[STEP_END:STEP-OP10]]`.
6. `CHANGE_LOG` receives entries for Control Plan update, Work Instruction creation/assignment if applicable, Work Instruction section update, and the row-level sync summary.
7. If the PFMEA row is inactive, the downstream artifacts are flagged inactive instead of being deleted.

## Safety controls in the MVP
- dry-run preview mode
- overwrite protection through `ALLOW_OVERWRITE`
- bulk sync confirmation through `CONFIRM_BULK_SYNC`
- spreadsheet backup before full sync
- Work Instruction backup before document writes
- `[[LOCKED:STEP_ID]]` markers for protected document sections

## What I still need from you
The MVP is built, but to make it production-ready for your plant I still need your real operating details:

1. Actual PFMEA and Control Plan column names if they differ from the seeded schema.
2. Your final Drive and document structure:
   - one Work Instruction doc per line
   - one doc per product family
   - one master document with many step sections
3. Your preferred overwrite policy:
   - always overwrite mapped downstream fields
   - update only blanks
   - approval before write
4. Your final Work Instruction layout preference:
   - one section per operation
   - one section per station
   - one section per control characteristic group
5. Any plant-specific rules for reaction plans, customer special characteristics, or escalation wording.

## Suggested v2 improvements
- approval workflow before bulk release
- dashboard sheet for sync health and unresolved mappings
- multi-template document routing by product family or production line
- email alerts for failed or skipped critical sync events
- risk-based sync rules driven by severity, occurrence, detection, or special characteristic flags
