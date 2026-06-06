const CONFIG = {
  LOCK_TIMEOUT_MS: 30000,
  TARGET_SHEET_PREFIX: '=',
  PROPERTY_PREFIX: 'lastValue_',
  CHANGE_MARKER_CELL: 'G4',
  FILTER_RANGE: 'B4:B',
  FILTER_COLUMN: 2, // B
  FIRST_DATA_ROW: 5,
  FIRST_BORDER_CLEANUP_ROW: 6,
  MAIN_COLUMN: 6, // F
  MERGE_COLUMNS: [1, 3, 4, 5, 6, 38], // A, C, D, E, F, AL
  COPY_TEMP_PREFIX: '__tmp_values_copy__',
  PENDING_RUN_PROPERTY: 'processSheetsPendingRun',
  QUEUED_TRIGGER_HANDLER: 'processSheetsQueuedRetry',
  QUEUED_TRIGGER_DELAY_MS: 60000,
};

function processSheets() {
  runProcessSheets_();
}

/**
 * Entry point for a delayed retry created when an edit trigger fires while
 * another processSheets run is still working.
 */
function processSheetsQueuedRetry() {
  deleteQueuedProcessSheetsTriggers();
  runProcessSheets_();
}

/**
 * Runs sheet processing under a lock.
 *
 * If another trigger already holds the lock, this function records that one
 * more run is needed and schedules a single delayed retry instead of throwing
 * a lock-timeout error.
 */
function runProcessSheets_() {
  const lock = LockService.getScriptLock();
  const properties = PropertiesService.getScriptProperties();
  let isLockAcquired = false;

  try {
    isLockAcquired = lock.tryLock(CONFIG.LOCK_TIMEOUT_MS);

    if (!isLockAcquired) {
      queueProcessSheetsRetry(properties);
      Logger.log('Пропуск запуска processSheets: другой запуск ещё обрабатывает таблицу. Повтор поставлен в очередь.');
      return;
    }

    properties.deleteProperty(CONFIG.PENDING_RUN_PROPERTY);

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();

    removeStaleTemporaryCopySheets(spreadsheet);
    removeStaleSheetProperties(properties, sheets);
    logTrackedSheetProperties(properties);

    sheets
      .filter(isTargetSheet)
      .forEach(sheet => processSheetIfChanged(sheet, properties, spreadsheet));

    Logger.log('=== Все листы обработаны ===');
  } catch (error) {
    Logger.log(`Ошибка при обработке листов: ${error && error.stack ? error.stack : error}`);
  } finally {
    if (isLockAcquired) {
      if (properties.getProperty(CONFIG.PENDING_RUN_PROPERTY) === '1') {
        ensureQueuedProcessSheetsTrigger();
        Logger.log('Во время обработки пришёл ещё один запуск. Дополнительная обработка поставлена в очередь.');
      }

      lock.releaseLock();
    }
  }
}

/**
 * Marks that a retry is needed and creates a delayed retry trigger when needed.
 *
 * @param {GoogleAppsScript.Properties.Properties} properties Script properties.
 */
function queueProcessSheetsRetry(properties) {
  properties.setProperty(CONFIG.PENDING_RUN_PROPERTY, '1');
  ensureQueuedProcessSheetsTrigger();
}

/**
 * Creates one delayed retry trigger if it does not already exist.
 */
function ensureQueuedProcessSheetsTrigger() {
  const hasQueuedTrigger = ScriptApp.getProjectTriggers()
    .some(trigger => trigger.getHandlerFunction() === CONFIG.QUEUED_TRIGGER_HANDLER);

  if (hasQueuedTrigger) {
    return;
  }

  ScriptApp.newTrigger(CONFIG.QUEUED_TRIGGER_HANDLER)
    .timeBased()
    .after(CONFIG.QUEUED_TRIGGER_DELAY_MS)
    .create();
}

/**
 * Deletes queued retry triggers for this script.
 */
function deleteQueuedProcessSheetsTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === CONFIG.QUEUED_TRIGGER_HANDLER)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Deletes temporary copy sheets that could remain after an interrupted previous run.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Active spreadsheet.
 */
function removeStaleTemporaryCopySheets(spreadsheet) {
  spreadsheet.getSheets()
    .filter(sheet => sheet.getName().startsWith(CONFIG.COPY_TEMP_PREFIX))
    .forEach(sheet => spreadsheet.deleteSheet(sheet));
}

/**
 * Removes cached values for sheets that no longer exist.
 *
 * @param {GoogleAppsScript.Properties.Properties} properties Script properties.
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} sheets Spreadsheet sheets.
 */
function removeStaleSheetProperties(properties, sheets) {
  const existingSheetNames = new Set(sheets.map(sheet => sheet.getName()));
  const allProperties = properties.getProperties();

  Object.keys(allProperties)
    .filter(key => key.startsWith(CONFIG.PROPERTY_PREFIX))
    .forEach(key => {
      const sheetName = key.slice(CONFIG.PROPERTY_PREFIX.length);

      if (!existingSheetNames.has(sheetName)) {
        properties.deleteProperty(key);
        Logger.log(`Удалён ключ "${key}", так как лист "${sheetName}" больше не существует.`);
      }
    });
}

/**
 * Writes the currently tracked cached values to the Apps Script log.
 *
 * @param {GoogleAppsScript.Properties.Properties} properties Script properties.
 */
function logTrackedSheetProperties(properties) {
  const trackedProperties = Object.entries(properties.getProperties())
    .filter(([key]) => key.startsWith(CONFIG.PROPERTY_PREFIX));

  if (trackedProperties.length === 0) {
    Logger.log(`В PropertiesService не осталось ключей ${CONFIG.PROPERTY_PREFIX}.`);
    return;
  }

  Logger.log('Актуальные ключи в PropertiesService:');
  trackedProperties.forEach(([key, value]) => Logger.log(` • ${key} = ${value}`));
}

/**
 * Checks whether the sheet should be processed by this script.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 * @returns {boolean} True when this is a target sheet.
 */
function isTargetSheet(sheet) {
  return sheet.getName().startsWith(CONFIG.TARGET_SHEET_PREFIX);
}

/**
 * Processes a sheet only if the value in the marker cell has changed.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 * @param {GoogleAppsScript.Properties.Properties} properties Script properties.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Active spreadsheet.
 */
function processSheetIfChanged(sheet, properties, spreadsheet) {
  const sheetName = sheet.getName();
  const currentValue = String(sheet.getRange(CONFIG.CHANGE_MARKER_CELL).getValue());
  const propertyKey = `${CONFIG.PROPERTY_PREFIX}${sheetName}`;
  const storedValue = properties.getProperty(propertyKey);

  if (currentValue === String(storedValue)) {
    Logger.log(`Пропуск листа "${sheetName}" — значение ${CONFIG.CHANGE_MARKER_CELL} не изменилось (${currentValue}).`);
    return;
  }

  Logger.log(`Обработка листа "${sheetName}". Старое значение: ${storedValue}, новое: ${currentValue}`);

  resetFilter(sheet);
  prepareDataRange(sheet);
  mergeGroupsByMainColumn(sheet);
  applyNotEmptyFilter(sheet);
  replaceValuesOnlyCopy(sheet, spreadsheet);

  properties.setProperty(propertyKey, currentValue);
  Logger.log(`Обработка листа "${sheetName}" завершена.`);
}

/**
 * Creates the filter when needed and clears criteria in the configured column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 */
function resetFilter(sheet) {
  const filter = getOrCreateFilter(sheet);

  filter.setColumnFilterCriteria(
    CONFIG.FILTER_COLUMN,
    SpreadsheetApp.newFilterCriteria().build()
  );
}

/**
 * Returns an existing sheet filter or creates a new one.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Filter} Sheet filter.
 */
function getOrCreateFilter(sheet) {
  const existingFilter = sheet.getFilter();

  if (existingFilter) {
    return existingFilter;
  }

  sheet.getRange(CONFIG.FILTER_RANGE).createFilter();
  return sheet.getFilter();
}

/**
 * Removes previous merges and borders from the data area.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 */
function prepareDataRange(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < CONFIG.FIRST_DATA_ROW || lastColumn === 0) {
    return;
  }

  const dataRowCount = lastRow - CONFIG.FIRST_DATA_ROW + 1;
  sheet.getRange(CONFIG.FIRST_DATA_ROW, 1, dataRowCount, lastColumn).breakApart();

  if (lastRow < CONFIG.FIRST_BORDER_CLEANUP_ROW) {
    return;
  }

  const borderCleanupRowCount = lastRow - CONFIG.FIRST_BORDER_CLEANUP_ROW + 1;
  sheet
    .getRange(CONFIG.FIRST_BORDER_CLEANUP_ROW, 1, borderCleanupRowCount, lastColumn)
    .setBorder(false, false, false, false, false, false);
}

/**
 * Merges configured columns for adjacent rows with equal values in the main column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 */
function mergeGroupsByMainColumn(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < CONFIG.FIRST_DATA_ROW || lastColumn === 0) {
    return;
  }

  const rowCount = lastRow - CONFIG.FIRST_DATA_ROW + 1;
  const mainColumnValues = sheet
    .getRange(CONFIG.FIRST_DATA_ROW, CONFIG.MAIN_COLUMN, rowCount, 1)
    .getValues()
    .flat();

  let groupStartRow = CONFIG.FIRST_DATA_ROW;

  for (let index = 1; index <= mainColumnValues.length; index += 1) {
    const previousValue = mainColumnValues[index - 1];
    const currentValue = mainColumnValues[index];
    const isLastValue = index === mainColumnValues.length;
    const isGroupEnd = isLastValue || currentValue !== previousValue;

    if (!isGroupEnd) {
      continue;
    }

    const groupEndRow = CONFIG.FIRST_DATA_ROW + index - 1;
    mergeGroupRows(sheet, groupStartRow, groupEndRow);
    addGroupTopBorder(sheet, groupStartRow, lastColumn);
    groupStartRow = groupEndRow + 1;
  }
}

/**
 * Merges cells vertically for one group in all configured columns.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 * @param {number} groupStartRow First row in the group.
 * @param {number} groupEndRow Last row in the group.
 */
function mergeGroupRows(sheet, groupStartRow, groupEndRow) {
  const groupRowCount = groupEndRow - groupStartRow + 1;

  if (groupRowCount < 2) {
    return;
  }

  CONFIG.MERGE_COLUMNS.forEach(column => {
    sheet.getRange(groupStartRow, column, groupRowCount, 1).mergeVertically();
  });
}

/**
 * Adds a dotted border above a group, except above the first data group.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 * @param {number} groupStartRow First row in the group.
 * @param {number} lastColumn Last used column in the sheet.
 */
function addGroupTopBorder(sheet, groupStartRow, lastColumn) {
  if (groupStartRow <= CONFIG.FIRST_DATA_ROW) {
    return;
  }

  sheet
    .getRange(groupStartRow, 1, 1, lastColumn)
    .setBorder(true, null, null, null, null, null, true, SpreadsheetApp.BorderStyle.DOTTED);
}

/**
 * Applies the final "not empty" filter criterion.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 */
function applyNotEmptyFilter(sheet) {
  getOrCreateFilter(sheet).setColumnFilterCriteria(
    CONFIG.FILTER_COLUMN,
    SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build()
  );
}

/**
 * Updates the values-only public copy of a processed sheet.
 *
 * The existing public copy is not deleted. A hidden temporary copy is prepared
 * first, formulas are replaced by values there, and only then the prepared
 * sheet state is transferred into the target sheet in one short hidden update.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet Processed source sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Active spreadsheet.
 */
function replaceValuesOnlyCopy(sourceSheet, spreadsheet) {
  const copyName = getValuesOnlyCopyName(sourceSheet.getName());

  if (!copyName) {
    Logger.log(`Пропуск создания копии для листа "${sourceSheet.getName()}" — пустое имя копии.`);
    return;
  }

  SpreadsheetApp.flush();

  const temporarySheet = createPreparedTemporaryCopy(sourceSheet, spreadsheet);
  let targetSheet = spreadsheet.getSheetByName(copyName);

  temporarySheet.hideSheet();
  freezeFormulasAsValues(temporarySheet);
  SpreadsheetApp.flush();

  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet(copyName);
  }

  try {
    updateSheetFromPreparedCopy(temporarySheet, targetSheet, sourceSheet, spreadsheet);
  } finally {
    spreadsheet.deleteSheet(temporarySheet);
    spreadsheet.setActiveSheet(sourceSheet);
  }

  Logger.log(`Копия листа "${sourceSheet.getName()}" обновлена листом "${copyName}" без формул.`);
}


/**
 * Transfers the prepared temporary sheet into the existing target sheet.
 *
 * The target sheet is hidden during the destructive part of the update, so the
 * user does not see clearing, resizing, copying and filter restoration steps.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} preparedSheet Hidden prepared values-only sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet Existing or newly created public copy sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} fallbackActiveSheet Sheet to activate while the target is hidden.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Active spreadsheet.
 */
function updateSheetFromPreparedCopy(preparedSheet, targetSheet, fallbackActiveSheet, spreadsheet) {
  const shouldRestoreTargetVisibility = !targetSheet.isSheetHidden();

  spreadsheet.setActiveSheet(fallbackActiveSheet);
  targetSheet.hideSheet();

  try {
    clearTargetSheetForSnapshot(targetSheet);
    resizeSheetLikePreparedCopy(targetSheet, preparedSheet);
    copyPreparedSheetContents(preparedSheet, targetSheet);
    copySheetDimensions(preparedSheet, targetSheet);
    copySheetViewSettings(preparedSheet, targetSheet);
    restorePreparedSheetFilter(preparedSheet, targetSheet);
    SpreadsheetApp.flush();
  } finally {
    if (shouldRestoreTargetVisibility) {
      targetSheet.showSheet();
    }
  }
}

/**
 * Clears previous target state before copying a fresh snapshot.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet Target sheet.
 */
function clearTargetSheetForSnapshot(targetSheet) {
  const targetFilter = targetSheet.getFilter();

  if (targetFilter) {
    targetFilter.remove();
  }

  targetSheet
    .getRange(1, 1, targetSheet.getMaxRows(), targetSheet.getMaxColumns())
    .breakApart();
  targetSheet.clear();
}

/**
 * Makes target row and column counts match the prepared sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet Target sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} preparedSheet Prepared sheet.
 */
function resizeSheetLikePreparedCopy(targetSheet, preparedSheet) {
  resizeRows(targetSheet, preparedSheet.getMaxRows());
  resizeColumns(targetSheet, preparedSheet.getMaxColumns());
}

/**
 * Changes the target row count to the requested value.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to resize.
 * @param {number} expectedRows Required row count.
 */
function resizeRows(sheet, expectedRows) {
  const currentRows = sheet.getMaxRows();

  if (currentRows < expectedRows) {
    sheet.insertRowsAfter(currentRows, expectedRows - currentRows);
    return;
  }

  if (currentRows > expectedRows) {
    sheet.deleteRows(expectedRows + 1, currentRows - expectedRows);
  }
}

/**
 * Changes the target column count to the requested value.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to resize.
 * @param {number} expectedColumns Required column count.
 */
function resizeColumns(sheet, expectedColumns) {
  const currentColumns = sheet.getMaxColumns();

  if (currentColumns < expectedColumns) {
    sheet.insertColumnsAfter(currentColumns, expectedColumns - currentColumns);
    return;
  }

  if (currentColumns > expectedColumns) {
    sheet.deleteColumns(expectedColumns + 1, currentColumns - expectedColumns);
  }
}

/**
 * Copies values-only cell content, formatting, notes, merged cells and validations.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} preparedSheet Prepared sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet Target sheet.
 */
function copyPreparedSheetContents(preparedSheet, targetSheet) {
  preparedSheet
    .getRange(1, 1, preparedSheet.getMaxRows(), preparedSheet.getMaxColumns())
    .copyTo(targetSheet.getRange(1, 1));
}

/**
 * Copies column widths and row heights from the prepared sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} preparedSheet Prepared sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet Target sheet.
 */
function copySheetDimensions(preparedSheet, targetSheet) {
  for (let column = 1; column <= preparedSheet.getMaxColumns(); column += 1) {
    targetSheet.setColumnWidth(column, preparedSheet.getColumnWidth(column));
  }

  for (let row = 1; row <= preparedSheet.getMaxRows(); row += 1) {
    targetSheet.setRowHeight(row, preparedSheet.getRowHeight(row));
  }
}

/**
 * Copies high-level sheet view settings from the prepared sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} preparedSheet Prepared sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet Target sheet.
 */
function copySheetViewSettings(preparedSheet, targetSheet) {
  targetSheet.setFrozenRows(preparedSheet.getFrozenRows());
  targetSheet.setFrozenColumns(preparedSheet.getFrozenColumns());
  targetSheet.setRightToLeft(preparedSheet.isRightToLeft());

  const tabColor = preparedSheet.getTabColorObject();
  targetSheet.setTabColorObject(tabColor || null);
}

/**
 * Recreates the prepared sheet filter range and criteria on the target sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} preparedSheet Prepared sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet Target sheet.
 */
function restorePreparedSheetFilter(preparedSheet, targetSheet) {
  const preparedFilter = preparedSheet.getFilter();

  if (!preparedFilter) {
    return;
  }

  const preparedFilterRange = preparedFilter.getRange();
  const targetFilterRange = targetSheet.getRange(
    preparedFilterRange.getRow(),
    preparedFilterRange.getColumn(),
    preparedFilterRange.getNumRows(),
    preparedFilterRange.getNumColumns()
  );

  targetFilterRange.createFilter();

  const targetFilter = targetSheet.getFilter();
  const filterStartColumn = preparedFilterRange.getColumn();
  const filterEndColumn = filterStartColumn + preparedFilterRange.getNumColumns() - 1;

  for (let column = filterStartColumn; column <= filterEndColumn; column += 1) {
    const criteria = preparedFilter.getColumnFilterCriteria(column);

    if (criteria) {
      targetFilter.setColumnFilterCriteria(column, criteria.copy().build());
    }
  }
}

/**
 * Builds the values-only copy sheet name by removing the target prefix.
 *
 * @param {string} sourceSheetName Source sheet name.
 * @returns {string} Copy sheet name.
 */
function getValuesOnlyCopyName(sourceSheetName) {
  return sourceSheetName.replace(new RegExp(`^${escapeRegExp(CONFIG.TARGET_SHEET_PREFIX)}+`), '').trim();
}

/**
 * Creates a temporary sheet copy and removes formulas from it before the final swap.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet Processed source sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Active spreadsheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Prepared temporary sheet.
 */
function createPreparedTemporaryCopy(sourceSheet, spreadsheet) {
  const temporarySheet = sourceSheet.copyTo(spreadsheet);
  temporarySheet.setName(makeUniqueTemporarySheetName(spreadsheet));
  return temporarySheet;
}

/**
 * Replaces every formula in the sheet with its currently calculated value.
 * Formatting, merged cells, borders, filters and dimensions stay from the copied sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to freeze.
 */
function freezeFormulasAsValues(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) {
    return;
  }

  const dataRange = sheet.getRange(1, 1, lastRow, lastColumn);
  dataRange.copyTo(dataRange, {contentsOnly: true});
}

/**
 * Builds a temporary sheet name that does not collide with existing sheets.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Active spreadsheet.
 * @returns {string} Unique temporary sheet name.
 */
function makeUniqueTemporarySheetName(spreadsheet) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmssSSS');
  let index = 0;
  let temporaryName = `${CONFIG.COPY_TEMP_PREFIX}${timestamp}`;

  while (spreadsheet.getSheetByName(temporaryName)) {
    index += 1;
    temporaryName = `${CONFIG.COPY_TEMP_PREFIX}${timestamp}_${index}`;
  }

  return temporaryName;
}

/**
 * Escapes a string for safe usage inside RegExp.
 *
 * @param {string} value Raw string.
 * @returns {string} RegExp-safe string.
 */
function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
