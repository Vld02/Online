const CONFIG = {
  LOCK_TIMEOUT_MS: 1000,
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
};

function processSheets() {
  const lock = LockService.getScriptLock();
  let isLockAcquired = false;

  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    isLockAcquired = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const properties = PropertiesService.getScriptProperties();

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
      lock.releaseLock();
    }
  }
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
 * Replaces the values-only public copy of a processed sheet.
 *
 * The old public copy remains in place while the temporary copy is being prepared.
 * After formulas are replaced by their calculated values, the script deletes the old
 * copy and renames the prepared sheet, keeping the visible replacement window short.
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
  const oldCopy = spreadsheet.getSheetByName(copyName);

  temporarySheet.hideSheet();
  freezeFormulasAsValues(temporarySheet);
  SpreadsheetApp.flush();

  if (oldCopy) {
    spreadsheet.deleteSheet(oldCopy);
  }

  temporarySheet.setName(copyName);
  temporarySheet.showSheet();
  spreadsheet.setActiveSheet(sourceSheet);

  Logger.log(`Копия листа "${sourceSheet.getName()}" заменена листом "${copyName}" без формул.`);
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
