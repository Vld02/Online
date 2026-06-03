function processSheets() {
  const lock = LockService.getScriptLock();
  let isLockAcquired = false;

  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    isLockAcquired = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const properties = PropertiesService.getScriptProperties();

    removeStaleSheetProperties(properties, sheets);
    logTrackedSheetProperties(properties);

    sheets
      .filter(isTargetSheet)
      .forEach(sheet => processSheetIfChanged(sheet, properties));

    Logger.log('=== Все листы обработаны ===');
  } catch (error) {
    Logger.log(`Ошибка при обработке листов: ${error && error.stack ? error.stack : error}`);
  } finally {
    if (isLockAcquired) {
      lock.releaseLock();
    }
  }
}

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
};

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
 */
function processSheetIfChanged(sheet, properties) {
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
  syncPlainCopySheet(sheet);

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
 * Creates or updates the plain copy for a processed target sheet.
 * The copy has the same name without the target prefix and keeps values and formatting only.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet Processed target sheet.
 */
function syncPlainCopySheet(sourceSheet) {
  const spreadsheet = sourceSheet.getParent();
  const sourceSheetName = sourceSheet.getName();
  const copySheetName = getPlainCopySheetName(sourceSheetName);

  if (!copySheetName) {
    Logger.log(`Копия для листа "${sourceSheetName}" не создана: имя без префикса пустое.`);
    return;
  }

  const existingCopySheet = spreadsheet.getSheetByName(copySheetName);
  const targetSheetIndex = existingCopySheet ? existingCopySheet.getIndex() : sourceSheet.getIndex() + 1;

  if (existingCopySheet) {
    spreadsheet.deleteSheet(existingCopySheet);
  }

  const copySheet = sourceSheet.copyTo(spreadsheet).setName(copySheetName);
  spreadsheet.setActiveSheet(copySheet);
  spreadsheet.moveActiveSheet(targetSheetIndex);

  removeSheetFilter(copySheet);
  replaceFormulasWithValues(copySheet);

  const actionText = existingCopySheet ? 'Обновлена' : 'Создана';
  Logger.log(`${actionText} копия листа "${sourceSheetName}" с именем "${copySheetName}".`);
}

/**
 * Returns the name of the plain copy for a target sheet.
 *
 * @param {string} sourceSheetName Source target sheet name.
 * @returns {string} Copy sheet name without the target prefix.
 */
function getPlainCopySheetName(sourceSheetName) {
  return sourceSheetName.slice(CONFIG.TARGET_SHEET_PREFIX.length).trim();
}


/**
 * Removes a filter from the destination sheet before replacing values.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet.
 */
function removeSheetFilter(sheet) {
  const filter = sheet.getFilter();

  if (filter) {
    filter.remove();
  }
}

/**
 * Replaces formulas in the copied sheet with their calculated values while preserving formatting.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Copied sheet.
 */
function replaceFormulasWithValues(sheet) {
  const dataRange = sheet.getDataRange();
  dataRange.copyTo(dataRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}
