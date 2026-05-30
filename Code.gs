/**
 * Rebuilds grouping/formatting on sheets whose names start with "=".
 *
 * The script only reprocesses a sheet when G4 changes. It is defensive about
 * short/empty sheets, creates a table-wide filter, avoids marking a sheet as
 * processed until all operations succeed, and cleans obsolete script
 * properties for deleted sheets.
 */
function processSheets() {
  const CONFIG = Object.freeze({
    sheetNamePrefix: '=',
    firstDataRow: 5,
    headerRow: 4,
    changeMarkerCell: 'G4',
    mainColumn: 6, // F
    mergeColumns: [1, 3, 4, 5, 6, 38], // A, C, D, E, F, AL
    filterColumn: 2, // B
    propertyPrefix: 'lastValue_',
    lockTimeoutMs: 30000,
  });

  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  let lockAcquired = false;

  try {
    lock.waitLock(CONFIG.lockTimeoutMs);
    lockAcquired = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const props = PropertiesService.getScriptProperties();

    removePropertiesForDeletedSheets_(props, sheets, CONFIG.propertyPrefix);

    sheets.forEach(sheet => {
      if (!sheet.getName().startsWith(CONFIG.sheetNamePrefix)) {
        return;
      }

      processSingleSheet_(sheet, props, CONFIG);
    });

    Logger.log('=== Все подходящие листы обработаны ===');
  } catch (error) {
    Logger.log('Ошибка processSheets: ' + formatError_(error));
    throw error;
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

/**
 * Processes one sheet if its change marker value differs from the stored value.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {GoogleAppsScript.Properties.Properties} props
 * @param {{changeMarkerCell:string,propertyPrefix:string,firstDataRow:number,headerRow:number,mainColumn:number,mergeColumns:number[],filterColumn:number}} config
 */
function processSingleSheet_(sheet, props, config) {
  const sheetName = sheet.getName();
  const propertyKey = config.propertyPrefix + sheetName;
  const currentMarker = normalizePropertyValue_(sheet.getRange(config.changeMarkerCell).getValue());
  const previousMarker = props.getProperty(propertyKey);

  if (currentMarker === previousMarker) {
    Logger.log('Пропуск листа "' + sheetName + '" — значение ' + config.changeMarkerCell + ' не изменилось.');
    return;
  }

  Logger.log(
    'Обработка листа "' + sheetName + '". Старое значение: ' + previousMarker + ', новое: ' + currentMarker
  );

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), Math.max.apply(null, config.mergeColumns));

  ensureFilter_(sheet, config.headerRow, lastColumn, config.filterColumn);
  clearFilterCriteria_(sheet, config.filterColumn);

  if (lastRow >= config.firstDataRow) {
    const dataRowCount = lastRow - config.firstDataRow + 1;
    sheet.getRange(config.firstDataRow, 1, dataRowCount, lastColumn).breakApart();

    if (lastRow > config.firstDataRow) {
      sheet
        .getRange(config.firstDataRow + 1, 1, lastRow - config.firstDataRow, lastColumn)
        .setBorder(false, false, false, false, false, false);
    }

    mergeGroupsByMainColumn_(sheet, config.firstDataRow, lastRow, lastColumn, config.mainColumn, config.mergeColumns);
  } else {
    Logger.log('Лист "' + sheetName + '" не содержит строк данных для группировки.');
  }

  applyNotEmptyFilter_(sheet, config.filterColumn);

  // Store the marker only after successful processing, so failed runs are retried.
  props.setProperty(propertyKey, currentMarker);
  Logger.log('Обработка листа "' + sheetName + '" завершена.');
}

/**
 * Deletes lastValue_* properties for sheets that no longer exist.
 * @param {GoogleAppsScript.Properties.Properties} props
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} sheets
 * @param {string} propertyPrefix
 */
function removePropertiesForDeletedSheets_(props, sheets, propertyPrefix) {
  const existingSheetNames = new Set(sheets.map(sheet => sheet.getName()));
  const allProps = props.getProperties();

  Object.keys(allProps)
    .filter(key => key.startsWith(propertyPrefix))
    .forEach(key => {
      const sheetName = key.slice(propertyPrefix.length);
      if (!existingSheetNames.has(sheetName)) {
        props.deleteProperty(key);
        Logger.log('Удалён ключ "' + key + '", так как лист "' + sheetName + '" больше не существует.');
      }
    });
}

/**
 * Ensures there is a filter covering the whole table from the header row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} headerRow
 * @param {number} lastColumn
 * @param {number} filterColumn
 */
function ensureFilter_(sheet, headerRow, lastColumn, filterColumn) {
  const maxRows = sheet.getMaxRows();
  const width = Math.max(lastColumn, filterColumn);
  const desiredRange = sheet.getRange(headerRow, 1, maxRows - headerRow + 1, width);
  const currentFilter = sheet.getFilter();

  if (!currentFilter) {
    desiredRange.createFilter();
    return;
  }

  const currentRange = currentFilter.getRange();
  const coversFilterColumn =
    currentRange.getColumn() <= filterColumn && currentRange.getLastColumn() >= filterColumn;
  const startsAtHeader = currentRange.getRow() === headerRow;

  if (!coversFilterColumn || !startsAtHeader) {
    currentFilter.remove();
    desiredRange.createFilter();
  }
}

/**
 * Clears criteria on the filter column without removing the filter itself.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} filterColumn
 */
function clearFilterCriteria_(sheet, filterColumn) {
  const filter = sheet.getFilter();
  if (filter) {
    filter.removeColumnFilterCriteria(filterColumn);
  }
}

/**
 * Applies the final "not empty" criterion to the filter column.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} filterColumn
 */
function applyNotEmptyFilter_(sheet, filterColumn) {
  const filter = sheet.getFilter();
  if (filter) {
    filter.setColumnFilterCriteria(
      filterColumn,
      SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build()
    );
  }
}

/**
 * Merges configured columns for consecutive groups with equal main-column values.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} firstDataRow
 * @param {number} lastRow
 * @param {number} lastColumn
 * @param {number} mainColumn
 * @param {number[]} mergeColumns
 */
function mergeGroupsByMainColumn_(sheet, firstDataRow, lastRow, lastColumn, mainColumn, mergeColumns) {
  const values = sheet
    .getRange(firstDataRow, mainColumn, lastRow - firstDataRow + 1, 1)
    .getValues()
    .map(row => normalizeGroupValue_(row[0]));

  let groupStartRow = firstDataRow;
  let previousValue = values[0];

  for (let index = 1; index <= values.length; index++) {
    const currentValue = values[index];
    const reachedEnd = index === values.length;

    if (reachedEnd || currentValue !== previousValue) {
      const groupEndRow = firstDataRow + index - 1;
      const groupHeight = groupEndRow - groupStartRow + 1;

      if (groupHeight > 1) {
        mergeColumns.forEach(column => {
          sheet.getRange(groupStartRow, column, groupHeight, 1).mergeVertically();
        });
      }

      if (groupStartRow > firstDataRow) {
        sheet
          .getRange(groupStartRow, 1, 1, lastColumn)
          .setBorder(true, null, null, null, null, null, true, SpreadsheetApp.BorderStyle.DOTTED);
      }

      groupStartRow = firstDataRow + index;
      previousValue = currentValue;
    }
  }
}

/**
 * Produces a stable string for values saved in PropertiesService.
 * @param {*} value
 * @return {string}
 */
function normalizePropertyValue_(value) {
  if (value instanceof Date) {
    return String(value.getTime());
  }

  return value === null || value === undefined ? '' : String(value);
}

/**
 * Produces a stable comparison value for grouping.
 * @param {*} value
 * @return {string}
 */
function normalizeGroupValue_(value) {
  if (value instanceof Date) {
    return String(value.getTime());
  }

  return value === null || value === undefined ? '' : String(value).trim();
}

/**
 * Formats an Apps Script error for logs.
 * @param {*} error
 * @return {string}
 */
function formatError_(error) {
  if (error && error.stack) {
    return error.stack;
  }

  return String(error);
}
