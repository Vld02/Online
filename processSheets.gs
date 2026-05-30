function processSheets() {
  const lock = LockService.getScriptLock();
  let hasLock = false;
  const hiddenSheets = [];
  let spreadsheet;
  let temporarySheet = null;
  let activeSheet = null;

  try {
    lock.waitLock(1000);
    hasLock = true;

    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    activeSheet = spreadsheet.getActiveSheet();

    const sheets = spreadsheet.getSheets();
    const mainColumn = 6; // F
    const mergeColumns = [1, 3, 4, 5, 6, 38]; // A, C, D, E, F, AL
    const filterColumn = 2; // B
    const props = PropertiesService.getScriptProperties();

    cleanupDeletedSheetProperties_(props, sheets);

    const sheetsToProcess = getSheetsToProcess_(sheets, props, mainColumn);

    if (sheetsToProcess.length === 0) {
      Logger.log('Нет листов с изменённым G4 для обработки.');
      return;
    }

    temporarySheet = createProcessingSheet_(spreadsheet);
    temporarySheet.activate();
    SpreadsheetApp.flush();

    sheetsToProcess.forEach(item => {
      if (!item.sheet.isSheetHidden()) {
        item.sheet.hideSheet();
        hiddenSheets.push(item.sheet);
      }
    });
    SpreadsheetApp.flush();

    const processedSheets = [];

    sheetsToProcess.forEach(item => {
      const sheet = item.sheet;
      Logger.log(
        `Обработка листа "${item.name}". Старое значение: ${item.storedValue}, новое: ${item.currentValue}`
      );

      processOneSheet_(sheet, item.lastRow, item.lastColumn, item.data, mainColumn, mergeColumns, filterColumn);
      processedSheets.push({ name: item.name, value: String(item.currentValue) });
      Logger.log(`Обработка листа "${item.name}" завершена.`);
    });

    SpreadsheetApp.flush();
    processedSheets.forEach(item => props.setProperty('lastValue_' + item.name, item.value));
    Logger.log(`=== Все листы обработаны скрыто: ${processedSheets.length} ===`);
  } catch (e) {
    Logger.log('Ошибка: ' + e);
  } finally {
    restoreVisibleState_(spreadsheet, activeSheet, hiddenSheets, temporarySheet);
    if (hasLock) lock.releaseLock();
  }
}

function getSheetsToProcess_(sheets, props, mainColumn) {
  const sheetsToProcess = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith('=')) return;

    const currentValue = sheet.getRange('G4').getValue();
    const storedValue = props.getProperty('lastValue_' + name);

    if (String(currentValue) === String(storedValue)) {
      Logger.log(`Пропуск листа "${name}" — значение G4 не изменилось (${currentValue}).`);
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 5 || lastColumn === 0) {
      props.setProperty('lastValue_' + name, String(currentValue));
      Logger.log(`Пропуск листа "${name}" — недостаточно строк для обработки.`);
      return;
    }

    sheetsToProcess.push({
      sheet,
      name,
      currentValue,
      storedValue,
      lastRow,
      lastColumn,
      data: sheet.getRange(5, mainColumn, lastRow - 4, 1).getValues().flat()
    });
  });

  return sheetsToProcess;
}

function processOneSheet_(sheet, lastRow, lastColumn, data, mainColumn, mergeColumns, filterColumn) {
  const filter = ensureFilterForColumn_(sheet, lastRow, lastColumn, filterColumn);
  const filterPosition = getFilterColumnPosition_(filter, filterColumn);

  filter.setColumnFilterCriteria(filterPosition, SpreadsheetApp.newFilterCriteria().build());

  sheet.getRange(5, 1, lastRow - 4, lastColumn).breakApart();

  if (lastRow >= 6) {
    sheet.getRange(6, 1, lastRow - 5, lastColumn).setBorder(false, false, false, false, false, false);
  }

  mergeGroups_(sheet, data, mergeColumns, lastColumn);

  filter.setColumnFilterCriteria(
    filterPosition,
    SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build()
  );
}

function ensureFilterForColumn_(sheet, lastRow, lastColumn, filterColumn) {
  let filter = sheet.getFilter();

  if (filter && isFilterColumnInsideRange_(filter, filterColumn)) {
    return filter;
  }

  if (filter) {
    filter.remove();
  }

  const filterRows = Math.max(lastRow - 3, 1);
  const filterColumns = Math.max(lastColumn, filterColumn);
  sheet.getRange(4, 1, filterRows, filterColumns).createFilter();
  return sheet.getFilter();
}

function isFilterColumnInsideRange_(filter, filterColumn) {
  const range = filter.getRange();
  const firstColumn = range.getColumn();
  const lastColumn = firstColumn + range.getNumColumns() - 1;

  return filterColumn >= firstColumn && filterColumn <= lastColumn;
}

function getFilterColumnPosition_(filter, filterColumn) {
  return filterColumn - filter.getRange().getColumn() + 1;
}

function mergeGroups_(sheet, data, mergeColumns, lastColumn) {
  let start = 5;

  for (let i = 1; i <= data.length; i++) {
    const curr = data[i];
    const prev = data[i - 1];
    const isGroupEnd = curr !== prev || i === data.length;

    if (!isGroupEnd) continue;

    const groupStart = start;
    const groupEnd = i + 4;

    if (groupEnd > groupStart) {
      mergeColumns.forEach(column => {
        sheet.getRange(groupStart, column, groupEnd - groupStart + 1, 1).mergeVertically();
      });
    }

    if (groupStart > 5) {
      sheet.getRange(groupStart, 1, 1, lastColumn)
        .setBorder(true, null, null, null, null, null, true, SpreadsheetApp.BorderStyle.DOTTED);
    }

    start = i + 5;
  }
}

function createProcessingSheet_(spreadsheet) {
  const sheetName = `__processing_${Date.now()}__`;
  const sheet = spreadsheet.insertSheet(sheetName, 0);

  sheet.getRange('A1')
    .setValue('Идёт обработка таблицы…')
    .setFontWeight('bold')
    .setFontSize(14);
  sheet.getRange('A2').setValue('Готовый результат появится автоматически после завершения скрипта.');
  sheet.setTabColor('#fbbc04');

  return sheet;
}

function restoreVisibleState_(spreadsheet, activeSheet, hiddenSheets, temporarySheet) {
  if (!spreadsheet) return;

  hiddenSheets.forEach(sheet => {
    try {
      sheet.showSheet();
    } catch (e) {
      Logger.log(`Не удалось снова показать лист "${sheet.getName()}": ${e}`);
    }
  });

  try {
    if (activeSheet && !activeSheet.isSheetHidden()) {
      activeSheet.activate();
    }
  } catch (e) {
    Logger.log('Не удалось вернуть активный лист: ' + e);
  }

  if (temporarySheet) {
    try {
      spreadsheet.deleteSheet(temporarySheet);
    } catch (e) {
      Logger.log(`Не удалось удалить временный лист "${temporarySheet.getName()}": ${e}`);
    }
  }
}

function cleanupDeletedSheetProperties_(props, sheets) {
  const allProps = props.getProperties();
  const sheetNames = sheets.map(sheet => sheet.getName());

  Object.keys(allProps).forEach(key => {
    if (!key.startsWith('lastValue_')) return;

    const sheetName = key.replace('lastValue_', '');
    if (!sheetNames.includes(sheetName)) {
      props.deleteProperty(key);
      Logger.log(`Удалён ключ "${key}", так как лист "${sheetName}" больше не существует.`);
    }
  });

  const remainingProps = props.getProperties();
  const keys = Object.keys(remainingProps).filter(key => key.startsWith('lastValue_'));

  if (keys.length > 0) {
    Logger.log('Актуальные ключи в PropertiesService:');
    keys.forEach(key => Logger.log(` • ${key} = ${remainingProps[key]}`));
  } else {
    Logger.log('В PropertiesService не осталось ключей lastValue_.');
  }
}
