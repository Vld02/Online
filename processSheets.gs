/**
 * Processes at most one changed sheet per run and applies all visual changes
 * in a single Sheets API batchUpdate request so users do not see intermediate
 * table states (unmerge -> clear borders -> merge -> filter) while it runs.
 *
 * Requirement: enable the Advanced Google service "Google Sheets API"
 * in Apps Script (Services -> Sheets API).
 */
function processSheets() {
  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    lock.waitLock(1000);
    hasLock = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const spreadsheetId = spreadsheet.getId();
    const mainColumn = 6; // F
    const mergeColumns = [1, 3, 4, 5, 6, 38]; // A, C, D, E, F, AL
    const filterColumn = 2; // B

    const props = PropertiesService.getScriptProperties();

    cleanupDeletedSheetProperties_(props, sheets);
    logActualSheetProperties_(props);

    for (const sheet of sheets) {
      const name = sheet.getName();
      if (!name.startsWith('=')) continue;

      const currentValue = sheet.getRange('G4').getValue();
      const storedValue = props.getProperty('lastValue_' + name);

      if (String(currentValue) === String(storedValue)) {
        Logger.log(`Пропуск листа "${name}" — значение G4 не изменилось (${currentValue}).`);
        continue;
      }

      Logger.log(`Обработка листа "${name}". Старое значение: ${storedValue}, новое: ${currentValue}`);

      processOneSheetInSingleBatch_(spreadsheetId, sheet, mainColumn, mergeColumns, filterColumn);
      props.setProperty('lastValue_' + name, String(currentValue));

      Logger.log(`Обработка листа "${name}" завершена одним batchUpdate.`);
      Logger.log('=== За этот запуск обработан один лист ===');
      return;
    }

    Logger.log('Нет листов с изменённым значением G4.');
  } catch (e) {
    Logger.log('Ошибка: ' + e);
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function cleanupDeletedSheetProperties_(props, sheets) {
  const allProps = props.getProperties();
  const sheetNames = sheets.map(sheet => sheet.getName());

  for (const key in allProps) {
    if (!key.startsWith('lastValue_')) continue;

    const sheetName = key.replace('lastValue_', '');
    if (!sheetNames.includes(sheetName)) {
      props.deleteProperty(key);
      Logger.log(`Удалён ключ "${key}", так как лист "${sheetName}" больше не существует.`);
    }
  }
}

function logActualSheetProperties_(props) {
  const remainingProps = props.getProperties();
  const keys = Object.keys(remainingProps).filter(key => key.startsWith('lastValue_'));

  if (keys.length > 0) {
    Logger.log('Актуальные ключи в PropertiesService:');
    keys.forEach(key => Logger.log(` • ${key} = ${remainingProps[key]}`));
  } else {
    Logger.log('В PropertiesService не осталось ключей lastValue_.');
  }
}

function processOneSheetInSingleBatch_(spreadsheetId, sheet, mainColumn, mergeColumns, filterColumn) {
  const sheetId = sheet.getSheetId();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 5 || lastColumn < 1) {
    Logger.log(`Лист "${sheet.getName()}" пропущен — недостаточно строк для обработки.`);
    return;
  }

  const requests = [];
  const dataRowCount = lastRow - 4;

  requests.push({
    setBasicFilter: {
      filter: {
        range: toGridRange_(sheetId, 4, filterColumn, lastRow, filterColumn),
        criteria: {
          [filterColumn - 1]: {
            condition: {
              type: 'NOT_BLANK'
            }
          }
        }
      }
    }
  });

  requests.push({
    unmergeCells: {
      range: toGridRange_(sheetId, 5, 1, lastRow, lastColumn)
    }
  });

  if (lastRow >= 6) {
    requests.push({
      updateBorders: Object.assign(
        {range: toGridRange_(sheetId, 6, 1, lastRow, lastColumn)},
        emptyBorders_()
      )
    });
  }

  const data = sheet.getRange(5, mainColumn, dataRowCount).getValues().flat();
  let start = 5;

  for (let i = 1; i <= data.length; i++) {
    const curr = data[i];
    const prev = data[i - 1];
    const isGroupEnd = curr !== prev || i === data.length;

    if (!isGroupEnd) continue;

    const groupStart = start;
    const groupEnd = i + 4;

    if (groupEnd > groupStart) {
      mergeColumns.forEach(col => {
        requests.push({
          mergeCells: {
            range: toGridRange_(sheetId, groupStart, col, groupEnd, col),
            mergeType: 'MERGE_COLUMNS'
          }
        });
      });
    }

    if (groupStart > 5) {
      requests.push({
        updateBorders: {
          range: toGridRange_(sheetId, groupStart, 1, groupStart, lastColumn),
          top: {
            style: 'DOTTED',
            width: 1,
            color: {red: 0, green: 0, blue: 0}
          }
        }
      });
    }

    start = i + 5;
  }

  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({requests}, spreadsheetId);
  }
}

function toGridRange_(sheetId, startRow, startColumn, endRow, endColumn) {
  return {
    sheetId,
    startRowIndex: startRow - 1,
    endRowIndex: endRow,
    startColumnIndex: startColumn - 1,
    endColumnIndex: endColumn
  };
}

function emptyBorders_() {
  const none = {style: 'NONE'};

  return {
    top: none,
    bottom: none,
    left: none,
    right: none,
    innerHorizontal: none,
    innerVertical: none
  };
}
