/**
 * Processes all sheets whose names start with "=" in one Sheets API batch.
 *
 * Important:
 * 1. Enable the Advanced Google service "Google Sheets API" in Apps Script:
 *    Services → + → Google Sheets API.
 * 2. Do not call SpreadsheetApp.flush() inside this function. All visual changes
 *    are sent through one batchUpdate request at the end of the preparation step.
 */
function processSheets() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(1000);

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    const sheets = spreadsheet.getSheets();
    const mainColumn = 6; // F
    const mergeColumns = [1, 3, 4, 5, 6, 38]; // A, C, D, E, F, AL
    const filterColumn = 2; // B
    const filterStartRow = 4;
    const props = PropertiesService.getScriptProperties();
    const requests = [];
    const pendingLastValues = [];

    removeDeletedSheetProperties_(props, sheets);
    logActualLastValueProperties_(props);

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
      if (lastRow < 5 || lastColumn < 1) {
        props.setProperty('lastValue_' + name, String(currentValue));
        Logger.log(`Пропуск листа "${name}" — недостаточно строк для обработки.`);
        return;
      }

      Logger.log(`Подготовка листа "${name}". Старое значение: ${storedValue}, новое: ${currentValue}`);

      const sheetId = sheet.getSheetId();
      const data = sheet.getRange(5, mainColumn, lastRow - 4).getValues().flat();
      const sheetRequests = buildSheetRequests_(sheetId, lastRow, lastColumn, data, {
        mergeColumns,
        filterColumn,
        filterStartRow,
      });

      requests.push(...sheetRequests);
      pendingLastValues.push({ key: 'lastValue_' + name, value: String(currentValue) });
      Logger.log(`Лист "${name}" добавлен в общий пакет обработки.`);
    });

    if (requests.length === 0) {
      Logger.log('Нет листов с изменениями для обработки.');
      return;
    }

    Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
    pendingLastValues.forEach(item => props.setProperty(item.key, item.value));
    Logger.log(`=== Все изменения применены одним пакетом: ${requests.length} запросов ===`);
  } catch (e) {
    Logger.log('Ошибка: ' + e);
  } finally {
    lock.releaseLock();
  }
}

function removeDeletedSheetProperties_(props, sheets) {
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
}

function logActualLastValueProperties_(props) {
  const remainingProps = props.getProperties();
  const keys = Object.keys(remainingProps).filter(key => key.startsWith('lastValue_'));

  if (keys.length > 0) {
    Logger.log('Актуальные ключи в PropertiesService:');
    keys.forEach(key => Logger.log(` • ${key} = ${remainingProps[key]}`));
  } else {
    Logger.log('В PropertiesService не осталось ключей lastValue_.');
  }
}

function buildSheetRequests_(sheetId, lastRow, lastColumn, data, options) {
  const requests = [];
  const mergeColumns = options.mergeColumns;
  const filterColumn = options.filterColumn;
  const filterStartRow = options.filterStartRow;

  const bodyRange = {
    sheetId,
    startRowIndex: 4,
    endRowIndex: lastRow,
    startColumnIndex: 0,
    endColumnIndex: lastColumn,
  };

  const borderRange = {
    sheetId,
    startRowIndex: 5,
    endRowIndex: lastRow,
    startColumnIndex: 0,
    endColumnIndex: lastColumn,
  };

  requests.push({
    clearBasicFilter: {
      sheetId,
    },
  });

  requests.push({
    unmergeCells: {
      range: bodyRange,
    },
  });

  if (lastRow >= 6) {
    requests.push({
      updateBorders: {
        range: borderRange,
        top: { style: 'NONE' },
        bottom: { style: 'NONE' },
        left: { style: 'NONE' },
        right: { style: 'NONE' },
        innerHorizontal: { style: 'NONE' },
        innerVertical: { style: 'NONE' },
      },
    });
  }

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
        if (column > lastColumn) return;

        requests.push({
          mergeCells: {
            range: {
              sheetId,
              startRowIndex: groupStart - 1,
              endRowIndex: groupEnd,
              startColumnIndex: column - 1,
              endColumnIndex: column,
            },
            mergeType: 'MERGE_COLUMNS',
          },
        });
      });
    }

    if (groupStart > 5) {
      requests.push({
        updateBorders: {
          range: {
            sheetId,
            startRowIndex: groupStart - 1,
            endRowIndex: groupStart,
            startColumnIndex: 0,
            endColumnIndex: lastColumn,
          },
          top: {
            style: 'DOTTED',
            width: 1,
          },
        },
      });
    }

    start = i + 5;
  }

  requests.push({
    setBasicFilter: {
      filter: {
        range: {
          sheetId,
          startRowIndex: filterStartRow - 1,
          endRowIndex: lastRow,
          startColumnIndex: filterColumn - 1,
          endColumnIndex: filterColumn,
        },
        criteria: {
          [filterColumn - 1]: {
            condition: {
              type: 'NOT_BLANK',
            },
          },
        },
      },
    },
  });

  return requests;
}
