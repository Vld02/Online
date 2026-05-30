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

    const requests = [];
    const processedSheets = [];

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

      Logger.log(`Подготовка листа "${name}". Старое значение: ${storedValue}, новое: ${currentValue}`);

      const sheetId = sheet.getSheetId();
      const mainValues = sheet.getRange(5, mainColumn, lastRow - 4, 1).getValues().flat();

      // Все визуальные изменения собираются в один batchUpdate, чтобы пользователь не видел
      // промежуточные состояния: снятые объединения, очищенные границы и частичные новые объединения.
      requests.push(createBasicFilterRequest_(sheetId, lastRow, filterColumn, null));
      requests.push(createUnmergeRequest_(sheetId, lastRow, lastColumn));
      requests.push(createClearBordersRequest_(sheetId, lastRow, lastColumn));

      buildGroupRequests_(requests, sheetId, mainValues, mergeColumns, lastColumn);

      requests.push(createBasicFilterRequest_(sheetId, lastRow, filterColumn, 'NOT_BLANK'));
      processedSheets.push({ name, value: String(currentValue) });
    });

    if (requests.length === 0) {
      Logger.log('Нет листов с изменённым G4 для обработки.');
      return;
    }

    sheetsBatchUpdate_(spreadsheetId, requests);
    processedSheets.forEach(sheet => props.setProperty('lastValue_' + sheet.name, sheet.value));
    Logger.log(`=== Все листы обработаны одним пакетным обновлением: ${requests.length} операций ===`);
  } catch (e) {
    Logger.log('Ошибка: ' + e);
  } finally {
    if (hasLock) lock.releaseLock();
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

function buildGroupRequests_(requests, sheetId, data, mergeColumns, lastColumn) {
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
        requests.push(createMergeRequest_(sheetId, groupStart, groupEnd, column));
      });
    }

    if (groupStart > 5) {
      requests.push(createDottedTopBorderRequest_(sheetId, groupStart, lastColumn));
    }

    start = i + 5;
  }
}

function createBasicFilterRequest_(sheetId, lastRow, filterColumn, conditionType) {
  const criteria = conditionType
    ? {
        [filterColumn - 1]: {
          condition: {
            type: conditionType
          }
        }
      }
    : {};

  return {
    setBasicFilter: {
      filter: {
        range: {
          sheetId,
          startRowIndex: 3,
          endRowIndex: lastRow,
          startColumnIndex: filterColumn - 1,
          endColumnIndex: filterColumn
        },
        criteria
      }
    }
  };
}

function createUnmergeRequest_(sheetId, lastRow, lastColumn) {
  return {
    unmergeCells: {
      range: {
        sheetId,
        startRowIndex: 4,
        endRowIndex: lastRow,
        startColumnIndex: 0,
        endColumnIndex: lastColumn
      }
    }
  };
}

function createClearBordersRequest_(sheetId, lastRow, lastColumn) {
  return {
    updateBorders: {
      range: {
        sheetId,
        startRowIndex: 5,
        endRowIndex: lastRow,
        startColumnIndex: 0,
        endColumnIndex: lastColumn
      },
      top: { style: 'NONE' },
      bottom: { style: 'NONE' },
      left: { style: 'NONE' },
      right: { style: 'NONE' },
      innerHorizontal: { style: 'NONE' },
      innerVertical: { style: 'NONE' }
    }
  };
}

function createMergeRequest_(sheetId, startRow, endRow, column) {
  return {
    mergeCells: {
      range: {
        sheetId,
        startRowIndex: startRow - 1,
        endRowIndex: endRow,
        startColumnIndex: column - 1,
        endColumnIndex: column
      },
      mergeType: 'MERGE_COLUMNS'
    }
  };
}

function createDottedTopBorderRequest_(sheetId, row, lastColumn) {
  return {
    updateBorders: {
      range: {
        sheetId,
        startRowIndex: row - 1,
        endRowIndex: row,
        startColumnIndex: 0,
        endColumnIndex: lastColumn
      },
      top: {
        style: 'DOTTED',
        width: 1,
        color: {
          red: 0,
          green: 0,
          blue: 0
        }
      }
    }
  };
}

function sheetsBatchUpdate_(spreadsheetId, requests) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify({ requests }),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Sheets API batchUpdate failed (${code}): ${response.getContentText()}`);
  }
}
