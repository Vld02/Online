/**
 * Обрабатывает листы без показа промежуточных состояний пользователю.
 *
 * Важно: включите Advanced Google Service "Google Sheets API"
 * (Services → Google Sheets API), потому что изменения применяются через
 * Sheets.Spreadsheets.batchUpdate одним пакетным запросом на каждый лист.
 */
function processSheets() {
  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    lock.waitLock(1000);
    hasLock = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const props = PropertiesService.getScriptProperties();

    // --- Очистка старых записей ---
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

    // Показать, что осталось после очистки
    const remainingProps = props.getProperties();
    const keys = Object.keys(remainingProps).filter(key => key.startsWith('lastValue_'));

    if (keys.length > 0) {
      Logger.log('Актуальные ключи в PropertiesService:');
      keys.forEach(key => Logger.log(` • ${key} = ${remainingProps[key]}`));
    } else {
      Logger.log('В PropertiesService не осталось ключей lastValue_.');
    }

    // --- Основная обработка ---
    sheets.forEach(sheet => processSingleSheet_(spreadsheet, sheet, props));

    Logger.log('=== Все листы обработаны ===');
  } catch (error) {
    Logger.log('Ошибка: ' + error);
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function processSingleSheet_(spreadsheet, sheet, props) {
  const name = sheet.getName();
  if (!name.startsWith('=')) return;

  const currentValue = sheet.getRange('G4').getValue();
  const storedValue = props.getProperty('lastValue_' + name);

  if (String(currentValue) === String(storedValue)) {
    Logger.log(`Пропуск листа "${name}" — значение G4 не изменилось (${currentValue}).`);
    return;
  }

  const requests = buildSheetBatchRequests_(sheet);
  if (requests.length === 0) {
    props.setProperty('lastValue_' + name, String(currentValue));
    Logger.log(`Пропуск листа "${name}" — нет строк для обработки. Значение G4 сохранено (${currentValue}).`);
    return;
  }

  Logger.log(`Обработка листа "${name}". Старое значение: ${storedValue}, новое: ${currentValue}`);

  // Один batchUpdate на лист: пользователь не видит сброс фильтра, снятие объединений
  // и очистку границ как отдельные промежуточные состояния.
  Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheet.getId());

  // Сохраняем новое значение только после успешного применения пакета изменений.
  props.setProperty('lastValue_' + name, String(currentValue));
  Logger.log(`Обработка листа "${name}" завершена.`);
}

function buildSheetBatchRequests_(sheet) {
  const mainColumn = 6; // F
  const mergeColumns = [1, 3, 4, 5, 6, 38]; // A, C, D, E, F, AL
  const filterColumn = 2; // B
  const firstDataRow = 5;
  const firstBorderClearRow = 6;

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const sheetId = sheet.getSheetId();

  if (lastRow < firstDataRow || lastColumn < 1) return [];

  const requests = [];

  // Финальный фильтр "не пустые" сразу задаётся в пакетном запросе.
  // Отдельный видимый сброс фильтра больше не нужен.
  requests.push({
    setBasicFilter: {
      filter: {
        range: {
          sheetId,
          startRowIndex: 3, // B4:B
          endRowIndex: sheet.getMaxRows(),
          startColumnIndex: filterColumn - 1,
          endColumnIndex: filterColumn,
        },
        criteria: {
          [String(filterColumn - 1)]: {
            condition: {
              type: 'NOT_BLANK',
            },
          },
        },
      },
    },
  });

  requests.push({
    unmergeCells: {
      range: {
        sheetId,
        startRowIndex: firstDataRow - 1,
        endRowIndex: lastRow,
        startColumnIndex: 0,
        endColumnIndex: lastColumn,
      },
    },
  });

  if (lastRow >= firstBorderClearRow) {
    requests.push({
      updateBorders: {
        range: {
          sheetId,
          startRowIndex: firstBorderClearRow - 1,
          endRowIndex: lastRow,
          startColumnIndex: 0,
          endColumnIndex: lastColumn,
        },
        top: { style: 'NONE' },
        bottom: { style: 'NONE' },
        left: { style: 'NONE' },
        right: { style: 'NONE' },
        innerHorizontal: { style: 'NONE' },
        innerVertical: { style: 'NONE' },
      },
    });
  }

  const data = sheet.getRange(firstDataRow, mainColumn, lastRow - firstDataRow + 1).getValues().flat();
  let start = firstDataRow;

  for (let index = 1; index <= data.length; index++) {
    const current = data[index];
    const previous = data[index - 1];
    const isGroupEnd = current !== previous || index === data.length;

    if (!isGroupEnd) continue;

    const groupStart = start;
    const groupEnd = index + firstDataRow - 1;

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
            mergeType: 'MERGE_ALL',
          },
        });
      });
    }

    if (groupStart > firstDataRow) {
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

    start = index + firstDataRow;
  }

  return requests;
}
