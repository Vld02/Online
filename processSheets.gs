/**
 * Обрабатывает листы через один атомарный Sheets API batchUpdate.
 *
 * Важно: включите Advanced Google service "Google Sheets API" в Apps Script:
 * Services → + → Google Sheets API.
 */
function processSheets() {
  const lock = LockService.getScriptLock();
  let lockTaken = false;

  try {
    lock.waitLock(1000);
    lockTaken = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    const sheets = spreadsheet.getSheets();
    const mainColumn = 6; // F
    const mergeColumns = [1, 3, 4, 5, 6, 38]; // A, C, D, E, F, AL
    const filterColumn = 2; // B
    const filterStartRow = 4;

    const props = PropertiesService.getScriptProperties();

    cleanupDeletedSheetProperties_(props, sheets);

    const requests = [];

    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (!name.startsWith('=')) return;

      const currentValue = sheet.getRange('G4').getValue();
      const propertyKey = 'lastValue_' + name;
      const storedValue = props.getProperty(propertyKey);

      if (String(currentValue) === String(storedValue)) {
        Logger.log(`Пропуск листа "${name}" — значение G4 не изменилось (${currentValue}).`);
        return;
      }

      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      if (lastRow < 5 || lastColumn < 2) {
        props.setProperty(propertyKey, String(currentValue));
        Logger.log(`Пропуск листа "${name}" — недостаточно строк для обработки.`);
        return;
      }

      Logger.log(`Подготовка листа "${name}". Старое значение: ${storedValue}, новое: ${currentValue}`);

      const sheetId = sheet.getSheetId();
      const data = sheet.getRange(5, mainColumn, lastRow - 4, 1).getValues().flat();

      requests.push(
        // 1. Сброс фильтра.
        buildBasicFilterRequest_(sheetId, lastRow, lastColumn, filterStartRow, filterColumn, false),

        // 2. Снятие объединений.
        {
          unmergeCells: {
            range: toGridRange_(sheetId, 5, 1, lastRow, lastColumn),
          },
        }
      );

      // 3. Очистка границ.
      if (lastRow >= 6) {
        requests.push({
          updateBorders: {
            range: toGridRange_(sheetId, 6, 1, lastRow, lastColumn),
            top: { style: 'NONE' },
            bottom: { style: 'NONE' },
            left: { style: 'NONE' },
            right: { style: 'NONE' },
            innerHorizontal: { style: 'NONE' },
            innerVertical: { style: 'NONE' },
          },
        });
      }

      // 4. Группировка по F: объединения и пунктирные границы.
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
            if (col <= lastColumn) {
              requests.push({
                mergeCells: {
                  range: toGridRange_(sheetId, groupStart, col, groupEnd, col),
                  mergeType: 'MERGE_COLUMNS',
                },
              });
            }
          });
        }

        if (groupStart > 5) {
          requests.push({
            updateBorders: {
              range: toGridRange_(sheetId, groupStart, 1, groupStart, lastColumn),
              top: {
                style: 'DOTTED',
                width: 1,
                color: { red: 0, green: 0, blue: 0 },
              },
            },
          });
        }

        start = i + 5;
      }

      // 5. Установить фильтр "не пустые".
      requests.push(buildBasicFilterRequest_(sheetId, lastRow, lastColumn, filterStartRow, filterColumn, true));

      props.setProperty(propertyKey, String(currentValue));
      Logger.log(`Лист "${name}" добавлен в общий пакет изменений.`);
    });

    if (requests.length === 0) {
      Logger.log('Нет изменений для применения.');
      return;
    }

    // Все визуальные изменения отправляются в таблицу одним запросом.
    Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);

    Logger.log(`=== Все листы обработаны одним batchUpdate. Запросов: ${requests.length} ===`);
  } catch (e) {
    Logger.log('Ошибка: ' + e);
    throw e;
  } finally {
    if (lockTaken) {
      lock.releaseLock();
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

function buildBasicFilterRequest_(sheetId, lastRow, lastColumn, filterStartRow, filterColumn, onlyNotEmpty) {
  const filter = {
    range: toGridRange_(sheetId, filterStartRow, 2, lastRow, lastColumn),
  };

  if (onlyNotEmpty) {
    filter.criteria = {
      [filterColumn - 1]: {
        condition: {
          type: 'NOT_BLANK',
        },
      },
    };
  }

  return { setBasicFilter: { filter } };
}

function toGridRange_(sheetId, startRow, startColumn, endRow, endColumn) {
  return {
    sheetId,
    startRowIndex: startRow - 1,
    endRowIndex: endRow,
    startColumnIndex: startColumn - 1,
    endColumnIndex: endColumn,
  };
}
