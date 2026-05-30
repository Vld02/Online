/**
 * Processes sheets whose names start with "=" without showing users the
 * intermediate table states.
 *
 * The function prepares every change for a sheet and sends it as a single
 * Sheets API batchUpdate request. That keeps break-apart, border cleanup,
 * merging, separator borders, and filter restoration in one visible update per
 * processed sheet instead of many separate SpreadsheetApp mutations.
 *
 * Prerequisite: enable the Advanced Google service "Google Sheets API" in Apps
 * Script (Services → + → Google Sheets API).
 */
function processSheets() {
  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    lock.waitLock(1000);
    hasLock = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    const sheets = spreadsheet.getSheets();
    const mainColumn = 6; // F
    const mergeColumns = [1, 3, 4, 5, 6, 38]; // A, C, D, E, F, AL
    const filterColumn = 2; // B
    const filterStartRow = 4;

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

      Logger.log(`Обработка листа "${name}". Старое значение: ${storedValue}, новое: ${currentValue}`);

      const sheetId = sheet.getSheetId();
      const requests = buildSheetRequests_(sheet, {
        sheetId,
        lastRow,
        lastColumn,
        mainColumn,
        mergeColumns,
        filterColumn,
        filterStartRow,
      });

      if (requests.length === 0) {
        props.setProperty('lastValue_' + name, String(currentValue));
        Logger.log(`Для листа "${name}" нет изменений для отправки.`);
        return;
      }

      // Один batchUpdate на лист = пользователь не видит промежуточные состояния.
      Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
      props.setProperty('lastValue_' + name, String(currentValue));

      Logger.log(`Обработка листа "${name}" завершена одним batchUpdate.`);
    });

    Logger.log('=== Все листы обработаны ===');
  } catch (e) {
    Logger.log('Ошибка: ' + e);
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Builds all formatting/filter requests for one sheet.
 * Rows/columns in Apps Script are 1-based, Sheets API indexes are 0-based and
 * end-exclusive.
 */
function buildSheetRequests_(sheet, config) {
  const {
    sheetId,
    lastRow,
    lastColumn,
    mainColumn,
    mergeColumns,
    filterColumn,
    filterStartRow,
  } = config;

  const requests = [];

  // 1. Заменяем текущий фильтр на фильтр "не пустые" сразу в финальном виде.
  requests.push({ clearBasicFilter: { sheetId } });
  requests.push({
    setBasicFilter: {
      filter: {
        range: {
          sheetId,
          startRowIndex: filterStartRow - 1,
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

  // 2. Снятие объединений.
  requests.push({
    unmergeCells: {
      range: {
        sheetId,
        startRowIndex: 4,
        endRowIndex: lastRow,
        startColumnIndex: 0,
        endColumnIndex: lastColumn,
      },
    },
  });

  // 3. Очистка границ.
  requests.push({
    updateBorders: {
      range: {
        sheetId,
        startRowIndex: 5,
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

  // 4. Группировка по F: объединения и пунктирные разделители групп.
  const data = sheet.getRange(5, mainColumn, lastRow - 4).getValues().flat();
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

  return requests;
}
