/**
 * Форматирует листы, имя которых начинается с "=":
 * - пропускает лист, если значение G4 не изменилось с прошлого запуска;
 * - сбрасывает фильтр, снимает объединения и очищает границы в рабочей области;
 * - объединяет одинаковые подряд группы по колонке F в выбранных колонках;
 * - возвращает фильтр по колонке B на условие "не пустые".
 */
function processSheets() {
  const CONFIG = {
    firstDataRow: 5,
    filterHeaderRow: 4,
    mainColumn: 6, // F
    mergeColumns: [1, 3, 4, 5, 6, 38], // A, C, D, E, F, AL
    filterColumn: 2, // B
    changeCell: 'G4',
    processedSheetPrefix: '=',
    propertyPrefix: 'lastValue_',
    lockTimeoutMs: 10 * 1000,
  };

  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    hasLock = lock.tryLock(CONFIG.lockTimeoutMs);
    if (!hasLock) {
      Logger.log('Другой запуск processSheets ещё выполняется. Текущий запуск пропущен.');
      return;
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const sheetNames = new Set(sheets.map(sheet => sheet.getName()));
    const props = PropertiesService.getScriptProperties();

    cleanupDeletedSheetProperties_(props, sheetNames, CONFIG.propertyPrefix);
    logActualProperties_(props, CONFIG.propertyPrefix);

    sheets
      .filter(sheet => sheet.getName().startsWith(CONFIG.processedSheetPrefix))
      .forEach(sheet => processOneSheet_(sheet, props, CONFIG));

    Logger.log('=== Все листы обработаны ===');
  } catch (error) {
    Logger.log('Ошибка processSheets: ' + (error && error.stack ? error.stack : error));
    throw error;
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Обрабатывает один подходящий лист.
 */
function processOneSheet_(sheet, props, config) {
  const sheetName = sheet.getName();
  const propertyKey = config.propertyPrefix + sheetName;
  const currentValue = String(sheet.getRange(config.changeCell).getDisplayValue());
  const storedValue = props.getProperty(propertyKey);

  if (currentValue === storedValue) {
    Logger.log('Пропуск листа "' + sheetName + '" — значение ' + config.changeCell + ' не изменилось (' + currentValue + ').');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < config.firstDataRow || lastColumn < 1) {
    props.setProperty(propertyKey, currentValue);
    Logger.log('Лист "' + sheetName + '" пропущен: нет строк данных для обработки.');
    return;
  }

  Logger.log('Обработка листа "' + sheetName + '". Старое значение: ' + storedValue + ', новое: ' + currentValue);

  const filter = resetFilter_(sheet, lastRow, lastColumn, config);
  resetFormatting_(sheet, lastRow, lastColumn, config);
  mergeGroupsByMainColumn_(sheet, lastRow, lastColumn, config);
  applyNotEmptyFilter_(filter, config.filterColumn);

  // Сохраняем новое значение только после успешной обработки листа.
  props.setProperty(propertyKey, currentValue);
  Logger.log('Обработка листа "' + sheetName + '" завершена.');
}

/**
 * Удаляет значения lastValue_* для листов, которых уже нет в таблице.
 */
function cleanupDeletedSheetProperties_(props, sheetNames, propertyPrefix) {
  const allProps = props.getProperties();

  Object.keys(allProps)
    .filter(key => key.startsWith(propertyPrefix))
    .forEach(key => {
      const sheetName = key.substring(propertyPrefix.length);
      if (!sheetNames.has(sheetName)) {
        props.deleteProperty(key);
        Logger.log('Удалён ключ "' + key + '", так как лист "' + sheetName + '" больше не существует.');
      }
    });
}

/**
 * Логирует актуальные сохранённые значения.
 */
function logActualProperties_(props, propertyPrefix) {
  const actualProps = props.getProperties();
  const keys = Object.keys(actualProps).filter(key => key.startsWith(propertyPrefix));

  if (keys.length === 0) {
    Logger.log('В PropertiesService не осталось ключей ' + propertyPrefix + '.');
    return;
  }

  Logger.log('Актуальные ключи в PropertiesService:');
  keys.sort().forEach(key => Logger.log(' • ' + key + ' = ' + actualProps[key]));
}

/**
 * Создаёт корректный фильтр на всю таблицу и очищает критерии колонки B.
 */
function resetFilter_(sheet, lastRow, lastColumn, config) {
  const filterRange = sheet.getRange(
    config.filterHeaderRow,
    1,
    lastRow - config.filterHeaderRow + 1,
    lastColumn
  );

  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  const filter = filterRange.createFilter();
  filter.removeColumnFilterCriteria(config.filterColumn);
  return filter;
}

/**
 * Снимает объединения и очищает границы в рабочей области.
 */
function resetFormatting_(sheet, lastRow, lastColumn, config) {
  const dataRowCount = lastRow - config.firstDataRow + 1;
  sheet.getRange(config.firstDataRow, 1, dataRowCount, lastColumn).breakApart();

  const borderStartRow = config.firstDataRow + 1;
  if (lastRow >= borderStartRow) {
    sheet.getRange(borderStartRow, 1, lastRow - borderStartRow + 1, lastColumn)
      .setBorder(false, false, false, false, false, false);
  }
}

/**
 * Объединяет подряд идущие одинаковые значения в основной колонке.
 */
function mergeGroupsByMainColumn_(sheet, lastRow, lastColumn, config) {
  const rowCount = lastRow - config.firstDataRow + 1;
  const values = sheet.getRange(config.firstDataRow, config.mainColumn, rowCount, 1)
    .getDisplayValues()
    .map(row => row[0]);

  let groupStartIndex = 0;

  for (let index = 1; index <= values.length; index++) {
    const isLastValue = index === values.length;
    const isNewGroup = !isLastValue && values[index] !== values[index - 1];

    if (!isLastValue && !isNewGroup) {
      continue;
    }

    const groupStartRow = config.firstDataRow + groupStartIndex;
    const groupEndRow = config.firstDataRow + index - 1;
    const groupRowCount = groupEndRow - groupStartRow + 1;

    if (groupRowCount > 1) {
      config.mergeColumns
        .filter(column => column <= lastColumn)
        .forEach(column => sheet.getRange(groupStartRow, column, groupRowCount, 1).mergeVertically());
    }

    if (groupStartRow > config.firstDataRow) {
      sheet.getRange(groupStartRow, 1, 1, lastColumn)
        .setBorder(true, null, null, null, null, null, true, SpreadsheetApp.BorderStyle.DOTTED);
    }

    groupStartIndex = index;
  }
}

/**
 * Включает фильтр "не пустые" для указанной колонки.
 */
function applyNotEmptyFilter_(filter, filterColumn) {
  filter.setColumnFilterCriteria(
    filterColumn,
    SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build()
  );
}
