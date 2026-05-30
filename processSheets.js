function processSheets() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(1000);

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const mainColumn = 6; // F
    const mergeColumns = [1, 3, 4, 5, 6, 38]; // A, C, D, E, F, AL
    const filterColumn = 2; // B
    const filterRange = 'B4:B';

    const props = PropertiesService.getScriptProperties();

    // --- Очистка старых записей ---
    const allProps = props.getProperties();
    const sheetNames = sheets.map(s => s.getName());
    for (let key in allProps) {
      if (key.startsWith("lastValue_")) {
        const sheetName = key.replace("lastValue_", "");
        if (!sheetNames.includes(sheetName)) {
          props.deleteProperty(key);
          Logger.log(`Удалён ключ "${key}", так как лист "${sheetName}" больше не существует.`);
        }
      }
    }

    // Показать, что осталось после очистки
    const remainingProps = props.getProperties();
    const keys = Object.keys(remainingProps).filter(k => k.startsWith("lastValue_"));
    if (keys.length > 0) {
      Logger.log("Актуальные ключи в PropertiesService:");
      keys.forEach(k => Logger.log(` • ${k} = ${remainingProps[k]}`));
    } else {
      Logger.log("В PropertiesService не осталось ключей lastValue_.");
    }

    // --- Основная обработка ---
    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (!name.startsWith('=')) return;

      const currentValue = sheet.getRange("G4").getValue();
      const storedValue = props.getProperty("lastValue_" + name);

      if (String(currentValue) === String(storedValue)) {
        Logger.log(`Пропуск листа "${name}" — значение G4 не изменилось (${currentValue}).`);
        return;
      }

      Logger.log(`Обработка листа "${name}". Старое значение: ${storedValue}, новое: ${currentValue}`);
      props.setProperty("lastValue_" + name, String(currentValue));

      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();

      // 1. Сброс фильтра
      let filter = sheet.getFilter();
      if (!filter) {
        sheet.getRange(filterRange).createFilter();
        filter = sheet.getFilter();
      }
      filter.setColumnFilterCriteria(filterColumn, SpreadsheetApp.newFilterCriteria().build());

      // 2. Снятие объединений
      sheet.getRange(5, 1, lastRow - 4, lastColumn).breakApart();

      // 3. Очистка границ
      sheet.getRange(6, 1, lastRow - 5 + 1, lastColumn)
        .setBorder(false, false, false, false, false, false);

      // 4. Группировка по F
      const data = sheet.getRange(5, mainColumn, lastRow - 4).getValues().flat();
      let start = 5;
      for (let i = 1; i <= data.length; i++) {
        const curr = data[i];
        const prev = data[i - 1];
        const isGroupEnd = curr !== prev || i === data.length;

        if (isGroupEnd) {
          const groupStart = start;
          const groupEnd = i + 4;

          if (groupEnd > groupStart) {
            mergeColumns.forEach(col => {
              sheet.getRange(groupStart, col, groupEnd - groupStart + 1).mergeVertically();
            });
          }

          if (groupStart > 5) {
            sheet.getRange(groupStart, 1, 1, lastColumn)
              .setBorder(true, null, null, null, null, null, true, SpreadsheetApp.BorderStyle.DOTTED);
          }

          start = i + 5;
        }
      }

      // 5. Установить фильтр "не пустые"
      filter.setColumnFilterCriteria(
        filterColumn,
        SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build()
      );

      // 6. Обновить лист-копию: форматирование как на =листе, значения без формул
      updateValuesCopy_(spreadsheet, sheet);

      Logger.log(`Обработка листа "${name}" завершена.`);
    });

    Logger.log(`=== Все листы обработаны ===`);
  } catch (e) {
    Logger.log("Ошибка: " + e);
  } finally {
    lock.releaseLock();
  }
}

function updateValuesCopy_(spreadsheet, sourceSheet) {
  const sourceName = sourceSheet.getName();
  const targetName = getValuesCopySheetName_(sourceName);
  let targetSheet = spreadsheet.getSheetByName(targetName);

  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet(targetName);
    Logger.log(`Создан лист-копия "${targetName}" для листа "${sourceName}".`);
  }

  const lastRow = Math.max(sourceSheet.getLastRow(), 1);
  const lastColumn = Math.max(sourceSheet.getLastColumn(), 1);

  resizeSheet_(targetSheet, lastRow, lastColumn);

  const targetFilter = targetSheet.getFilter();
  if (targetFilter) {
    targetFilter.remove();
  }

  const fullTargetRange = targetSheet.getRange(1, 1, targetSheet.getMaxRows(), targetSheet.getMaxColumns());
  fullTargetRange.breakApart();
  fullTargetRange.clear({ contentsOnly: false });

  const sourceRange = sourceSheet.getRange(1, 1, lastRow, lastColumn);
  const targetRange = targetSheet.getRange(1, 1, lastRow, lastColumn);

  sourceRange.copyTo(targetRange, { formatOnly: true });
  targetRange.breakApart();
  targetRange.setValues(sourceRange.getValues());

  copySheetLayout_(sourceSheet, targetSheet, lastRow, lastColumn);
  copyMergedRanges_(sourceSheet, targetSheet, lastRow, lastColumn);

  Logger.log(`Лист-копия "${targetName}" обновлён значениями и форматированием листа "${sourceName}".`);
}

function getValuesCopySheetName_(sourceName) {
  const copyName = sourceName.replace(/^=+/, '').trim();
  return copyName || `${sourceName}_values`;
}

function resizeSheet_(sheet, targetRows, targetColumns) {
  const currentRows = sheet.getMaxRows();
  if (currentRows < targetRows) {
    sheet.insertRowsAfter(currentRows, targetRows - currentRows);
  } else if (currentRows > targetRows) {
    sheet.deleteRows(targetRows + 1, currentRows - targetRows);
  }

  const currentColumns = sheet.getMaxColumns();
  if (currentColumns < targetColumns) {
    sheet.insertColumnsAfter(currentColumns, targetColumns - currentColumns);
  } else if (currentColumns > targetColumns) {
    sheet.deleteColumns(targetColumns + 1, currentColumns - targetColumns);
  }
}

function copySheetLayout_(sourceSheet, targetSheet, lastRow, lastColumn) {
  targetSheet.setFrozenRows(sourceSheet.getFrozenRows());
  targetSheet.setFrozenColumns(sourceSheet.getFrozenColumns());
  targetSheet.setTabColor(sourceSheet.getTabColor());

  for (let row = 1; row <= lastRow; row++) {
    targetSheet.setRowHeight(row, sourceSheet.getRowHeight(row));
  }

  for (let column = 1; column <= lastColumn; column++) {
    targetSheet.setColumnWidth(column, sourceSheet.getColumnWidth(column));
  }
}

function copyMergedRanges_(sourceSheet, targetSheet, lastRow, lastColumn) {
  sourceSheet
    .getRange(1, 1, lastRow, lastColumn)
    .getMergedRanges()
    .forEach(range => {
      targetSheet
        .getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns())
        .merge();
    });
}
