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

      updateValuesOnlyCopy_(spreadsheet, sheet);

      Logger.log(`Обработка листа "${name}" завершена.`);
    });

    Logger.log(`=== Все листы обработаны ===`);
  } catch (e) {
    Logger.log("Ошибка: " + e);
  } finally {
    lock.releaseLock();
  }
}

function updateValuesOnlyCopy_(spreadsheet, sourceSheet) {
  const sourceName = sourceSheet.getName();
  const copyName = getValuesOnlyCopyName_(sourceName);

  if (!copyName) {
    Logger.log(`Копия для листа "${sourceName}" не создана: имя копии пустое.`);
    return;
  }

  const valuesOnlyCopy = getOrCreateValuesOnlyCopy_(spreadsheet, sourceSheet, copyName);
  valuesOnlyCopy.showSheet();

  replaceSheetWithValuesAndFormatting_(sourceSheet, valuesOnlyCopy);

  Logger.log(`Копия листа "${sourceName}" обновлена как "${copyName}" — оставлены только значения и форматирование.`);
}

function getOrCreateValuesOnlyCopy_(spreadsheet, sourceSheet, copyName) {
  const existingCopy = spreadsheet.getSheetByName(copyName);

  if (existingCopy) {
    return existingCopy;
  }

  const activeSheet = spreadsheet.getActiveSheet();
  const valuesOnlyCopy = sourceSheet.copyTo(spreadsheet).setName(copyName);

  valuesOnlyCopy.showSheet();
  spreadsheet.setActiveSheet(valuesOnlyCopy);
  spreadsheet.moveActiveSheet(sourceSheet.getIndex() + 1);

  if (activeSheet && !activeSheet.isSheetHidden()) {
    spreadsheet.setActiveSheet(activeSheet);
  }

  Logger.log(`Создана копия листа "${sourceSheet.getName()}" с именем "${copyName}".`);
  return valuesOnlyCopy;
}

function replaceSheetWithValuesAndFormatting_(sourceSheet, targetSheet) {
  const filter = targetSheet.getFilter();
  if (filter) {
    filter.remove();
  }

  targetSheet.getRange(1, 1, targetSheet.getMaxRows(), targetSheet.getMaxColumns()).breakApart();
  targetSheet.clear();

  resizeSheet_(targetSheet, sourceSheet.getMaxRows(), sourceSheet.getMaxColumns());

  const rowCount = sourceSheet.getMaxRows();
  const columnCount = sourceSheet.getMaxColumns();
  const sourceRange = sourceSheet.getRange(1, 1, rowCount, columnCount);
  const targetRange = targetSheet.getRange(1, 1, rowCount, columnCount);

  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  targetSheet.setFrozenRows(sourceSheet.getFrozenRows());
  targetSheet.setFrozenColumns(sourceSheet.getFrozenColumns());
  copyRowHeights_(sourceSheet, targetSheet, rowCount);
  copyColumnWidths_(sourceSheet, targetSheet, columnCount);
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

function copyRowHeights_(sourceSheet, targetSheet, rowCount) {
  for (let row = 1; row <= rowCount; row++) {
    targetSheet.setRowHeight(row, sourceSheet.getRowHeight(row));
  }
}

function copyColumnWidths_(sourceSheet, targetSheet, columnCount) {
  for (let column = 1; column <= columnCount; column++) {
    targetSheet.setColumnWidth(column, sourceSheet.getColumnWidth(column));
  }
}

function getValuesOnlyCopyName_(sourceName) {
  return sourceName.replace(/^=+/, '').trim();
}
