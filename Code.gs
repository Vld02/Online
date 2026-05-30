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
      const copyName = getValueCopySheetName_(name);
      const hasCopySheet = Boolean(spreadsheet.getSheetByName(copyName));

      if (String(currentValue) === String(storedValue)) {
        if (!hasCopySheet) {
          Logger.log(`Значение G4 листа "${name}" не изменилось (${currentValue}), но лист-копия "${copyName}" отсутствует.`);
          updateValueCopySheet_(spreadsheet, sheet);
          return;
        }

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

      updateValueCopySheet_(spreadsheet, sheet);

      Logger.log(`Обработка листа "${name}" завершена.`);
    });

    Logger.log(`=== Все листы обработаны ===`);
  } catch (e) {
    Logger.log("Ошибка: " + e);
  } finally {
    lock.releaseLock();
  }
}

function updateValueCopySheet_(spreadsheet, sourceSheet) {
  const sourceName = sourceSheet.getName();
  const copyName = getValueCopySheetName_(sourceName);
  let copySheet = spreadsheet.getSheetByName(copyName);

  if (!copySheet) {
    copySheet = spreadsheet.insertSheet(copyName);
    Logger.log(`Создан лист-копия "${copyName}" для листа "${sourceName}".`);
  }

  syncSheetSize_(copySheet, sourceSheet.getMaxRows(), sourceSheet.getMaxColumns());

  const copyFilter = copySheet.getFilter();
  if (copyFilter) {
    copyFilter.remove();
  }

  copySheet.getCharts().forEach(chart => copySheet.removeChart(chart));
  copySheet.getImages().forEach(image => image.remove());
  copySheet.getDrawings().forEach(drawing => drawing.remove());

  const maxRows = copySheet.getMaxRows();
  const maxColumns = copySheet.getMaxColumns();
  const fullCopyRange = copySheet.getRange(1, 1, maxRows, maxColumns);
  fullCopyRange.breakApart();
  fullCopyRange.clear();

  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();

  if (lastRow > 0 && lastColumn > 0) {
    const sourceRange = sourceSheet.getRange(1, 1, lastRow, lastColumn);
    const copyRange = copySheet.getRange(1, 1, lastRow, lastColumn);

    copyRange.setValues(sourceRange.getValues());
    sourceRange.copyTo(copyRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    copyRange.breakApart();

    sourceRange.getMergedRanges().forEach(mergedRange => {
      copySheet
        .getRange(
          mergedRange.getRow(),
          mergedRange.getColumn(),
          mergedRange.getNumRows(),
          mergedRange.getNumColumns()
        )
        .merge();
    });
  }

  copySheet.setFrozenRows(sourceSheet.getFrozenRows());
  copySheet.setFrozenColumns(sourceSheet.getFrozenColumns());
  copySheet.setHiddenGridlines(sourceSheet.hasHiddenGridlines());
  copySheet.setTabColor(sourceSheet.getTabColor());

  for (let row = 1; row <= sourceSheet.getMaxRows(); row++) {
    copySheet.setRowHeight(row, sourceSheet.getRowHeight(row));
  }

  for (let column = 1; column <= sourceSheet.getMaxColumns(); column++) {
    copySheet.setColumnWidth(column, sourceSheet.getColumnWidth(column));
  }

  Logger.log(`Лист-копия "${copyName}" обновлён значениями и форматированием листа "${sourceName}".`);
}

function getValueCopySheetName_(sourceName) {
  const baseName = sourceName.replace(/^=+/, '').trim();
  const copyName = baseName || `${sourceName}_values`;
  return copyName.substring(0, 100);
}

function syncSheetSize_(sheet, requiredRows, requiredColumns) {
  const currentRows = sheet.getMaxRows();
  if (currentRows < requiredRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
  } else if (currentRows > requiredRows) {
    sheet.deleteRows(requiredRows + 1, currentRows - requiredRows);
  }

  const currentColumns = sheet.getMaxColumns();
  if (currentColumns < requiredColumns) {
    sheet.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
  } else if (currentColumns > requiredColumns) {
    sheet.deleteColumns(requiredColumns + 1, currentColumns - requiredColumns);
  }
}
