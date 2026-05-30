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

      const copyName = name.substring(1);
      let copySheet = spreadsheet.getSheetByName(copyName);
      const copyMissing = !copySheet;
      const currentValue = sheet.getRange("G4").getValue();
      const storedValue = props.getProperty("lastValue_" + name);

      if (String(currentValue) === String(storedValue) && !copyMissing) {
        Logger.log(`Пропуск листа "${name}" — значение G4 не изменилось (${currentValue}).`);
        return;
      }

      if (copyMissing) {
        copySheet = spreadsheet.insertSheet(copyName);
        Logger.log(`Создан лист-копия "${copyName}" для листа "${name}".`);
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

      updateValuesOnlyCopy_(sheet, copySheet);

      Logger.log(`Обработка листа "${name}" завершена. Лист-копия "${copyName}" обновлён.`);
    });

    Logger.log(`=== Все листы обработаны ===`);
  } catch (e) {
    Logger.log("Ошибка: " + e);
  } finally {
    lock.releaseLock();
  }
}

function updateValuesOnlyCopy_(sourceSheet, copySheet) {
  const lastRow = Math.max(sourceSheet.getLastRow(), 1);
  const lastColumn = Math.max(sourceSheet.getLastColumn(), 1);

  const oldFilter = copySheet.getFilter();
  if (oldFilter) oldFilter.remove();

  copySheet.getRange(1, 1, copySheet.getMaxRows(), copySheet.getMaxColumns()).breakApart();
  syncSheetSize_(copySheet, sourceSheet.getMaxRows(), sourceSheet.getMaxColumns());

  copySheet.clear({ contentsOnly: false });
  copySheet.clearNotes();
  copySheet.getRange(1, 1, copySheet.getMaxRows(), copySheet.getMaxColumns()).clearDataValidations();
  copySheet.clearConditionalFormatRules();

  const sourceRange = sourceSheet.getRange(1, 1, lastRow, lastColumn);
  const targetRange = copySheet.getRange(1, 1, lastRow, lastColumn);

  sourceRange.copyTo(targetRange, { formatOnly: true });
  targetRange.setValues(sourceRange.getValues());

  copySheet.setFrozenRows(sourceSheet.getFrozenRows());
  copySheet.setFrozenColumns(sourceSheet.getFrozenColumns());
  copySheet.setHiddenGridlines(sourceSheet.hasHiddenGridlines());
  copySheet.setTabColor(sourceSheet.getTabColor());

  for (let column = 1; column <= lastColumn; column++) {
    copySheet.setColumnWidth(column, sourceSheet.getColumnWidth(column));
    if (sourceSheet.isColumnHiddenByUser(column)) {
      copySheet.hideColumns(column);
    } else {
      copySheet.showColumns(column);
    }
  }

  for (let row = 1; row <= lastRow; row++) {
    copySheet.setRowHeight(row, sourceSheet.getRowHeight(row));
    if (sourceSheet.isRowHiddenByUser(row)) {
      copySheet.hideRows(row);
    } else {
      copySheet.showRows(row);
    }
  }

  sourceSheet.getRange(1, 1, lastRow, lastColumn).getMergedRanges().forEach(range => {
    copySheet
      .getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns())
      .merge();
  });

  const filter = sourceSheet.getFilter();
  if (filter) {
    const filterRange = filter.getRange();
    copySheet
      .getRange(filterRange.getRow(), filterRange.getColumn(), filterRange.getNumRows(), filterRange.getNumColumns())
      .createFilter();
  }
}

function syncSheetSize_(sheet, rows, columns) {
  const currentRows = sheet.getMaxRows();
  const currentColumns = sheet.getMaxColumns();

  if (currentRows < rows) {
    sheet.insertRowsAfter(currentRows, rows - currentRows);
  } else if (currentRows > rows) {
    sheet.deleteRows(rows + 1, currentRows - rows);
  }

  if (currentColumns < columns) {
    sheet.insertColumnsAfter(currentColumns, columns - currentColumns);
  } else if (currentColumns > columns) {
    sheet.deleteColumns(columns + 1, currentColumns - columns);
  }
}
