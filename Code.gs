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
      if (key.startsWith('lastValue_')) {
        const sheetName = key.replace('lastValue_', '');
        if (!sheetNames.includes(sheetName)) {
          props.deleteProperty(key);
          Logger.log(`Удалён ключ "${key}", так как лист "${sheetName}" больше не существует.`);
        }
      }
    }

    // Показать, что осталось после очистки
    const remainingProps = props.getProperties();
    const keys = Object.keys(remainingProps).filter(k => k.startsWith('lastValue_'));
    if (keys.length > 0) {
      Logger.log('Актуальные ключи в PropertiesService:');
      keys.forEach(k => Logger.log(` • ${k} = ${remainingProps[k]}`));
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

      Logger.log(`Обработка листа "${name}". Старое значение: ${storedValue}, новое: ${currentValue}`);
      props.setProperty('lastValue_' + name, String(currentValue));

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

    Logger.log('=== Все листы обработаны ===');
  } catch (e) {
    Logger.log('Ошибка: ' + e);
  } finally {
    lock.releaseLock();
  }
}

function updateValuesOnlyCopy_(spreadsheet, sourceSheet) {
  const sourceName = sourceSheet.getName();
  const copyName = getValuesOnlyCopyName_(sourceName);
  let copySheet = spreadsheet.getSheetByName(copyName);

  if (!copySheet) {
    copySheet = spreadsheet.insertSheet(copyName);
    Logger.log(`Создан лист-копия "${copyName}" для листа "${sourceName}".`);
  }

  const sourceRows = sourceSheet.getMaxRows();
  const sourceColumns = sourceSheet.getMaxColumns();
  resizeSheetGrid_(copySheet, sourceRows, sourceColumns);

  const existingFilter = copySheet.getFilter();
  if (existingFilter) existingFilter.remove();

  copySheet.clear({ contentsOnly: false });
  copySheet.clearConditionalFormatRules();
  copySheet.getRange(1, 1, copySheet.getMaxRows(), copySheet.getMaxColumns()).breakApart();

  const sourceRange = sourceSheet.getRange(1, 1, sourceRows, sourceColumns);
  const copyRange = copySheet.getRange(1, 1, sourceRows, sourceColumns);

  copyRange.setValues(sourceRange.getValues());
  sourceRange.copyTo(copyRange, { formatOnly: true });
  copyMergedRanges_(sourceRange, copySheet);
  copySheetSettings_(sourceSheet, copySheet, sourceRows, sourceColumns);
  copyConditionalFormatRules_(sourceSheet, copySheet);
  copyFilter_(sourceSheet, copySheet);

  Logger.log(`Лист-копия "${copyName}" обновлён значениями и форматированием листа "${sourceName}".`);
}

function getValuesOnlyCopyName_(sourceName) {
  const copyName = sourceName.replace(/^=+/, '').trim();
  return copyName || sourceName + ' values';
}

function resizeSheetGrid_(sheet, targetRows, targetColumns) {
  const rows = sheet.getMaxRows();
  if (rows < targetRows) {
    sheet.insertRowsAfter(rows, targetRows - rows);
  } else if (rows > targetRows) {
    sheet.deleteRows(targetRows + 1, rows - targetRows);
  }

  const columns = sheet.getMaxColumns();
  if (columns < targetColumns) {
    sheet.insertColumnsAfter(columns, targetColumns - columns);
  } else if (columns > targetColumns) {
    sheet.deleteColumns(targetColumns + 1, columns - targetColumns);
  }
}

function copyMergedRanges_(sourceRange, copySheet) {
  sourceRange.getMergedRanges().forEach(range => {
    copySheet
      .getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns())
      .merge();
  });
}

function copySheetSettings_(sourceSheet, copySheet, sourceRows, sourceColumns) {
  copySheet.setFrozenRows(sourceSheet.getFrozenRows());
  copySheet.setFrozenColumns(sourceSheet.getFrozenColumns());
  copySheet.setRightToLeft(sourceSheet.isRightToLeft());
  copySheet.setHiddenGridlines(sourceSheet.hasHiddenGridlines());

  const tabColor = sourceSheet.getTabColor();
  copySheet.setTabColor(tabColor);

  for (let row = 1; row <= sourceRows; row++) {
    copySheet.setRowHeight(row, sourceSheet.getRowHeight(row));
    if (sourceSheet.isRowHiddenByUser(row)) {
      copySheet.hideRows(row);
    } else {
      copySheet.showRows(row);
    }
  }

  for (let column = 1; column <= sourceColumns; column++) {
    copySheet.setColumnWidth(column, sourceSheet.getColumnWidth(column));
    if (sourceSheet.isColumnHiddenByUser(column)) {
      copySheet.hideColumns(column);
    } else {
      copySheet.showColumns(column);
    }
  }
}

function copyConditionalFormatRules_(sourceSheet, copySheet) {
  const copiedRules = sourceSheet.getConditionalFormatRules().map(rule => {
    const copiedRanges = rule.getRanges().map(range => copySheet.getRange(
      range.getRow(),
      range.getColumn(),
      range.getNumRows(),
      range.getNumColumns()
    ));
    return rule.copy().setRanges(copiedRanges).build();
  });

  copySheet.setConditionalFormatRules(copiedRules);
}

function copyFilter_(sourceSheet, copySheet) {
  const sourceFilter = sourceSheet.getFilter();
  if (!sourceFilter) return;

  const sourceRange = sourceFilter.getRange();
  const copyRange = copySheet.getRange(
    sourceRange.getRow(),
    sourceRange.getColumn(),
    sourceRange.getNumRows(),
    sourceRange.getNumColumns()
  );
  copyRange.createFilter();

  const copyFilter = copySheet.getFilter();
  const firstColumn = sourceRange.getColumn();
  const lastColumn = sourceRange.getLastColumn();

  for (let column = firstColumn; column <= lastColumn; column++) {
    const criteria = sourceFilter.getColumnFilterCriteria(column);
    if (criteria) {
      copyFilter.setColumnFilterCriteria(column, criteria.copy().build());
    }
  }
}
