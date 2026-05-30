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

      Logger.log(`Обработка листа "${name}" завершена.`);
    });

    // --- Создание листов-копий без формул ---
    createValueOnlyCopies_(spreadsheet);

    Logger.log('=== Все листы обработаны ===');
  } catch (e) {
    Logger.log('Ошибка: ' + e);
  } finally {
    lock.releaseLock();
  }
}

function createValueOnlyCopies_(spreadsheet) {
  const props = PropertiesService.getScriptProperties();
  const copyNamesProperty = 'valueOnlyCopyNames';
  const sourceSheets = spreadsheet.getSheets().filter(sheet => sheet.getName().startsWith('='));
  const sourceNames = sourceSheets.map(sheet => sheet.getName());
  const targetNames = sourceSheets
    .map(sheet => sheet.getName().substring(1))
    .filter(name => name);
  const activeSheet = spreadsheet.getActiveSheet();

  // Удаляем копии, которые были созданы прошлым запуском, но для них больше нет исходного =листа.
  const previousCopyNames = JSON.parse(props.getProperty(copyNamesProperty) || '[]');
  previousCopyNames.forEach(name => {
    if (targetNames.includes(name) || sourceNames.includes(name)) return;

    const oldCopy = spreadsheet.getSheetByName(name);
    if (oldCopy && spreadsheet.getSheets().length > 1) {
      spreadsheet.deleteSheet(oldCopy);
      Logger.log(`Удалена устаревшая копия "${name}".`);
    }
  });

  sourceSheets.forEach(sourceSheet => {
    const sourceName = sourceSheet.getName();
    const targetName = sourceName.substring(1);
    if (!targetName) {
      Logger.log(`Пропуск листа "${sourceName}" — имя копии без "=" пустое.`);
      return;
    }

    const oldTargetSheet = spreadsheet.getSheetByName(targetName);
    const sourceIndex = sourceSheet.getIndex();

    // Самый надёжный способ сохранить всё форматирование, размеры, фильтры и объединения —
    // создать копию листа, а затем заменить формулы в ней на текущие значения.
    if (oldTargetSheet && spreadsheet.getSheets().length > 1) {
      spreadsheet.deleteSheet(oldTargetSheet);
    }

    const targetSheet = sourceSheet.copyTo(spreadsheet).setName(targetName);
    spreadsheet.setActiveSheet(targetSheet);
    spreadsheet.moveActiveSheet(sourceIndex + 1);

    const range = targetSheet.getDataRange();
    range.copyTo(range, { contentsOnly: true });

    Logger.log(`Создана копия "${targetName}" без формул для листа "${sourceName}".`);
  });

  props.setProperty(copyNamesProperty, JSON.stringify(targetNames));

  if (activeSheet && spreadsheet.getSheetByName(activeSheet.getName())) {
    spreadsheet.setActiveSheet(activeSheet);
  } else if (sourceSheets.length > 0 && spreadsheet.getSheetByName(sourceSheets[0].getName())) {
    spreadsheet.setActiveSheet(sourceSheets[0]);
  }
}
