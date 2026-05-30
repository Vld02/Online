/**
 * Обрабатывает листы одной пакетной операцией на лист.
 *
 * Чтобы пользователь не видел промежуточные состояния (снятые объединения,
 * очищенные границы и т.п.), все визуальные изменения для каждого листа
 * собираются в массив requests и применяются одним вызовом Sheets API
 * batchUpdate().
 *
 * В Apps Script нужно включить Advanced Google Services:
 * Services → Google Sheets API → Add.
 */
function processSheets() {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    lock.waitLock(1000);
    lockAcquired = true;

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
    const sheetNames = sheets.map(s => s.getName());
    for (const key in allProps) {
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

      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      if (lastRow < 5 || lastColumn < 1) {
        props.setProperty('lastValue_' + name, String(currentValue));
        Logger.log(`Пропуск визуальной обработки листа "${name}" — недостаточно строк.`);
        return;
      }

      const sheetId = sheet.getSheetId();
      const requests = [];

      // 1. Финальное состояние фильтра: показывать только непустые ячейки в B.
      // Отдельный сброс фильтра больше не нужен: пользователь увидит только итог.
      requests.push({
        setBasicFilter: {
          filter: {
            range: {
              sheetId,
              startRowIndex: filterStartRow - 1,
              endRowIndex: sheet.getMaxRows(),
              startColumnIndex: filterColumn - 1,
              endColumnIndex: filterColumn
            },
            criteria: {
              [filterColumn - 1]: {
                condition: {
                  type: 'NOT_BLANK'
                }
              }
            }
          }
        }
      });

      // 2. Снятие объединений.
      requests.push({
        unmergeCells: {
          range: {
            sheetId,
            startRowIndex: 4,
            endRowIndex: lastRow,
            startColumnIndex: 0,
            endColumnIndex: lastColumn
          }
        }
      });

      // 3. Очистка границ.
      if (lastRow >= 6) {
        requests.push({
          updateBorders: {
            range: {
              sheetId,
              startRowIndex: 5,
              endRowIndex: lastRow,
              startColumnIndex: 0,
              endColumnIndex: lastColumn
            },
            top: {style: 'NONE'},
            bottom: {style: 'NONE'},
            left: {style: 'NONE'},
            right: {style: 'NONE'},
            innerHorizontal: {style: 'NONE'},
            innerVertical: {style: 'NONE'}
          }
        });
      }

      // 4. Группировка по F: заранее рассчитываем все объединения и границы.
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
              if (col <= lastColumn) {
                requests.push({
                  mergeCells: {
                    range: {
                      sheetId,
                      startRowIndex: groupStart - 1,
                      endRowIndex: groupEnd,
                      startColumnIndex: col - 1,
                      endColumnIndex: col
                    },
                    mergeType: 'MERGE_ALL'
                  }
                });
              }
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
                  endColumnIndex: lastColumn
                },
                top: {
                  style: 'DOTTED',
                  color: {red: 0, green: 0, blue: 0}
                }
              }
            });
          }

          start = i + 5;
        }
      }

      if (requests.length > 0) {
        Sheets.Spreadsheets.batchUpdate({requests}, spreadsheetId);
      }

      // Обновляем lastValue_ только после успешного batchUpdate().
      props.setProperty('lastValue_' + name, String(currentValue));
      Logger.log(`Обработка листа "${name}" завершена одним batchUpdate.`);
    });

    Logger.log('=== Все листы обработаны ===');
  } catch (e) {
    Logger.log('Ошибка: ' + e);
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}
