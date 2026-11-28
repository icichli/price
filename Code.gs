/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT: Серверная синхронизация для веб-приложения сравнения прайсов
 * ==============================================================================
 * 
 * ПОШАГОВАЯ ИНСТРУКЦИЯ ПО РАЗВЕРТЫВАНИЮ:
 * 1. Создайте новую Google Таблицу (или откройте существующую) на Google Диске.
 * 2. В верхнем меню выберите: «Расширения» -> «Apps Script» (Extensions -> Apps Script).
 * 3. Полностью удалите стандартный код из файла Code.gs и вставьте данный код целиком.
 * 4. Нажмите синюю кнопку «Развернуть» (Deploy) в правом верхнем углу -> «Новое развертывание» (New deployment).
 * 5. Нажмите на значок шестеренки «Выберите тип» -> выберите «Веб-приложение» (Web app).
 * 6. Заполните параметры развертывания:
 *    - Описание: «Синхронизация прайсов» (или любое другое)
 *    - Выполнять от имени: «От моего имени» (Me / ваш аккаунт)
 *    - У КОГО ЕСТЬ ДОСТУП: «ВСЕ» (Anyone) — КРИТИЧЕСКИ ВАЖНО!
 *      (Это позволяет веб-приложению GitHub Pages обращаться к скрипту без сложной Google OAuth авторизации).
 * 7. Нажмите «Развернуть» (Deploy). При первом запуске Google запросит разрешения на доступ к таблице — предоставьте их.
 * 8. Скопируйте полученный «URL веб-приложения» (вида: https://script.google.com/macros/s/.../exec).
 * 9. Вставьте этот URL и ваш придуманный секретный PIN/токен в веб-приложении в модальном окне «Настройки облака».
 * ==============================================================================
 */

// Имя листа для хранения данных
const SHEET_NAME = 'Stock_Snapshots';

// Заголовки таблицы
const HEADERS = [
  'SupplierID',
  'SupplierName',
  'UpdatedAt',
  'MappingJSON',
  'CurrentSnapshotJSON',
  'PreviousSnapshotJSON'
];

/**
 * Получение или автоматическое создание технического листа Stock_Snapshots
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    
    // Форматирование шапки
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#1e293b');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setColumnWidth(1, 120); // SupplierID
    sheet.setColumnWidth(2, 180); // SupplierName
    sheet.setColumnWidth(3, 160); // UpdatedAt
    sheet.setColumnWidth(4, 200); // MappingJSON
    sheet.setColumnWidth(5, 300); // CurrentSnapshotJSON
    sheet.setColumnWidth(6, 300); // PreviousSnapshotJSON
  }
  
  return sheet;
}

/**
 * Валидация токена доступа.
 * Токен можно сохранить в Script Properties (ключ 'SECRET_TOKEN').
 * Если свойство 'SECRET_TOKEN' еще не создано, скрипт автоматически запоминает
 * первый переданный непустой токен в свойства скрипта.
 */
function validateToken(token) {
  if (!token || typeof token !== 'string' || token.trim() === '') {
    return false;
  }
  
  const props = PropertiesService.getScriptProperties();
  let savedToken = props.getProperty('SECRET_TOKEN');
  
  if (!savedToken) {
    // Автоматическая инициализация токена при первой настройке
    props.setProperty('SECRET_TOKEN', token.trim());
    return true;
  }
  
  return savedToken === token.trim();
}

/**
 * Вспомогательный метод создания JSON-ответа
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Безопасное форматирование даты/метки времени из ячейки Google Sheets в ISO-строку
 */
function formatTimestamp(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? '' : val.toISOString();
  }
  return String(val).trim();
}

/**
 * Безопасная запись длинных строк JSON в ячейки с обходом лимита 50 000 символов.
 * Поддерживает произвольное количество чанков (не только 2):
 * - Основная колонка col (5 для Current, 6 для Previous)
 * - Колонки переполнения: col + 2, col + 4, col + 6, ...
 *   (Current -> 7, 9, 11...; Previous -> 8, 10, 12...)
 * Полностью обратно совместимо со старым форматом CHUNK:: (2 чанка).
 */
function setCellSafe(sheet, row, col, str) {
  const MAX_SINGLE = 48000;
  const CHUNK_SIZE = 45000;
  const lastCol = Math.max(sheet.getLastColumn(), col + 2);

  if (!str) {
    sheet.getRange(row, col).setNumberFormat('@').setValue('');
    clearOverflowCols(sheet, row, col, 1, lastCol);
    return;
  }

  if (str.length <= MAX_SINGLE) {
    sheet.getRange(row, col).setNumberFormat('@').setValue(str);
    clearOverflowCols(sheet, row, col, 1, lastCol);
    return;
  }

  // Разбиваем строку на произвольное число чанков по CHUNK_SIZE символов
  const chunks = [];
  for (let offset = 0; offset < str.length; offset += CHUNK_SIZE) {
    chunks.push(str.substring(offset, offset + CHUNK_SIZE));
  }

  const maxNeededCol = col + (chunks.length - 1) * 2;
  if (maxNeededCol > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxNeededCol - sheet.getMaxColumns() + 2);
  }

  // Сначала записываем колонки переполнения (1..N-1), затем основную ячейку с заголовком CHUNKS::N::
  for (let k = 1; k < chunks.length; k++) {
    const overflowCol = col + k * 2;
    sheet.getRange(row, overflowCol).setNumberFormat('@').setValue(chunks[k]);
  }
  sheet.getRange(row, col).setNumberFormat('@').setValue('CHUNKS::' + chunks.length + '::' + chunks[0]);

  // Очищаем возможные лишние старые колонки переполнения
  clearOverflowCols(sheet, row, col, chunks.length, lastCol);
}

/**
 * Очистка неиспользуемых колонок переполнения для заданной основной колонки
 */
function clearOverflowCols(sheet, row, col, startStep, lastCol) {
  const maxCols = sheet.getMaxColumns();
  const limitCol = Math.min(lastCol, maxCols);
  for (let c = col + startStep * 2; c <= limitCol; c += 2) {
    const cell = sheet.getRange(row, c);
    if (cell.getValue() !== '') {
      cell.setValue('');
    }
  }
}

/**
 * Безопасное чтение длинной строки JSON из ячейки и ее колонок переполнения.
 * Возвращает { text, corrupted }.
 */
function getCellSafeInfo(row, colZeroBased) {
  const val = row[colZeroBased] !== undefined && row[colZeroBased] !== null ? String(row[colZeroBased]) : '';

  if (val.startsWith('CHUNKS::')) {
    const secondSep = val.indexOf('::', 8);
    if (secondSep === -1) {
      return { text: '', corrupted: true };
    }
    const totalChunks = parseInt(val.substring(8, secondSep), 10);
    if (isNaN(totalChunks) || totalChunks < 1) {
      return { text: '', corrupted: true };
    }
    let result = val.substring(secondSep + 2);
    for (let k = 1; k < totalChunks; k++) {
      const idx = colZeroBased + k * 2;
      const part = (idx < row.length && row[idx] !== undefined && row[idx] !== null) ? String(row[idx]) : '';
      if (!part) {
        return { text: result, corrupted: true };
      }
      result += part;
    }
    return { text: result, corrupted: false };
  }

  if (val.startsWith('CHUNK::')) {
    const part1 = val.substring(7);
    const idx = colZeroBased + 2;
    const part2 = (idx < row.length && row[idx] !== undefined && row[idx] !== null) ? String(row[idx]) : '';
    return { text: part1 + part2, corrupted: false };
  }

  return { text: val, corrupted: false };
}

function getCellSafe(row, colZeroBased) {
  return getCellSafeInfo(row, colZeroBased).text;
}

/**
 * Безопасный разбор JSON снимка с возвратом флага повреждения
 */
function parseSnapshotJsonSafe(info) {
  if (!info || !info.text) {
    return { data: [], corrupted: Boolean(info && info.corrupted) };
  }
  try {
    const parsed = JSON.parse(info.text);
    return { data: parsed || [], corrupted: Boolean(info.corrupted) };
  } catch (err) {
    return { data: [], corrupted: true };
  }
}

/**
 * Проверка, содержит ли JSON-строка непустой снимок (массив или объект v2 с непустым d)
 */
function isNonEmptySnapshotJson(jsonStr) {
  if (!jsonStr || jsonStr === '[]' || jsonStr === '{}') return false;
  try {
    const parsed = JSON.parse(jsonStr);
    if (Array.isArray(parsed)) {
      if (parsed.length === 0) return false;
      if (parsed[0] && !Array.isArray(parsed[0]) && parsed[0]._v === 2) {
        return parsed.length > 1;
      }
      return true;
    }
    if (parsed && typeof parsed === 'object' && Array.isArray(parsed.d)) return parsed.d.length > 0;
    return false;
  } catch (err) {
    return false;
  }
}

/**
 * Обработчик GET-запросов (чтение базы)
 * URL: https://script.google.com/macros/s/.../exec?token=YOUR_TOKEN
 */
function doGet(e) {
  try {
    const token = e && e.parameter ? e.parameter.token : null;
    
    if (!validateToken(token)) {
      return createJsonResponse({
        success: false,
        error: 'Неверный или отсутствующий токен доступа (PIN)'
      });
    }
    
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return createJsonResponse({
        success: true,
        suppliers: []
      });
    }
    
    const suppliers = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const supplierId = row[0];
      if (!supplierId) continue;
      
      let mapping = null;
      const currentInfo = getCellSafeInfo(row, 4);
      const previousInfo = getCellSafeInfo(row, 5);
      const parsedCurrent = parseSnapshotJsonSafe(currentInfo);
      const parsedPrevious = parseSnapshotJsonSafe(previousInfo);
      
      try { mapping = row[3] ? JSON.parse(String(row[3])) : null; } catch(err) {}
      
      suppliers.push({
        id: String(supplierId),
        name: String(row[1] || ''),
        updatedAt: formatTimestamp(row[2]),
        mapping: mapping,
        currentSnapshot: parsedCurrent.data,
        previousSnapshot: parsedPrevious.data,
        corruptedCurrent: parsedCurrent.corrupted,
        corruptedPrevious: parsedPrevious.corrupted
      });
    }
    
    return createJsonResponse({
      success: true,
      suppliers: suppliers
    });
    
  } catch (err) {
    return createJsonResponse({
      success: false,
      error: 'Ошибка сервера: ' + err.toString()
    });
  }
}

/**
 * Обработчик POST-запросов (запись, откат, управление поставщиками)
 * Запросы принимаются с Content-Type: text/plain;charset=utf-8
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ success: false, error: 'Пустое тело запроса' });
    }
    
    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (err) {
      return createJsonResponse({ success: false, error: 'Ошибка разбора JSON: ' + err.toString() });
    }
    
    const token = payload.token;
    if (!validateToken(token)) {
      return createJsonResponse({ success: false, error: 'Неверный или отсутствующий токен доступа (PIN)' });
    }
    
    const action = payload.action;
    
    // 1. Проверка соединения (не требует блокировки таблицы)
    if (action === 'test_connection') {
      return createJsonResponse({
        success: true,
        message: 'Соединение с Google Apps Script успешно установлено!'
      });
    }

    // Блокировка на время записи в таблицу для защиты от параллельных запросов (гонок)
    const lock = LockService.getScriptLock();
    const gotLock = lock.tryLock(15000);
    if (!gotLock) {
      return createJsonResponse({
        success: false,
        error: 'Таблица временно занята другой операцией синхронизации. Повторите через пару секунд.'
      });
    }

    try {
      const sheet = getOrCreateSheet();
      const data = sheet.getDataRange().getValues();
      
      // 2. Сохранение нового снимка (save_snapshot)
      if (action === 'save_snapshot') {
        const supplierId = String(payload.supplierId || '');
        const supplierName = String(payload.supplierName || 'Поставщик');
        const newSnapshot = payload.snapshot !== undefined ? payload.snapshot : [];
        const hasExplicitPrevious = payload.previousSnapshot !== undefined && payload.previousSnapshot !== null;
        const explicitPreviousJson = hasExplicitPrevious ? JSON.stringify(payload.previousSnapshot) : null;
        const syncMode = payload.syncMode === 'sync' ? 'sync' : 'commit';
        const mapping = payload.mapping || null;
        const nowIso = new Date().toISOString();
        
        if (!supplierId) {
          return createJsonResponse({ success: false, error: 'supplierId обязателен' });
        }
        
        let targetRowIndex = -1;
        let existingCurrentInfo = { text: '[]', corrupted: false };
        let existingPreviousInfo = { text: '[]', corrupted: false };
        
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === supplierId) {
            targetRowIndex = i + 1; // в Sheets строки с 1
            existingCurrentInfo = getCellSafeInfo(data[i], 4);
            existingPreviousInfo = getCellSafeInfo(data[i], 5);
            break;
          }
        }
        
        const newSnapshotJson = JSON.stringify(newSnapshot);
        const mappingJson = mapping ? JSON.stringify(mapping) : '';
        
        // Вычисляем итоговый PreviousSnapshotJSON безопасно (без затирания истории при повторном пуше того же снимка)
        let nextPreviousJson = existingPreviousInfo.corrupted ? '[]' : (existingPreviousInfo.text || '[]');
        
        if (syncMode === 'sync') {
          // Режим синхронизации/восстановления: сохраняем переданный локальный previous (если передан),
          // и никогда не сдвигаем current в previous!
          if (explicitPreviousJson !== null) {
            nextPreviousJson = explicitPreviousJson;
          }
        } else {
          // Режим commit: переносим старый current в previous ТОЛЬКО если он валидный, непустой и отличается от нового
          const existingCurrentValid = !existingCurrentInfo.corrupted && isNonEmptySnapshotJson(existingCurrentInfo.text);
          if (existingCurrentValid && existingCurrentInfo.text !== newSnapshotJson) {
            nextPreviousJson = existingCurrentInfo.text;
          } else if (!existingCurrentValid && explicitPreviousJson !== null && isNonEmptySnapshotJson(explicitPreviousJson)) {
            nextPreviousJson = explicitPreviousJson;
          }
        }
        
        if (targetRowIndex > 0) {
          // Атомарный порядок записи: сначала все данные и чанки, и ТОЛЬКО В КОНЦЕ — UpdatedAt!
          sheet.getRange(targetRowIndex, 2).setValue(supplierName);
          if (mappingJson) {
            sheet.getRange(targetRowIndex, 4).setNumberFormat('@').setValue(mappingJson);
          }
          setCellSafe(sheet, targetRowIndex, 6, nextPreviousJson);
          setCellSafe(sheet, targetRowIndex, 5, newSnapshotJson);
          sheet.getRange(targetRowIndex, 3).setNumberFormat('@').setValue(nowIso);
        } else {
          // Новый поставщик: создаем строку с пустым UpdatedAt, записываем чанки, затем ставим UpdatedAt
          sheet.appendRow([
            supplierId,
            supplierName,
            '',
            mappingJson,
            '',
            ''
          ]);
          const newRowIndex = sheet.getLastRow();
          setCellSafe(sheet, newRowIndex, 6, explicitPreviousJson !== null ? explicitPreviousJson : '[]');
          setCellSafe(sheet, newRowIndex, 5, newSnapshotJson);
          sheet.getRange(newRowIndex, 3).setNumberFormat('@').setValue(nowIso);
        }
        
        SpreadsheetApp.flush();
        
        return createJsonResponse({
          success: true,
          updatedAt: nowIso,
          message: 'Снимок успешно сохранен в Google Sheets'
        });
      }
      
      // 3. Откат к предыдущему снимку (rollback)
      if (action === 'rollback') {
        const supplierId = String(payload.supplierId || '');
        if (!supplierId) {
          return createJsonResponse({ success: false, error: 'supplierId обязателен' });
        }
        
        let targetRowIndex = -1;
        let previousSnapshotInfo = { text: '', corrupted: false };
        
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === supplierId) {
            targetRowIndex = i + 1;
            previousSnapshotInfo = getCellSafeInfo(data[i], 5);
            break;
          }
        }
        
        if (targetRowIndex === -1) {
          return createJsonResponse({ success: false, error: 'Поставщик не найден' });
        }
        
        const previousSnapshotJson = previousSnapshotInfo.text;
        if (previousSnapshotInfo.corrupted || !isNonEmptySnapshotJson(previousSnapshotJson)) {
          return createJsonResponse({ success: false, error: 'Предыдущий снимок отсутствует или поврежден. Откат невозможен.' });
        }
        
        const nowIso = new Date().toISOString();
        // Восстанавливаем PreviousSnapshot в CurrentSnapshot, UpdatedAt обновляем в конце
        setCellSafe(sheet, targetRowIndex, 5, previousSnapshotJson);
        setCellSafe(sheet, targetRowIndex, 6, '[]');
        sheet.getRange(targetRowIndex, 3).setNumberFormat('@').setValue(nowIso);
        SpreadsheetApp.flush();
        
        let restoredData = [];
        try { restoredData = JSON.parse(previousSnapshotJson); } catch (e) {}
        
        return createJsonResponse({
          success: true,
          updatedAt: nowIso,
          currentSnapshot: restoredData,
          message: 'База успешно откачена к предыдущему снимку'
        });
      }
      
      // 3.1. Очистка базы снимка поставщика (clear_snapshot)
      if (action === 'clear_snapshot') {
        const supplierId = String(payload.supplierId || '');
        if (!supplierId) {
          return createJsonResponse({ success: false, error: 'supplierId обязателен' });
        }
        
        let targetRowIndex = -1;
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === supplierId) {
            targetRowIndex = i + 1;
            break;
          }
        }
        
        const nowIso = new Date().toISOString();
        if (targetRowIndex > 0) {
          setCellSafe(sheet, targetRowIndex, 5, '[]');
          setCellSafe(sheet, targetRowIndex, 6, '[]');
          
          // Очищаем lastFileName в MappingJSON, если он был сохранен
          let mapping = null;
          try { mapping = data[targetRowIndex - 1][3] ? JSON.parse(String(data[targetRowIndex - 1][3])) : null; } catch (e) {}
          if (mapping) {
            delete mapping.lastFileName;
            delete mapping.lastCommitTime;
            sheet.getRange(targetRowIndex, 4).setNumberFormat('@').setValue(JSON.stringify(mapping));
          }
          sheet.getRange(targetRowIndex, 3).setNumberFormat('@').setValue(nowIso);
          SpreadsheetApp.flush();
        }
        
        return createJsonResponse({
          success: true,
          updatedAt: nowIso,
          message: 'База остатков поставщика успешно очищена в Google Sheets'
        });
      }
      
      // 4. Управление поставщиками (создание, переименование, удаление)
      if (action === 'manage_supplier') {
        const subAction = payload.subAction; // 'create', 'rename', 'delete'
        const supplierId = String(payload.supplierId || '');
        const supplierName = String(payload.supplierName || '');
        
        if (!supplierId) {
          return createJsonResponse({ success: false, error: 'supplierId обязателен' });
        }
        
        let targetRowIndex = -1;
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === supplierId) {
            targetRowIndex = i + 1;
            break;
          }
        }
        
        if (subAction === 'delete') {
          if (targetRowIndex > 0) {
            sheet.deleteRow(targetRowIndex);
            SpreadsheetApp.flush();
          }
          return createJsonResponse({ success: true, message: 'Поставщик удален' });
        }
        
        if (subAction === 'rename') {
          if (targetRowIndex > 0) {
            sheet.getRange(targetRowIndex, 2).setValue(supplierName);
            SpreadsheetApp.flush();
          }
          return createJsonResponse({ success: true, message: 'Поставщик переименован' });
        }
        
        if (subAction === 'create') {
          if (targetRowIndex === -1) {
            const nowIso = new Date().toISOString();
            sheet.appendRow([supplierId, supplierName, '', '', '[]', '[]']);
            const newRowIndex = sheet.getLastRow();
            sheet.getRange(newRowIndex, 3).setNumberFormat('@').setValue(nowIso);
            SpreadsheetApp.flush();
          }
          return createJsonResponse({ success: true, message: 'Поставщик создан' });
        }
        
        return createJsonResponse({ success: false, error: 'Неизвестное действие: ' + subAction });
      }
      
      return createJsonResponse({ success: false, error: 'Неизвестное действие: ' + action });
    } finally {
      lock.releaseLock();
    }
    
  } catch (err) {
    return createJsonResponse({
      success: false,
      error: 'Ошибка при обработке запроса: ' + err.toString()
    });
  }
}
