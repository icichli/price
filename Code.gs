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
      let currentSnapshot = [];
      let previousSnapshot = [];
      
      try { mapping = row[3] ? JSON.parse(row[3]) : null; } catch(err) {}
      try { currentSnapshot = row[4] ? JSON.parse(row[4]) : []; } catch(err) {}
      try { previousSnapshot = row[5] ? JSON.parse(row[5]) : []; } catch(err) {}
      
      suppliers.push({
        id: String(supplierId),
        name: String(row[1] || ''),
        updatedAt: row[2] ? String(row[2]) : '',
        mapping: mapping,
        currentSnapshot: currentSnapshot,
        previousSnapshot: previousSnapshot
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
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    // 1. Проверка соединения
    if (action === 'test_connection') {
      return createJsonResponse({
        success: true,
        message: 'Соединение с Google Apps Script успешно установлено!'
      });
    }
    
    // 2. Сохранение нового снимка (save_snapshot)
    if (action === 'save_snapshot') {
      const supplierId = String(payload.supplierId || '');
      const supplierName = String(payload.supplierName || 'Поставщик');
      const newSnapshot = payload.snapshot || [];
      const mapping = payload.mapping || null;
      const nowIso = new Date().toISOString();
      
      if (!supplierId) {
        return createJsonResponse({ success: false, error: 'supplierId обязателен' });
      }
      
      let targetRowIndex = -1;
      let existingCurrentJson = '[]';
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === supplierId) {
          targetRowIndex = i + 1; // в Sheets строки с 1
          existingCurrentJson = data[i][4] || '[]';
          break;
        }
      }
      
      const newSnapshotJson = JSON.stringify(newSnapshot);
      const mappingJson = mapping ? JSON.stringify(mapping) : '';
      
      if (targetRowIndex > 0) {
        // Переносим текущее в предыдущее (откат на 1 шаг), пишем новое в текущее
        sheet.getRange(targetRowIndex, 2).setValue(supplierName);
        sheet.getRange(targetRowIndex, 3).setValue(nowIso);
        if (mappingJson) {
          sheet.getRange(targetRowIndex, 4).setValue(mappingJson);
        }
        sheet.getRange(targetRowIndex, 5).setValue(newSnapshotJson);
        sheet.getRange(targetRowIndex, 6).setValue(existingCurrentJson);
      } else {
        // Новый поставщик
        sheet.appendRow([
          supplierId,
          supplierName,
          nowIso,
          mappingJson,
          newSnapshotJson,
          '[]'
        ]);
      }
      
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
      let previousSnapshotJson = '';
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === supplierId) {
          targetRowIndex = i + 1;
          previousSnapshotJson = data[i][5] || '';
          break;
        }
      }
      
      if (targetRowIndex === -1) {
        return createJsonResponse({ success: false, error: 'Поставщик не найден' });
      }
      
      if (!previousSnapshotJson || previousSnapshotJson === '[]') {
        return createJsonResponse({ success: false, error: 'Предыдущий снимок отсутствует. Откат невозможен.' });
      }
      
      const nowIso = new Date().toISOString();
      // Восстанавливаем PreviousSnapshot в CurrentSnapshot
      sheet.getRange(targetRowIndex, 3).setValue(nowIso);
      sheet.getRange(targetRowIndex, 5).setValue(previousSnapshotJson);
      sheet.getRange(targetRowIndex, 6).setValue(''); // очищаем previous после отката
      
      let restoredData = [];
      try { restoredData = JSON.parse(previousSnapshotJson); } catch (e) {}
      
      return createJsonResponse({
        success: true,
        updatedAt: nowIso,
        currentSnapshot: restoredData,
        message: 'База успешно откачена к предыдущему снимку'
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
        }
        return createJsonResponse({ success: true, message: 'Поставщик удален' });
      }
      
      if (subAction === 'rename') {
        if (targetRowIndex > 0) {
          sheet.getRange(targetRowIndex, 2).setValue(supplierName);
        }
        return createJsonResponse({ success: true, message: 'Поставщик переименован' });
      }
      
      if (subAction === 'create') {
        if (targetRowIndex === -1) {
          const nowIso = new Date().toISOString();
          sheet.appendRow([supplierId, supplierName, nowIso, '', '[]', '[]']);
        }
        return createJsonResponse({ success: true, message: 'Поставщик создан' });
      }
      
      return createJsonResponse({ success: false, error: 'Неизвестное действие: ' + subAction });
    }
    
    return createJsonResponse({ success: false, error: 'Неизвестное действие: ' + action });
    
  } catch (err) {
    return createJsonResponse({
      success: false,
      error: 'Ошибка при обработке запроса: ' + err.toString()
    });
  }
}
