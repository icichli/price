/**
 * storage.js - Управление локальным хранилищем (IndexedDB, localStorage) и облачной синхронизацией (Google Apps Script)
 */

class StorageManager {
  constructor() {
    this.dbName = 'PriceCompareDB';
    this.dbVersion = 1;
    this.db = null;
    this.syncListeners = [];
    this.syncStatus = 'offline'; // 'synced' | 'syncing' | 'offline' | 'error' | 'unconfigured'
    this.lastSyncMessage = '';
    // Предустановленный URL Google Apps Script для автоматического подключения
    this.DEFAULT_CLOUD_URL = 'https://script.google.com/macros/s/AKfycbzbltwk1a7gvStn-K0F82UVn3kqtnJzYSQOhZSVyPIen8Mtap0_QYkfMZGzLLmFpuUGrA/exec';
  }

  /**
   * Инициализация IndexedDB и дефолтных данных
   */
  async init() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, this.dbVersion);

      request.onupgradeneeded = (event) => {
        const db = event.target.result;

        // Хранилище поставщиков: { id, name, order, mapping, updatedAt }
        if (!db.objectStoreNames.contains('suppliers')) {
          const supplierStore = db.createObjectStore('suppliers', { keyPath: 'id' });
          supplierStore.createIndex('order', 'order', { unique: false });
        }

        // Хранилище снимков остатков: { supplierId, current, previous, updatedAt }
        if (!db.objectStoreNames.contains('snapshots')) {
          db.createObjectStore('snapshots', { keyPath: 'supplierId' });
        }
      };

      request.onsuccess = async (event) => {
        this.db = event.target.result;
        await this._seedDefaultSuppliers();
        resolve(this.db);
      };

      request.onerror = (event) => {
        console.error('IndexedDB error:', event.target.error);
        reject(event.target.error);
      };
    });
  }

  /**
   * Создание поставщиков «ПП» и «Бир и Ко» по умолчанию
   */
  async _seedDefaultSuppliers() {
    const suppliers = await this.getAllSuppliers();
    if (!suppliers || suppliers.length === 0) {
      const defaultSuppliers = [
        { id: 'supplier_1', name: 'ПП', order: 0, mapping: null, updatedAt: new Date().toISOString() },
        { id: 'supplier_2', name: 'Бир и Ко', order: 1, mapping: null, updatedAt: new Date().toISOString() }
      ];

      for (const sup of defaultSuppliers) {
        await this.saveSupplier(sup);
        await this.saveSnapshotRecord({
          supplierId: sup.id,
          current: [],
          previous: [],
          updatedAt: sup.updatedAt
        });
      }
    } else {
      // Автоматическое переименование дефолтных поставщиков, если они назывались «Поставщик 1» / «Поставщик 2»
      for (const sup of suppliers) {
        if (sup.id === 'supplier_1' && (sup.name === 'Поставщик 1' || sup.name === 'Поставщик')) {
          sup.name = 'ПП';
          await this.saveSupplier(sup);
        } else if (sup.id === 'supplier_2' && (sup.name === 'Поставщик 2')) {
          sup.name = 'Бир и Ко';
          await this.saveSupplier(sup);
        }
      }
    }
  }

  // ==========================================
  // МЕТОДЫ РАБОТЫ С INDEXEDDB
  // ==========================================

  async getAllSuppliers() {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('suppliers', 'readonly');
      const store = tx.objectStore('suppliers');
      const req = store.getAll();

      req.onsuccess = () => {
        const list = req.result || [];
        list.sort((a, b) => (a.order || 0) - (b.order || 0));
        resolve(list);
      };
      req.onerror = () => reject(req.error);
    });
  }

  async getSupplier(id) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('suppliers', 'readonly');
      const store = tx.objectStore('suppliers');
      const req = store.get(id);

      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  }

  async saveSupplier(supplier) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('suppliers', 'readwrite');
      const store = tx.objectStore('suppliers');
      const req = store.put(supplier);

      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async deleteSupplier(id) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(['suppliers', 'snapshots'], 'readwrite');
      tx.objectStore('suppliers').delete(id);
      tx.objectStore('snapshots').delete(id);

      tx.oncomplete = () => resolve(true);
      tx.onerror = () => reject(tx.error);
    });
  }

  async getSnapshotRecord(supplierId) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('snapshots', 'readonly');
      const store = tx.objectStore('snapshots');
      const req = store.get(supplierId);

      req.onsuccess = () => resolve(req.result || { supplierId, current: [], previous: [], updatedAt: '' });
      req.onerror = () => reject(req.error);
    });
  }

  async saveSnapshotRecord(record) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('snapshots', 'readwrite');
      const store = tx.objectStore('snapshots');
      const req = store.put(record);

      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  /**
   * Локальное сохранение нового снимка остатков (сдвиг текущего в предыдущий)
   */
  async commitNewSnapshot(supplierId, newSnapshotItems, mapping = null) {
    const existing = await this.getSnapshotRecord(supplierId);
    const nowIso = new Date().toISOString();

    const updatedRecord = {
      supplierId: supplierId,
      current: newSnapshotItems,
      previous: (existing.current && existing.current.length > 0) ? existing.current : (existing.previous || []),
      updatedAt: nowIso
    };

    await this.saveSnapshotRecord(updatedRecord);

    // Обновляем метку времени и маппинг в профиле поставщика
    const supplier = await this.getSupplier(supplierId);
    if (supplier) {
      supplier.updatedAt = nowIso;
      if (mapping) supplier.mapping = mapping;
      await this.saveSupplier(supplier);
    }

    return updatedRecord;
  }

  /**
   * Локальный откат снимка остатков к previousSnapshot
   */
  async rollbackLocalSnapshot(supplierId) {
    const existing = await this.getSnapshotRecord(supplierId);
    if (!existing || !existing.previous || existing.previous.length === 0) {
      throw new Error('Предыдущий снимок отсутствует. Откат невозможен.');
    }

    const nowIso = new Date().toISOString();
    const updatedRecord = {
      supplierId: supplierId,
      current: existing.previous,
      previous: [], // очищаем после отката
      updatedAt: nowIso
    };

    await this.saveSnapshotRecord(updatedRecord);

    const supplier = await this.getSupplier(supplierId);
    if (supplier) {
      supplier.updatedAt = nowIso;
      await this.saveSupplier(supplier);
    }

    return updatedRecord;
  }

  // ==========================================
  // LOCALSTORAGE (НАСТРОЙКИ UI И ОБЛАКА)
  // ==========================================

  getCloudConfig() {
    const savedUrl = localStorage.getItem('price_gas_url');
    return {
      url: (savedUrl && savedUrl.trim()) ? savedUrl.trim() : this.DEFAULT_CLOUD_URL,
      token: localStorage.getItem('price_gas_token') || sessionStorage.getItem('price_gas_token') || ''
    };
  }

  setCloudConfig(url, token, remember = true) {
    const cleanUrl = (url || this.DEFAULT_CLOUD_URL).trim();
    const cleanToken = (token || '').trim();

    localStorage.setItem('price_gas_url', cleanUrl);

    if (remember) {
      if (cleanToken) {
        localStorage.setItem('price_gas_token', cleanToken);
      } else {
        localStorage.removeItem('price_gas_token');
      }
      localStorage.setItem('price_remember_pin', 'true');
    } else {
      localStorage.removeItem('price_gas_token');
      localStorage.setItem('price_remember_pin', 'false');
    }

    if (cleanToken) {
      sessionStorage.setItem('price_gas_token', cleanToken);
      sessionStorage.setItem('price_session_auth', 'true');
    } else {
      sessionStorage.removeItem('price_gas_token');
      sessionStorage.removeItem('price_session_auth');
    }

    this._evaluateCloudStatus();
  }

  hasAuthenticatedPin() {
    const token = localStorage.getItem('price_gas_token') || sessionStorage.getItem('price_gas_token');
    return Boolean(token && token.trim().length > 0);
  }

  clearCloudAuth() {
    localStorage.removeItem('price_gas_token');
    sessionStorage.removeItem('price_gas_token');
    sessionStorage.removeItem('price_session_auth');
    this._evaluateCloudStatus();
  }

  getThresholds() {
    try {
      const saved = localStorage.getItem('price_thresholds');
      if (saved) return JSON.parse(saved);
    } catch (e) {}
    return { growthPct: 50, growthQty: 10, lowQty: 5 };
  }

  setThresholds(thresholds) {
    localStorage.setItem('price_thresholds', JSON.stringify(thresholds));
  }

  getActiveSupplierId() {
    return localStorage.getItem('price_active_supplier') || 'supplier_1';
  }

  setActiveSupplierId(id) {
    localStorage.setItem('price_active_supplier', id);
  }

  getTheme() {
    return localStorage.getItem('price_theme') || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  }

  setTheme(theme) {
    localStorage.setItem('price_theme', theme);
  }

  // ==========================================
  // СИНХРОНИЗАЦИЯ С GOOGLE APPS SCRIPT
  // ==========================================

  onSyncStatusChange(callback) {
    this.syncListeners.push(callback);
    callback(this.syncStatus, this.lastSyncMessage);
  }

  _notifyStatus(status, message = '') {
    this.syncStatus = status;
    this.lastSyncMessage = message;
    for (const listener of this.syncListeners) {
      try {
        listener(status, message);
      } catch (e) {
        console.error(e);
      }
    }
  }

  _evaluateCloudStatus() {
    const { url, token } = this.getCloudConfig();
    if (!token) {
      this._notifyStatus('unconfigured', 'Требуется ввод PIN');
    } else if (!navigator.onLine) {
      this._notifyStatus('offline', 'Нет интернета');
    } else {
      this._notifyStatus('synced', 'Готово к синхронизации');
    }
  }

  /**
   * Тест подключения к Google Apps Script
   */
  async testCloudConnection(url, token) {
    if (!url || !token) {
      throw new Error('Укажите URL веб-приложения и токен (PIN)');
    }

    try {
      this._notifyStatus('syncing', 'Проверка соединения...');
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify({
          action: 'test_connection',
          token: token.trim()
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP ошибка: ${response.status} ${response.statusText}`);
      }

      const result = await response.json();
      if (!result.success) {
        throw new Error(result.error || 'Ошибка проверки авторизации');
      }

      this._notifyStatus('synced', 'Подключение успешно!');
      return result;
    } catch (err) {
      this._notifyStatus('error', err.message);
      throw err;
    }
  }

  /**
   * Фоновый запрос свежей базы из Google Apps Script (GET)
   */
  async syncFromCloud() {
    const { url, token } = this.getCloudConfig();
    if (!url || !token) {
      this._notifyStatus('unconfigured', 'Облако не настроено');
      return { hasUpdates: false, error: 'Облако не настроено' };
    }

    if (!navigator.onLine) {
      this._notifyStatus('offline', 'Офлайн режим');
      return { hasUpdates: false, offline: true };
    }

    try {
      this._notifyStatus('syncing', 'Синхронизация с облаком...');

      const reqUrl = `${url}?token=${encodeURIComponent(token.trim())}&_t=${Date.now()}`;
      const response = await fetch(reqUrl);

      if (!response.ok) {
        throw new Error(`Ошибка облака HTTP ${response.status}`);
      }

      const res = await response.json();
      if (!res.success) {
        throw new Error(res.error || 'Ошибка получения данных из облака');
      }

      let hasUpdates = false;
      const remoteSuppliers = res.suppliers || [];

      for (const remote of remoteSuppliers) {
        const localSup = await this.getSupplier(remote.id);
        const localSnap = await this.getSnapshotRecord(remote.id);

        const remoteTime = remote.updatedAt ? new Date(remote.updatedAt).getTime() : 0;
        const localTime = localSnap.updatedAt ? new Date(localSnap.updatedAt).getTime() : 0;

        // Если в облаке снимок новее, чем локально
        if (remoteTime > localTime || !localSup) {
          hasUpdates = true;

          // Сохраняем/обновляем поставщика
          await this.saveSupplier({
            id: remote.id,
            name: remote.name,
            order: localSup ? localSup.order : 99,
            mapping: remote.mapping || (localSup ? localSup.mapping : null),
            updatedAt: remote.updatedAt
          });

          // Сохраняем снимок
          await this.saveSnapshotRecord({
            supplierId: remote.id,
            current: remote.currentSnapshot || [],
            previous: remote.previousSnapshot || [],
            updatedAt: remote.updatedAt
          });
        }
      }

      this._notifyStatus('synced', 'Синхронизировано');
      return { hasUpdates, suppliers: remoteSuppliers };

    } catch (err) {
      console.warn('Sync failed:', err);
      this._notifyStatus('error', err.message);
      return { hasUpdates: false, error: err.message };
    }
  }

  /**
   * Отправка нового снимка в Google Apps Script (POST)
   */
  async syncSnapshotToCloud(supplierId, supplierName, mapping, snapshot) {
    const { url, token } = this.getCloudConfig();
    if (!url || !token) return null; // Офлайн режим

    try {
      this._notifyStatus('syncing', 'Отправка в Google Таблицу...');

      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8' // Предотвращает CORS preflight в браузере!
        },
        body: JSON.stringify({
          action: 'save_snapshot',
          token: token.trim(),
          supplierId,
          supplierName,
          mapping,
          snapshot
        })
      });

      const res = await response.json();
      if (!res.success) {
        throw new Error(res.error || 'Ошибка сохранения в Google Apps Script');
      }

      this._notifyStatus('synced', 'Сохранено в облаке');
      return res;

    } catch (err) {
      console.error('Save to cloud failed:', err);
      this._notifyStatus('error', 'Ошибка облака: ' + err.message);
      throw err;
    }
  }

  /**
   * Откат базы в Google Apps Script (POST)
   */
  async rollbackInCloud(supplierId) {
    const { url, token } = this.getCloudConfig();
    if (!url || !token) return null;

    try {
      this._notifyStatus('syncing', 'Откат в Google Таблице...');

      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify({
          action: 'rollback',
          token: token.trim(),
          supplierId
        })
      });

      const res = await response.json();
      if (!res.success) {
        throw new Error(res.error || 'Ошибка отката в Google Apps Script');
      }

      this._notifyStatus('synced', 'База откачена');
      return res;

    } catch (err) {
      this._notifyStatus('error', 'Ошибка отката в облаке: ' + err.message);
      throw err;
    }
  }

  /**
   * Управление поставщиком в облаке (создание, переименование, удаление)
   */
  async manageSupplierInCloud(subAction, supplierId, supplierName = '') {
    const { url, token } = this.getCloudConfig();
    if (!url || !token) return null;

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify({
          action: 'manage_supplier',
          subAction,
          token: token.trim(),
          supplierId,
          supplierName
        })
      });

      return await response.json();
    } catch (err) {
      console.warn('Cloud supplier management failed:', err);
      return null;
    }
  }
}

// Экспорт глобального синглтона
window.PriceStorage = new StorageManager();
