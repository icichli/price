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
    this.isSyncing = false;
    this.isMutating = false;
    this.lastSyncAt = 0;
    this.SYNC_COOLDOWN_MS = 10000;
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
        { id: 'supplier_1', name: 'ПП', order: 0, mapping: null, updatedAt: null },
        { id: 'supplier_2', name: 'Бир и Ко', order: 1, mapping: null, updatedAt: null }
      ];

      for (const sup of defaultSuppliers) {
        await this.saveSupplier(sup);
        await this.saveSnapshotRecord({
          supplierId: sup.id,
          current: [],
          previous: [],
          updatedAt: null
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

      req.onsuccess = () => {
        const res = req.result || { supplierId, current: [], previous: [], updatedAt: '' };
        res.current = this._decompactSnapshot(res.current || []);
        res.previous = this._decompactSnapshot(res.previous || []);
        resolve(res);
      };
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

    const sanitizedItems = this._sanitizeSnapshotItems(newSnapshotItems);

    const updatedRecord = {
      supplierId: supplierId,
      current: sanitizedItems,
      previous: (existing.current && existing.current.length > 0) ? existing.current : (existing.previous || []),
      updatedAt: nowIso,
      syncPending: true,
      clearedLocally: false
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
      updatedAt: nowIso,
      syncPending: true,
      clearedLocally: false
    };

    await this.saveSnapshotRecord(updatedRecord);

    const supplier = await this.getSupplier(supplierId);
    if (supplier) {
      supplier.updatedAt = nowIso;
      await this.saveSupplier(supplier);
    }

    return updatedRecord;
  }

  /**
   * Локальная очистка снимка остатков поставщика (сброс текущих и предыдущих остатков)
   */
  async clearLocalSnapshot(supplierId) {
    const nowIso = new Date().toISOString();
    const emptyRecord = {
      supplierId: supplierId,
      current: [],
      previous: [],
      updatedAt: nowIso,
      syncPending: true,
      clearedLocally: true
    };

    await this.saveSnapshotRecord(emptyRecord);

    const supplier = await this.getSupplier(supplierId);
    if (supplier) {
      supplier.updatedAt = nowIso;
      if (supplier.mapping) {
        delete supplier.mapping.lastFileName;
        delete supplier.mapping.lastCommitTime;
      }
      await this.saveSupplier(supplier);
    }

    return emptyRecord;
  }

  /**
   * Фиксация серверного timestamp и снятие флага syncPending после успешной записи в Google Sheets
   */
  async markSnapshotSynced(supplierId, serverUpdatedAt) {
    const existing = await this.getSnapshotRecord(supplierId);
    const syncedAt = serverUpdatedAt || existing.updatedAt || new Date().toISOString();

    await this.saveSnapshotRecord({
      supplierId: supplierId,
      current: existing.current || [],
      previous: existing.previous || [],
      updatedAt: syncedAt,
      syncPending: false,
      clearedLocally: false
    });

    const supplier = await this.getSupplier(supplierId);
    if (supplier) {
      supplier.updatedAt = syncedAt;
      await this.saveSupplier(supplier);
    }

    return syncedAt;
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
   * Безопасный парсинг временной метки (возвращает 0 при пустом или некорректном значении)
   */
  _parseTimestampMs(val) {
    if (!val) return 0;
    const ms = new Date(val).getTime();
    return isNaN(ms) ? 0 : ms;
  }

  /**
   * Фоновый запрос свежей базы из Google Apps Script (GET)
   * @param {boolean} force - принудительное обновление локальной базы данными из облака
   */
  async syncFromCloud(force = false) {
    const { url, token } = this.getCloudConfig();
    if (!url || !token) {
      this._notifyStatus('unconfigured', 'Облако не настроено');
      return { hasUpdates: false, error: 'Облако не настроено' };
    }

    if (!navigator.onLine) {
      this._notifyStatus('offline', 'Офлайн режим');
      return { hasUpdates: false, offline: true };
    }

    // Защита от параллельных вызовов синхронизации (гонок)
    if (this.isSyncing) {
      return { hasUpdates: false, skipped: true };
    }

    this.isSyncing = true;
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
      const remoteIds = new Set();

      for (const remote of remoteSuppliers) {
        remoteIds.add(remote.id);

        const localSup = await this.getSupplier(remote.id);
        const localSnap = await this.getSnapshotRecord(remote.id);

        const currentDecompacted = this._decompactSnapshot(remote.currentSnapshot || []);
        const previousDecompacted = this._decompactSnapshot(remote.previousSnapshot || []);

        const remoteTime = this._parseTimestampMs(remote.updatedAt);
        const localTime = this._parseTimestampMs(localSnap && localSnap.updatedAt);
        const localHasData = Boolean(localSnap && localSnap.current && localSnap.current.length > 0);
        const remoteHasData = currentDecompacted.length > 0;
        const localPending = Boolean(localSnap && localSnap.syncPending);
        const localCleared = Boolean(localSnap && localSnap.clearedLocally);
        const remoteCorrupted = Boolean(remote.corruptedCurrent);

        // 0. Защита от повреждения JSON в облаке: не затираем валидную локальную базу пустотой!
        if (remoteCorrupted) {
          console.warn(`Cloud snapshot for ${remote.id} appears corrupted. Keeping local copy.`);
          if (!force && localHasData) {
            try {
              await this.syncSnapshotToCloud(
                remote.id,
                localSup ? localSup.name : remote.name,
                localSup ? localSup.mapping : remote.mapping,
                localSnap.current,
                { previousSnapshot: localSnap.previous || [], syncMode: 'sync' }
              );
            } catch (healErr) {
              console.warn('Auto-heal corrupted cloud snapshot failed:', healErr);
            }
          }
          continue;
        }

        // 1. Явная синхронизация несохраненных локальных изменений в облако (только при syncPending
        // или если в облаке вообще никогда не было данных). Никогда не пушим вслепую из-за разницы часов!
        if (!force && localPending) {
          if (localCleared && localTime >= remoteTime) {
            try {
              await this.clearSnapshotInCloud(remote.id, localSup ? localSup.name : remote.name);
            } catch (clearErr) {
              console.warn('Auto-push local clear to cloud failed:', clearErr);
            }
            continue;
          }
          if (localHasData && (localTime >= remoteTime || !remoteHasData)) {
            try {
              await this.syncSnapshotToCloud(
                remote.id,
                localSup ? localSup.name : remote.name,
                localSup ? localSup.mapping : remote.mapping,
                localSnap.current,
                { previousSnapshot: localSnap.previous || [], syncMode: 'sync' }
              );
            } catch (pushErr) {
              console.warn('Auto-push pending local snapshot to cloud failed:', pushErr);
            }
            continue;
          }
        }

        // Если в облаке пустая/новая строка без метки времени, а локально есть база — аккуратно загружаем её с сохранением previous
        if (!force && localHasData && !remoteHasData && remoteTime === 0) {
          try {
            await this.syncSnapshotToCloud(
              remote.id,
              localSup ? localSup.name : remote.name,
              localSup ? localSup.mapping : remote.mapping,
              localSnap.current,
              { previousSnapshot: localSnap.previous || [], syncMode: 'sync' }
            );
          } catch (pushErr) {
            console.warn('Initial push local snapshot to cloud failed:', pushErr);
          }
          continue;
        }

        // 2. Обновляем локальную базу, если:
        // - Принудительная синхронизация (force: true)
        // - Поставщик еще не создан локально
        // - В облаке снимок строго новее локального (remoteTime > localTime)
        // - Локальная база девственно пустая (localTime === 0 && !localCleared), а в облаке есть данные
        const shouldUpdate = force ||
          !localSup ||
          (remoteTime > localTime) ||
          (!localHasData && !localCleared && localTime === 0 && remoteHasData);

        if (shouldUpdate) {
          hasUpdates = true;

          // Сохраняем/обновляем поставщика
          await this.saveSupplier({
            id: remote.id,
            name: remote.name,
            order: localSup ? localSup.order : 99,
            mapping: remote.mapping || (localSup ? localSup.mapping : null),
            updatedAt: remote.updatedAt
          });

          await this.saveSnapshotRecord({
            supplierId: remote.id,
            current: currentDecompacted,
            previous: remote.corruptedPrevious ? (localSnap.previous || []) : previousDecompacted,
            updatedAt: remote.updatedAt,
            syncPending: false,
            clearedLocally: false
          });
        }
      }

      // 3. Проверяем локальных поставщиков, которых еще нет в Google Таблице
      if (!force) {
        const localSuppliers = await this.getAllSuppliers();
        for (const localSup of localSuppliers) {
          if (remoteIds.has(localSup.id)) continue;
          const localSnap = await this.getSnapshotRecord(localSup.id);
          const localHasData = Boolean(localSnap && localSnap.current && localSnap.current.length > 0);
          if (localHasData && localSnap.syncPending) {
            try {
              await this.syncSnapshotToCloud(
                localSup.id,
                localSup.name,
                localSup.mapping,
                localSnap.current,
                { previousSnapshot: localSnap.previous || [], syncMode: 'sync' }
              );
            } catch (pushErr) {
              console.warn('Auto-push local-only supplier failed:', pushErr);
            }
          }
        }
      }

      this.lastSyncAt = Date.now();
      this._notifyStatus('synced', 'Синхронизировано');
      return { hasUpdates, suppliers: remoteSuppliers };

    } catch (err) {
      console.warn('Sync failed:', err);
      this._notifyStatus('error', err.message);
      return { hasUpdates: false, error: err.message };
    } finally {
      this.isSyncing = false;
    }
  }

  /**
   * Отправка нового снимка в Google Apps Script (POST)
   * @param {object} options - { previousSnapshot?: Array, syncMode?: 'commit' | 'sync' }
   */
  async syncSnapshotToCloud(supplierId, supplierName, mapping, snapshot, options = {}) {
    const { url, token } = this.getCloudConfig();
    if (!url || !token) return null; // Офлайн режим

    try {
      this._notifyStatus('syncing', 'Отправка в Google Таблицу...');

      // Компактная сериализация v2 со словарем категорий/брендов
      const sanitized = this._sanitizeSnapshotItems(snapshot);
      const compactSnapshot = this._compactSnapshot(sanitized);

      const payload = {
        action: 'save_snapshot',
        token: token.trim(),
        supplierId,
        supplierName,
        mapping,
        snapshot: compactSnapshot,
        syncMode: options.syncMode === 'sync' ? 'sync' : 'commit'
      };

      if (options.previousSnapshot !== undefined) {
        const sanitizedPrev = this._sanitizeSnapshotItems(options.previousSnapshot);
        payload.previousSnapshot = this._compactSnapshot(sanitizedPrev);
      }

      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8' // Предотвращает CORS preflight в браузере!
        },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        throw new Error(`HTTP ошибка облака: ${response.status}`);
      }

      const res = await response.json();
      if (!res.success) {
        throw new Error(res.error || 'Ошибка сохранения в Google Apps Script');
      }

      // Синхронизируем локальную метку времени с серверной, чтобы избежать ложных авто-пушей
      if (res.updatedAt) {
        await this.markSnapshotSynced(supplierId, res.updatedAt);
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
   * Принудительная отправка локальной базы поставщика в Google Sheets (с сохранением previous)
   */
  async pushLocalSnapshotToCloud(supplierId) {
    const { url, token } = this.getCloudConfig();
    if (!url || !token) throw new Error('Облако не настроено');

    const supplier = await this.getSupplier(supplierId);
    const snap = await this.getSnapshotRecord(supplierId);

    const supName = supplier ? supplier.name : 'Поставщик';
    const mapping = supplier ? supplier.mapping : null;
    const items = (snap && snap.current) ? snap.current : [];
    const prevItems = (snap && snap.previous) ? snap.previous : [];

    return await this.syncSnapshotToCloud(supplierId, supName, mapping, items, {
      previousSnapshot: prevItems,
      syncMode: 'sync'
    });
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

      if (!response.ok) {
        throw new Error(`HTTP ошибка облака: ${response.status}`);
      }

      const res = await response.json();
      if (!res.success) {
        throw new Error(res.error || 'Ошибка отката в Google Apps Script');
      }

      if (res.updatedAt) {
        await this.markSnapshotSynced(supplierId, res.updatedAt);
      }

      this._notifyStatus('synced', 'База откачена');
      return res;

    } catch (err) {
      this._notifyStatus('error', 'Ошибка отката в облаке: ' + err.message);
      throw err;
    }
  }

  /**
   * Очистка базы остатков поставщика в Google Apps Script (POST)
   */
  async clearSnapshotInCloud(supplierId, supplierName = '') {
    const { url, token } = this.getCloudConfig();
    if (!url || !token) return null;

    try {
      this._notifyStatus('syncing', 'Очистка базы в Google Таблице...');

      // 1. Попытка вызвать специализированное действие clear_snapshot
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify({
          action: 'clear_snapshot',
          token: token.trim(),
          supplierId
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP ошибка облака: ${response.status}`);
      }

      const res = await response.json();
      if (res.success) {
        if (res.updatedAt) {
          await this.markSnapshotSynced(supplierId, res.updatedAt);
        }
        this._notifyStatus('synced', 'База в облаке очищена');
        return res;
      }

      // 2. Fallback для обратной совместимости со старыми развертываниями Code.gs
      if (res.error && res.error.includes('Неизвестное действие')) {
        await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({
            action: 'save_snapshot',
            token: token.trim(),
            supplierId,
            supplierName: supplierName || 'Поставщик',
            mapping: null,
            snapshot: []
          })
        });
        const fbRes = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({
            action: 'save_snapshot',
            token: token.trim(),
            supplierId,
            supplierName: supplierName || 'Поставщик',
            mapping: null,
            snapshot: []
          })
        });
        const fbJson = await fbRes.json();
        if (fbJson && fbJson.updatedAt) {
          await this.markSnapshotSynced(supplierId, fbJson.updatedAt);
        }
        this._notifyStatus('synced', 'База в облаке очищена');
        return fbJson;
      }

      throw new Error(res.error || 'Ошибка очистки базы в Google Apps Script');

    } catch (err) {
      console.error('Clear cloud snapshot failed:', err);
      this._notifyStatus('error', 'Ошибка очистки в облаке: ' + err.message);
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

  // ==========================================
  // СЖАТИЕ И СЕРИАЛИЗАЦИЯ СНИМКОВ ОСТАТКОВ
  // ==========================================

  _sanitizeSnapshotItems(items) {
    if (!Array.isArray(items)) return [];
    return items.map(item => {
      const clean = {
        name: String(item.name || '').trim(),
        stock: typeof item.stock === 'number' ? item.stock : (parseFloat(item.stock) || 0),
        code: item.code ? String(item.code).trim() : '',
        isTextStock: Boolean(item.isTextStock),
        textStockNote: item.textStockNote ? String(item.textStockNote).trim() : '',
        packageType: item.packageType || '',
        packageSubType: item.packageSubType || '',
        packageLabel: item.packageLabel || '',
        packageIcon: item.packageIcon || '',
        category: item.category || '',
        subCategory: item.subCategory || ''
      };
      return clean;
    });
  }

  _compactSnapshot(items) {
    if (!Array.isArray(items) || items.length === 0) return [];

    // Формат v2 со словарем категорий и подкатегорий:
    // Сокращает объем JSON для крупных поставщиков (например, ПП) на ~25-30%
    // и сохраняет совместимость с проверками Array.isArray() && length > 0.
    const cats = [];
    const catIndexMap = new Map();
    const subs = [];
    const subIndexMap = new Map();

    const pkgToCode = { draft: 1, bottle: 2, other: 3 };

    const rows = items.map(i => {
      let ci = -1;
      if (i.category) {
        if (!catIndexMap.has(i.category)) {
          catIndexMap.set(i.category, cats.length);
          cats.push(i.category);
        }
        ci = catIndexMap.get(i.category);
      }

      let si = -1;
      if (i.subCategory) {
        if (!subIndexMap.has(i.subCategory)) {
          subIndexMap.set(i.subCategory, subs.length);
          subs.push(i.subCategory);
        }
        si = subIndexMap.get(i.subCategory);
      }

      const pkgCode = pkgToCode[i.packageType] || 0;
      const code = i.code || '';
      const isText = i.isTextStock ? 1 : 0;
      const note = i.textStockNote || '';

      // Кортеж v2: [name, stock, catIdx?, subIdx?, pkgCode?, code?, isText?, note?]
      const arr = [i.name, i.stock];
      if (ci >= 0 || si >= 0 || pkgCode || code || isText || note) {
        arr.push(ci, si, pkgCode);
        if (code || isText || note) {
          arr.push(code);
          if (isText || note) {
            arr.push(isText);
            if (note) {
              arr.push(note);
            }
          }
        }
      }
      return arr;
    });

    rows.unshift({ _v: 2, c: cats, s: subs });
    return rows;
  }

  _decompactSnapshot(items) {
    if (!Array.isArray(items) || items.length === 0) return [];

    // Формат v2 со словарем категорий в 0-м элементе
    if (items[0] && !Array.isArray(items[0]) && items[0]._v === 2) {
      const header = items[0];
      const cats = Array.isArray(header.c) ? header.c : [];
      const subs = Array.isArray(header.s) ? header.s : [];
      const codeToPkg = { 1: 'draft', 2: 'bottle', 3: 'other' };

      const result = [];
      for (let idx = 1; idx < items.length; idx++) {
        const i = items[idx];
        if (!Array.isArray(i)) continue;
        const ci = typeof i[2] === 'number' ? i[2] : -1;
        const si = typeof i[3] === 'number' ? i[3] : -1;
        result.push({
          name: i[0] || '',
          stock: typeof i[1] === 'number' ? i[1] : (parseFloat(i[1]) || 0),
          category: (ci >= 0 && ci < cats.length) ? cats[ci] : '',
          subCategory: (si >= 0 && si < subs.length) ? subs[si] : '',
          packageType: codeToPkg[i[4]] || '',
          code: typeof i[5] === 'string' ? i[5] : '',
          isTextStock: i[6] === 1,
          textStockNote: typeof i[7] === 'string' ? i[7] : ''
        });
      }
      return result;
    }

    // Формат v1 (кортежи) и v0 (объекты)
    return items.map(i => {
      if (Array.isArray(i)) {
        const isV1WithPkg = typeof i[5] === 'string' && ['draft', 'bottle', 'other', ''].includes(i[5]);
        return {
          name: i[0] || '',
          stock: typeof i[1] === 'number' ? i[1] : (parseFloat(i[1]) || 0),
          code: i[2] || '',
          isTextStock: isV1WithPkg ? (i[3] === 1) : (i[4] === 1),
          textStockNote: isV1WithPkg
            ? (typeof i[4] === 'string' ? i[4] : '')
            : (typeof i[5] === 'string' ? i[5] : ''),
          packageType: (i[5] && typeof i[5] === 'string' && ['draft', 'bottle', 'other'].includes(i[5])) ? i[5] : '',
          category: typeof i[6] === 'string' ? i[6] : '',
          subCategory: typeof i[7] === 'string' ? i[7] : ''
        };
      }
      return {
        name: i.name || '',
        stock: typeof i.stock === 'number' ? i.stock : (parseFloat(i.stock) || 0),
        code: i.code || '',
        isTextStock: Boolean(i.isTextStock),
        textStockNote: i.textStockNote || '',
        packageType: i.packageType || '',
        packageSubType: i.packageSubType || '',
        packageLabel: i.packageLabel || '',
        packageIcon: i.packageIcon || '',
        category: i.category || '',
        subCategory: i.subCategory || ''
      };
    });
  }
}

// Экспорт глобального синглтона
window.PriceStorage = new StorageManager();
