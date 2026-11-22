/**
 * app.js - Основной модуль приложения: Diff Engine, управление поставщиками,
 * фильтрация, пороги, экспорт и обработка событий UI.
 */

(function () {
  'use strict';

  // ==========================================================================
  // СОСТОЯНИЕ ПРИЛОЖЕНИЯ (APP STATE)
  // ==========================================================================
  const state = {
    suppliers: [],
    activeSupplierId: null,
    
    // Текущий загруженный файл и распарсенные данные
    currentWorkbook: null,
    currentFileName: '',
    currentSheetName: '',
    currentSheetData: [],
    currentHeaderRowIndex: 0,
    currentMapping: null,
    parsedItems: [],

    // Эталонная база остатков для активного поставщика
    baselineSnapshot: {
      current: [],
      previous: [],
      updatedAt: ''
    },

    // Результаты сравнения (Diff)
    diffItems: [],
    diffStats: {
      all: 0,
      new: 0,
      growth: 0,
      drop: 0,
      out: 0
    },

    // Фильтрация, сортировка и поиск
    activeFilter: 'all', // 'all' | 'new' | 'growth' | 'drop' | 'out'
    searchQuery: '',
    sortField: 'stockDelta',
    sortDir: 'desc',

    // Пороговые значения
    thresholds: {
      growthPct: 50,
      growthQty: 10,
      lowQty: 5
    },

    // Текущий режим редактирования поставщика
    supplierModalMode: 'add', // 'add' | 'edit'
    supplierModalTargetId: null
  };

  // ==========================================================================
  // DOM ЭЛЕМЕНТЫ
  // ==========================================================================
  const dom = {
    // Шапка
    cloudStatusBadge: document.getElementById('cloudStatusBadge'),
    cloudStatusText: document.getElementById('cloudStatusText'),
    btnOpenMapping: document.getElementById('btnOpenMapping'),
    btnOpenCloudSettings: document.getElementById('btnOpenCloudSettings'),
    btnThemeToggle: document.getElementById('btnThemeToggle'),
    themeIconSun: document.getElementById('themeIconSun'),
    themeIconMoon: document.getElementById('themeIconMoon'),

    // Поставщики
    supplierTabsContainer: document.getElementById('supplierTabsContainer'),
    btnAddSupplier: document.getElementById('btnAddSupplier'),

    // Загрузка файла
    dropzone: document.getElementById('dropzone'),
    fileInput: document.getElementById('fileInput'),
    fileCard: document.getElementById('fileCard'),
    loadedFileName: document.getElementById('loadedFileName'),
    loadedFileStats: document.getElementById('loadedFileStats'),
    loadedFileBaseline: document.getElementById('loadedFileBaseline'),
    btnTuneColumns: document.getElementById('btnTuneColumns'),
    btnChangeFile: document.getElementById('btnChangeFile'),

    // Пороги
    threshGrowthPct: document.getElementById('threshGrowthPct'),
    threshGrowthQty: document.getElementById('threshGrowthQty'),
    threshLowQty: document.getElementById('threshLowQty'),

    // Фильтры и действия
    filterTabs: document.querySelectorAll('.filter-tab'),
    badgeAll: document.getElementById('badgeAll'),
    badgeNew: document.getElementById('badgeNew'),
    badgeGrowth: document.getElementById('badgeGrowth'),
    badgeDrop: document.getElementById('badgeDrop'),
    badgeOut: document.getElementById('badgeOut'),
    searchInput: document.getElementById('searchInput'),

    btnCommitSnapshot: document.getElementById('btnCommitSnapshot'),
    btnRollbackSnapshot: document.getElementById('btnRollbackSnapshot'),
    btnExportMenuToggle: document.getElementById('btnExportMenuToggle'),
    exportDropdownMenu: document.getElementById('exportDropdownMenu'),
    btnExportExcel: document.getElementById('btnExportExcel'),
    btnExportClipboard: document.getElementById('btnExportClipboard'),

    // Таблица
    tableBody: document.getElementById('tableBody'),
    tableEmptyState: document.getElementById('tableEmptyState'),
    tableHeaders: document.querySelectorAll('.data-table th.sortable'),
    footerStats: document.getElementById('footerStats'),
    footerBaselineInfo: document.getElementById('footerBaselineInfo'),

    // Модалка облака
    modalCloudSettings: document.getElementById('modalCloudSettings'),
    cloudUrlInput: document.getElementById('cloudUrlInput'),
    cloudPinInput: document.getElementById('cloudPinInput'),
    cloudTestResult: document.getElementById('cloudTestResult'),
    btnTestCloudConnection: document.getElementById('btnTestCloudConnection'),
    btnSyncCloudNow: document.getElementById('btnSyncCloudNow'),
    btnSaveCloudSettings: document.getElementById('btnSaveCloudSettings'),

    // Модалка колонок
    modalColumnMapping: document.getElementById('modalColumnMapping'),
    sheetSelector: document.getElementById('sheetSelector'),
    headerRowSelector: document.getElementById('headerRowSelector'),
    previewTableHead: document.getElementById('previewTableHead'),
    previewTableBody: document.getElementById('previewTableBody'),
    btnApplyMapping: document.getElementById('btnApplyMapping'),

    // Модалка поставщика
    modalSupplier: document.getElementById('modalSupplier'),
    modalSupplierTitle: document.getElementById('modalSupplierTitle'),
    supplierNameInput: document.getElementById('supplierNameInput'),
    btnSaveSupplier: document.getElementById('btnSaveSupplier'),

    // Тосты
    toastContainer: document.getElementById('toastContainer')
  };

  // ==========================================================================
  // ИНИЦИАЛИЗАЦИЯ ПРИЛОЖЕНИЯ
  // ==========================================================================
  async function initApp() {
    initTheme();
    initThresholds();
    setupEventListeners();
    setupCloudStatusListener();

    try {
      // 1. Инициализируем локальную БД IndexedDB
      await window.PriceStorage.init();
      
      // 2. Загружаем поставщиков
      state.suppliers = await window.PriceStorage.getAllSuppliers();
      const savedActiveId = window.PriceStorage.getActiveSupplierId();
      
      if (state.suppliers.some(s => s.id === savedActiveId)) {
        state.activeSupplierId = savedActiveId;
      } else if (state.suppliers.length > 0) {
        state.activeSupplierId = state.suppliers[0].id;
      }

      renderSupplierTabs();
      await loadActiveSupplierSnapshot();

      // 3. Фоновая проверка и синхронизация с облаком (если настроено)
      backgroundCloudSync();

    } catch (err) {
      console.error('Init error:', err);
      showToast('Ошибка инициализации хранилища: ' + err.message, 'error');
    }
  }

  // ==========================================================================
  // ТЕМА ОФОРМЛЕНИЯ (DARK / LIGHT)
  // ==========================================================================
  function initTheme() {
    const theme = window.PriceStorage.getTheme();
    applyTheme(theme);
  }

  function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    window.PriceStorage.setTheme(theme);
    if (theme === 'dark') {
      dom.themeIconSun.style.display = 'block';
      dom.themeIconMoon.style.display = 'none';
    } else {
      dom.themeIconSun.style.display = 'none';
      dom.themeIconMoon.style.display = 'block';
    }
  }

  function toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme') || 'light';
    const next = current === 'dark' ? 'light' : 'dark';
    applyTheme(next);
  }

  // ==========================================================================
  // ПОРОГОВЫЕ ЗНАЧЕНИЯ
  // ==========================================================================
  function initThresholds() {
    state.thresholds = window.PriceStorage.getThresholds();
    dom.threshGrowthPct.value = state.thresholds.growthPct;
    dom.threshGrowthQty.value = state.thresholds.growthQty;
    dom.threshLowQty.value = state.thresholds.lowQty;
  }

  function onThresholdChange() {
    state.thresholds = {
      growthPct: Math.max(1, parseFloat(dom.threshGrowthPct.value) || 50),
      growthQty: Math.max(1, parseFloat(dom.threshGrowthQty.value) || 10),
      lowQty: Math.max(0, parseFloat(dom.threshLowQty.value) || 5)
    };
    window.PriceStorage.setThresholds(state.thresholds);

    // Мгновенный пересчет без перезагрузки файла!
    if (state.parsedItems.length > 0 || state.baselineSnapshot.current.length > 0) {
      recomputeDiff();
      renderTable();
    }
  }

  // ==========================================================================
  // ПОСТАВЩИКИ (TABS MANAGEMENT)
  // ==========================================================================
  function renderSupplierTabs() {
    dom.supplierTabsContainer.innerHTML = '';

    state.suppliers.forEach(sup => {
      const tab = document.createElement('div');
      tab.className = `supplier-tab ${sup.id === state.activeSupplierId ? 'active' : ''}`;
      tab.dataset.id = sup.id;

      const nameSpan = document.createElement('span');
      nameSpan.textContent = sup.name;
      nameSpan.title = 'Нажмите для переключения';
      tab.appendChild(nameSpan);

      // Иконка переименования
      const editIcon = document.createElement('span');
      editIcon.className = 'tab-action-icon';
      editIcon.title = 'Переименовать поставщика';
      editIcon.innerHTML = `
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M12 20h9"></path>
          <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
        </svg>
      `;
      editIcon.addEventListener('click', (e) => {
        e.stopPropagation();
        openEditSupplierModal(sup);
      });
      tab.appendChild(editIcon);

      // Иконка удаления (если поставщиков больше 1)
      if (state.suppliers.length > 1) {
        const delIcon = document.createElement('span');
        delIcon.className = 'tab-action-icon';
        delIcon.title = 'Удалить поставщика';
        delIcon.innerHTML = `
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="18" y1="6" x2="6" y2="18"></line>
            <line x1="6" y1="6" x2="18" y2="18"></line>
          </svg>
        `;
        delIcon.addEventListener('click', (e) => {
          e.stopPropagation();
          deleteSupplier(sup);
        });
        tab.appendChild(delIcon);
      }

      tab.addEventListener('click', () => switchSupplier(sup.id));
      dom.supplierTabsContainer.appendChild(tab);
    });
  }

  async function switchSupplier(supplierId) {
    if (state.activeSupplierId === supplierId) return;

    state.activeSupplierId = supplierId;
    window.PriceStorage.setActiveSupplierId(supplierId);
    renderSupplierTabs();

    // Сброс загруженного файла для нового поставщика
    resetCurrentFile();
    await loadActiveSupplierSnapshot();
  }

  async function loadActiveSupplierSnapshot() {
    if (!state.activeSupplierId) return;

    const snap = await window.PriceStorage.getSnapshotRecord(state.activeSupplierId);
    state.baselineSnapshot = snap;

    updateBaselineInfo();
    updateRollbackButtonState();

    if (state.parsedItems.length > 0) {
      recomputeDiff();
      renderTable();
    } else if (snap.current && snap.current.length > 0) {
      // Если файл еще не загружен, но есть сохраненная база — отобразим ее
      displayBaselineAsCurrent();
    } else {
      state.diffItems = [];
      updateBadges({ all: 0, new: 0, growth: 0, drop: 0, out: 0 });
      renderTable();
    }
  }

  function displayBaselineAsCurrent() {
    state.diffItems = state.baselineSnapshot.current.map(item => ({
      status: 'normal',
      statusLabel: 'В базе',
      code: item.code || '',
      name: item.name,
      prevStock: item.stock,
      currStock: item.stock,
      stockDelta: 0,
      stockDeltaPct: 0,
      prevPrice: item.price,
      currPrice: item.price,
      priceDelta: 0,
      isTextStock: item.isTextStock,
      textStockNote: item.textStockNote
    }));

    const stats = {
      all: state.diffItems.length,
      new: 0,
      growth: 0,
      drop: 0,
      out: 0
    };
    updateBadges(stats);
    renderTable();
  }

  function updateBaselineInfo() {
    const sup = state.suppliers.find(s => s.id === state.activeSupplierId);
    const dateStr = state.baselineSnapshot.updatedAt
      ? new Date(state.baselineSnapshot.updatedAt).toLocaleString('ru-RU')
      : 'нет данных';

    const count = (state.baselineSnapshot.current || []).length;
    const infoText = `База: ${count} поз. (от ${dateStr})`;
    
    dom.loadedFileBaseline.textContent = infoText;
    dom.footerBaselineInfo.textContent = `Эталон: ${sup ? sup.name : ''} | ${infoText}`;
  }

  function updateRollbackButtonState() {
    const hasPrevious = state.baselineSnapshot.previous && state.baselineSnapshot.previous.length > 0;
    dom.btnRollbackSnapshot.disabled = !hasPrevious;
  }

  function openAddSupplierModal() {
    state.supplierModalMode = 'add';
    state.supplierModalTargetId = null;
    dom.modalSupplierTitle.textContent = 'Новый поставщик';
    dom.supplierNameInput.value = '';
    openModal('modalSupplier');
  }

  function openEditSupplierModal(sup) {
    state.supplierModalMode = 'edit';
    state.supplierModalTargetId = sup.id;
    dom.modalSupplierTitle.textContent = 'Переименовать поставщика';
    dom.supplierNameInput.value = sup.name;
    openModal('modalSupplier');
  }

  async function saveSupplierFromModal() {
    const name = dom.supplierNameInput.value.trim();
    if (!name) {
      showToast('Введите название поставщика', 'error');
      return;
    }

    if (state.supplierModalMode === 'add') {
      const newId = 'supplier_' + Date.now();
      const newSup = {
        id: newId,
        name: name,
        order: state.suppliers.length,
        mapping: null,
        updatedAt: new Date().toISOString()
      };

      await window.PriceStorage.saveSupplier(newSup);
      await window.PriceStorage.saveSnapshotRecord({
        supplierId: newId,
        current: [],
        previous: [],
        updatedAt: newSup.updatedAt
      });

      // Фоновое создание в Google Таблице
      window.PriceStorage.manageSupplierInCloud('create', newId, name);

      state.suppliers.push(newSup);
      state.activeSupplierId = newId;
      window.PriceStorage.setActiveSupplierId(newId);
      showToast(`Поставщик «${name}» добавлен`, 'success');

    } else {
      const sup = state.suppliers.find(s => s.id === state.supplierModalTargetId);
      if (sup) {
        sup.name = name;
        await window.PriceStorage.saveSupplier(sup);
        window.PriceStorage.manageSupplierInCloud('rename', sup.id, name);
        showToast('Название обновлено', 'success');
      }
    }

    closeModal('modalSupplier');
    renderSupplierTabs();
    await loadActiveSupplierSnapshot();
  }

  async function deleteSupplier(sup) {
    if (state.suppliers.length <= 1) {
      showToast('Нельзя удалить единственного поставщика', 'error');
      return;
    }

    if (!confirm(`Удалить поставщика «${sup.name}» и все связанные остатки?`)) {
      return;
    }

    await window.PriceStorage.deleteSupplier(sup.id);
    window.PriceStorage.manageSupplierInCloud('delete', sup.id);

    state.suppliers = state.suppliers.filter(s => s.id !== sup.id);
    if (state.activeSupplierId === sup.id) {
      state.activeSupplierId = state.suppliers[0].id;
      window.PriceStorage.setActiveSupplierId(state.activeSupplierId);
    }

    showToast(`Поставщик «${sup.name}» удален`, 'info');
    renderSupplierTabs();
    resetCurrentFile();
    await loadActiveSupplierSnapshot();
  }

  // ==========================================
  // ОБРАБОТКА И ПАРСИНГ EXCEL ФАЙЛА
  // ==========================================
  async function handleFileSelect(file) {
    if (!file) return;

    const validExts = ['.xlsx', '.xls'];
    const lowerName = file.name.toLowerCase();
    if (!validExts.some(ext => lowerName.endsWith(ext))) {
      showToast('Пожалуйста, выберите файл в формате .xlsx или .xls', 'error');
      return;
    }

    try {
      showToast('Чтение Excel файла...', 'info');

      const result = await window.ExcelParser.readWorkbook(file);
      state.currentWorkbook = result.workbook;
      state.currentFileName = file.name;
      state.currentSheetName = result.defaultSheetName;

      // Читаем сырые данные листа
      state.currentSheetData = window.ExcelParser.getSheetData(result.workbook, state.currentSheetName);

      if (!state.currentSheetData || state.currentSheetData.length === 0) {
        throw new Error('Выбранный лист Excel пуст');
      }

      // Определяем строку заголовков и авто-маппинг
      await processSheetDataAndAutoMap();

      dom.fileCard.classList.add('visible');
      dom.loadedFileName.textContent = file.name;
      dom.btnOpenMapping.disabled = false;
      dom.btnCommitSnapshot.disabled = false;

    } catch (err) {
      console.error('File parsing error:', err);
      showToast('Ошибка разбора файла: ' + err.message, 'error');
    }
  }

  async function processSheetDataAndAutoMap() {
    // 1. Ищем строку заголовка
    const headerDetection = window.ExcelParser.detectHeaderRow(state.currentSheetData);
    state.currentHeaderRowIndex = headerDetection.rowIndex;
    const headerRow = state.currentSheetData[state.currentHeaderRowIndex] || [];
    const previousRow = headerDetection.isMultiLineHeader && state.currentHeaderRowIndex > 0
      ? state.currentSheetData[state.currentHeaderRowIndex - 1]
      : null;

    // 2. Проверяем сохраненный маппинг поставщика
    const activeSup = state.suppliers.find(s => s.id === state.activeSupplierId);
    let mapping = activeSup ? activeSup.mapping : null;

    // Если нет сохраненного маппинга или он не валиден — выполняем автомаппинг
    if (!mapping || mapping.nameCol < 0 || mapping.stockCol < 0) {
      mapping = window.ExcelParser.detectColumns(headerRow, previousRow);
    }

    state.currentMapping = mapping;

    // Если автоопределение не нашло обязательные колонки — открываем окно ручной настройки
    if (mapping.nameCol < 0 || mapping.stockCol < 0) {
      showToast('Не удалось автоматически распознать все колонки. Настройте их вручную.', 'info');
      openMappingModal();
      return;
    }

    // Извлекаем позиции
    extractAndDiff();
  }

  function extractAndDiff() {
    try {
      const extracted = window.ExcelParser.extractItems(
        state.currentSheetData,
        state.currentHeaderRowIndex,
        state.currentMapping
      );

      state.parsedItems = extracted.items;

      dom.loadedFileStats.textContent = `Лист: ${state.currentSheetName} | ${extracted.items.length} поз.` +
        (extracted.skippedCategoriesCount > 0 ? ` (пропущено подзаголовков: ${extracted.skippedCategoriesCount})` : '');

      showToast(`Успешно обработано: ${extracted.items.length} позиций`, 'success');

      recomputeDiff();
      renderTable();

    } catch (err) {
      showToast(err.message, 'error');
      openMappingModal();
    }
  }

  function resetCurrentFile() {
    state.currentWorkbook = null;
    state.currentFileName = '';
    state.currentSheetName = '';
    state.currentSheetData = [];
    state.parsedItems = [];
    state.diffItems = [];

    dom.fileCard.classList.remove('visible');
    dom.fileInput.value = '';
    dom.btnOpenMapping.disabled = true;
    dom.btnCommitSnapshot.disabled = true;
  }

  // ==========================================================================
  // DIFF ENGINE (СРАВНЕНИЕ ОСТАТКОВ И КЛАССИФИКАЦИЯ)
  // ==========================================================================
  function recomputeDiff() {
    const currentList = state.parsedItems;
    const baselineList = state.baselineSnapshot.current || [];

    // 1. Проверяем, есть ли артикулы в обоих наборах данных
    const hasBaselineCode = baselineList.some(i => i.code && i.code.trim().length > 0);
    const hasCurrentCode = currentList.some(i => i.code && i.code.trim().length > 0);
    const useCodeAndName = hasBaselineCode && hasCurrentCode;

    // Вспомогательная функция формирования уникального ключа позиции
    const getItemKey = (item) => {
      const normName = window.ExcelParser.normalizeItemName(item.name);
      if (useCodeAndName && item.code && item.code.trim()) {
        const normCode = item.code.trim().toLowerCase();
        return `${normCode}:::${normName}`;
      }
      return normName;
    };

    // Строим словарь предыдущих позиций
    const prevMap = new Map();
    baselineList.forEach(item => {
      const key = getItemKey(item);
      prevMap.set(key, item);
    });

    const diffResults = [];
    const seenPrevKeys = new Set();

    let newCount = 0;
    let growthCount = 0;
    let dropCount = 0;
    let outCount = 0;

    // 2. Обрабатываем позиции из свежего файла
    for (const curr of currentList) {
      const key = getItemKey(curr);
      const prev = prevMap.get(key);
      seenPrevKeys.add(key);

      const prevStock = prev ? prev.stock : 0;
      const currStock = curr.stock;
      const stockDelta = currStock - prevStock;

      let stockDeltaPct = 0;
      if (prevStock > 0) {
        stockDeltaPct = Math.round(((currStock - prevStock) / prevStock) * 100);
      } else if (currStock > 0) {
        stockDeltaPct = 100;
      }

      const prevPrice = prev ? prev.price : null;
      const currPrice = curr.price;
      const priceDelta = (currPrice !== null && prevPrice !== null)
        ? Math.round((currPrice - prevPrice) * 100) / 100
        : null;

      // КЛАССИФИКАЦИЯ СТАТУСА
      let status = 'normal';
      let statusLabel = 'В наличии';

      // 1. Новинка: не было в прошлом снимке или был 0, а стал > 0
      if (!prev || (prevStock === 0 && currStock > 0)) {
        status = 'new';
        statusLabel = 'Новинка';
        newCount++;
      }
      // 2. Выбыло: был > 0, стал 0
      else if (currStock === 0 && prevStock > 0) {
        status = 'out';
        statusLabel = 'Выбыло (0)';
        outCount++;
      }
      // 3. Сильный рост / Поступление: вырос больше чем на порог % ИЛИ порог ед.
      else if (stockDelta > 0 && (stockDeltaPct >= state.thresholds.growthPct || stockDelta >= state.thresholds.growthQty)) {
        status = 'growth';
        statusLabel = 'Поступление';
        growthCount++;
      }
      // 4. Резкое падение / Мало: остаток упал или стал меньше критического порога
      else if (stockDelta < 0 && (currStock < state.thresholds.lowQty || stockDeltaPct <= -30)) {
        status = 'drop';
        statusLabel = 'Мало / Спад';
        dropCount++;
      }

      diffResults.push({
        status,
        statusLabel,
        code: curr.code || (prev ? prev.code : ''),
        name: curr.name,
        prevStock,
        currStock,
        stockDelta,
        stockDeltaPct,
        prevPrice,
        currPrice,
        priceDelta,
        isTextStock: curr.isTextStock,
        textStockNote: curr.textStockNote
      });
    }

    // 3. Обрабатываем позиции, которые БЫЛИ в базе с остатком > 0, но ИСЧЕЗЛИ из нового файла
    for (const [key, prev] of prevMap.entries()) {
      if (!seenPrevKeys.has(key) && prev.stock > 0) {
        diffResults.push({
          status: 'out',
          statusLabel: 'Выбыло (нет в прайсе)',
          code: prev.code || '',
          name: prev.name,
          prevStock: prev.stock,
          currStock: 0,
          stockDelta: -prev.stock,
          stockDeltaPct: -100,
          prevPrice: prev.price,
          currPrice: null,
          priceDelta: null,
          isTextStock: false,
          textStockNote: ''
        });
        outCount++;
      }
    }

    state.diffItems = diffResults;
    state.diffStats = {
      all: diffResults.length,
      new: newCount,
      growth: growthCount,
      drop: dropCount,
      out: outCount
    };

    updateBadges(state.diffStats);
  }

  function updateBadges(stats) {
    dom.badgeAll.textContent = stats.all;
    dom.badgeNew.textContent = stats.new;
    dom.badgeGrowth.textContent = stats.growth;
    dom.badgeDrop.textContent = stats.drop;
    dom.badgeOut.textContent = stats.out;
  }

  // ==========================================================================
  // РЕНДЕР ТАБЛИЦЫ РЕЗУЛЬТАТОВ
  // ==========================================================================
  function renderTable() {
    let items = state.diffItems;

    // 1. Фильтрация по активной вкладке
    if (state.activeFilter !== 'all') {
      items = items.filter(i => i.status === state.activeFilter);
    }

    // 2. Поиск по строке (наименование или артикул)
    if (state.searchQuery) {
      const q = state.searchQuery.toLowerCase();
      items = items.filter(i =>
        i.name.toLowerCase().includes(q) ||
        (i.code && i.code.toLowerCase().includes(q))
      );
    }

    // 3. Сортировка
    items.sort((a, b) => {
      let valA = a[state.sortField];
      let valB = b[state.sortField];

      if (valA === null || valA === undefined) valA = '';
      if (valB === null || valB === undefined) valB = '';

      if (typeof valA === 'string') {
        valA = valA.toLowerCase();
        valB = String(valB).toLowerCase();
        return state.sortDir === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
      } else {
        return state.sortDir === 'asc' ? (valA - valB) : (valB - valA);
      }
    });

    // 4. Отрисовка
    dom.tableBody.innerHTML = '';

    if (items.length === 0) {
      dom.tableEmptyState.style.display = 'flex';
      dom.footerStats.textContent = 'Показано 0 строк';
      return;
    }

    dom.tableEmptyState.style.display = 'none';
    dom.footerStats.textContent = `Показано ${items.length} из ${state.diffItems.length} позиций`;

    const fragment = document.createDocumentFragment();

    items.forEach(item => {
      const tr = document.createElement('tr');

      // Статус
      const tdStatus = document.createElement('td');
      tdStatus.innerHTML = `<span class="status-badge badge-${item.status}">${item.statusLabel}</span>`;
      tr.appendChild(tdStatus);

      // Артикул
      const tdCode = document.createElement('td');
      tdCode.textContent = item.code || '—';
      tdCode.style.color = 'var(--text-secondary)';
      tdCode.style.fontSize = '12px';
      tr.appendChild(tdCode);

      // Наименование
      const tdName = document.createElement('td');
      tdName.textContent = item.name;
      tdName.style.fontWeight = '500';
      tr.appendChild(tdName);

      // Было
      const tdPrev = document.createElement('td');
      tdPrev.style.textAlign = 'right';
      tdPrev.textContent = item.prevStock;
      tr.appendChild(tdPrev);

      // Стало
      const tdCurr = document.createElement('td');
      tdCurr.style.textAlign = 'right';
      tdCurr.innerHTML = `${item.currStock}${item.isTextStock ? `<span class="text-stock-tag" title="Текстовый маркер">${item.textStockNote}</span>` : ''}`;
      tr.appendChild(tdCurr);

      // Изменение
      const tdDelta = document.createElement('td');
      tdDelta.style.textAlign = 'right';

      let deltaClass = 'delta-neutral';
      let sign = '';
      if (item.stockDelta > 0) {
        deltaClass = 'delta-positive';
        sign = '+';
      } else if (item.stockDelta < 0) {
        deltaClass = 'delta-negative';
      }

      const pctStr = item.stockDeltaPct !== 0 ? ` (${sign}${item.stockDeltaPct}%)` : '';
      tdDelta.innerHTML = `<span class="delta-val ${deltaClass}">${sign}${item.stockDelta}${pctStr}</span>`;
      tr.appendChild(tdDelta);

      // Цена и дельта цены
      const tdPrice = document.createElement('td');
      tdPrice.style.textAlign = 'right';

      if (item.currPrice !== null) {
        let priceHtml = `${item.currPrice.toLocaleString('ru-RU')} ₽`;
        if (item.priceDelta !== null && item.priceDelta !== 0) {
          const pSign = item.priceDelta > 0 ? '+' : '';
          const pClass = item.priceDelta > 0 ? 'delta-negative' : 'delta-positive'; // рост цены обычно красный
          priceHtml += ` <br><span class="delta-val ${pClass}" style="font-size:11px;">${pSign}${item.priceDelta} ₽</span>`;
        }
        tdPrice.innerHTML = priceHtml;
      } else if (item.prevPrice !== null) {
        tdPrice.innerHTML = `<span style="text-decoration:line-through;color:var(--text-muted);">${item.prevPrice} ₽</span>`;
      } else {
        tdPrice.textContent = '—';
      }
      tr.appendChild(tdPrice);

      fragment.appendChild(tr);
    });

    dom.tableBody.appendChild(fragment);
  }

  // ==========================================================================
  // ДЕЙСТВИЯ: ПРИНЯТЬ, ОТКАТИТЬ, ВЫГРУЗИТЬ
  // ==========================================================================
  async function commitSnapshot() {
    if (!state.activeSupplierId || state.parsedItems.length === 0) {
      showToast('Нет распарсенных данных для сохранения', 'error');
      return;
    }

    const sup = state.suppliers.find(s => s.id === state.activeSupplierId);
    const supName = sup ? sup.name : 'Поставщик';

    try {
      dom.btnCommitSnapshot.disabled = true;
      showToast('Сохранение новой базы в IndexedDB...', 'info');

      // 1. Локальное сохранение
      const updatedRecord = await window.PriceStorage.commitNewSnapshot(
        state.activeSupplierId,
        state.parsedItems,
        state.currentMapping
      );

      state.baselineSnapshot = updatedRecord;
      updateBaselineInfo();
      updateRollbackButtonState();

      // 2. Облачная отправка (POST)
      const cloudRes = await window.PriceStorage.syncSnapshotToCloud(
        state.activeSupplierId,
        supName,
        state.currentMapping,
        state.parsedItems
      );

      if (cloudRes) {
        showToast('База успешно зафиксирована и отправлена в Google Sheets!', 'success');
      } else {
        showToast('База сохранена локально (офлайн)', 'success');
      }

      // Пересчитываем diff относительно новой базы
      recomputeDiff();
      renderTable();

    } catch (err) {
      showToast('Ошибка при фиксации базы: ' + err.message, 'error');
    } finally {
      dom.btnCommitSnapshot.disabled = false;
    }
  }

  async function rollbackSnapshot() {
    if (!state.activeSupplierId) return;

    if (!confirm('Вы уверены, что хотите восстановить предыдущий снимок базы остатков?')) {
      return;
    }

    try {
      dom.btnRollbackSnapshot.disabled = true;
      showToast('Откат базы к предыдущему снимку...', 'info');

      // 1. Локальный откат
      const updatedRecord = await window.PriceStorage.rollbackLocalSnapshot(state.activeSupplierId);
      state.baselineSnapshot = updatedRecord;
      updateBaselineInfo();
      updateRollbackButtonState();

      // 2. Облачный откат
      await window.PriceStorage.rollbackInCloud(state.activeSupplierId);

      showToast('База успешно возвращена к предыдущему состоянию!', 'success');

      if (state.parsedItems.length > 0) {
        recomputeDiff();
      } else {
        displayBaselineAsCurrent();
      }
      renderTable();

    } catch (err) {
      showToast('Ошибка при откате базы: ' + err.message, 'error');
    } finally {
      dom.btnRollbackSnapshot.disabled = false;
    }
  }

  // ==========================================================================
  // ВЫГРУЗКА СПИСКА (.XLSX И БУФЕР ОБМЕНА)
  // ==========================================================================
  function getFilteredItemsForExport() {
    let items = state.diffItems;

    if (state.activeFilter !== 'all') {
      items = items.filter(i => i.status === state.activeFilter);
    }

    if (state.searchQuery) {
      const q = state.searchQuery.toLowerCase();
      items = items.filter(i =>
        i.name.toLowerCase().includes(q) ||
        (i.code && i.code.toLowerCase().includes(q))
      );
    }

    return items;
  }

  function exportToExcel() {
    const items = getFilteredItemsForExport();
    if (items.length === 0) {
      showToast('Нет позиций для экспорта в текущей вкладке', 'error');
      return;
    }

    const sup = state.suppliers.find(s => s.id === state.activeSupplierId);
    const supName = sup ? sup.name.replace(/[^a-zа-я0-9]/gi, '_') : 'supplier';
    const filterName = state.activeFilter;
    const dateStr = new Date().toISOString().slice(0, 10);

    // Подготовка колонок для Excel таблицы
    const rows = items.map(item => ({
      'Статус': item.statusLabel,
      'Артикул': item.code || '',
      'Наименование': item.name,
      'Остаток было': item.prevStock,
      'Остаток стало': item.currStock,
      'Дельта остатка': item.stockDelta,
      'Изменение %': item.stockDeltaPct + '%',
      'Цена': item.currPrice !== null ? item.currPrice : (item.prevPrice !== null ? item.prevPrice : ''),
      'Дельта цены': item.priceDelta !== null ? item.priceDelta : ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Выгрузка');

    const fileName = `${supName}_${filterName}_${dateStr}.xlsx`;
    XLSX.writeFile(workbook, fileName);

    showToast(`Файл «${fileName}» сформирован (${items.length} строк)`, 'success');
    closeExportDropdown();
  }

  async function exportToClipboard() {
    const items = getFilteredItemsForExport();
    if (items.length === 0) {
      showToast('Нет позиций для копирования', 'error');
      return;
    }

    // Формируем TSV (табуляция), который идеально вставляется в Excel и Google Sheets
    const headers = ['Статус', 'Артикул', 'Наименование', 'Было', 'Стало', 'Изменение', 'Изм %', 'Цена', 'Дельта цены'];
    const lines = [headers.join('\t')];

    items.forEach(item => {
      lines.push([
        item.statusLabel,
        item.code || '',
        item.name,
        item.prevStock,
        item.currStock,
        item.stockDelta,
        item.stockDeltaPct + '%',
        item.currPrice !== null ? item.currPrice : '',
        item.priceDelta !== null ? item.priceDelta : ''
      ].join('\t'));
    });

    const textToCopy = lines.join('\n');

    try {
      await navigator.clipboard.writeText(textToCopy);
      showToast(`Скопировано в буфер: ${items.length} позиций`, 'success');
    } catch (e) {
      showToast('Не удалось скопировать в буфер: ' + e.message, 'error');
    }

    closeExportDropdown();
  }

  function toggleExportDropdown(e) {
    e.stopPropagation();
    dom.exportDropdownMenu.classList.toggle('show');
  }

  function closeExportDropdown() {
    dom.exportDropdownMenu.classList.remove('show');
  }

  // ==========================================================================
  // МОДАЛЬНОЕ ОКНО: НАСТРОЙКА КОЛОНОК И ПРЕДПРОСМОТР
  // ==========================================================================
  function openMappingModal() {
    if (!state.currentWorkbook) {
      showToast('Сначала загрузите Excel файл', 'error');
      return;
    }

    // 1. Заполняем селектор листов книги
    dom.sheetSelector.innerHTML = '';
    state.currentWorkbook.SheetNames.forEach(sheetName => {
      const opt = document.createElement('option');
      opt.value = sheetName;
      opt.textContent = sheetName;
      if (sheetName === state.currentSheetName) opt.selected = true;
      dom.sheetSelector.appendChild(opt);
    });

    // 2. Заполняем селектор номера строки заголовков (1..25)
    dom.headerRowSelector.innerHTML = '';
    const maxRows = Math.min(25, state.currentSheetData.length);
    for (let r = 0; r < maxRows; r++) {
      const opt = document.createElement('option');
      opt.value = r;
      opt.textContent = `Строка ${r + 1}`;
      if (r === state.currentHeaderRowIndex) opt.selected = true;
      dom.headerRowSelector.appendChild(opt);
    }

    renderPreviewTable();
    openModal('modalColumnMapping');
  }

  function renderPreviewTable() {
    const previewData = window.ExcelParser.getPreviewRows(state.currentSheetData, 10);
    const rows = previewData.rows;
    const maxCols = previewData.maxCols;

    const currentMapping = state.currentMapping || { nameCol: -1, stockCol: -1, codeCol: -1, priceCol: -1 };

    // Шапка с селекторами для каждой колонки
    dom.previewTableHead.innerHTML = '';
    const headerTr = document.createElement('tr');

    for (let c = 0; c < maxCols; c++) {
      const th = document.createElement('th');
      th.innerHTML = `
        <div>Колонка ${c + 1}</div>
        <select class="preview-col-select" data-col-index="${c}">
          <option value="none">[ Пропустить ]</option>
          <option value="name" ${c === currentMapping.nameCol ? 'selected' : ''}>Наименование (обяз.)</option>
          <option value="stock" ${c === currentMapping.stockCol ? 'selected' : ''}>Остаток (обяз.)</option>
          <option value="code" ${c === currentMapping.codeCol ? 'selected' : ''}>Артикул / Код</option>
          <option value="price" ${c === currentMapping.priceCol ? 'selected' : ''}>Цена</option>
        </select>
      `;
      headerTr.appendChild(th);
    }
    dom.previewTableHead.appendChild(headerTr);

    // Тело предпросмотра (до 10 строк)
    dom.previewTableBody.innerHTML = '';
    rows.forEach((row, rIdx) => {
      const tr = document.createElement('tr');
      if (rIdx === state.currentHeaderRowIndex) {
        tr.style.backgroundColor = 'var(--accent-light)';
        tr.style.fontWeight = 'bold';
      }

      for (let c = 0; c < maxCols; c++) {
        const td = document.createElement('td');
        td.textContent = row[c] !== undefined && row[c] !== null ? String(row[c]) : '';
        tr.appendChild(td);
      }
      dom.previewTableBody.appendChild(tr);
    });
  }

  async function applyCustomMapping() {
    const selects = dom.previewTableHead.querySelectorAll('.preview-col-select');
    const newMapping = {
      nameCol: -1,
      stockCol: -1,
      codeCol: -1,
      priceCol: -1
    };

    selects.forEach(sel => {
      const colIdx = parseInt(sel.dataset.colIndex, 10);
      const val = sel.value;
      if (val === 'name') newMapping.nameCol = colIdx;
      if (val === 'stock') newMapping.stockCol = colIdx;
      if (val === 'code') newMapping.codeCol = colIdx;
      if (val === 'price') newMapping.priceCol = colIdx;
    });

    if (newMapping.nameCol < 0) {
      showToast('Пожалуйста, выберите колонку «Наименование»', 'error');
      return;
    }
    if (newMapping.stockCol < 0) {
      showToast('Пожалуйста, выберите колонку «Остаток»', 'error');
      return;
    }

    state.currentMapping = newMapping;
    state.currentHeaderRowIndex = parseInt(dom.headerRowSelector.value, 10);

    // Сохраняем маппинг в профиле поставщика в IndexedDB
    const activeSup = state.suppliers.find(s => s.id === state.activeSupplierId);
    if (activeSup) {
      activeSup.mapping = newMapping;
      await window.PriceStorage.saveSupplier(activeSup);
    }

    closeModal('modalColumnMapping');
    showToast('Маппинг сохранен для этого поставщика', 'success');

    extractAndDiff();
  }

  function onSheetChange() {
    state.currentSheetName = dom.sheetSelector.value;
    state.currentSheetData = window.ExcelParser.getSheetData(state.currentWorkbook, state.currentSheetName);
    
    // Переопределяем заголовок для нового листа
    const headerDetection = window.ExcelParser.detectHeaderRow(state.currentSheetData);
    state.currentHeaderRowIndex = headerDetection.rowIndex;
    dom.headerRowSelector.value = state.currentHeaderRowIndex;

    renderPreviewTable();
  }

  function onHeaderRowChange() {
    state.currentHeaderRowIndex = parseInt(dom.headerRowSelector.value, 10);
    renderPreviewTable();
  }

  // ==========================================================================
  // МОДАЛЬНОЕ ОКНО: НАСТРОЙКИ ОБЛАКА (GAS)
  // ==========================================================================
  function openCloudSettingsModal() {
    const config = window.PriceStorage.getCloudConfig();
    dom.cloudUrlInput.value = config.url;
    dom.cloudPinInput.value = config.token;
    dom.cloudTestResult.style.display = 'none';

    openModal('modalCloudSettings');
  }

  async function testCloudConnection() {
    const url = dom.cloudUrlInput.value.trim();
    const token = dom.cloudPinInput.value.trim();

    if (!url || !token) {
      showCloudTestNotice('Введите URL и токен', 'error');
      return;
    }

    try {
      showCloudTestNotice('Проверка соединения...', 'info');
      dom.btnTestCloudConnection.disabled = true;

      const res = await window.PriceStorage.testCloudConnection(url, token);
      showCloudTestNotice(res.message || 'Соединение успешно!', 'success');

    } catch (err) {
      showCloudTestNotice('Ошибка: ' + err.message, 'error');
    } finally {
      dom.btnTestCloudConnection.disabled = false;
    }
  }

  async function syncCloudNow() {
    const url = dom.cloudUrlInput.value.trim();
    const token = dom.cloudPinInput.value.trim();

    if (!url || !token) {
      showCloudTestNotice('Сначала сохраните настройки подключения', 'error');
      return;
    }

    window.PriceStorage.setCloudConfig(url, token);

    try {
      dom.btnSyncCloudNow.disabled = true;
      showCloudTestNotice('Запрос свежей базы из Google Таблицы...', 'info');

      const res = await window.PriceStorage.syncFromCloud();
      if (res.hasUpdates) {
        showCloudTestNotice('База обновлена из облака!', 'success');
        state.suppliers = await window.PriceStorage.getAllSuppliers();
        renderSupplierTabs();
        await loadActiveSupplierSnapshot();
      } else {
        showCloudTestNotice('Локальная база уже актуальна', 'success');
      }

    } catch (err) {
      showCloudTestNotice('Ошибка синхронизации: ' + err.message, 'error');
    } finally {
      dom.btnSyncCloudNow.disabled = false;
    }
  }

  function saveCloudSettings() {
    const url = dom.cloudUrlInput.value.trim();
    const token = dom.cloudPinInput.value.trim();

    window.PriceStorage.setCloudConfig(url, token);
    closeModal('modalCloudSettings');
    showToast('Настройки облака сохранены', 'success');

    // Пробуем выполнить синхронизацию
    backgroundCloudSync();
  }

  function showCloudTestNotice(message, type = 'info') {
    dom.cloudTestResult.className = `modal-notice ${type}`;
    dom.cloudTestResult.textContent = message;
    dom.cloudTestResult.style.display = 'block';
  }

  async function backgroundCloudSync() {
    const { url, token } = window.PriceStorage.getCloudConfig();
    if (!url || !token) return;

    try {
      const res = await window.PriceStorage.syncFromCloud();
      if (res.hasUpdates) {
        state.suppliers = await window.PriceStorage.getAllSuppliers();
        renderSupplierTabs();
        await loadActiveSupplierSnapshot();
        showToast('Облачные остатки синхронизированы', 'success');
      }
    } catch (e) {
      console.warn('Background sync failed:', e);
    }
  }

  function setupCloudStatusListener() {
    window.PriceStorage.onSyncStatusChange((status, message) => {
      dom.cloudStatusBadge.className = `cloud-status-badge ${status}`;

      let label = 'Облако: ';
      if (status === 'synced') label += 'Синхронизировано';
      else if (status === 'syncing') label += 'Синхронизация...';
      else if (status === 'offline') label += 'Офлайн';
      else if (status === 'error') label += 'Ошибка синхронизации';
      else label += 'Не настроено';

      dom.cloudStatusText.textContent = label;
      dom.cloudStatusBadge.title = message || label;
    });
  }

  // ==========================================================================
  // МОДАЛЬНЫЕ ОКНА И ТОСТЫ (УТИЛИТЫ)
  // ==========================================================================
  function openModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) modal.classList.add('open');
  }

  function closeModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) modal.classList.remove('open');
  }

  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;

    dom.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => toast.remove(), 250);
    }, 3500);
  }

  // ==========================================================================
  // СЛУШАТЕЛИ СОБЫТИЙ (EVENT LISTENERS)
  // ==========================================================================
  function setupEventListeners() {
    // Тема
    dom.btnThemeToggle.addEventListener('click', toggleTheme);

    // Модалки шапки
    dom.cloudStatusBadge.addEventListener('click', openCloudSettingsModal);
    dom.btnOpenCloudSettings.addEventListener('click', openCloudSettingsModal);
    dom.btnOpenMapping.addEventListener('click', openMappingModal);
    dom.btnTuneColumns.addEventListener('click', openMappingModal);

    // Закрытие модалок по кнопкам с атрибутом [data-close-modal]
    document.querySelectorAll('[data-close-modal]').forEach(btn => {
      btn.addEventListener('click', () => {
        closeModal(btn.dataset.closeModal);
      });
    });

    // Закрытие модалок по клику на фон
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) overlay.classList.remove('open');
      });
    });

    // Добавление поставщика
    dom.btnAddSupplier.addEventListener('click', openAddSupplierModal);
    dom.btnSaveSupplier.addEventListener('click', saveSupplierFromModal);

    // Drag & Drop
    dom.dropzone.addEventListener('click', () => dom.fileInput.click());
    dom.btnChangeFile.addEventListener('click', () => dom.fileInput.click());

    dom.fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleFileSelect(e.target.files[0]);
      }
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      dom.dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dom.dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dom.dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dom.dropzone.classList.remove('dragover');
      });
    });

    dom.dropzone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      if (dt.files && dt.files[0]) {
        handleFileSelect(dt.files[0]);
      }
    });

    // Пороги
    dom.threshGrowthPct.addEventListener('input', onThresholdChange);
    dom.threshGrowthQty.addEventListener('input', onThresholdChange);
    dom.threshLowQty.addEventListener('input', onThresholdChange);

    // Вкладки-фильтры
    dom.filterTabs.forEach(tab => {
      tab.addEventListener('click', () => {
        dom.filterTabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        state.activeFilter = tab.dataset.filter;
        renderTable();
      });
    });

    // Поиск
    dom.searchInput.addEventListener('input', (e) => {
      state.searchQuery = e.target.value.trim();
      renderTable();
    });

    // Сортировка заголовков таблицы
    dom.tableHeaders.forEach(th => {
      th.addEventListener('click', () => {
        const field = th.dataset.sort;
        if (state.sortField === field) {
          state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          state.sortField = field;
          state.sortDir = (field === 'name' || field === 'code') ? 'asc' : 'desc';
        }

        // Обновляем индикаторы стрелочек
        dom.tableHeaders.forEach(h => {
          const ind = h.querySelector('.sort-indicator');
          if (h === th) {
            ind.textContent = state.sortDir === 'asc' ? '▲' : '▼';
          } else {
            ind.textContent = '';
          }
        });

        renderTable();
      });
    });

    // Кнопки действий
    dom.btnCommitSnapshot.addEventListener('click', commitSnapshot);
    dom.btnRollbackSnapshot.addEventListener('click', rollbackSnapshot);

    // Меню выгрузки
    dom.btnExportMenuToggle.addEventListener('click', toggleExportDropdown);
    dom.btnExportExcel.addEventListener('click', exportToExcel);
    dom.btnExportClipboard.addEventListener('click', exportToClipboard);
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.dropdown-container')) {
        closeExportDropdown();
      }
    });

    // Настройки облака
    dom.btnTestCloudConnection.addEventListener('click', testCloudConnection);
    dom.btnSyncCloudNow.addEventListener('click', syncCloudNow);
    dom.btnSaveCloudSettings.addEventListener('click', saveCloudSettings);

    // Настройка маппинга
    dom.sheetSelector.addEventListener('change', onSheetChange);
    dom.headerRowSelector.addEventListener('change', onHeaderRowChange);
    dom.btnApplyMapping.addEventListener('click', applyCustomMapping);
  }

  // Запуск при готовности DOM
  document.addEventListener('DOMContentLoaded', initApp);

})();
