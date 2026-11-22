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
    activeFilter: 'all',          // 'all' | 'new' | 'growth' | 'drop' | 'out'
    subFilter: 'all',             // 'all' | 'draft' | 'bottle'
    selectedMainCategory: 'all',  // 'all' | 'КРАФТ фасовка' | etc.
    selectedSubCategory: 'all',   // 'all' | 'Аляска Бревери' | etc.
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
    btnSupplierMenuToggle: document.getElementById('btnSupplierMenuToggle'),
    supplierDropdownMenu: document.getElementById('supplierDropdownMenu'),
    btnMenuAddSupplier: document.getElementById('btnMenuAddSupplier'),
    btnMenuDeleteSupplier: document.getElementById('btnMenuDeleteSupplier'),
    btnMenuDeleteSupplierText: document.getElementById('btnMenuDeleteSupplierText'),

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
    thresholdsCard: document.getElementById('thresholdsCard'),
    thresholdsHeader: document.getElementById('thresholdsHeader'),
    thresholdsSummary: document.getElementById('thresholdsSummary'),
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
    subFilterRow: document.getElementById('subFilterRow'),
    subFilterTabs: document.querySelectorAll('.sub-filter-tab'),
    subBadgeAll: document.getElementById('subBadgeAll'),
    subBadgeDraft: document.getElementById('subBadgeDraft'),
    subBadgeBottle: document.getElementById('subBadgeBottle'),
    groupChipsWrapper: document.getElementById('groupChipsWrapper'),
    groupMainChips: document.getElementById('groupMainChips'),
    groupSubChipsContainer: document.getElementById('groupSubChipsContainer'),
    groupSubChips: document.getElementById('groupSubChips'),
    searchInput: document.getElementById('searchInput'),

    btnCommitSnapshot: document.getElementById('btnCommitSnapshot'),
    btnRollbackSnapshot: document.getElementById('btnRollbackSnapshot'),
    btnClearSnapshot: document.getElementById('btnClearSnapshot'),
    btnExportMenuToggle: document.getElementById('btnExportMenuToggle'),
    exportDropdownMenu: document.getElementById('exportDropdownMenu'),
    btnExportExcel: document.getElementById('btnExportExcel'),
    btnExportClipboard: document.getElementById('btnExportClipboard'),

    // Таблица
    tableCard: document.querySelector('.table-card'),
    tableWrapper: document.querySelector('.table-wrapper'),
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
    btnPushCloudNow: document.getElementById('btnPushCloudNow'),
    btnSyncCloudNow: document.getElementById('btnSyncCloudNow'),
    btnSaveCloudSettings: document.getElementById('btnSaveCloudSettings'),
    btnClearPin: document.getElementById('btnClearPin'),

    // Модалка ввода PIN при входе
    modalPinPrompt: document.getElementById('modalPinPrompt'),
    sitePinInput: document.getElementById('sitePinInput'),
    btnTogglePinVisibility: document.getElementById('btnTogglePinVisibility'),
    rememberPinCheckbox: document.getElementById('rememberPinCheckbox'),
    pinPromptError: document.getElementById('pinPromptError'),
    btnSubmitPin: document.getElementById('btnSubmitPin'),
    btnSubmitPinText: document.getElementById('btnSubmitPinText'),
    btnSkipPin: document.getElementById('btnSkipPin'),

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

      // 3. Запрос PIN при входе на сайт (если еще не введен) или фоновая синхронизация
      checkAndPromptPinOnStartup();

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
      tab.title = 'Нажмите для переключения, удерживайте для переименования';

      const nameSpan = document.createElement('span');
      nameSpan.className = 'supplier-tab-name';
      nameSpan.textContent = sup.name;
      tab.appendChild(nameSpan);

      // Логика удержания (long-press) для переименования поставщика
      let pressTimer = null;
      let isLongPress = false;
      let startX = 0;
      let startY = 0;

      const startPress = (e) => {
        if (e.button && e.button !== 0) return;

        isLongPress = false;
        startX = e.clientX ?? (e.touches && e.touches[0] ? e.touches[0].clientX : 0);
        startY = e.clientY ?? (e.touches && e.touches[0] ? e.touches[0].clientY : 0);

        tab.classList.add('holding');

        pressTimer = setTimeout(() => {
          isLongPress = true;
          tab.classList.remove('holding');
          if (navigator.vibrate) {
            try { navigator.vibrate(40); } catch (_) { }
          }
          openEditSupplierModal(sup);
        }, 500);
      };

      const cancelPress = () => {
        tab.classList.remove('holding');
        if (pressTimer) {
          clearTimeout(pressTimer);
          pressTimer = null;
        }
      };

      const movePress = (e) => {
        if (!pressTimer) return;
        const currentX = e.clientX ?? (e.touches && e.touches[0] ? e.touches[0].clientX : 0);
        const currentY = e.clientY ?? (e.touches && e.touches[0] ? e.touches[0].clientY : 0);
        if (Math.abs(currentX - startX) > 8 || Math.abs(currentY - startY) > 8) {
          cancelPress();
        }
      };

      // Мышь
      tab.addEventListener('mousedown', startPress);
      tab.addEventListener('mouseup', cancelPress);
      tab.addEventListener('mouseleave', cancelPress);
      tab.addEventListener('mousemove', movePress);

      // Сенсорные экраны
      tab.addEventListener('touchstart', startPress, { passive: true });
      tab.addEventListener('touchend', cancelPress, { passive: true });
      tab.addEventListener('touchcancel', cancelPress, { passive: true });
      tab.addEventListener('touchmove', movePress, { passive: true });

      // Предотвращение контекстного меню при долгом тапе
      tab.addEventListener('contextmenu', (e) => {
        if (isLongPress) {
          e.preventDefault();
        }
      });

      // Переименование по двойному клику на ПК
      tab.addEventListener('dblclick', (e) => {
        e.preventDefault();
        cancelPress();
        openEditSupplierModal(sup);
      });

      // Обычный короткий клик - переключение поставщика
      tab.addEventListener('click', (e) => {
        if (isLongPress) {
          e.preventDefault();
          e.stopPropagation();
          isLongPress = false;
          return;
        }
        switchSupplier(sup.id);
      });

      dom.supplierTabsContainer.appendChild(tab);
    });

    // Синхронизируем кнопку удаления активного поставщика в меню
    const activeSup = state.suppliers.find(s => s.id === state.activeSupplierId);
    if (dom.btnMenuDeleteSupplierText && activeSup) {
      dom.btnMenuDeleteSupplierText.textContent = `Удалить «${activeSup.name}»`;
    }
    if (dom.btnMenuDeleteSupplier) {
      const canDelete = state.suppliers.length > 1;
      dom.btnMenuDeleteSupplier.disabled = !canDelete;
      dom.btnMenuDeleteSupplier.style.opacity = canDelete ? '1' : '0.4';
      dom.btnMenuDeleteSupplier.style.cursor = canDelete ? 'pointer' : 'not-allowed';
      dom.btnMenuDeleteSupplier.title = canDelete ? '' : 'Нельзя удалить единственного поставщика';
    }
  }

  function toggleSupplierDropdown() {
    if (dom.supplierDropdownMenu) {
      dom.supplierDropdownMenu.classList.toggle('show');
    }
  }

  function closeSupplierDropdown() {
    if (dom.supplierDropdownMenu) {
      dom.supplierDropdownMenu.classList.remove('show');
    }
  }

  function updateThresholdsSummary() {
    if (!dom.thresholdsSummary) return;
    const pct = dom.threshGrowthPct ? dom.threshGrowthPct.value : state.thresholds.growthPct;
    const qty = dom.threshGrowthQty ? dom.threshGrowthQty.value : state.thresholds.growthQty;
    const low = dom.threshLowQty ? dom.threshLowQty.value : state.thresholds.lowQty;
    dom.thresholdsSummary.textContent = `Рост ≥ ${pct}% или ≥ ${qty} шт. | Мало < ${low} шт.`;
  }

  async function switchSupplier(supplierId) {
    if (state.activeSupplierId === supplierId) return;

    state.activeSupplierId = supplierId;
    window.PriceStorage.setActiveSupplierId(supplierId);
    renderSupplierTabs();

    // Сброс загруженного файла для нового поставщика
    resetCurrentFile();
    await loadActiveSupplierSnapshot();

    // Фоновая проверка обновлений поставщика из облака
    if (window.PriceStorage.hasAuthenticatedPin()) {
      backgroundCloudSync();
    }
  }

  async function loadActiveSupplierSnapshot() {
    if (!state.activeSupplierId) return;

    const snap = await window.PriceStorage.getSnapshotRecord(state.activeSupplierId);
    state.baselineSnapshot = snap;

    updateBaselineInfo();
    updateRollbackButtonState();

    if (state.parsedItems.length > 0) {
      // Есть выбранный пользователем свежий файл
      recomputeDiff();
      renderTable();
    } else if (snap.current && snap.current.length > 0) {
      // Вариант 1: Показываем изменения последнего принятого прихода (current vs previous)
      recomputeDiff();
      renderTable();
    } else {
      state.diffItems = [];
      renderGroupChips();
      updateBadges({ all: 0, new: 0, growth: 0, drop: 0, out: 0 });
      dom.fileCard.classList.remove('visible');
      renderTable();
    }
  }

  function displayBaselineAsCurrent() {
    state.diffItems = state.baselineSnapshot.current.map(item => {
      const pkgInfo = (window.ExcelParser && typeof window.ExcelParser.detectPackageType === 'function')
        ? window.ExcelParser.detectPackageType(item.name, item.rawUnit || item.unit, item.category)
        : (item.packageType ? { type: item.packageType, subType: item.packageSubType || item.packageType, label: item.packageLabel || 'Товар', icon: item.packageIcon || '' } : { type: 'other', subType: 'other', label: 'Товар', icon: '📦' });

      return {
        status: 'normal',
        statusLabel: 'В базе',
        code: item.code || '',
        name: item.name,
        category: item.category || '',
        subCategory: item.subCategory || '',
        packageType: pkgInfo.type,
        packageSubType: pkgInfo.subType,
        packageLabel: pkgInfo.label,
        packageIcon: pkgInfo.icon,
        prevStock: item.stock,
        currStock: item.stock,
        stockDelta: 0,
        stockDeltaPct: 0,
        isTextStock: item.isTextStock,
        textStockNote: item.textStockNote
      };
    });

    const stats = {
      all: state.diffItems.length,
      new: 0,
      growth: 0,
      drop: 0,
      out: 0
    };
    renderGroupChips();
    updateBadges(stats);
    renderTable();
  }

  function updateBaselineInfo() {
    const sup = state.suppliers.find(s => s.id === state.activeSupplierId);
    const hasData = Boolean(state.baselineSnapshot && state.baselineSnapshot.current && state.baselineSnapshot.current.length > 0);

    let dateStr = 'нет данных';
    if (state.baselineSnapshot && state.baselineSnapshot.updatedAt) {
      try {
        const d = new Date(state.baselineSnapshot.updatedAt);
        if (!isNaN(d.getTime())) {
          dateStr = d.toLocaleString('ru-RU');
        }
      } catch (e) { }
    }

    const count = hasData ? state.baselineSnapshot.current.length : 0;
    const infoText = hasData ? `База: ${count} поз. (от ${dateStr})` : 'База: пустая (нет данных)';

    dom.loadedFileBaseline.textContent = infoText;
    dom.footerBaselineInfo.textContent = `Эталон: ${sup ? sup.name : ''} | ${infoText}`;
  }

  function updateRollbackButtonState() {
    const hasPrevious = state.baselineSnapshot && state.baselineSnapshot.previous && state.baselineSnapshot.previous.length > 0;
    const hasCurrent = state.baselineSnapshot && state.baselineSnapshot.current && state.baselineSnapshot.current.length > 0;

    if (dom.btnRollbackSnapshot) {
      dom.btnRollbackSnapshot.disabled = !hasPrevious;
    }
    if (dom.btnClearSnapshot) {
      dom.btnClearSnapshot.disabled = !(hasCurrent || hasPrevious);
    }
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
      dom.btnChangeFile.querySelector('span').textContent = 'Сменить файл';
      dom.btnOpenMapping.disabled = false;
      dom.btnCommitSnapshot.disabled = false;
      dom.btnCommitSnapshot.title = 'Зафиксировать текущий прайс как эталон в IndexedDB и отправить в Google Таблицу';

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
    state.selectedMainCategory = 'all';
    state.selectedSubCategory = 'all';
    state.subFilter = 'all';
    if (dom.groupChipsWrapper) dom.groupChipsWrapper.style.display = 'none';
    dom.fileInput.value = '';
    dom.btnCommitSnapshot.disabled = true;

    // Вариант 1: возвращаемся к показу зафиксированного последнего прихода
    recomputeDiff();
    renderTable();
  }

  // ==========================================================================
  // DIFF ENGINE (СРАВНЕНИЕ ОСТАТКОВ И КЛАССИФИКАЦИЯ)
  // ==========================================================================
  function recomputeDiff() {
    let currentList = [];
    let baselineList = [];
    let isShowingCommittedShipment = false;

    if (state.parsedItems.length > 0) {
      // Режим 1: Сравнение выбранного свежего файла с текущей базой
      currentList = state.parsedItems;
      baselineList = state.baselineSnapshot.current || [];
      isShowingCommittedShipment = false;
    } else if (state.baselineSnapshot.current && state.baselineSnapshot.current.length > 0) {
      // Режим 2 (Вариант 1): Показ изменений последнего принятого прихода (current vs previous)
      currentList = state.baselineSnapshot.current;
      baselineList = state.baselineSnapshot.previous || [];
      isShowingCommittedShipment = true;
    } else {
      state.diffItems = [];
      renderGroupChips();
      updateBadges({ all: 0, new: 0, growth: 0, drop: 0, out: 0 });
      dom.fileCard.classList.remove('visible');
      return;
    }

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

    // 2. Обрабатываем позиции из текущего списка (свежего или последнего прихода)
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

      // КЛАССИФИКАЦИЯ СТАТУСА
      let status = 'normal';
      let statusLabel = 'В наличии';

      // 1. Появилось (новинка): не было в прошлом снимке или был 0, а стал > 0
      if (!prev || (prevStock === 0 && currStock > 0)) {
        status = 'new';
        statusLabel = 'Появилось';
        newCount++;
      }
      // 2. Выбыло: был > 0, стал 0
      else if (currStock === 0 && prevStock > 0) {
        status = 'out';
        statusLabel = 'Выбыло (0)';
        outCount++;
      }
      // 3. Сильный рост / Поступление (+): вырос больше чем на порог % ИЛИ порог ед.
      else if (stockDelta > 0 && (stockDeltaPct >= state.thresholds.growthPct || stockDelta >= state.thresholds.growthQty)) {
        status = 'growth';
        statusLabel = '+';
        growthCount++;
      }
      // 4. Резкое падение / Мало (-): остаток упал или стал меньше критического порога
      else if (stockDelta < 0 && (currStock < state.thresholds.lowQty || stockDeltaPct <= -30)) {
        status = 'drop';
        statusLabel = '-';
        dropCount++;
      }

      // Определение тары (Разливное / Бутылочное)
      const pkgInfo = (window.ExcelParser && typeof window.ExcelParser.detectPackageType === 'function')
        ? window.ExcelParser.detectPackageType(curr.name, curr.rawUnit || curr.unit, curr.category)
        : (curr.packageType ? { type: curr.packageType, subType: curr.packageSubType || curr.packageType, label: curr.packageLabel || 'Товар', icon: curr.packageIcon || '' } : { type: 'other', subType: 'other', label: 'Товар', icon: '📦' });

      diffResults.push({
        status,
        statusLabel,
        code: curr.code || (prev ? prev.code : ''),
        name: curr.name,
        category: curr.category || (prev ? prev.category : ''),
        subCategory: curr.subCategory || (prev ? prev.subCategory : ''),
        packageType: pkgInfo.type,
        packageSubType: pkgInfo.subType,
        packageLabel: pkgInfo.label,
        packageIcon: pkgInfo.icon,
        prevStock,
        currStock,
        stockDelta,
        stockDeltaPct,
        isTextStock: curr.isTextStock,
        textStockNote: curr.textStockNote
      });
    }

    // 3. Обрабатываем позиции, которые БЫЛИ в базе с остатком > 0, но ИСЧЕЗЛИ
    for (const [key, prev] of prevMap.entries()) {
      if (!seenPrevKeys.has(key) && prev.stock > 0) {
        const pkgInfo = (window.ExcelParser && typeof window.ExcelParser.detectPackageType === 'function')
          ? window.ExcelParser.detectPackageType(prev.name, prev.rawUnit || prev.unit, prev.category)
          : (prev.packageType ? { type: prev.packageType, subType: prev.packageSubType || prev.packageType, label: prev.packageLabel || 'Товар', icon: prev.packageIcon || '' } : { type: 'other', subType: 'other', label: 'Товар', icon: '📦' });

        diffResults.push({
          status: 'out',
          statusLabel: 'Выбыло (нет в прайсе)',
          code: prev.code || '',
          name: prev.name,
          category: prev.category || '',
          subCategory: prev.subCategory || '',
          packageType: pkgInfo.type,
          packageSubType: pkgInfo.subType,
          packageLabel: pkgInfo.label,
          packageIcon: pkgInfo.icon,
          prevStock: prev.stock,
          currStock: 0,
          stockDelta: -prev.stock,
          stockDeltaPct: -100,
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

    renderGroupChips();
    updateBadges(state.diffStats);

    if (isShowingCommittedShipment) {
      updateCommittedShipmentCard(currentList.length);
    }
  }

  function updateCommittedShipmentCard(itemCount) {
    dom.fileCard.classList.add('visible');

    const activeSup = state.suppliers.find(s => s.id === state.activeSupplierId);
    const lastFileName = (activeSup && activeSup.mapping && activeSup.mapping.lastFileName)
      ? activeSup.mapping.lastFileName
      : (state.currentFileName || 'Последний принятый приход');

    dom.loadedFileName.textContent = lastFileName;
    dom.loadedFileStats.textContent = `Приход: ${itemCount} поз. (в облаке)`;

    let dateStr = '';
    if (state.baselineSnapshot && state.baselineSnapshot.updatedAt) {
      try {
        const d = new Date(state.baselineSnapshot.updatedAt);
        if (!isNaN(d.getTime())) {
          dateStr = d.toLocaleString('ru-RU');
        }
      } catch (e) { }
    }
    dom.loadedFileBaseline.textContent = dateStr ? `Зафиксировано: ${dateStr}` : 'В облачной базе';

    dom.btnCommitSnapshot.disabled = true;
    dom.btnCommitSnapshot.title = 'Текущий приход уже зафиксирован в базе. Выберите свежий файл для следующего прихода';
    if (dom.btnChangeFile.querySelector('span')) {
      dom.btnChangeFile.querySelector('span').textContent = 'Загрузить свежий приход';
    }
    dom.btnOpenMapping.disabled = false;
  }

  function updateBadges(stats) {
    dom.badgeAll.textContent = stats.all;
    dom.badgeNew.textContent = stats.new;
    dom.badgeGrowth.textContent = stats.growth;
    dom.badgeDrop.textContent = stats.drop;
    dom.badgeOut.textContent = stats.out;
    updateSubFilterBadges();
  }

  function updateSubFilterBadges() {
    if (!dom.subFilterRow) return;

    let baseItems = state.diffItems;
    if (state.activeFilter !== 'all') {
      baseItems = baseItems.filter(i => i.status === state.activeFilter);
    }

    // Учитываем фильтр по разделу и пивоварне (Вариант 1)
    if (state.selectedMainCategory && state.selectedMainCategory !== 'all') {
      baseItems = baseItems.filter(i => i.category === state.selectedMainCategory);
      if (state.selectedSubCategory && state.selectedSubCategory !== 'all') {
        baseItems = baseItems.filter(i => i.subCategory === state.selectedSubCategory);
      }
    }

    let draftCount = 0;
    let bottleCount = 0;

    for (const item of baseItems) {
      if (item.packageType === 'draft') {
        draftCount++;
      } else {
        bottleCount++;
      }
    }

    if (dom.subBadgeAll) dom.subBadgeAll.textContent = baseItems.length;
    if (dom.subBadgeDraft) dom.subBadgeDraft.textContent = draftCount;
    if (dom.subBadgeBottle) dom.subBadgeBottle.textContent = bottleCount;

    // Синхронизируем подсветку активного чипса
    if (dom.subFilterTabs) {
      dom.subFilterTabs.forEach(t => {
        t.classList.toggle('active', t.dataset.subFilter === state.subFilter);
      });
    }

    // Подфильтр отображается для вкладок 'new', 'growth' и 'all' при наличии позиций
    const isApplicable = ['new', 'growth', 'all'].includes(state.activeFilter);
    if (isApplicable && state.diffItems.length > 0) {
      dom.subFilterRow.classList.add('visible');
    } else {
      dom.subFilterRow.classList.remove('visible');
    }
  }

  /**
   * Интерактивные чипсы разделов и пивоварен из файла (Вариант 1)
   */
  function renderGroupChips() {
    if (!dom.groupChipsWrapper || !dom.groupMainChips) return;

    // Считаем распределение по категориям в контексте активного таба (Появилось, +, -, Выбыло или Все)
    let activeItems = state.diffItems;
    if (state.activeFilter !== 'all') {
      activeItems = activeItems.filter(i => i.status === state.activeFilter);
    }

    // Собираем все категории и подкатегории
    const categoryMap = new Map(); // mainCat -> Map(subCat -> totalAllCount)

    state.diffItems.forEach(item => {
      const mainCat = (item.category || '').trim();
      const subCat = (item.subCategory || '').trim();
      if (mainCat) {
        if (!categoryMap.has(mainCat)) {
          categoryMap.set(mainCat, new Map());
        }
        if (subCat) {
          const subMap = categoryMap.get(mainCat);
          if (!subMap.has(subCat)) {
            subMap.set(subCat, 0);
          }
        }
      }
    });

    // Если в файле вообще нет разделов — полностью скрываем блок чипсов
    if (categoryMap.size === 0) {
      dom.groupChipsWrapper.style.display = 'none';
      state.selectedMainCategory = 'all';
      state.selectedSubCategory = 'all';
      return;
    }

    dom.groupChipsWrapper.style.display = 'flex';

    // Считаем актуальные количества для текущей активной вкладки
    const mainCounts = new Map();
    const subCounts = new Map(); // "mainCat:::subCat" -> count

    activeItems.forEach(item => {
      const mainCat = (item.category || '').trim();
      const subCat = (item.subCategory || '').trim();
      if (mainCat) {
        mainCounts.set(mainCat, (mainCounts.get(mainCat) || 0) + 1);
        if (subCat) {
          const subKey = `${mainCat}:::${subCat}`;
          subCounts.set(subKey, (subCounts.get(subKey) || 0) + 1);
        }
      }
    });

    // 1. РЕНДЕР ЯРУСА 1: Главные разделы
    dom.groupMainChips.innerHTML = '';

    // Чипс "Все разделы"
    const totalActiveWithCat = activeItems.length;
    const btnAll = document.createElement('button');
    btnAll.type = 'button';
    btnAll.className = `group-chip ${state.selectedMainCategory === 'all' ? 'active' : ''}`;
    btnAll.innerHTML = `<span class="chip-icon">📁</span><span>Все разделы</span><span class="chip-badge">${totalActiveWithCat}</span>`;
    btnAll.addEventListener('click', () => {
      state.selectedMainCategory = 'all';
      state.selectedSubCategory = 'all';
      renderGroupChips();
      updateSubFilterBadges();
      renderTable();
    });
    dom.groupMainChips.appendChild(btnAll);

    // Чипсы для каждого главного раздела
    categoryMap.forEach((subMap, mainCat) => {
      const count = mainCounts.get(mainCat) || 0;
      const icon = getCategoryIcon(mainCat);
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `group-chip ${state.selectedMainCategory === mainCat ? 'active' : ''}`;
      btn.innerHTML = `<span class="chip-icon">${icon}</span><span>${escapeHtml(mainCat)}</span><span class="chip-badge">${count}</span>`;
      btn.addEventListener('click', () => {
        state.selectedMainCategory = mainCat;
        state.selectedSubCategory = 'all'; // сброс фильтра пивоварни при смене категории
        renderGroupChips();
        updateSubFilterBadges();
        renderTable();
      });
      dom.groupMainChips.appendChild(btn);
    });

    // 2. РЕНДЕР ЯРУСА 2: Подразделы / Бренды
    // Отображается только если выбран конкретный раздел И у него есть подкатегории
    if (state.selectedMainCategory !== 'all' && categoryMap.has(state.selectedMainCategory)) {
      const subMap = categoryMap.get(state.selectedMainCategory);
      if (subMap && subMap.size > 0) {
        dom.groupSubChipsContainer.style.display = 'flex';
        dom.groupSubChips.innerHTML = '';

        // Кнопка "Все в этом разделе"
        const mainCatCount = mainCounts.get(state.selectedMainCategory) || 0;
        const btnSubAll = document.createElement('button');
        btnSubAll.type = 'button';
        btnSubAll.className = `group-sub-chip ${state.selectedSubCategory === 'all' ? 'active' : ''}`;
        btnSubAll.innerHTML = `<span>Все в разделе</span><span class="sub-chip-badge">${mainCatCount}</span>`;
        btnSubAll.addEventListener('click', () => {
          state.selectedSubCategory = 'all';
          renderGroupChips();
          updateSubFilterBadges();
          renderTable();
        });
        dom.groupSubChips.appendChild(btnSubAll);

        // Кнопки для каждого бренда/пивоварни
        subMap.forEach((_, subCat) => {
          const subKey = `${state.selectedMainCategory}:::${subCat}`;
          const count = subCounts.get(subKey) || 0;
          const btnSub = document.createElement('button');
          btnSub.type = 'button';
          btnSub.className = `group-sub-chip ${state.selectedSubCategory === subCat ? 'active' : ''}`;
          btnSub.innerHTML = `<span>${escapeHtml(subCat)}</span><span class="sub-chip-badge">${count}</span>`;
          btnSub.addEventListener('click', () => {
            state.selectedSubCategory = subCat;
            renderGroupChips();
            updateSubFilterBadges();
            renderTable();
          });
          dom.groupSubChips.appendChild(btnSub);
        });
      } else {
        dom.groupSubChipsContainer.style.display = 'none';
        state.selectedSubCategory = 'all';
      }
    } else {
      dom.groupSubChipsContainer.style.display = 'none';
      state.selectedSubCategory = 'all';
    }
  }

  function getCategoryIcon(catName) {
    const lower = String(catName || '').toLowerCase();
    if (lower.includes('разлив') || lower.includes('кег')) return '🍺';
    if (lower.includes('фасовк') || lower.includes('бут') || lower.includes('стекл') || lower.includes('ж/б') || lower.includes('ж.б')) return '📦';
    if (lower.includes('вод')) return '💧';
    if (lower.includes('сидр')) return '🍏';
    if (lower.includes('снек') || lower.includes('чипс') || lower.includes('грен')) return '🥨';
    if (lower.includes('росси')) return '🇷🇺';
    if (lower.includes('крафт')) return '✨';
    if (lower.includes('лимонад') || lower.includes('сок') || lower.includes('напит')) return '🥤';
    if (lower.includes('импорт')) return '🌍';
    return '📁';
  }

  function escapeHtml(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
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

    // 1.1. Фильтрация по таре (Разливное / Бутылочное)
    if (state.subFilter === 'draft') {
      items = items.filter(i => i.packageType === 'draft');
    } else if (state.subFilter === 'bottle') {
      items = items.filter(i => i.packageType !== 'draft');
    }

    // 1.2. Фильтрация по разделу и пивоварне (Вариант 1)
    if (state.selectedMainCategory && state.selectedMainCategory !== 'all') {
      items = items.filter(i => i.category === state.selectedMainCategory);
      if (state.selectedSubCategory && state.selectedSubCategory !== 'all') {
        items = items.filter(i => i.subCategory === state.selectedSubCategory);
      }
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

      // 1. Статус
      const tdStatus = document.createElement('td');
      tdStatus.className = 'col-status';
      tdStatus.innerHTML = `<span class="status-badge badge-${item.status}">${item.statusLabel}</span>`;
      tr.appendChild(tdStatus);

      // 2. Артикул
      const tdCode = document.createElement('td');
      tdCode.className = 'col-code cell-code';
      tdCode.textContent = item.code || '—';
      tr.appendChild(tdCode);

      // 3. Наименование + маркер тары + бейдж категории + мобильный 2-й ярус (статус + тара + артикул)
      const tdName = document.createElement('td');
      tdName.className = 'col-name';

      let pkgBadge = '';
      if (item.packageType === 'draft') {
        pkgBadge = `<span class="package-badge package-draft" title="Разливное пиво (кег)">🍺 Разлив</span>`;
      } else if (item.packageSubType === 'can') {
        pkgBadge = `<span class="package-badge package-can" title="Банка (ж/б)">🥫 Банка</span>`;
      } else if (item.packageSubType === 'bottle') {
        pkgBadge = `<span class="package-badge package-bottle" title="Бутылочное / стекло">🍾 Бутылка</span>`;
      } else if (item.packageSubType === 'other' && item.packageLabel) {
        pkgBadge = `<span class="package-badge package-other" title="Фасованный товар">${item.packageIcon || '📦'} ${item.packageLabel}</span>`;
      }

      const groupTag = item.subCategory
        ? `<span class="item-category-tag" title="Раздел: ${escapeHtml(item.category || '')} / ${escapeHtml(item.subCategory)}">${escapeHtml(item.subCategory)}</span>`
        : (item.category ? `<span class="item-category-tag" title="Раздел: ${escapeHtml(item.category)}">${escapeHtml(item.category)}</span>` : '');

      tdName.innerHTML = `
        <div class="product-name-row">
          <span class="product-name">${item.name}</span>
          ${pkgBadge}
          ${groupTag}
        </div>
        <div class="mobile-item-meta">
          <span class="status-badge badge-${item.status}">${item.statusLabel}</span>
          ${pkgBadge}
          ${groupTag}
          ${item.code ? `<span class="mobile-sku">Арт: ${item.code}</span>` : ''}
        </div>
      `;
      tr.appendChild(tdName);

      // 4. Было
      const tdPrev = document.createElement('td');
      tdPrev.className = 'col-prev';
      tdPrev.style.textAlign = 'right';
      tdPrev.textContent = item.prevStock;
      tr.appendChild(tdPrev);

      // 5. Стало
      const tdCurr = document.createElement('td');
      tdCurr.className = 'col-curr';
      tdCurr.style.textAlign = 'right';
      tdCurr.innerHTML = `${item.currStock}${item.isTextStock ? `<span class="text-stock-tag" title="Текстовый маркер">${item.textStockNote}</span>` : ''}`;
      tr.appendChild(tdCurr);

      // 6. Изменение остатка (на мобильном: 1-й ярус "Было ➔ Стало", 2-й ярус Дельта)
      const tdDelta = document.createElement('td');
      tdDelta.className = 'col-delta';
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
      const textStockBadge = item.isTextStock ? `<span class="text-stock-tag">${item.textStockNote}</span>` : '';

      tdDelta.innerHTML = `
        <div class="mobile-stock-flow">${item.prevStock} ➔ ${item.currStock}${textStockBadge}</div>
        <span class="delta-val ${deltaClass}">${sign}${item.stockDelta}${pctStr}</span>
      `;
      tr.appendChild(tdDelta);

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
      showToast('Сохранение новой базы в IndexedDB и Google Sheets...', 'info');

      // Сохраняем имя файла в объекте маппинга для синхронизации в облако
      const mapping = state.currentMapping ? { ...state.currentMapping } : {};
      mapping.lastFileName = state.currentFileName || 'price.xlsx';
      mapping.lastCommitTime = new Date().toISOString();
      state.currentMapping = mapping;

      // 1. Локальное сохранение
      const updatedRecord = await window.PriceStorage.commitNewSnapshot(
        state.activeSupplierId,
        state.parsedItems,
        mapping
      );

      state.baselineSnapshot = updatedRecord;
      updateBaselineInfo();
      updateRollbackButtonState();

      // 2. Облачная отправка (POST)
      let cloudSuccess = false;
      try {
        const cloudRes = await window.PriceStorage.syncSnapshotToCloud(
          state.activeSupplierId,
          supName,
          mapping,
          state.baselineSnapshot.current
        );
        if (cloudRes) cloudSuccess = true;
      } catch (cloudErr) {
        console.warn('Cloud sync error on commit:', cloudErr);
        showToast('База зафиксирована локально (ошибка облака: ' + cloudErr.message + ')', 'warning');
      }

      if (cloudSuccess) {
        showToast('База успешно зафиксирована и отправлена в Google Sheets!', 'success');
      }

      // Вариант 1: очищаем буфер свежего файла, переключая систему
      // на отображение изменений принятого прихода (current vs previous)
      state.parsedItems = [];
      state.currentWorkbook = null;
      state.currentSheetData = [];
      dom.fileInput.value = '';

      // Пересчитываем diff: теперь он показывает разницу между только что принятым приходом и previous!
      try {
        recomputeDiff();
        renderTable();
      } catch (renderErr) {
        console.error('Render error after commit:', renderErr);
      }

    } catch (err) {
      showToast('Ошибка при фиксации базы: ' + err.message, 'error');
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
      updateRollbackButtonState();
    }
  }

  async function clearSnapshot() {
    if (!state.activeSupplierId) return;

    const sup = state.suppliers.find(s => s.id === state.activeSupplierId);
    const supName = sup ? sup.name : 'поставщика';

    const confirmed = confirm(
      `Вы действительно хотите полностью очистить базу остатков для поставщика «${supName}»?\n\n` +
      `Все зафиксированные остатки (текущие и предыдущие) будут безвозвратно удалены из браузера и Google Таблицы.`
    );
    if (!confirmed) return;

    try {
      if (dom.btnClearSnapshot) dom.btnClearSnapshot.disabled = true;
      if (dom.btnRollbackSnapshot) dom.btnRollbackSnapshot.disabled = true;
      showToast('Очистка базы в IndexedDB и Google Sheets...', 'info');

      // 1. Локальная очистка в IndexedDB
      const clearedRecord = await window.PriceStorage.clearLocalSnapshot(state.activeSupplierId);
      state.baselineSnapshot = clearedRecord;

      // 2. Сброс состояния в памяти
      state.parsedItems = [];
      state.currentWorkbook = null;
      state.currentSheetData = [];
      state.currentFileName = '';
      state.diffItems = [];
      state.selectedMainCategory = 'all';
      state.selectedSubCategory = 'all';
      state.subFilter = 'all';
      if (dom.groupChipsWrapper) dom.groupChipsWrapper.style.display = 'none';
      dom.fileInput.value = '';

      if (sup && sup.mapping) {
        delete sup.mapping.lastFileName;
        delete sup.mapping.lastCommitTime;
      }

      // 3. Очистка в облаке Google Sheets
      await window.PriceStorage.clearSnapshotInCloud(state.activeSupplierId, supName);

      // 4. Сброс UI
      dom.fileCard.classList.remove('visible');
      updateBadges({ all: 0, new: 0, growth: 0, drop: 0, out: 0 });
      updateBaselineInfo();
      updateRollbackButtonState();
      renderTable();

      showToast(`База для поставщика «${supName}» успешно очищена!`, 'success');

    } catch (err) {
      console.error('Clear snapshot error:', err);
      showToast('Ошибка при очистке базы: ' + err.message, 'error');
      updateRollbackButtonState();
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

    if (state.subFilter === 'draft') {
      items = items.filter(i => i.packageType === 'draft');
    } else if (state.subFilter === 'bottle') {
      items = items.filter(i => i.packageType !== 'draft');
    }

    if (state.selectedMainCategory && state.selectedMainCategory !== 'all') {
      items = items.filter(i => i.category === state.selectedMainCategory);
      if (state.selectedSubCategory && state.selectedSubCategory !== 'all') {
        items = items.filter(i => i.subCategory === state.selectedSubCategory);
      }
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
    const subFilterSuffix = state.subFilter !== 'all' ? `_${state.subFilter}` : '';
    const dateStr = new Date().toISOString().slice(0, 10);

    // Подготовка колонок для Excel таблицы
    const rows = items.map(item => {
      return {
        'Статус': item.statusLabel,
        'Тара': item.packageLabel || '',
        'Раздел': item.category || '',
        'Подраздел / Бренд': item.subCategory || '',
        'Артикул': item.code || '',
        'Наименование': item.name,
        'Остаток было': item.prevStock,
        'Остаток стало': item.currStock,
        'Дельта остатка': item.stockDelta,
        'Изменение %': item.stockDeltaPct + '%'
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Выгрузка');

    const fileName = `${supName}_${filterName}${subFilterSuffix}_${dateStr}.xlsx`;
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
    const headers = ['Статус', 'Тара', 'Раздел', 'Подраздел', 'Артикул', 'Наименование', 'Было', 'Стало', 'Изменение', 'Изм %'];
    const lines = [headers.join('\t')];

    items.forEach(item => {
      const parts = [
        item.statusLabel,
        item.packageLabel || '',
        item.category || '',
        item.subCategory || '',
        item.code || '',
        item.name,
        item.prevStock,
        item.currStock,
        item.stockDelta,
        item.stockDeltaPct + '%'
      ];
      lines.push(parts.join('\t'));
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

    const currentMapping = state.currentMapping || { nameCol: -1, stockCol: -1, codeCol: -1 };

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
      codeCol: -1
    };

    selects.forEach(sel => {
      const colIdx = parseInt(sel.dataset.colIndex, 10);
      const val = sel.value;
      if (val === 'name') newMapping.nameCol = colIdx;
      if (val === 'stock') newMapping.stockCol = colIdx;
      if (val === 'code') newMapping.codeCol = colIdx;
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

  async function pushCloudNow() {
    const url = dom.cloudUrlInput.value.trim();
    const token = dom.cloudPinInput.value.trim();

    if (!url || !token) {
      showCloudTestNotice('Сначала сохраните настройки подключения', 'error');
      return;
    }

    if (!state.activeSupplierId) {
      showCloudTestNotice('Выберите поставщика для выгрузки', 'error');
      return;
    }

    window.PriceStorage.setCloudConfig(url, token);

    try {
      if (dom.btnPushCloudNow) dom.btnPushCloudNow.disabled = true;
      showCloudTestNotice('Выгрузка локальной базы в Google Таблицу...', 'info');

      const res = await window.PriceStorage.pushLocalSnapshotToCloud(state.activeSupplierId);
      if (res && res.success) {
        showCloudTestNotice('Локальная база успешно выгружена в Google Таблицу!', 'success');
        showToast('База поставщика успешно сохранена в Google Sheets!', 'success');
      } else {
        showCloudTestNotice('Ошибка выгрузки: ' + (res ? res.error : 'неизвестно'), 'error');
      }
    } catch (err) {
      showCloudTestNotice('Ошибка выгрузки в облако: ' + err.message, 'error');
    } finally {
      if (dom.btnPushCloudNow) dom.btnPushCloudNow.disabled = false;
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

      const res = await window.PriceStorage.syncFromCloud(true);
      if (res.hasUpdates || (res.suppliers && res.suppliers.length > 0)) {
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

  async function backgroundCloudSync(force = false) {
    const { url, token } = window.PriceStorage.getCloudConfig();
    if (!url || !token) return;

    try {
      const res = await window.PriceStorage.syncFromCloud(force);
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

  // ==========================================================================
  // АВТОРИЗАЦИЯ И ЗАПРОС PIN ПРИ ВХОДЕ НА САЙТ
  // ==========================================================================
  function checkAndPromptPinOnStartup() {
    const hasPin = window.PriceStorage.hasAuthenticatedPin();
    if (!hasPin) {
      // Если PIN не сохранен или сессия не активна — запрашиваем PIN при входе
      openPinPromptModal();
    } else {
      // Если PIN уже авторизован — запускаем фоновую синхронизацию
      backgroundCloudSync();
    }
  }

  function openPinPromptModal() {
    if (dom.sitePinInput) dom.sitePinInput.value = '';
    showPinPromptError('');
    openModal('modalPinPrompt');
    setTimeout(() => {
      if (dom.sitePinInput) dom.sitePinInput.focus();
    }, 120);
  }

  function closePinPromptModal() {
    closeModal('modalPinPrompt');
  }

  function showPinPromptError(msg) {
    if (!dom.pinPromptError) return;
    if (msg) {
      dom.pinPromptError.textContent = msg;
      dom.pinPromptError.style.display = 'block';
    } else {
      dom.pinPromptError.textContent = '';
      dom.pinPromptError.style.display = 'none';
    }
  }

  function setPinSubmitLoading(isLoading) {
    if (dom.btnSubmitPin) dom.btnSubmitPin.disabled = isLoading;
    if (dom.btnSubmitPinText) {
      dom.btnSubmitPinText.textContent = isLoading ? 'Проверка PIN...' : 'Войти и подключить облако';
    }
  }

  async function submitPinFromPrompt() {
    const pin = dom.sitePinInput ? dom.sitePinInput.value.trim() : '';
    if (!pin) {
      showPinPromptError('Пожалуйста, введите PIN-код доступа');
      if (dom.sitePinInput) dom.sitePinInput.focus();
      return;
    }

    const config = window.PriceStorage.getCloudConfig();
    const url = config.url || window.PriceStorage.DEFAULT_CLOUD_URL;
    const remember = dom.rememberPinCheckbox ? dom.rememberPinCheckbox.checked : true;

    try {
      setPinSubmitLoading(true);
      showPinPromptError('');

      // 1. Проверяем токен через тестовый запрос к Google Apps Script
      await window.PriceStorage.testCloudConnection(url, pin);

      // 2. Сохраняем проверенный PIN
      window.PriceStorage.setCloudConfig(url, pin, remember);

      // 3. Закрываем модальное окно
      closePinPromptModal();
      showToast('PIN подтвержден! Облако успешно подключено', 'success');

      // 4. Запускаем принудительную синхронизацию остатков из Google Таблицы
      const res = await window.PriceStorage.syncFromCloud(true);
      if (res.hasUpdates || (res.suppliers && res.suppliers.length > 0)) {
        state.suppliers = await window.PriceStorage.getAllSuppliers();
        renderSupplierTabs();
        await loadActiveSupplierSnapshot();
      }

    } catch (err) {
      console.error('PIN verification error:', err);
      showPinPromptError(err.message || 'Неверный PIN-код или ошибка подключения');
      if (dom.sitePinInput) {
        dom.sitePinInput.focus();
        dom.sitePinInput.select();
      }
    } finally {
      setPinSubmitLoading(false);
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
  // СВАЙП-НАВИГАЦИЯ И ФИЛЬТРЫ (SWIPE & FILTER NAVIGATION)
  // ==========================================================================
  const FILTER_ORDER = ['all', 'new', 'growth', 'drop', 'out'];

  /**
   * Активация фильтра с синхронизацией UI и опциональной анимацией сдвига
   * @param {string} filterName - 'all' | 'new' | 'growth' | 'drop' | 'out'
   * @param {number} direction - 0 (клик), 1 (свайп влево/следующий), -1 (свайп вправо/предыдущий)
   */
  function setActiveFilter(filterName, direction = 0) {
    if (!FILTER_ORDER.includes(filterName)) return;

    state.activeFilter = filterName;
    state.subFilter = 'all'; // сброс в "Все форматы", чтобы сразу видеть полный срез новой вкладки

    // Подсветка активной вкладки и ее автопрокрутка в центр экрана на мобильных устройствах
    dom.filterTabs.forEach(t => {
      const isActive = t.dataset.filter === filterName;
      t.classList.toggle('active', isActive);
      if (isActive) {
        t.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
      }
    });

    renderGroupChips(); // обновляет счетчики на чипсах категорий
    updateSubFilterBadges();

    // Сброс вертикального скролла таблицы к верху списка при смене фильтра
    if (dom.tableWrapper) {
      dom.tableWrapper.scrollTop = 0;
    }

    // Анимация плавного сдвига контента при свайпе
    if (direction !== 0 && dom.tableWrapper) {
      const animClass = direction > 0 ? 'swipe-slide-left' : 'swipe-slide-right';
      dom.tableWrapper.classList.remove('swipe-slide-left', 'swipe-slide-right', 'swipe-bounce-left', 'swipe-bounce-right');
      void dom.tableWrapper.offsetWidth; // принудительный reflow для перезапуска CSS-анимации
      dom.tableWrapper.classList.add(animClass);
      setTimeout(() => {
        dom.tableWrapper.classList.remove(animClass);
      }, 250);
    }

    renderTable();
  }

  /**
   * Переключение на соседний фильтр по шагу (+1 или -1)
   */
  function switchFilterByStep(step) {
    const currentIndex = FILTER_ORDER.indexOf(state.activeFilter);
    if (currentIndex === -1) return false;

    const nextIndex = currentIndex + step;
    if (nextIndex < 0 || nextIndex >= FILTER_ORDER.length) {
      // Эффект отпружинивания (bounce) при попытке свайпа за пределы крайних вкладок
      if (dom.tableWrapper) {
        const bounceClass = step > 0 ? 'swipe-bounce-left' : 'swipe-bounce-right';
        dom.tableWrapper.classList.remove('swipe-slide-left', 'swipe-slide-right', 'swipe-bounce-left', 'swipe-bounce-right');
        void dom.tableWrapper.offsetWidth;
        dom.tableWrapper.classList.add(bounceClass);
        setTimeout(() => {
          dom.tableWrapper.classList.remove(bounceClass);
        }, 240);
      }
      return false;
    }

    // Тактильный виброотклик на смартфонах
    if (typeof navigator !== 'undefined' && navigator.vibrate) {
      try {
        navigator.vibrate(10);
      } catch (_) {}
    }

    const nextFilter = FILTER_ORDER[nextIndex];
    setActiveFilter(nextFilter, step);
    return true;
  }

  /**
   * Инициализация мобильного жестового свайпа по области таблицы
   */
  function initTouchSwipe() {
    const target = dom.tableCard || dom.tableWrapper;
    if (!target) return;

    let touchStartX = 0;
    let touchStartY = 0;
    let touchStartTime = 0;
    let isSwiping = false;

    target.addEventListener('touchstart', (e) => {
      // Отслеживаем только одиночные касания
      if (e.touches.length !== 1) return;

      // Не обрабатываем свайпы, если открыто модальное окно
      if (document.querySelector('.modal-overlay.open')) {
        isSwiping = false;
        return;
      }

      const touch = e.touches[0];
      touchStartX = touch.clientX;
      touchStartY = touch.clientY;
      touchStartTime = Date.now();
      isSwiping = true;

      // Защита от системных жестов истории "Назад/Вперед" в iOS Safari и Android Chrome
      const edgeMargin = 25;
      if (touchStartX < edgeMargin || touchStartX > window.innerWidth - edgeMargin) {
        isSwiping = false;
        return;
      }

      // Не перехватываем свайп, если касание началось на интерактивном элементе ввода, меню или кнопке
      const el = e.target;
      if (el && el.closest('input, select, textarea, button, a, .dropdown-menu, .modal-card, .btn-icon')) {
        isSwiping = false;
        return;
      }
    }, { passive: true });

    target.addEventListener('touchmove', (e) => {
      if (!isSwiping || e.touches.length !== 1) return;

      const touch = e.touches[0];
      const deltaX = touch.clientX - touchStartX;
      const deltaY = touch.clientY - touchStartY;

      // Если пользователь листает таблицу вертикально, отменяем свайп фильтров
      if (Math.abs(deltaY) > 25 && Math.abs(deltaY) > Math.abs(deltaX)) {
        isSwiping = false;
      }
    }, { passive: true });

    target.addEventListener('touchend', (e) => {
      if (!isSwiping) return;
      isSwiping = false;

      if (!e.changedTouches || e.changedTouches.length === 0) return;
      const touch = e.changedTouches[0];
      const deltaX = touch.clientX - touchStartX;
      const deltaY = touch.clientY - touchStartY;
      const duration = Date.now() - touchStartTime;

      // Критерии валидного горизонтального свайпа:
      // 1. Длина смещения по X не менее 45px
      // 2. Угол жеста выраженно горизонтальный: |deltaX| > |deltaY| * 1.5
      // 3. Быстрый жест (не более 500мс)
      if (
        Math.abs(deltaX) >= 45 &&
        Math.abs(deltaX) > Math.abs(deltaY) * 1.5 &&
        duration <= 500
      ) {
        if (deltaX < 0) {
          // Свайп влево (палец движется справа налево) -> Следующий фильтр
          switchFilterByStep(1);
        } else {
          // Свайп вправо (палец движется слева направо) -> Предыдущий фильтр
          switchFilterByStep(-1);
        }
      }
    }, { passive: true });

    target.addEventListener('touchcancel', () => {
      isSwiping = false;
    }, { passive: true });
  }

  // ==========================================================================
  // СЛУШАТЕЛИ СОБЫТИЙ (EVENT LISTENERS)
  // ==========================================================================
  function setupEventListeners() {
    // Тема
    dom.btnThemeToggle.addEventListener('click', toggleTheme);

    // Модалки шапки
    dom.cloudStatusBadge.addEventListener('click', () => {
      if (!window.PriceStorage.hasAuthenticatedPin()) {
        openPinPromptModal();
      } else {
        openCloudSettingsModal();
      }
    });
    dom.btnOpenCloudSettings.addEventListener('click', openCloudSettingsModal);
    dom.btnOpenMapping.addEventListener('click', openMappingModal);
    dom.btnTuneColumns.addEventListener('click', openMappingModal);

    // Авторизация и ввод PIN при входе
    if (dom.btnSubmitPin) {
      dom.btnSubmitPin.addEventListener('click', submitPinFromPrompt);
    }
    if (dom.sitePinInput) {
      dom.sitePinInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') submitPinFromPrompt();
      });
    }
    if (dom.btnTogglePinVisibility) {
      dom.btnTogglePinVisibility.addEventListener('click', () => {
        const isPass = dom.sitePinInput.type === 'password';
        dom.sitePinInput.type = isPass ? 'text' : 'password';
      });
    }
    if (dom.btnSkipPin) {
      dom.btnSkipPin.addEventListener('click', () => {
        closePinPromptModal();
        showToast('Работа в офлайн режиме (локальное хранилище)', 'info');
      });
    }
    if (dom.btnClearPin) {
      dom.btnClearPin.addEventListener('click', () => {
        if (confirm('Сбросить сохраненный PIN-код и выйти из облака?')) {
          window.PriceStorage.clearCloudAuth();
          closeModal('modalCloudSettings');
          showToast('PIN сброшен. Облако отключено', 'info');
          openPinPromptModal();
        }
      });
    }

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

    // Меню управления поставщиками
    if (dom.btnSupplierMenuToggle) {
      dom.btnSupplierMenuToggle.addEventListener('click', (e) => {
        e.stopPropagation();
        toggleSupplierDropdown();
      });
    }
    if (dom.btnMenuAddSupplier) {
      dom.btnMenuAddSupplier.addEventListener('click', () => {
        closeSupplierDropdown();
        openAddSupplierModal();
      });
    }
    if (dom.btnMenuDeleteSupplier) {
      dom.btnMenuDeleteSupplier.addEventListener('click', () => {
        closeSupplierDropdown();
        const activeSup = state.suppliers.find(s => s.id === state.activeSupplierId);
        if (activeSup && state.suppliers.length > 1) {
          deleteSupplier(activeSup);
        }
      });
    }
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

    // Пороги: сворачивание/разворачивание
    if (dom.thresholdsHeader) {
      dom.thresholdsHeader.addEventListener('click', () => {
        if (dom.thresholdsCard) dom.thresholdsCard.classList.toggle('expanded');
      });
    }

    const handleThresholdInput = (e) => {
      onThresholdChange(e);
      updateThresholdsSummary();
    };
    dom.threshGrowthPct.addEventListener('input', handleThresholdInput);
    dom.threshGrowthQty.addEventListener('input', handleThresholdInput);
    dom.threshLowQty.addEventListener('input', handleThresholdInput);
    updateThresholdsSummary();

    // Вкладки-фильтры
    dom.filterTabs.forEach(tab => {
      tab.addEventListener('click', () => {
        setActiveFilter(tab.dataset.filter, 0);
      });
    });

    // Свайп-навигация между фильтрами на мобильных устройствах
    initTouchSwipe();

    // Подфильтры по формату тары (Разливное / Бутылочное)
    if (dom.subFilterTabs) {
      dom.subFilterTabs.forEach(tab => {
        tab.addEventListener('click', () => {
          dom.subFilterTabs.forEach(t => t.classList.remove('active'));
          tab.classList.add('active');
          state.subFilter = tab.dataset.subFilter;
          renderTable();
        });
      });
    }

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
    if (dom.btnClearSnapshot) {
      dom.btnClearSnapshot.addEventListener('click', clearSnapshot);
    }

    // Меню выгрузки
    dom.btnExportMenuToggle.addEventListener('click', toggleExportDropdown);
    dom.btnExportExcel.addEventListener('click', exportToExcel);
    dom.btnExportClipboard.addEventListener('click', exportToClipboard);
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.dropdown-container')) {
        closeExportDropdown();
        closeSupplierDropdown();
      }
    });

    // Настройки облака
    if (dom.btnTestCloudConnection) dom.btnTestCloudConnection.addEventListener('click', testCloudConnection);
    if (dom.btnPushCloudNow) dom.btnPushCloudNow.addEventListener('click', pushCloudNow);
    if (dom.btnSyncCloudNow) dom.btnSyncCloudNow.addEventListener('click', syncCloudNow);
    if (dom.btnSaveCloudSettings) dom.btnSaveCloudSettings.addEventListener('click', saveCloudSettings);

    // Автоматическая периодическая синхронизация (каждые 30 сек при активной вкладке)
    setInterval(() => {
      if (document.visibilityState === 'visible' && window.PriceStorage.hasAuthenticatedPin()) {
        backgroundCloudSync();
      }
    }, 30000);

    // Синхронизация при возвращении на вкладку (переключение между браузером и другими окнами/приложениями)
    window.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible' && window.PriceStorage.hasAuthenticatedPin()) {
        backgroundCloudSync();
      }
    });

    window.addEventListener('focus', () => {
      if (window.PriceStorage.hasAuthenticatedPin()) {
        backgroundCloudSync();
      }
    });

    // Настройка маппинга
    dom.sheetSelector.addEventListener('change', onSheetChange);
    dom.headerRowSelector.addEventListener('change', onHeaderRowChange);
    dom.btnApplyMapping.addEventListener('click', applyCustomMapping);
  }

  // Запуск при готовности DOM
  document.addEventListener('DOMContentLoaded', initApp);

})();
