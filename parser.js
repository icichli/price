/**
 * parser.js - Умный парсер Excel файлов (.xlsx, .xls) через SheetJS
 * С поддержкой отчетов 1С, многострочных шапок, автопоиском колонок и фильтрацией категорий
 */

class ExcelParser {
  constructor() {
    // Словари ключевых слов для распознавания колонок (в нижнем регистре)
    this.KEYWORD_DICTIONARY = {
      name: [
        'номенклатура', 'наименование', 'товар', 'название', 'позиция',
        'описание', 'продукт', 'наименование товара', 'наименование номенклатуры',
        'name', 'title', 'item', 'description', 'product'
      ],
      stock: [
        'излишек', 'излишки', 'остаток', 'кол-во', 'количество', 'доступно', 'наличие',
        'свободный', 'остатки', 'баланс', 'склад', 'свободно', 'факт',
        'в наличии', 'свободный остаток', 'доступный остаток', 'кол-во свободных',
        'конечный остаток', 'текущий остаток',
        'qty', 'stock', 'count', 'amount', 'available', 'balance'
      ],
      code: [
        'артикул', 'код', 'код товара', 'арт', 'номенклатурный номер',
        'код номенклатуры', 'код 1с', 'штрихкод', 'баркод',
        'sku', 'id', 'item id', 'code', 'art', 'barcode', 'articul'
      ],
      unit: [
        'ед. изм.', 'ед.изм.', 'ед. изм', 'ед изм', 'ед', 'базовая единица', 'unit', 'uom'
      ]
    };

    // Строки, которые являются заголовками отчетов или параметрами, а не шапкой таблицы
    this.REPORT_TITLE_MARKERS = [
      'остатки и доступность товаров',
      'параметры:',
      'отбор:',
      'прайс-лист',
      'ведомость по остаткам',
      'отчет об остатках'
    ];
  }

  /**
   * Чтение сырого файла Excel в объект книги SheetJS
   * Оптимизировано: native arrayBuffer() и dense: true (экономия до 70% памяти и кратное ускорение парсинга)
   */
  async readWorkbook(file) {
    try {
      const buffer = file.arrayBuffer
        ? await file.arrayBuffer()
        : await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = () => reject(new Error('Ошибка чтения файла с диска'));
            reader.readAsArrayBuffer(file);
          });

      const data = new Uint8Array(buffer);
      const workbook = XLSX.read(data, {
        type: 'array',
        dense: true,
        cellDates: true,
        cellNF: false,
        cellText: false,
        raw: true
      });

      const sheetNames = workbook.SheetNames || [];
      if (sheetNames.length === 0) {
        throw new Error('Книга Excel не содержит листов');
      }

      return {
        workbook,
        sheetNames,
        defaultSheetName: this._findFirstNonEmptySheet(workbook, sheetNames)
      };
    } catch (err) {
      throw new Error('Не удалось прочитать Excel файл: ' + err.message);
    }
  }

  /**
   * Поиск первой непустой вкладки в книге
   */
  _findFirstNonEmptySheet(workbook, sheetNames) {
    for (const name of sheetNames) {
      const sheet = workbook.Sheets[name];
      if (sheet && sheet['!ref']) {
        const range = XLSX.utils.decode_range(sheet['!ref']);
        if (range.e.r >= 0 && range.e.c >= 0) {
          return name;
        }
      }
    }
    return sheetNames[0];
  }

  /**
   * Преобразование листа книги в 2D-массив строк
   */
  getSheetData(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return [];

    const rawRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      blankrows: false,
      raw: true
    });

    return rawRows;
  }

  /**
   * Автоматический поиск строки заголовков в первых 25 строках файла.
   * Поддерживает как однострочные, так и двухстрочные шапки 1С (merged headers).
   */
  detectHeaderRow(sheetData, maxScanRows = 25) {
    const limit = Math.min(sheetData.length, maxScanRows);
    let bestRowIndex = 0;
    let highestScore = -1;
    let isMultiLineHeader = false;

    // Вспомогательная функция скоринга массива ячеек
    const scoreHeaderCells = (rowCells) => {
      let hasName = false;
      let hasStock = false;
      let hasCode = false;
      let hasPrice = false;
      let isReportTitle = false;
      let matchCount = 0;

      for (let c = 0; c < rowCells.length; c++) {
        const cellText = String(rowCells[c] || '').trim().toLowerCase();
        if (!cellText) continue;

        // Проверяем, не заголовок ли это отчета 1С (типа "Остатки и доступность товаров")
        for (const marker of this.REPORT_TITLE_MARKERS) {
          if (cellText === marker || cellText.startsWith(marker)) {
            isReportTitle = true;
            break;
          }
        }

        if (!hasName && this._matchesKeywords(cellText, this.KEYWORD_DICTIONARY.name)) {
          hasName = true;
          matchCount += 5;
        }
        if (!hasStock && this._matchesKeywords(cellText, this.KEYWORD_DICTIONARY.stock)) {
          hasStock = true;
          matchCount += 5;
        }
        if (!hasCode && this._matchesKeywords(cellText, this.KEYWORD_DICTIONARY.code)) {
          hasCode = true;
          matchCount += 2;
        }
        if (!hasPrice && this._matchesKeywords(cellText, this.KEYWORD_DICTIONARY.price)) {
          hasPrice = true;
          matchCount += 2;
        }
      }

      if (isReportTitle) {
        return -50; // штраф заголовкам отчетов
      }

      let score = matchCount;
      if (hasName && hasStock) score += 20;
      else if (hasName) score += 2;
      else if (hasStock) score += 2;

      return score;
    };

    // 1. Проверяем пары смежных строк (многострочные шапки 1С, например строки 5 и 6)
    for (let r = 0; r < limit - 1; r++) {
      const row1 = sheetData[r] || [];
      const row2 = sheetData[r + 1] || [];
      const maxCols = Math.max(row1.length, row2.length);

      // Объединяем ячейки двух строк
      const combined = [];
      for (let c = 0; c < maxCols; c++) {
        const val1 = String(row1[c] || '').trim();
        const val2 = String(row2[c] || '').trim();
        combined.push(val1 ? (val2 ? `${val1} ${val2}` : val1) : val2);
      }

      const score = scoreHeaderCells(combined);
      if (score > highestScore) {
        highestScore = score;
        bestRowIndex = r + 1; // Заголовок заканчивается на r+1, данные начинаются со строки r+2
        isMultiLineHeader = true;
      }
    }

    // 2. Проверяем также одиночные строки (для стандартных прайсов с 1 строкой заголовков)
    for (let r = 0; r < limit; r++) {
      const row = sheetData[r] || [];
      const score = scoreHeaderCells(row);

      // Одиночная строка выигрывает только если ее счет явно выше
      if (score > highestScore) {
        highestScore = score;
        bestRowIndex = r;
        isMultiLineHeader = false;
      }
    }

    return {
      rowIndex: bestRowIndex,
      isMultiLineHeader: isMultiLineHeader,
      confidence: highestScore >= 20 ? 'high' : (highestScore >= 5 ? 'medium' : 'low')
    };
  }

  /**
   * Проверка совпадения текста ячейки со словарем ключевых слов
   */
  _matchesKeywords(text, keywords) {
    const clean = text.replace(/[^a-zа-я0-9\s]/gi, ' ').replace(/\s+/g, ' ').trim();
    return keywords.some(k => {
      if (clean === k) return true;
      if (clean.startsWith(k + ' ') || clean.endsWith(' ' + k) || clean.includes(' ' + k + ' ')) return true;
      return false;
    });
  }

  /**
   * Автомаппинг колонок по найденной строке (или паре строк) заголовков
   */
  detectColumns(headerRow, previousRow = null) {
    const mapping = {
      nameCol: -1,
      stockCol: -1,
      codeCol: -1,
      unitCol: -1
    };

    // Если шапка была двухстрочной, объединяем предыдущую строку и текущую
    let effectiveRow = headerRow;
    if (previousRow && Array.isArray(previousRow)) {
      const maxCols = Math.max(headerRow.length, previousRow.length);
      effectiveRow = [];
      for (let c = 0; c < maxCols; c++) {
        const v1 = String(previousRow[c] || '').trim();
        const v2 = String(headerRow[c] || '').trim();
        effectiveRow.push(v1 ? (v2 ? `${v1} ${v2}` : v1) : v2);
      }
    }

    if (!Array.isArray(effectiveRow)) return mapping;

    const findBestCol = (field, excludeIndices = []) => {
      const keywords = this.KEYWORD_DICTIONARY[field];
      let bestIndex = -1;
      let bestScore = 0;

      for (let c = 0; c < effectiveRow.length; c++) {
        if (excludeIndices.includes(c)) continue;
        const text = String(effectiveRow[c] || '').trim().toLowerCase();
        if (!text) continue;

        const clean = text.replace(/[^a-zа-я0-9\s]/gi, ' ').replace(/\s+/g, ' ').trim();

        for (const kw of keywords) {
          let score = 0;
          if (clean === kw) {
            score = 100 - kw.length;
          } else if (clean.startsWith(kw) || clean.includes(kw)) {
            score = 50;
          }

          if (score > bestScore) {
            bestScore = score;
            bestIndex = c;
          }
        }
      }

      return bestIndex;
    };

    mapping.nameCol = findBestCol('name');
    mapping.stockCol = findBestCol('stock', [mapping.nameCol]);
    mapping.codeCol = findBestCol('code', [mapping.nameCol, mapping.stockCol]);
    mapping.unitCol = findBestCol('unit', [mapping.nameCol, mapping.stockCol, mapping.codeCol]);

    return mapping;
  }

  /**
   * Очистка и парсинг числового значения остатка:
   * - "15 шт", "4 кег", "20,5", ">10"
   * - "есть", "в наличии", "+" -> 999 (с пометкой isTextStock)
   * - "нет", "-", "отсутствует", пусто -> 0
   */
  parseStockValue(rawVal) {
    if (rawVal === null || rawVal === undefined) {
      return { stock: 0, isTextStock: false, raw: '' };
    }

    if (typeof rawVal === 'number') {
      return {
        stock: isNaN(rawVal) ? 0 : rawVal,
        isTextStock: false,
        raw: String(rawVal)
      };
    }

    const str = String(rawVal).trim().toLowerCase();
    if (!str) {
      return { stock: 0, isTextStock: false, raw: '' };
    }

    // Текстовые маркеры "в наличии"
    const inStockMarkers = ['есть', 'в наличии', 'да', 'вналичии', '+', 'много', 'доступен'];
    if (inStockMarkers.includes(str) || str.startsWith('в наличии') || str === '++' || str === '+++') {
      return { stock: 999, isTextStock: true, textStockNote: 'В наличии', raw: String(rawVal) };
    }

    // Текстовые маркеры "нет в наличии"
    const outOfStockMarkers = ['нет', '-', '—', 'отсутствует', 'под заказ', 'заказ', '0 шт', 'нет на складе'];
    if (outOfStockMarkers.includes(str) || str === '0') {
      return { stock: 0, isTextStock: false, raw: String(rawVal) };
    }

    let cleanStr = str.replace(/[><≥≤]/g, '').trim();
    cleanStr = cleanStr.replace(',', '.');
    cleanStr = cleanStr.replace(/(\d)\s+(\d)/g, '$1$2');

    const match = cleanStr.match(/[-+]?\d+(?:\.\d+)?/);
    if (match) {
      const num = parseFloat(match[0]);
      return {
        stock: isNaN(num) ? 0 : num,
        isTextStock: false,
        raw: String(rawVal)
      };
    }

    return { stock: 0, isTextStock: false, raw: String(rawVal) };
  }

  /**
   * Определение типа тары (Разливное / Бутылочное / Банка / Прочее)
   * Гибридный алгоритм: анализ наименования, единиц измерения и родительской категории 1С
   */
  detectPackageType(nameStr, rawUnit = '', categoryName = '') {
    const name = String(nameStr || '').toLowerCase();
    const unit = String(rawUnit || '').toLowerCase();
    const cat = String(categoryName || '').toLowerCase();

    const isExplicitCan = /(?:ж\/?б|ж\.б|жест[а-я]*|бан(?:к[а-я]*)?|\bcan\b)/i.test(name);
    const isExplicitBottle = /(?:бут(?:ыл[а-я]*)?|ст\/бут|ст\.бут|ст[.]|\bст\b|стекл[а-я]*|\bbottle\b)/i.test(name);
    const hasSmallVol = /(?:(?:\b|[\s_/-])(?:0[.,]\d+|0[.,]25|0[.,]33|0[.,]375|0[.,]44|0[.,]45|0[.,]5|0[.,]75|1[.,]0|1[.,]5)\s*(?:л|литр[а-я]*|l)?\b)|(?:\b(?:250|330|355|440|450|500|650|750)\s*(?:мл|ml)\b)/i.test(name);
    const hasPackNotation = /(?:(?:\b|[\s_/-])1\s*[/xх*]\s*(?:6|8|12|16|20|24)\b)|(?:(?:\b|[\s_/-])(?:12|20|24)\s*[/xх*]\s*0[.,]\d+)/i.test(name);

    const isExplicitKeg = /(?:кег[а-я]*|кэг[а-я]*|\bkegs?\b|пэт-?кег[а-я]*|фитинг[а-я]*|бочк[а-я]*)/i.test(name) || unit.includes('кег');
    const isDraftWord = /(?:разливн[а-я]*|на разлив)/i.test(name);
    const hasKegVol = /(?:(?:\b|[\s_/-])(?:20|25|30|50)\s*(?:л|литр[а-я]*|l)\b)|(?:\b1\s*[/xх*]\s*(?:25|30|50)\b)/i.test(name);
    const hasKegUnit = unit === 'кег' || unit === 'дал';

    // 1. Если есть явный маркер банки, бутылки или малого объема, и нет слова КЕГ
    if ((isExplicitCan || isExplicitBottle || hasSmallVol || hasPackNotation) && !isExplicitKeg) {
      if (isExplicitCan) {
        return { type: 'bottle', subType: 'can', label: 'Банка', icon: '🥫' };
      }
      return { type: 'bottle', subType: 'bottle', label: 'Бутылочное', icon: '🍾' };
    }

    // 2. Если есть явный маркер кега, разливного или кегового объема (20, 25, 30, 50л)
    if (isExplicitKeg || isDraftWord || hasKegVol || hasKegUnit) {
      return { type: 'draft', subType: 'keg', label: 'Разливное', icon: '🍺' };
    }

    // 3. Проверяем родительскую категорию из отчета 1С
    if (cat) {
      if (/(?:разлив|кег|кэг|бочк)/i.test(cat)) {
        return { type: 'draft', subType: 'keg', label: 'Разливное', icon: '🍺' };
      }
      if (/(?:бутыл|фасов|стекл|ж\/?б|банк)/i.test(cat)) {
        return { type: 'bottle', subType: 'bottle', label: 'Бутылочное', icon: '🍾' };
      }
    }

    // 4. Проверяем единицы измерения
    if (unit === 'кег' || unit === 'дал' || unit === 'л' || unit === 'литр') {
      return { type: 'draft', subType: 'keg', label: 'Разливное', icon: '🍺' };
    }
    if (unit === 'бут' || unit === 'бан' || unit === 'шт') {
      return { type: 'bottle', subType: 'bottle', label: 'Бутылочное', icon: '🍾' };
    }

    return { type: 'other', subType: 'other', label: 'Прочее', icon: '📦' };
  }

  /**
   * Проверка, является ли строка категорией / подзаголовком / группировкой 1С
   */
  isCategoryOrSummaryRow(nameStr, rawStock, rawUnit) {
    const lower = nameStr.toLowerCase().trim();

    // 1. Итоговые строки 1С
    if (lower === 'итого' || lower.startsWith('итого') || lower === 'всего') {
      return true;
    }

    // 2. Строки категорий, где остаток полностью пустой
    const isStockEmpty = (rawStock === null || rawStock === undefined || String(rawStock).trim() === '');
    if (isStockEmpty) {
      return true;
    }

    // 3. Специфика отчетов 1С с промежуточными суммами по группам
    if (rawUnit !== null && rawUnit !== undefined) {
      const unitStr = String(rawUnit).trim().toLowerCase();
      if (!unitStr) {
        const hasPackagingMarker = /(?:\d+(?:[\.,]\d+)?\s*(?:л|мл|кг|г|шт|ж\/б|ст|пэт|кег)|кег|пэт|ж\/б)/i.test(nameStr);
        if (!hasPackagingMarker) {
          return true; // Это строка группировки (например: "Вода питьевая", "КРАФТ разливное", "Аляска")
        }
      }
    }

    return false;
  }

  /**
   * Извлечение позиций номенклатуры из листа по настроенному маппингу
   */
  extractItems(sheetData, headerRowIndex, mapping) {
    const { nameCol, stockCol, codeCol, unitCol } = mapping;

    if (nameCol < 0 || stockCol < 0) {
      throw new Error('Колонки «Наименование» и «Остаток» обязательны для сопоставления');
    }

    const items = [];
    let skippedCategoriesCount = 0;
    let currentCategory = '';

    for (let r = headerRowIndex + 1; r < sheetData.length; r++) {
      const row = sheetData[r];
      if (!Array.isArray(row) || row.length === 0) continue;

      const rawName = row[nameCol];
      const rawStock = row[stockCol];
      const rawCode = codeCol >= 0 ? row[codeCol] : null;
      const rawUnit = unitCol >= 0 ? row[unitCol] : null;

      const nameStr = String(rawName || '').trim();
      if (!nameStr) continue;

      // Пропускаем категории, группировки и строки Итого, фиксируя текущий раздел
      if (this.isCategoryOrSummaryRow(nameStr, rawStock, rawUnit)) {
        skippedCategoriesCount++;
        const lowerCat = nameStr.toLowerCase().trim();
        if (lowerCat !== 'итого' && !lowerCat.startsWith('итого') && lowerCat !== 'всего') {
          currentCategory = nameStr;
        }
        continue;
      }

      const parsedStock = this.parseStockValue(rawStock);
      const codeStr = rawCode !== null && rawCode !== undefined ? String(rawCode).trim() : '';
      const normalizedName = this.normalizeItemName(nameStr);
      const packageInfo = this.detectPackageType(nameStr, rawUnit, currentCategory);

      items.push({
        id: `item_${r}_${Date.now()}`,
        code: codeStr,
        name: nameStr,
        normalizedName: normalizedName,
        category: currentCategory,
        packageType: packageInfo.type,       // 'draft' | 'bottle' | 'other'
        packageSubType: packageInfo.subType, // 'keg' | 'bottle' | 'can' | 'other'
        packageLabel: packageInfo.label,     // 'Разливное' | 'Бутылочное' | 'Банка'
        packageIcon: packageInfo.icon,       // '🍺' | '🍾' | '🥫' | '📦'
        stock: parsedStock.stock,
        isTextStock: parsedStock.isTextStock,
        textStockNote: parsedStock.textStockNote || '',
        rawStock: parsedStock.raw,
        rawUnit: rawUnit !== null && rawUnit !== undefined ? String(rawUnit).trim() : ''
      });
    }

    return {
      items,
      skippedCategoriesCount,
      totalRowsScanned: sheetData.length - headerRowIndex - 1
    };
  }

  /**
   * Нормализация строки наименования
   */
  normalizeItemName(str) {
    if (!str) return '';
    return str
      .toLowerCase()
      .replace(/ё/g, 'е')
      .replace(/[«»""'']/g, '"')
      .replace(/[–—]/g, '-')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /**
   * Получение первых N строк листа для таблицы предпросмотра
   */
  getPreviewRows(sheetData, maxRows = 10) {
    const count = Math.min(sheetData.length, maxRows);
    const rows = [];
    let maxCols = 0;

    for (let i = 0; i < count; i++) {
      const row = Array.isArray(sheetData[i]) ? sheetData[i] : [];
      rows.push(row);
      if (row.length > maxCols) maxCols = row.length;
    }

    return { rows, maxCols, totalRowsInSheet: sheetData.length };
  }
}

// Экспорт глобального синглтона
window.ExcelParser = new ExcelParser();
