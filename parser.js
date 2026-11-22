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
      price: [
        'цена', 'прайс', 'стоимость', 'опт', 'розница', 'закупка',
        'цена руб', 'цена без ндс', 'цена с ндс', 'цена опт', 'колонка цены',
        'price', 'cost', 'rate', 'price_rub'
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
   */
  async readWorkbook(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: true,
            cellNF: false,
            cellText: false,
            raw: true
          });

          const sheetNames = workbook.SheetNames || [];
          if (sheetNames.length === 0) {
            throw new Error('Книга Excel не содержит листов');
          }

          resolve({
            workbook,
            sheetNames,
            defaultSheetName: this._findFirstNonEmptySheet(workbook, sheetNames)
          });
        } catch (err) {
          reject(new Error('Не удалось прочитать Excel файл: ' + err.message));
        }
      };

      reader.onerror = () => reject(new Error('Ошибка чтения файла с диска'));
      reader.readAsArrayBuffer(file);
    });
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
      priceCol: -1,
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
    mapping.priceCol = findBestCol('price', [mapping.nameCol, mapping.stockCol, mapping.codeCol]);
    mapping.unitCol = findBestCol('unit', [mapping.nameCol, mapping.stockCol, mapping.codeCol, mapping.priceCol]);

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
   * Очистка и парсинг цены
   */
  parsePriceValue(rawVal) {
    if (rawVal === null || rawVal === undefined || rawVal === '') {
      return null;
    }

    if (typeof rawVal === 'number') {
      return isNaN(rawVal) ? null : Math.round(rawVal * 100) / 100;
    }

    let str = String(rawVal).trim().toLowerCase();
    if (!str || str === '-' || str === '—') return null;

    str = str.replace(/[₽$€]|руб|руб\.|usd|eur/gi, '').trim();
    str = str.replace(/(\d)\s+(\d)/g, '$1$2');
    str = str.replace(',', '.');

    const match = str.match(/[-+]?\d+(?:\.\d+)?/);
    if (match) {
      const val = parseFloat(match[0]);
      return isNaN(val) ? null : Math.round(val * 100) / 100;
    }

    return null;
  }

  /**
   * Проверка, является ли строка категорией / подзаголовком / группировкой 1С
   */
  isCategoryOrSummaryRow(nameStr, rawStock, rawPrice, rawUnit) {
    const lower = nameStr.toLowerCase().trim();

    // 1. Итоговые строки 1С
    if (lower === 'итого' || lower.startsWith('итого') || lower === 'всего') {
      return true;
    }

    // 2. Строки категорий, где остаток и цена полностью пустые
    const isStockEmpty = (rawStock === null || rawStock === undefined || String(rawStock).trim() === '');
    const isPriceEmpty = (rawPrice === null || rawPrice === undefined || String(rawPrice).trim() === '');

    if (isStockEmpty && isPriceEmpty) {
      return true;
    }

    // 3. Специфика отчетов 1С с промежуточными суммами по группам:
    // Если есть колонка Ед. изм. (rawUnit !== null), но в этой строке Ед. изм. ПУСТАЯ,
    // а у настоящих товаров всегда указано "шт", "л", "кег", и в названии нет объема:
    if (rawUnit !== null && rawUnit !== undefined) {
      const unitStr = String(rawUnit).trim().toLowerCase();
      if (!unitStr) {
        // Проверяем, похоже ли название на категорию (нет литража/штук/тары)
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
    const { nameCol, stockCol, codeCol, priceCol, unitCol } = mapping;

    if (nameCol < 0 || stockCol < 0) {
      throw new Error('Колонки «Наименование» и «Остаток» обязательны для сопоставления');
    }

    const items = [];
    let skippedCategoriesCount = 0;

    for (let r = headerRowIndex + 1; r < sheetData.length; r++) {
      const row = sheetData[r];
      if (!Array.isArray(row) || row.length === 0) continue;

      const rawName = row[nameCol];
      const rawStock = row[stockCol];
      const rawCode = codeCol >= 0 ? row[codeCol] : null;
      const rawPrice = priceCol >= 0 ? row[priceCol] : null;
      const rawUnit = unitCol >= 0 ? row[unitCol] : null;

      const nameStr = String(rawName || '').trim();
      if (!nameStr) continue;

      // Пропускаем категории, группировки и строки Итого
      if (this.isCategoryOrSummaryRow(nameStr, rawStock, rawPrice, rawUnit)) {
        skippedCategoriesCount++;
        continue;
      }

      const parsedStock = this.parseStockValue(rawStock);
      const parsedPrice = this.parsePriceValue(rawPrice);
      const codeStr = rawCode !== null && rawCode !== undefined ? String(rawCode).trim() : '';

      const normalizedName = this.normalizeItemName(nameStr);

      items.push({
        id: `item_${r}_${Date.now()}`,
        code: codeStr,
        name: nameStr,
        normalizedName: normalizedName,
        stock: parsedStock.stock,
        isTextStock: parsedStock.isTextStock,
        textStockNote: parsedStock.textStockNote || '',
        price: parsedPrice,
        rawStock: parsedStock.raw,
        rawPrice: rawPrice !== null && rawPrice !== undefined ? String(rawPrice) : ''
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
