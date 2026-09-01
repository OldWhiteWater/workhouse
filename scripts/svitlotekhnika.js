const staticTiles = [];

const CATEGORY_FIELD   = 'Вид изделия';
const CARD_ID_FIELD    = 'Картка';
const CARD_CODE_FIELD  = 'Код за каталогом';
const CARD_NAME_FIELD  = 'Найменування товару';
const FULL_NAME_FIELD  = 'Товар';
const ERRORS_FIELD     = 'Ошибки';
const PROPERTIES_FIELD = 'Свойства';

const GROUP_KEY_FIELD_PRIMARY  = 'Номер группы ch3';
const GROUP_KEY_FIELD_FALLBACK = 'Номер группы ch2';

const FILTER_FIELDS = [
  'Маркування B','Маркування C','Маркування D','Маркування H','Маркування M',
  'Маркування P','Маркування R','Маркування S','Маркування T','Маркування W',
  'Маркування','Джерело світла','Кількість світлодіодів','Тип світлодіодів',
  'Цоколь B','Цоколь P','Цоколь PG','Цоколь PK','Цоколь PX','Цоколь S',
  'Цоколь W','Цоколь X','Цоколь FesToon','Довжина','Напруга, В','Потужність, Вт',
  'Колірна температура, К','Застосування авто','Колір світіння','Кількість контактів',
  'Виконання','Особливості'
];

const VIRTUAL_FILTER_COLUMN_PATTERNS = {};

const OVERVIEW_FIELDS = [
  'Джерело світла','Тип світлодіодів','Напруга, В',
  'Потужність, Вт','Колірна температура, К','Колір світіння'
];

const HIERARCHY = [
  { key: 'parent', field: 'Родитель',  imageOwn: 'картинка_ родитель', label: 'Категорія' },
  { key: 'group1', field: 'группа1',   imageOwn: 'картинка_ группа1',  label: 'Підкатегорія' },
];

const GROUP2_FIELD       = 'группа2';
const GROUP3_FIELD       = 'группа3';
const GROUP3_IMAGE_SMALL = 'картинка_ группа3_small';
const GROUP3_IMAGE_BIG   = 'картинка_ группа3_big';
const GROUP3_DESCRIPTION = 'описание_ группа3';

const HIDDEN_EXTRA_FIELDS_BASE = new Set([
  CARD_ID_FIELD, CARD_CODE_FIELD, CARD_NAME_FIELD, FULL_NAME_FIELD,
  ERRORS_FIELD, PROPERTIES_FIELD, CATEGORY_FIELD,
  GROUP_KEY_FIELD_PRIMARY, GROUP_KEY_FIELD_FALLBACK,
  ...OVERVIEW_FIELDS,
  ...HIERARCHY.map(h => h.field),
  ...HIERARCHY.map(h => h.imageOwn).filter(Boolean),
  GROUP2_FIELD, GROUP3_FIELD, GROUP3_IMAGE_SMALL, GROUP3_IMAGE_BIG, GROUP3_DESCRIPTION
]);

const HIDDEN_EXTRA_FIELDS = new Set(
  [...HIDDEN_EXTRA_FIELDS_BASE].map(normalizeHeaderKey)
);

let allHeaders = [];
let headerIndex = {};
let allRows = [];
let deepHierarchyAvailable = false;

let tiles = [...staticTiles];
let catalogDetails = {};

let navStack = [];
let selection = {};

// === ПРАВИЛЬНЫЕ DOM-ЭЛЕМЕНТЫ ===
const tileSection = document.getElementById('tileSection');
const resultsContainer = document.getElementById('resultsContainer');
const sidePanel = document.getElementById('sidePanel');
const sideTitle = document.getElementById('sideTitle');
const sideList = document.getElementById('sideList');
const backButton = document.getElementById('backButton');
const catalogToolbar = document.getElementById('catalogToolbar');


// === РЕНДЕР ТОВАРОВ ===
function renderProducts(rows) {
    tileSection.style.display = 'none';
    resultsContainer.style.display = 'flex';
    resultsContainer.innerHTML = '';

    catalogToolbar.style.display = 'flex';
    sidePanel.style.display = 'block';

    const grid = document.createElement('div');
    grid.className = 'results-cards-grid';

    rows.forEach(row => {
        const card = renderProductCard(row);
        grid.appendChild(card);
    });

    resultsContainer.appendChild(grid);
}


// === ОТКРЫТИЕ ГРУППЫ 3 ===
function showGroup3Detail(groupKey, groupName) {
    const rows = allRows.filter(r => {
        const key = cell(r, GROUP_KEY_FIELD_PRIMARY) || cell(r, GROUP_KEY_FIELD_FALLBACK);
        return key == groupKey;
    });

    sideTitle.textContent = groupName;

    renderProducts(rows);
}


// === КНОПКА НАЗАД ===
backButton.addEventListener('click', () => {
    resultsContainer.style.display = 'none';
    tileSection.style.display = 'grid';
    catalogToolbar.style.display = 'none';
    sidePanel.style.display = 'none';
});




function normalizeCell(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function normalizeHeaderKey(name) {
  return String(name)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/_\s+/g, '_')
    .replace(/\s+_/g, '_');
}

function resolveIdx(canonicalName) {
  return headerIndex[normalizeHeaderKey(canonicalName)];
}

function cell(row, canonicalName) {
  const i = resolveIdx(canonicalName);
  if (i === undefined) return '';
  return normalizeCell(row[i]);
}

function hasField(canonicalName) {
  return resolveIdx(canonicalName) !== undefined;
}

function resolveGroupKeyField() {
  if (hasField(GROUP_KEY_FIELD_PRIMARY)) return GROUP_KEY_FIELD_PRIMARY;
  if (hasField(GROUP_KEY_FIELD_FALLBACK)) return GROUP_KEY_FIELD_FALLBACK;
  return null;
}

function parseNumericLike(v) {
  const s = String(v).trim().replace(',', '.');
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  const fracMatch = s.match(/^(\d+)\s*\/\s*(\d+)$/);
  if (fracMatch) return Number(fracMatch[1]) / Number(fracMatch[2]);
  return NaN;
}

function sortValues(values) {
  return [...values].sort((a, b) => {
    const first = parseNumericLike(a);
    const second = parseNumericLike(b);
    if (!Number.isNaN(first) && !Number.isNaN(second)) return first - second;
    return String(a).localeCompare(String(b), 'uk');
  });
}


function getFieldValuesForRow(fieldName, row) {
  const value = cell(row, fieldName);

  if (!value) return [];

  if (typeof value === 'string' && value.includes(',')) {
    return value.split(',').map(v => v.trim()).filter(v => v.length > 0);
  }

  return [value];
}



function getFallbackIcon() {
  return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" style="opacity:0.3;">
    <circle cx="50" cy="50" r="40" fill="none" stroke="#64748b" stroke-width="3" stroke-dasharray="4 4"/>
    <path d="M40 30 A10 10 0 0 1 60 30 V50 H40 Z" fill="#0f172a"/>
    <rect x="44" y="50" width="12" height="16" rx="3" fill="#64748b"/>
  </svg>`;
}

function getCategoryIconFallback(categoryName) {
  const label = categoryName.trim().toLowerCase();
  if (label.includes('led') || label.includes('світлодіод')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <circle cx="50" cy="40" r="16" fill="#22c55e"/>
      <path d="M34 40 A16 16 0 0 1 66 40" fill="none" stroke="#16a34a" stroke-width="3"/>
      <rect x="44" y="56" width="12" height="18" rx="3" fill="#0f172a"/>
    </svg>`;
  }
  if (label.includes('цоколь') || label.includes('socket')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <rect x="32" y="26" width="36" height="30" rx="8" fill="#0f172a"/>
      <rect x="38" y="56" width="24" height="18" rx="4" fill="#64748b"/>
      <rect x="30" y="74" width="40" height="8" rx="4" fill="#cbd5e1"/>
    </svg>`;
  }
  if (label.includes('фар') || label.includes('headlight')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <path d="M30 30 Q50 20 70 30 V70 Q50 80 30 70 Z" fill="#0f172a"/>
      <path d="M70 40 L88 36" stroke="#facc15" stroke-width="3" stroke-linecap="round"/>
      <path d="M70 50 L88 50" stroke="#facc15" stroke-width="3" stroke-linecap="round"/>
      <path d="M70 60 L88 64" stroke="#facc15" stroke-width="3" stroke-linecap="round"/>
    </svg>`;
  }
  return getFallbackIcon();
}


const TREE_XLSX_PATH = './source/svitlotekhnika_tree.xlsx';
const PRODUCT_XLSX_PATHS = ['./source/auto_bulb.xlsx'];

const HIERARCHY_TREE_FIELDS = [
  'Родитель','картинка_ родитель',
  'группа1','картинка_ группа1',
  'группа2',
  'группа3','описание_ группа3','картинка_ группа3_small','картинка_ группа3_big'
];

async function fetchAndParseXlsx(path) {
  if (!window.XLSX) throw new Error('SheetJS XLSX library is not loaded.');

  const fileUrl = new URL(path, window.location.href);
  fileUrl.searchParams.set('v', Date.now());
  const response = await fetch(fileUrl.toString(), { cache: 'no-store' });
  if (!response.ok) throw new Error(`Cannot load ${path}: ${response.status} ${response.statusText}`);

  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const knownSheetNames = ['Результат','Аркуш1'];
  const sheetName = knownSheetNames.find(n => workbook.SheetNames.includes(n)) || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const sheetRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
  return sheetRows;
}

async function loadTreeMap() {
  const sheetRows = await fetchAndParseXlsx(TREE_XLSX_PATH);
  const headerRow = sheetRows[0] || [];
  const idx = {};
  headerRow.forEach((h, i) => {
    const k = normalizeHeaderKey(h);
    if (!(k in idx)) idx[k] = i;
  });
  const tcell = (row, name) => {
    const i = idx[normalizeHeaderKey(name)];
    return i === undefined ? '' : normalizeCell(row[i]);
  };

  const map = new Map();
  sheetRows.slice(1).forEach(row => {
    const ch3 = tcell(row, GROUP_KEY_FIELD_PRIMARY);
    if (!ch3) return;
    const entry = {};
    HIERARCHY_TREE_FIELDS.forEach(f => { entry[f] = tcell(row, f); });
    map.set(ch3, entry);
  });
  return map;
}

async function loadMergedProductRows() {
  const combinedHeader = [];
  const combinedIndex = {};
  const ensureColumn = (fieldName) => {
    const key = normalizeHeaderKey(fieldName);
    if (key in combinedIndex) return combinedIndex[key];
    const i = combinedHeader.length;
    combinedHeader.push(fieldName);
    combinedIndex[key] = i;
    return i;
  };

  const combinedRows = [];
  let filesLoaded = 0;

  for (const path of PRODUCT_XLSX_PATHS) {
    let sheetRows;
    try {
      sheetRows = await fetchAndParseXlsx(path);
    } catch (err) {
      continue;
    }
    filesLoaded++;
    const fileHeader = sheetRows[0] || [];
    const fileIdxToCombinedIdx = fileHeader.map(h => ensureColumn(h));
    sheetRows.slice(1).forEach(row => {
      const newRow = new Array(combinedHeader.length).fill('');
      row.forEach((val, i) => {
        const ci = fileIdxToCombinedIdx[i];
        if (ci !== undefined) newRow[ci] = val;
      });
      combinedRows.push(newRow);
    });
  }

  if (filesLoaded === 0) {
    throw new Error('Жоден товарний файл не вдалося завантажити.');
  }
  return { header: combinedHeader, rows: combinedRows };
}

function hydrateRowsWithTree(header, rows, treeMap) {
  const idx = {};
  header.forEach((h, i) => {
    const k = normalizeHeaderKey(h);
    if (!(k in idx)) idx[k] = i;
  });
  const ensureColumn = (fieldName) => {
    const key = normalizeHeaderKey(fieldName);
    if (key in idx) return idx[key];
    const i = header.length;
    header.push(fieldName);
    idx[key] = i;
    return i;
  };

  const fieldIdx = {};
  HIERARCHY_TREE_FIELDS.forEach(f => { fieldIdx[f] = ensureColumn(f); });
  const ch3Idx = ensureColumn(GROUP_KEY_FIELD_PRIMARY);

  rows.forEach(row => {
    while (row.length < header.length) row.push('');
    const ch3 = normalizeCell(row[ch3Idx]);
    if (!ch3) return;
    const entry = treeMap.get(ch3);
    if (!entry) return;
    HIERARCHY_TREE_FIELDS.forEach(f => { row[fieldIdx[f]] = entry[f]; });
  });
}

function indexHeaders(headerRow) {
  allHeaders = headerRow.map((h, i) => normalizeCell(h) || String.fromCharCode(65 + i));
  headerIndex = {};
  allHeaders.forEach((name, i) => {
    const key = normalizeHeaderKey(name);
    if (!(key in headerIndex)) headerIndex[key] = i;
  });
}


function rowsMatchingSelection(uptoLevelIndex) {
  return allRows.filter(row =>
    HIERARCHY.slice(0, uptoLevelIndex).every(h => cell(row, h.field) === selection[h.key])
  );
}

function distinctValuesWithSample(rows, field) {
  const map = new Map();
  rows.forEach(row => {
    const value = cell(row, field);
    if (!value) return;
    if (!map.has(value)) map.set(value, { count: 0, sampleRow: row });
    map.get(value).count += 1;
  });
  return map;
}

function tileImageHtml(imageSrc, altText, fallbackHtml) {
  if (!imageSrc) return fallbackHtml;
  // ВАЖЛИВО: onerror обгорнуто в ОДИНАРНІ лапки, бо JSON.stringify()
  // повертає рядок у ПОДВІЙНИХ лапках — якщо атрибут теж у подвійних,
  // HTML-парсер обриває атрибут на першій внутрішній лапці і "хвіст"
  // SVG/тексту витікає як видимий вміст плитки (саме цей баг був помічений).
  return `<img src="images/${imageSrc}" alt="${altText}" onerror='console.warn("Не вдалося завантажити зображення:", this.src); this.outerHTML = ${JSON.stringify(fallbackHtml)};'>`;
}

function renderBreadcrumbs(target) {
  const bar = document.createElement('div');
  bar.style.cssText = 'margin-bottom:16px;color:#64748b;font-size:13px;display:flex;flex-wrap:wrap;gap:6px;align-items:center;';

  const rootLink = document.createElement('span');
  rootLink.textContent = 'Каталог';
  rootLink.style.cssText = 'cursor:pointer;font-weight:600;color:#334155;';
  rootLink.onclick = () => goToLevel(0);
  bar.appendChild(rootLink);

  HIERARCHY.forEach((h, i) => {
    if (selection[h.key] === undefined) return;
    const sep = document.createElement('span');
    sep.textContent = '›';
    bar.appendChild(sep);

    const crumb = document.createElement('span');
    crumb.textContent = selection[h.key];
    crumb.style.cssText = 'cursor:pointer;color:#334155;font-weight:600;';
    if (i === HIERARCHY.length - 1) {
      // Останній рівень (группа1) -> повернутися на екран секцій группа2+группа3
      crumb.onclick = () => renderGroup2SectionsScreen();
    } else {
      crumb.onclick = () => goToLevel(i + 1);
    }
    bar.appendChild(crumb);
  });

  target.appendChild(bar);
}

function goToLevel(levelIndex) {
  HIERARCHY.forEach((h, i) => { if (i >= levelIndex) delete selection[h.key]; });
  navStack = [];
  for (let i = 0; i < levelIndex; i++) navStack.push({ type: 'levelTiles', levelIndex: i });
  renderLevelTiles(levelIndex);
}

function renderLevelTiles(levelIndex) {
  sidePanel.style.display = 'none';

  resultsContainer.style.display = 'none';
  resultsContainer.innerHTML = '';

  catalogToolbar.style.display = levelIndex > 0 ? 'flex' : 'none';

  tileSection.innerHTML = '';
  tileSection.style.display = 'block';

  const wrap = document.createElement('div');
  renderBreadcrumbs(wrap);

  if (levelIndex === 0) {
    const skipBtn = document.createElement('button');
    skipBtn.type = 'button';
    skipBtn.textContent = `Пропустити вибір категорій — показати всі товари (${allRows.length})`;
    skipBtn.style.cssText = 'display:inline-flex;align-items:center;gap:8px;padding:10px 18px;border:1px solid #e2e8f0;border-radius:10px;background:#ffffff;color:#334155;font-size:14px;font-weight:600;font-family:inherit;cursor:pointer;box-shadow:0 1px 2px rgba(15,23,42,0.06);margin-bottom:20px;';
    skipBtn.onmouseenter = () => { skipBtn.style.background = '#f8fafc'; skipBtn.style.borderColor = '#cbd5e1'; };
    skipBtn.onmouseleave = () => { skipBtn.style.background = '#ffffff'; skipBtn.style.borderColor = '#e2e8f0'; };
    skipBtn.onclick = () => {
      navStack.push({ type: 'levelTiles', levelIndex: 0 });
      renderFullCatalogListing();
    };
    wrap.appendChild(skipBtn);
  }

  tileSection.appendChild(wrap);

  const grid = document.createElement('section');
  const gridClassByLevel = ['tile-container parent-grid', 'tile-container grid-6col'];
  grid.className = gridClassByLevel[levelIndex] || 'tile-container';
  tileSection.appendChild(grid);

  const levelDef = HIERARCHY[levelIndex];
  const rows = rowsMatchingSelection(levelIndex);
  const valueMap = distinctValuesWithSample(rows, levelDef.field);

  if (valueMap.size === 0) {
    showCatalogMessage(grid, `Немає значень у стовпці "${levelDef.field}" для цього вибору.`, 'error');
    return;
  }

  [...valueMap.entries()].sort((a, b) => a[0].localeCompare(b[0], 'uk')).forEach(([value, info]) => {
    const div = document.createElement('div');
    const tileClassByLevel = ['tile tile-parent', 'tile tile-medium'];
    div.className = tileClassByLevel[levelIndex] || 'tile';

    let imageSrc = '';
    if (levelDef.imageOwn) {
      imageSrc = cell(info.sampleRow, levelDef.imageOwn);
    } else {
      const scopeRows = allRows.filter(r =>
        HIERARCHY.slice(0, levelIndex + 1).every(h => cell(r, h.field) === (h.key === levelDef.key ? value : selection[h.key]))
      );
      const withImage = scopeRows.find(r => cell(r, GROUP3_IMAGE_SMALL));
      if (withImage) imageSrc = cell(withImage, GROUP3_IMAGE_SMALL);
    }

    const fallback = getCategoryIconFallback(value);
    div.innerHTML = `
      ${tileImageHtml(imageSrc, value, fallback)}
      <span>${value}</span>
      <span style="color:#94a3b8; font-size:12px; margin-top:4px;">${info.count} товарів</span>
    `;
    div.onclick = () => {
      selection[levelDef.key] = value;
      navStack.push({ type: 'levelTiles', levelIndex });
      if (levelIndex + 1 < HIERARCHY.length) {
        renderLevelTiles(levelIndex + 1);
      } else {
        renderGroup2SectionsScreen();
      }
    };
    grid.appendChild(div);
  });
}

function renderGroup2SectionsScreen() {
  sidePanel.style.display = 'none';

  resultsContainer.style.display = 'none';
  resultsContainer.innerHTML = '';

  catalogToolbar.style.display = 'flex';

  tileSection.style.display = 'block';
  tileSection.innerHTML = '';

  const wrap = document.createElement('div');
  renderBreadcrumbs(wrap);
  tileSection.appendChild(wrap);

  const scopeRows = rowsMatchingSelection(HIERARCHY.length);
  const g2map = distinctValuesWithSample(scopeRows, GROUP2_FIELD);

  if (g2map.size === 0) {
    showCatalogMessage(tileSection, `Немає значень у стовпці "${GROUP2_FIELD}" для цієї підкатегорії.`, 'error');
    return;
  }

  [...g2map.entries()].sort((a, b) => a[0].localeCompare(b[0], 'uk')).forEach(([group2Value, g2info], sectionIndex) => {
    const heading = document.createElement('h2');
    heading.textContent = group2Value;
    heading.style.cssText = `margin:${sectionIndex === 0 ? 0 : 32}px 0 16px 0;font-size:20px;color:#0f172a;`;
    tileSection.appendChild(heading);

    const grid = document.createElement('section');
    grid.className = 'tile-container grid-6col';
    tileSection.appendChild(grid);

    const group2Rows = scopeRows.filter(r => cell(r, GROUP2_FIELD) === group2Value);
    const g3map = distinctValuesWithSample(group2Rows, GROUP3_FIELD);

    if (g3map.size === 0) {
      showCatalogMessage(grid, `Немає значень у стовпці "${GROUP3_FIELD}" для цієї групи.`, 'error');
      return;
    }

    [...g3map.entries()].sort((a, b) => a[0].localeCompare(b[0], 'uk')).forEach(([group3Value, info]) => {
      const exactRows = group2Rows.filter(r => cell(r, GROUP3_FIELD) === group3Value);
      const sampleRow = exactRows[0];
      const imageSrc = cell(sampleRow, GROUP3_IMAGE_SMALL);
      const fallback = getCategoryIconFallback(group3Value);

      const div = document.createElement('div');
      div.className = 'tile tile-thumb';
      div.innerHTML = `
        ${tileImageHtml(imageSrc, group3Value, fallback)}
        <span class="tile-hover-label">${group3Value}<br><small>${info.count} товарів</small></span>
      `;
      div.onclick = () => {
        navStack.push({ type: 'group2sections' });
        renderGroup3Detail(group3Value, sampleRow, exactRows);
      };
      grid.appendChild(div);
    });
  });
}

function renderGroup3Detail(group3Value, sampleRow, scopeRows) {
  sidePanel.style.display = 'none';

  resultsContainer.style.display = 'none';
  resultsContainer.innerHTML = '';

  catalogToolbar.style.display = 'flex';

  tileSection.style.display = 'block';
  tileSection.innerHTML = '';

  renderBreadcrumbs(tileSection);

  const bigImage = sampleRow ? cell(sampleRow, GROUP3_IMAGE_BIG) : '';
  const description = sampleRow ? cell(sampleRow, GROUP3_DESCRIPTION) : '';

  const box = document.createElement('div');
  box.className = 'group3-detail-box';

  let html = '';
  if (bigImage) {
    html += `<img src="images/${bigImage}" alt="${group3Value}" class="group3-detail-image" onerror='console.warn("Не вдалося завантажити велике зображення:", this.src); this.style.display="none";'>`;
  }
  html += `<h2 style="margin:0 0 14px 0;font-size:20px;color:#0f172a;">${group3Value}</h2>`;
  if (description) {
    html += `<p style="white-space:pre-line;color:#334155;font-size:14px;line-height:1.6;margin:0 0 24px 0;">${description}</p>`;
  } else {
    html += `<p style="color:#94a3b8;font-size:13px;font-style:italic;margin:0 0 24px 0;">Опис ще не заповнено.</p>`;
  }
  box.innerHTML = html;

  const btn = document.createElement('button');
  btn.className = 'load-more-button';
  btn.textContent = 'Продовжити';
  btn.onclick = () => {
    navStack.push({ type: 'group3detail', group3Value, sampleRow, scopeRows });
    renderProductListing(group3Value, sampleRow, scopeRows);
  };
  box.appendChild(btn);

  tileSection.appendChild(box);
}

function renderFullCatalogListing() {
  tileSection.style.display = 'none';

  catalogToolbar.style.display = 'flex';

  sideTitle.textContent = `Усі товари (${allRows.length})`;
  sideList.innerHTML = '';
  sidePanel.style.display = 'block';

  resultsContainer.style.display = 'flex';
  resultsContainer.innerHTML = '';
  resultsContainer.className = 'results-panel';

  const crumbWrap = document.createElement('div');
  renderBreadcrumbs(crumbWrap);
  resultsContainer.appendChild(crumbWrap);

  renderCategoryDetails({
    rows: allRows,
    filterHeaders: FILTER_FIELDS,
    filterValues: buildFilterValueSets(allRows)
  });
}


function renderProductListing(group3Value, sampleRow, fallbackRows) {
  let rows = fallbackRows && fallbackRows.length > 0 ? fallbackRows : null;

  if (!rows) {
    const groupKeyField = resolveGroupKeyField();
    const groupKeyValue = (groupKeyField && sampleRow) ? cell(sampleRow, groupKeyField) : '';
    rows = (groupKeyField && groupKeyValue)
      ? allRows.filter(row => cell(row, groupKeyField) === groupKeyValue)
      : [];
  }

  // Плитки скрываем
  tileSection.style.display = 'none';

  // Панель инструментов показываем
  catalogToolbar.style.display = 'flex';

  // Заголовок и фильтры
  sideTitle.textContent = group3Value;
  sideList.innerHTML = '';
  sidePanel.style.display = 'block';

  // Панель результатов
  resultsContainer.style.display = 'flex';
  resultsContainer.innerHTML = '';
  resultsContainer.className = 'results-panel';

  const crumbWrap = document.createElement('div');
  renderBreadcrumbs(crumbWrap);
  resultsContainer.appendChild(crumbWrap);

  // ВАЖНО: вместо старого рендера — вызываем твой новый механизм
  renderCategoryDetails({
    rows,
    filterHeaders: FILTER_FIELDS,
    filterValues: buildFilterValueSets(rows)
  });
}

function buildFilterValueSets(rows) {
  const sets = {};
  FILTER_FIELDS.forEach(f => { sets[f] = new Set(); });
  rows.forEach(row => {
    FILTER_FIELDS.forEach(f => {
      getFieldValuesForRow(f, row).forEach(v => sets[f].add(v));
    });
  });
  const result = {};
  FILTER_FIELDS.forEach(f => { result[f] = sortValues(sets[f]); });
  return result;
}

function renderDisplayValues(resultContainer, filteredData) {
  resultContainer.innerHTML = '';
  resultContainer.className = 'results-cards-grid';

  if (!filteredData.rows || filteredData.rows.length === 0) {
    const empty = document.createElement('div');
    empty.textContent = 'Немає товарів для цього набору фільтрів.';
    empty.style.color = '#666';
    empty.style.padding = '20px';
    resultContainer.appendChild(empty);
    return;
  }

  filteredData.rows.forEach(row => {
    const card = document.createElement('div');
    card.className = 'product-card';

    const cardId = cell(row, CARD_ID_FIELD);
    const cardCode = cell(row, CARD_CODE_FIELD);
    const name = cell(row, CARD_NAME_FIELD) || cell(row, FULL_NAME_FIELD);

    const header = document.createElement('div');
    header.className = 'card-header';

    const nameElem = document.createElement('div');
    nameElem.className = 'card-name';
    nameElem.innerHTML = `<span class="product-title-text">${name || 'Без назви'}</span>`;
    header.appendChild(nameElem);

    const specsBox = document.createElement('div');
    specsBox.className = 'card-specs-box';

    const idElem = document.createElement('div');
    idElem.className = 'card-line';
    idElem.innerHTML = `<span class="card-line-label">Картка:</span> <span class="card-line-value">${cardId || 'N/A'}</span>`;
    specsBox.appendChild(idElem);

    const codeElem = document.createElement('div');
    codeElem.className = 'card-line';
    codeElem.innerHTML = `<span class="card-line-label">ОЕМ:</span> <span class="card-line-value highlight">${cardCode || 'N/A'}</span>`;
    specsBox.appendChild(codeElem);

    header.appendChild(specsBox);
    card.appendChild(header);

    card.addEventListener('click', () => openModal(row));
    resultContainer.appendChild(card);
  });
}

const modalOverlay = document.getElementById('modalOverlay');
const modalContent = document.getElementById('modalContent');
const modalClose   = document.getElementById('modalClose');

function modalRow(label, value, opts = {}) {
  const cls = opts.warning ? 'modal-value warning' : `modal-value${value ? '' : ' empty'}`;
  return `<div class="modal-row">
    <span class="modal-label">${label}</span>
    <span class="${cls}">${value || '—'}</span>
  </div>`;
}

function buildPortSections(row) {
  const portNumbers = new Set();
  allHeaders.forEach(h => {
    const m = h.match(/^Порт (\d+) /);
    if (m) portNumbers.add(Number(m[1]));
  });

  let html = '';
  [...portNumbers].sort((a, b) => a - b).forEach(n => {
    const prefix = `Порт ${n} `;
    const fields = allHeaders.filter(h => h.startsWith(prefix));
    const rowsHtml = fields
      .map(h => [h.slice(prefix.length), cell(row, h)])
      .filter(([, val]) => val !== '')
      .map(([label, val]) => modalRow(label.charAt(0).toUpperCase() + label.slice(1), val));

    if (rowsHtml.length > 0) {
      html += `<div class="modal-section">
        <p class="modal-section-title">Порт ${n}</p>
        ${rowsHtml.join('')}
      </div><hr class="modal-divider">`;
    }
  });
  return html;
}

function buildExtraSection(row) {
  const extraFields = allHeaders.filter(h => {
    if (HIDDEN_EXTRA_FIELDS.has(normalizeHeaderKey(h))) return false;
    if (/^Порт \d+ /.test(h)) return false;
    return cell(row, h) !== '';
  });
  if (extraFields.length === 0) return '';

  const rowsHtml = extraFields.map(h => modalRow(h, cell(row, h))).join('');
  return `<div class="modal-section">
    <p class="modal-section-title">Інше</p>
    ${rowsHtml}
  </div><hr class="modal-divider">`;
}

function openModal(row) {
  const name = cell(row, CARD_NAME_FIELD);
  const fullName = cell(row, FULL_NAME_FIELD);
  const errors = cell(row, ERRORS_FIELD);

  let html = `<div class="modal-header-container no-image">
    <div class="modal-header-text">
      <p class="modal-product-name">${name || fullName || 'Без назви'}</p>
    </div>
  </div>`;

  if (fullName && fullName !== name) {
    html += `<p style="color:#64748b; font-size:13px; margin: -8px 0 16px 0;">${fullName}</p>`;
  }

  html += `<div class="modal-section">
    <p class="modal-section-title">Ідентифікація</p>
    ${modalRow(CARD_ID_FIELD, cell(row, CARD_ID_FIELD))}
    ${modalRow(CARD_CODE_FIELD, cell(row, CARD_CODE_FIELD))}
  </div><hr class="modal-divider">`;

  const overviewRows = OVERVIEW_FIELDS
    .map(f => [f, cell(row, f)])
    .filter(([, val]) => val !== '')
    .map(([label, val]) => modalRow(label, val));
  if (overviewRows.length > 0) {
    html += `<div class="modal-section">
      <p class="modal-section-title">Загальні характеристики</p>
      ${overviewRows.join('')}
    </div><hr class="modal-divider">`;
  }

  html += buildPortSections(row);
  html += buildExtraSection(row);

  if (errors) {
    html += `<div class="modal-section">
      <p class="modal-section-title" style="color:#b91c1c;">Помилки контролю якості</p>
      ${modalRow('Ошибки', errors, { warning: true })}
    </div>`;
  }

  modalContent.innerHTML = html;
  modalOverlay.classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeModal() {
  modalOverlay.classList.remove('open');
  document.body.style.overflow = '';
}


modalClose.addEventListener('click', closeModal);
modalOverlay.addEventListener('click', e => { if (e.target === modalOverlay) closeModal(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

function getSelectedFilters(filterPanel) {
  const filters = {};
  filterPanel.querySelectorAll('input[type="checkbox"][data-filter-group]').forEach(cb => {
    const group = cb.dataset.filterGroup;
    if (!filters[group]) filters[group] = [];
    if (cb.checked) filters[group].push(cb.value);
  });
  return filters;
}

function getFilteredDisplayValues(categoryDetail, filters) {
  const active = Object.entries(filters).filter(([, v]) => v.length > 0);
  const filteredRows = categoryDetail.rows.filter(row =>
    active.every(([fieldName, values]) => {
      const rowValues = getFieldValuesForRow(fieldName, row);
      return values.some(v => rowValues.includes(v));
    })
  );
  return { rows: filteredRows };
}

function getAvailableValuesForGroup(categoryDetail, filters, targetGroup) {
  const otherActive = Object.entries(filters).filter(([g, v]) => g !== targetGroup && v.length > 0);
  const subset = categoryDetail.rows.filter(row =>
    otherActive.every(([fieldName, values]) => {
      const rowValues = getFieldValuesForRow(fieldName, row);
      return values.some(v => rowValues.includes(v));
    })
  );
  const available = new Set();
  subset.forEach(row => {
    getFieldValuesForRow(targetGroup, row).forEach(v => available.add(v));
  });
  return available;
}

function renderActiveTags(tagsBar, filters, filterPanel, categoryDetail, onUpdate) {
  tagsBar.innerHTML = '';
  const active = Object.entries(filters).filter(([, v]) => v.length > 0);
  if (active.length === 0) return;

  active.forEach(([group, values]) => {
    values.forEach(value => {
      const tag = document.createElement('span');
      tag.className = 'active-tag';
      tag.title = `Прибрати фільтр: ${group} = ${value}`;

      const label = document.createElement('span');
      label.textContent = `${group}: ${value}`;

      const x = document.createElement('span');
      x.className = 'tag-remove';
      x.textContent = '×';

      tag.appendChild(label);
      tag.appendChild(x);

      tag.addEventListener('click', () => {
        const cb = filterPanel.querySelector(
          `input[type="checkbox"][data-filter-group="${CSS.escape(group)}"][value="${CSS.escape(value)}"]`
        );
        if (cb) { cb.checked = false; onUpdate(); }
      });
      tagsBar.appendChild(tag);
    });
  });

  const resetLink = document.createElement('button');
  resetLink.type = 'button';
  resetLink.className = 'reset-filters-button';
  resetLink.textContent = 'Скинути всі';
  resetLink.style.marginLeft = 'auto';
  resetLink.onclick = () => {
    filterPanel.querySelectorAll('input[type="checkbox"][data-filter-group]').forEach(cb => cb.checked = false);
    onUpdate();
  };
  tagsBar.appendChild(resetLink);
}

function updateFilterAvailability(filterPanel, categoryDetail, filters) {
  FILTER_FIELDS.forEach(group => {
    const available = getAvailableValuesForGroup(categoryDetail, filters, group);

    const summary = filterPanel.querySelector(`summary[data-group="${CSS.escape(group)}"]`);
    if (summary) {
      const selected = (filters[group] || []).length;
      let badge = summary.querySelector('.filter-group-count');
      if (selected > 0) {
        if (!badge) { badge = document.createElement('span'); badge.className = 'filter-group-count'; summary.appendChild(badge); }
        badge.textContent = selected;
      } else if (badge) {
        badge.remove();
      }
    }

    filterPanel.querySelectorAll(`input[type="checkbox"][data-filter-group="${CSS.escape(group)}"]`).forEach(cb => {
      const li = cb.closest('li');
      if (!li) return;
      if (!available.has(cb.value) && !cb.checked) {
        li.classList.add('filter-option-disabled');
      } else {
        li.classList.remove('filter-option-disabled');
      }
    });
  });
}

function createFilterSection(categoryDetail, onFilterChange) {
  const filterSection = document.createElement('div');
  filterSection.style.marginBottom = '20px';

  const title = document.createElement('h4');
  title.textContent = 'Фільтри';
  title.style.margin = '0 0 10px 0';
  filterSection.appendChild(title);

  FILTER_FIELDS.forEach(headerName => {
    const values = categoryDetail.filterValues[headerName];

    if (!values || values.length <= 1) return;

    const details = document.createElement('details');
    details.className = 'details-filter-panel';
    details.open = true;

    const summary = document.createElement('summary');
    summary.innerHTML = `<span class="summary-text">${headerName}</span>`;
    summary.dataset.group = headerName;
    details.appendChild(summary);

    const ul = document.createElement('ul');
    ul.style.paddingLeft = '0';
    ul.style.listStyle = 'none';

    values.forEach(value => {
      const li = document.createElement('li');
      const label = document.createElement('label');
      label.style.cursor = 'pointer';
      label.style.display = 'block';
      label.style.marginBottom = '5px';

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = value;
      checkbox.style.marginRight = '8px';
      checkbox.dataset.filterGroup = headerName;
      checkbox.dataset.filterValue = value;
      checkbox.addEventListener('click', e => e.stopPropagation());
      checkbox.addEventListener('change', onFilterChange);

      const count = categoryDetail.rows.filter(row => getFieldValuesForRow(headerName, row).includes(value)).length;

      label.appendChild(checkbox);
      label.appendChild(document.createTextNode(value));

      const countSpan = document.createElement('span');
      countSpan.textContent = ` (${count})`;
      countSpan.style.color = '#94a3b8';
      countSpan.style.fontSize = '12px';
      countSpan.style.marginLeft = '4px';
      label.appendChild(countSpan);

      li.appendChild(label);
      ul.appendChild(li);
    });

    details.appendChild(ul);
    filterSection.appendChild(details);
  });

  return filterSection;
}


function renderCategoryDetails(categoryDetail) {
  let itemsToShow = 12;
  let searchQuery = '';

  // Панель фильтров
  const filterPanel = document.createElement('div');
  filterPanel.style.marginBottom = '22px';

  // Поиск
  const searchContainer = document.createElement('div');
  searchContainer.className = 'search-container';

  const searchInput = document.createElement('input');
  searchInput.type = 'text';
  searchInput.className = 'search-input';
  searchInput.placeholder = 'Швидкий пошук за карткою, ОЕМ або найменуванням...';
  searchContainer.appendChild(searchInput);

  searchInput.addEventListener('input', () => {
    searchQuery = searchInput.value.toLowerCase().trim();
    itemsToShow = 12;
    updateResults();
  });

  // Активные теги
  const tagsBar = document.createElement('div');
  tagsBar.className = 'active-tags-bar';

  // Счётчик
  const countLabel = document.createElement('div');
  countLabel.className = 'results-count';

  // Контейнер карточек
  const displayValuesContainer = document.createElement('div');

  const updateResults = () => {
    const filters = getSelectedFilters(filterPanel);
    const filtered = getFilteredDisplayValues(categoryDetail, filters);

    // Поиск
    if (searchQuery) {
      filtered.rows = filtered.rows.filter(row => {
        const id = cell(row, CARD_ID_FIELD).toLowerCase();
        const code = cell(row, CARD_CODE_FIELD).toLowerCase();
        const name = cell(row, CARD_NAME_FIELD).toLowerCase();
        const fullName = cell(row, FULL_NAME_FIELD).toLowerCase();
        return id.includes(searchQuery) || code.includes(searchQuery)
          || name.includes(searchQuery) || fullName.includes(searchQuery);
      });
    }

    const totalFilteredCount = filtered.rows.length;
    const paginatedRows = filtered.rows.slice(0, itemsToShow);

    renderActiveTags(tagsBar, filters, filterPanel, categoryDetail, () => {
      itemsToShow = 12;
      updateResults();
    });

    updateFilterAvailability(filterPanel, categoryDetail, filters);

    renderDisplayValues(displayValuesContainer, { rows: paginatedRows });

    if (totalFilteredCount === 0) {
      countLabel.textContent = 'Немає товарів, що відповідають вибраним фільтрам.';
      countLabel.style.color = '#dc2626';
    } else {
      countLabel.textContent = `Показано: ${Math.min(itemsToShow, totalFilteredCount)} з ${totalFilteredCount} товарів`;
      countLabel.style.color = '#64748b';
    }

    // Кнопка "Показати ще"
    let loadMoreBtn = resultsContainer.querySelector('.load-more-container');
    if (loadMoreBtn) loadMoreBtn.remove();

    if (totalFilteredCount > itemsToShow) {
      const loadMoreContainer = document.createElement('div');
      loadMoreContainer.className = 'load-more-container';

      const btn = document.createElement('button');
      btn.className = 'load-more-button';
      btn.textContent = 'Показати ще';
      btn.onclick = () => {
        itemsToShow += 12;
        updateResults();
      };

      loadMoreContainer.appendChild(btn);
      resultsContainer.appendChild(loadMoreContainer);
    }
  };

  // Генерация фильтров
  const filtersEl = createFilterSection(categoryDetail, () => {
    itemsToShow = 12;
    updateResults();
  });
  filterPanel.appendChild(filtersEl);

  // Добавляем элементы в resultsContainer
  resultsContainer.appendChild(searchContainer);
  resultsContainer.appendChild(tagsBar);
  resultsContainer.appendChild(countLabel);
  resultsContainer.appendChild(displayValuesContainer);

  updateResults();

  // Панель фильтров в боковой панели
  sideList.appendChild(filterPanel);
}

function showCatalogMessage(target, text, type = '') {
  const message = document.createElement('div');
  message.className = `catalog-message ${type}`.trim();
  message.textContent = text;
  target.appendChild(message);
}

function buildFlatCatalog(rows) {
  const categoryRows = {};
  const details = {};
  const categoryOrder = [];

  rows.forEach(row => {
    const category = cell(row, CATEGORY_FIELD);
    if (!category) return;
    if (!categoryRows[category]) {
      categoryRows[category] = [];
      categoryOrder.push(category);
    }
    categoryRows[category].push(row);
  });

  categoryOrder.forEach(category => {
    const rowsForCategory = categoryRows[category];
    details[category] = {
      rows: rowsForCategory,
      filterHeaders: FILTER_FIELDS,
      filterValues: buildFilterValueSets(rowsForCategory)
    };
  });

  return {
    tiles: categoryOrder.map(label => ({ label, items: [] })),
    details
  };
}

function renderFlatTiles(target) {
  target.innerHTML = '';
  target.style.display = 'grid';

  tiles.forEach(tile => {
    const div = document.createElement('div');
    div.className = 'tile';

    const count = catalogDetails[tile.label]
      ? catalogDetails[tile.label].rows.length
      : 0;

    div.innerHTML = `
      ${getCategoryIconFallback(tile.label)}
      <span>${tile.label}</span>
      <span style="color:#94a3b8; font-size:12px; margin-top:4px;">${count} товарів</span>
    `;

    div.onclick = () => showFlatTileDetails(tile);
    target.appendChild(div);
  });
}

function showFlatTileDetails(tile) {
  tileSection.style.display = 'none';

  resultsContainer.style.display = 'flex';
  resultsContainer.innerHTML = '';
  resultsContainer.className = 'results-panel';

  sideTitle.textContent = tile.label;
  sideList.innerHTML = '';
  sidePanel.style.display = 'block';

  catalogToolbar.style.display = 'flex';

  renderCategoryDetails(catalogDetails[tile.label]);
}

async function initializeCatalog() {
  try {
    const treeMap = await loadTreeMap();
    const { header, rows } = await loadMergedProductRows();
    hydrateRowsWithTree(header, rows, treeMap);

    indexHeaders(header);
    allRows = rows;

    deepHierarchyAvailable = hasField(HIERARCHY[0].field);

    if (deepHierarchyAvailable) {
      selection = {};
      navStack = [];
      renderLevelTiles(0);
    } else {
      console.warn(
        `Стовпця "${HIERARCHY[0].field}" не знайдено — використовую плаский режим за "${CATEGORY_FIELD}".`
      );

      const flat = buildFlatCatalog(allRows);
      catalogDetails = flat.details;
      tiles = [...staticTiles, ...flat.tiles];

      renderFlatTiles(tileSection);

      if (flat.tiles.length === 0) {
        showCatalogMessage(
          tileSection,
          `У файлі не знайдено значень ні у "${HIERARCHY[0].field}", ні у "${CATEGORY_FIELD}".`,
          'error'
        );
      }
    }
  } catch (error) {
    console.error(error);

    tileSection.innerHTML = '';
    tileSection.style.display = 'grid';

    showCatalogMessage(
      tileSection,
      `Не вдалося завантажити довідник дерева (${TREE_XLSX_PATH}) або товарні файли (${PRODUCT_XLSX_PATHS.join(', ')}). Відкрийте сторінку через локальний сервер і перевірте, що файли існують у папці source.`,
      'error'
    );
  }
}


backButton.onclick = () => {
  if (!deepHierarchyAvailable) {
    // Плоский режим
    sidePanel.style.display = 'none';
    catalogToolbar.style.display = 'none';
    resultsContainer.style.display = 'none';

    tileSection.style.display = 'grid';
    renderFlatTiles(tileSection);
    return;
  }

  // Иерархический режим
  if (navStack.length === 0) {
    goToLevel(0);
    return;
  }

  const prevScreen = navStack.pop();

  if (prevScreen.type === 'group2sections') {
    renderGroup2SectionsScreen();
  } else if (prevScreen.type === 'group3detail') {
    renderGroup3Detail(prevScreen.group3Value, prevScreen.sampleRow, prevScreen.scopeRows);
  } else {
    delete selection[HIERARCHY[prevScreen.levelIndex].key];
    renderLevelTiles(prevScreen.levelIndex);
  }
};

function ensureXLSXAndInit() {
  function doInitWhenReady() {
    if (window.XLSX) return initializeCatalog();

    const existing = Array.from(document.getElementsByTagName('script'))
      .find(s => s.src && s.src.indexOf('xlsx.full.min.js') !== -1);

    if (existing) {
      if (existing.getAttribute('data-xlsx-ready') === '1') {
        return initializeCatalog();
      }

      existing.addEventListener('load', () => {
        existing.setAttribute('data-xlsx-ready', '1');
        initializeCatalog();
      });

      setTimeout(() => {
        if (window.XLSX) initializeCatalog();
      }, 50);

      return;
    }

    const script = document.createElement('script');
    script.src = 'scripts/xlsx.full.min.js';
    script.async = false;

    script.addEventListener('load', () => {
      script.setAttribute('data-xlsx-ready', '1');
      initializeCatalog();
    });

    script.addEventListener('error', () => {
      console.error('Failed to load xlsx.full.min.js');
      initializeCatalog();
    });

    document.head.appendChild(script);
  }

  doInitWhenReady();
}

ensureXLSXAndInit();

