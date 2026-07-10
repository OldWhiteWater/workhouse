// =============================================================================
// Каталог з'єднувачів — 6 екранів навігації:
//   Екран 1: Родитель (плитки, 2 колонки, 75% ширини)
//   Екран 2: группа1 (плитки, в межах обраного Родитель, 6 колонок)
//   Екран 3: группа2 (плитки, в межах группа1, 6 колонок; своєї картинки
//            немає -> прев'ю бере картинка_ группа3_small довільної дочірньої
//            группа3, просто щоб плитка не була порожньою)
//   Екран 3.5: группа3 (плитки з картинка_ группа3_small, в межах группа2,
//            6 колонок) — ПОВЕРНУВ цей рівень: одна группа2 може містити
//            КІЛЬКА різних группа3 (перевірено на реальних даних — наприклад
//            "...пластиковые" містить одразу 3: прямий/кутовий/тристоронній
//            цанговий, кожен зі своєю парою small/big картинок і описом)
//   Екран 4: інфо-сторінка обраної группа3 (назва, картинка_ группа3_big,
//            описание_ группа3, кнопка "Продовжити"), 70% ширини екрана
//   Екран 5: фінальна вибірка карток (картки/фільтри/пошук/модалка)
//
// Назви полів узято ТОЧНО як в реальному файлі (включно з пробілом після "_" —
// "картинка_ родитель" тощо), пошук заголовків стійкий до пробілів/регістру.
//
// Якщо стовпця "Родитель" у файлі немає (старий connector.xlsx без ієрархії)
// — автоматично працює попередній плаский режим за "Вид изделия".
//
// ПРИПУЩЕННЯ, які варто звірити візуально:
//   - Фінальна вибірка карток фільтрується за "Номер группы ch3" (з фолбеком
//     на "Номер группы ch2" для старих файлів); якщо ключ групи порожній —
//     використовуються всі рядки в межах обраної группа3.
// =============================================================================

const staticTiles = [];

const XLSX_PATH = './source/connector.xlsx';

// ---- Канонічні поля (за назвою колонки, стійко до пробілів навколо "_") ----
const CATEGORY_FIELD   = 'Вид изделия';           // фолбек-режим (старі дані)
const CARD_ID_FIELD    = 'Номер карточки';
const CARD_CODE_FIELD  = 'ОЕМ';
const CARD_NAME_FIELD  = 'Товарное наименование';
const FULL_NAME_FIELD  = 'Полное техническое наименование';
const ERRORS_FIELD     = 'Ошибки';
const PROPERTIES_FIELD = 'Свойства';
const GROUP_KEY_FIELD_PRIMARY  = 'Номер группы ch3';
const GROUP_KEY_FIELD_FALLBACK = 'Номер группы ch2';

const FILTER_FIELDS = ['Диаметр трубки, мм', 'Диаметр резьбы', 'Материал корпуса', 'Геометрия', 'Типы портов', 'Конструктивные признаки'];

// Деякі фільтри збираються не з однієї колонки, а з декількох "Порт N ..."
// одразу (наприклад діаметр трубки є в кожного порту окремо). Для таких
// фільтрів рядок вважається таким, що має значення V, якщо V зустрічається
// хоча б в одному з портів цього рядка.
const VIRTUAL_FILTER_COLUMN_PATTERNS = {
  'Диаметр трубки, мм': /^Порт \d+ диаметр трубки, мм$/i,
  'Диаметр резьбы': /^Порт \d+ диаметр резьбы$/i
};

// Повертає МАСИВ значень поля для рядка (для звичайних полів — 0 або 1
// значення; для віртуальних мультиколонкових фільтрів — усі непорожні
// значення з відповідних "Порт N ..." колонок цього рядка).
// Дехто зберігає числові значення як "8", дехто як "8.0" — без нормалізації
// це створило б два різних чекбокси для того самого діаметра.
function normalizeNumericToken(v) {
  if (v === '') return v;
  const n = Number(v.replace(',', '.'));
  if (!Number.isNaN(n) && /^-?\d+([.,]\d+)?$/.test(v)) {
    return String(n);
  }
  return v;
}

function getFieldValuesForRow(fieldName, row) {
  const pattern = VIRTUAL_FILTER_COLUMN_PATTERNS[fieldName];
  if (pattern) {
    const values = new Set();
    allHeaders.forEach(h => {
      if (pattern.test(h)) {
        const v = normalizeNumericToken(cell(row, h));
        if (v) values.add(v);
      }
    });
    return [...values];
  }
  const v = cell(row, fieldName);
  return v ? [v] : [];
}

const OVERVIEW_FIELDS = [
  'Вид изделия', 'Среда применения', 'Материал корпуса',
  'Геометрия', 'Типы портов', 'Количество портов', 'Конструктивные признаки'
];

const HIERARCHY = [
  { key: 'parent', field: 'Родитель',  imageOwn: 'картинка_ родитель', label: 'Категорія' },
  { key: 'group1', field: 'группа1',   imageOwn: 'картинка_ группа1',  label: 'Підкатегорія' },
  { key: 'group2', field: 'группа2',   imageOwn: null,                 label: 'Група' },
  { key: 'group3', field: 'группа3',   imageOwn: 'картинка_ группа3_small', label: 'Виріб' },
];
const GROUP3_IMAGE_BIG   = 'картинка_ группа3_big';
const GROUP3_DESCRIPTION = 'описание_ группа3';

const HIDDEN_EXTRA_FIELDS_BASE = new Set([
  CARD_ID_FIELD, CARD_CODE_FIELD, CARD_NAME_FIELD, FULL_NAME_FIELD,
  ERRORS_FIELD, PROPERTIES_FIELD, CATEGORY_FIELD,
  GROUP_KEY_FIELD_PRIMARY, GROUP_KEY_FIELD_FALLBACK,
  ...OVERVIEW_FIELDS,
  ...HIERARCHY.map(h => h.field), ...HIERARCHY.map(h => h.imageOwn).filter(Boolean),
  GROUP3_IMAGE_BIG, GROUP3_DESCRIPTION
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

let container = document.getElementById('tileSection');
const resultsContainer = document.getElementById('resultsContainer');
const sidePanel = document.getElementById('sidePanel');
const sideTitle = document.getElementById('sideTitle');
const sideList = document.getElementById('sideList');
const backButton = document.getElementById('backButton');
const catalogToolbar = document.getElementById('catalogToolbar');

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

function getFallbackIcon() {
  return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" style="opacity:0.3;">
    <circle cx="50" cy="50" r="40" fill="none" stroke="#64748b" stroke-width="3" stroke-dasharray="4 4"/>
    <path d="M30 50 H70 M50 30 V70" stroke="#64748b" stroke-width="3" stroke-linecap="round"/>
  </svg>`;
}

function getCategoryIconFallback(categoryName) {
  const label = categoryName.trim().toLowerCase();
  if (label.includes('тройник') || label.includes('крестовин')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <circle cx="50" cy="50" r="10" fill="#475569"/>
      <rect x="44" y="6" width="12" height="38" rx="4" fill="none" stroke="#0f172a" stroke-width="4"/>
      <rect x="44" y="56" width="12" height="38" rx="4" fill="none" stroke="#0f172a" stroke-width="4"/>
      <rect x="6" y="44" width="38" height="12" rx="4" fill="none" stroke="#0f172a" stroke-width="4"/>
      <rect x="56" y="44" width="38" height="12" rx="4" fill="none" stroke="#0f172a" stroke-width="4"/>
    </svg>`;
  }
  if (label.includes('кран')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <rect x="10" y="42" width="80" height="16" rx="6" fill="none" stroke="#0f172a" stroke-width="4"/>
      <circle cx="50" cy="50" r="14" fill="none" stroke="#0f172a" stroke-width="4"/>
      <rect x="46" y="14" width="8" height="26" rx="3" fill="#475569"/>
      <rect x="30" y="8" width="40" height="10" rx="4" fill="#475569"/>
    </svg>`;
  }
  if (label.includes('адаптер') || label.includes('перех')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <rect x="8" y="38" width="34" height="24" rx="4" fill="none" stroke="#0f172a" stroke-width="4"/>
      <rect x="58" y="30" width="34" height="40" rx="4" fill="none" stroke="#0f172a" stroke-width="4"/>
      <path d="M42 50 H58" stroke="#475569" stroke-width="4"/>
    </svg>`;
  }
  if (label.includes('з’єдну') || label.includes('соедин')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <rect x="8" y="40" width="36" height="20" rx="4" fill="none" stroke="#0f172a" stroke-width="4"/>
      <rect x="56" y="40" width="36" height="20" rx="4" fill="none" stroke="#0f172a" stroke-width="4"/>
      <rect x="40" y="44" width="20" height="12" fill="#475569"/>
    </svg>`;
  }
  return getFallbackIcon();
}

async function loadWorkbookRows() {
  if (!window.XLSX) throw new Error('SheetJS XLSX library is not loaded.');

  const fileUrl = new URL(XLSX_PATH, window.location.href);
  fileUrl.searchParams.set('v', Date.now());
  const response = await fetch(fileUrl.toString(), { cache: 'no-store' });
  if (!response.ok) throw new Error(`Cannot load ${XLSX_PATH}: ${response.status} ${response.statusText}`);

  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = workbook.SheetNames.includes('Результат') ? 'Результат' : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const sheetRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
  console.debug('Read sheet:', sheetName, 'rows:', sheetRows.length);
  return sheetRows;
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
    crumb.onclick = () => goToLevel(i + 1);
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
  if (resultsContainer) { resultsContainer.style.display = 'none'; resultsContainer.innerHTML = ''; }
  if (catalogToolbar) catalogToolbar.style.display = levelIndex > 0 ? 'flex' : 'none';

  container.innerHTML = '';
  container.style.display = 'block';

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

  container.appendChild(wrap);

  const grid = document.createElement('section');
  const gridClassByLevel = ['tile-container parent-grid', 'tile-container grid-6col', 'tile-container grid-6col', 'tile-container grid-6col'];
  grid.className = gridClassByLevel[levelIndex] || 'tile-container';
  container.appendChild(grid);

  const levelDef = HIERARCHY[levelIndex];
  const rows = rowsMatchingSelection(levelIndex);
  const valueMap = distinctValuesWithSample(rows, levelDef.field);

  if (valueMap.size === 0) {
    showCatalogMessage(grid, `Немає значень у стовпці "${levelDef.field}" для цього вибору.`, 'error');
    return;
  }

  [...valueMap.entries()].sort((a, b) => a[0].localeCompare(b[0], 'uk')).forEach(([value, info]) => {
    const div = document.createElement('div');
    const tileClassByLevel = ['tile tile-parent', 'tile tile-medium', 'tile tile-medium', 'tile tile-medium'];
    div.className = tileClassByLevel[levelIndex] || 'tile';

    let imageSrc = '';
    if (levelDef.imageOwn) {
      imageSrc = cell(info.sampleRow, levelDef.imageOwn);
    } else {
      // Рівень без власної картинки (наразі це "группа2") -> беремо картинку
      // власної колонки НАСТУПНОГО рівня з будь-якого дочірнього рядка,
      // де вона задана (тобто прев'ю однієї з дочірніх "группа3").
      const nextLevel = HIERARCHY[levelIndex + 1];
      if (nextLevel && nextLevel.imageOwn) {
        const scopeRows = allRows.filter(r =>
          HIERARCHY.slice(0, levelIndex + 1).every(h => cell(r, h.field) === (h.key === levelDef.key ? value : selection[h.key]))
        );
        const withImage = scopeRows.find(r => cell(r, nextLevel.imageOwn));
        if (withImage) imageSrc = cell(withImage, nextLevel.imageOwn);
      }
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
        // Останній рівень (группа3) обрано -> відкриваємо ЇЇ власну інфо-сторінку
        // (саме групу3, що відповідає значенню обраної плитки, а не першу-ліпшу)
        const scopeRows = rowsMatchingSelection(HIERARCHY.length);
        const sampleRow = scopeRows[0];
        renderGroup3Detail(value, sampleRow, scopeRows);
      }
    };
    grid.appendChild(div);
  });
}

function renderGroup3Detail(group3Value, sampleRow, scopeRows) {
  sidePanel.style.display = 'none';
  if (resultsContainer) { resultsContainer.style.display = 'none'; resultsContainer.innerHTML = ''; }
  if (catalogToolbar) catalogToolbar.style.display = 'flex';

  container.style.display = 'block';
  container.innerHTML = '';

  renderBreadcrumbs(container);

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

  container.appendChild(box);
}

function renderFullCatalogListing() {
  container.style.display = 'none';
  if (catalogToolbar) catalogToolbar.style.display = 'flex';

  sideTitle.textContent = `Усі товари (${allRows.length})`;
  sideList.innerHTML = '';
  sidePanel.style.display = 'block';

  resultsContainer.style.display = 'block';
  resultsContainer.innerHTML = '';
  resultsContainer.className = 'results-panel';

  const crumbWrap = document.createElement('div');
  renderBreadcrumbs(crumbWrap);
  resultsContainer.appendChild(crumbWrap);

  renderCategoryDetails({ rows: allRows, filterHeaders: FILTER_FIELDS, filterValues: buildFilterValueSets(allRows) });
}

function renderProductListing(group3Value, sampleRow, fallbackRows) {
  const groupKeyField = resolveGroupKeyField();
  const groupKeyValue = (groupKeyField && sampleRow) ? cell(sampleRow, groupKeyField) : '';

  let rows;
  if (groupKeyField && groupKeyValue) {
    rows = allRows.filter(row => cell(row, groupKeyField) === groupKeyValue);
  } else {
    rows = fallbackRows || [];
  }

  container.style.display = 'none';
  if (catalogToolbar) catalogToolbar.style.display = 'flex';

  sideTitle.textContent = group3Value;
  sideList.innerHTML = '';
  sidePanel.style.display = 'block';

  resultsContainer.style.display = 'block';
  resultsContainer.innerHTML = '';
  resultsContainer.className = 'results-panel';

  const crumbWrap = document.createElement('div');
  renderBreadcrumbs(crumbWrap);
  resultsContainer.appendChild(crumbWrap);

  renderCategoryDetails({ rows, filterHeaders: FILTER_FIELDS, filterValues: buildFilterValueSets(rows) });
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
    if (!values || values.length === 0) return;

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

  const filterPanel = document.createElement('div');
  filterPanel.style.marginBottom = '22px';

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

  const tagsBar = document.createElement('div');
  tagsBar.className = 'active-tags-bar';

  const countLabel = document.createElement('div');
  countLabel.className = 'results-count';

  const displayValuesContainer = document.createElement('div');

  const updateResults = () => {
    const filters = getSelectedFilters(filterPanel);
    const filtered = getFilteredDisplayValues(categoryDetail, filters);

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

    renderActiveTags(tagsBar, filters, filterPanel, categoryDetail, () => { itemsToShow = 12; updateResults(); });
    updateFilterAvailability(filterPanel, categoryDetail, filters);
    renderDisplayValues(displayValuesContainer, { rows: paginatedRows });

    if (totalFilteredCount === 0) {
      countLabel.textContent = 'Немає товарів, що відповідають вибраним фільтрам.';
      countLabel.style.color = '#dc2626';
    } else {
      countLabel.textContent = `Показано: ${Math.min(itemsToShow, totalFilteredCount)} з ${totalFilteredCount} товарів`;
      countLabel.style.color = '#64748b';
    }

    let loadMoreBtn = resultsContainer.querySelector('.load-more-container');
    if (loadMoreBtn) loadMoreBtn.remove();

    if (totalFilteredCount > itemsToShow) {
      const loadMoreContainer = document.createElement('div');
      loadMoreContainer.className = 'load-more-container';
      const btn = document.createElement('button');
      btn.className = 'load-more-button';
      btn.textContent = 'Показати ще';
      btn.onclick = () => { itemsToShow += 12; updateResults(); };
      loadMoreContainer.appendChild(btn);
      resultsContainer.appendChild(loadMoreContainer);
    }
  };

  const filtersEl = createFilterSection(categoryDetail, () => { itemsToShow = 12; updateResults(); });
  filterPanel.appendChild(filtersEl);

  resultsContainer.appendChild(searchContainer);
  resultsContainer.appendChild(tagsBar);
  resultsContainer.appendChild(countLabel);
  resultsContainer.appendChild(displayValuesContainer);

  updateResults();

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
    if (!categoryRows[category]) { categoryRows[category] = []; categoryOrder.push(category); }
    categoryRows[category].push(row);
  });

  categoryOrder.forEach(category => {
    const rowsForCategory = categoryRows[category];
    details[category] = { rows: rowsForCategory, filterHeaders: FILTER_FIELDS, filterValues: buildFilterValueSets(rowsForCategory) };
  });

  return { tiles: categoryOrder.map(label => ({ label, items: [] })), details };
}

function renderFlatTiles(target) {
  target.innerHTML = '';
  target.style.display = 'grid';
  tiles.forEach(tile => {
    const div = document.createElement('div');
    div.className = 'tile';
    const count = catalogDetails[tile.label] ? catalogDetails[tile.label].rows.length : 0;
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
  container.style.display = 'none';
  resultsContainer.style.display = 'block';
  resultsContainer.innerHTML = '';
  resultsContainer.className = 'results-panel';

  sideTitle.textContent = tile.label;
  sideList.innerHTML = '';
  sidePanel.style.display = 'block';
  if (catalogToolbar) catalogToolbar.style.display = 'flex';

  renderCategoryDetails(catalogDetails[tile.label]);
}

async function initializeCatalog() {
  try {
    const sheetRows = await loadWorkbookRows();
    const headerRow = sheetRows[0] || [];
    indexHeaders(headerRow);
    allRows = sheetRows.slice(1);

    deepHierarchyAvailable = hasField(HIERARCHY[0].field);

    if (deepHierarchyAvailable) {
      selection = {};
      navStack = [];
      renderLevelTiles(0);
    } else {
      console.warn(`Стовпця "${HIERARCHY[0].field}" не знайдено — використовую плаский режим за "${CATEGORY_FIELD}".`);
      const flat = buildFlatCatalog(allRows);
      catalogDetails = flat.details;
      tiles = [...staticTiles, ...flat.tiles];
      renderFlatTiles(container);
      if (flat.tiles.length === 0) {
        showCatalogMessage(container, `У файлі не знайдено значень ні у "${HIERARCHY[0].field}", ні у "${CATEGORY_FIELD}".`, 'error');
      }
    }
  } catch (error) {
    console.error(error);
    container.innerHTML = '';
    container.style.display = 'grid';
    showCatalogMessage(
      container,
      `Не вдалося завантажити ${XLSX_PATH}. Відкрийте сторінку через локальний сервер і перевірте, що файл існує в папці source.`,
      'error'
    );
  }
}

backButton.onclick = () => {
  if (!deepHierarchyAvailable) {
    sidePanel.style.display = 'none';
    if (catalogToolbar) catalogToolbar.style.display = 'none';
    if (resultsContainer) resultsContainer.style.display = 'none';
    container.style.display = 'grid';
    renderFlatTiles(container);
    return;
  }

  if (navStack.length === 0) {
    goToLevel(0);
    return;
  }

  const prevScreen = navStack.pop();
  if (prevScreen.type === 'group3detail') {
    renderGroup3Detail(prevScreen.group3Value, prevScreen.sampleRow, prevScreen.scopeRows);
  } else {
    delete selection[HIERARCHY[prevScreen.levelIndex].key];
    renderLevelTiles(prevScreen.levelIndex);
  }
};

function ensureXLSXAndInit() {
  function doInitWhenReady() {
    if (window.XLSX) return initializeCatalog();

    const existing = Array.from(document.getElementsByTagName('script')).find(s => s.src && s.src.indexOf('xlsx.full.min.js') !== -1);
    if (existing) {
      if (existing.getAttribute('data-xlsx-ready') === '1') return initializeCatalog();
      existing.addEventListener('load', () => {
        existing.setAttribute('data-xlsx-ready', '1');
        initializeCatalog();
      });
      setTimeout(() => { if (window.XLSX) initializeCatalog(); }, 50);
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
