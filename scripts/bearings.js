const staticTiles = [];

const XLSX_PATH = './source/bearings.xlsx';
const COLUMN = {
  a: 0, b: 1, c: 2,
  d: 3, e: 4, f: 5,
  g: 6, h: 7, i: 8, j: 9, k: 10, l: 11,
  m: 12,
  v: 21
};

const DISPLAY_COLUMNS = [COLUMN.a, COLUMN.b, COLUMN.c];
const FILTER_COLUMNS  = [COLUMN.d, COLUMN.e, COLUMN.f];
const EXTRA_COLUMNS   = [COLUMN.g, COLUMN.h, COLUMN.i, COLUMN.j, COLUMN.k, COLUMN.l];

let allHeaders = [];
let tiles = [...staticTiles];
let bearingDetails = {};

let container = document.getElementById('tileSection');
const resultsContainer = document.getElementById('resultsContainer');
const sidePanel = document.getElementById('sidePanel');
const sideTitle = document.getElementById('sideTitle');
const sideList = document.getElementById('sideList');
const backButton = document.getElementById('backButton');
const catalogToolbar = document.getElementById('catalogToolbar');

function normalizeCell(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function getCategoryIcon(categoryName) {
  const label = categoryName.trim().toLowerCase();

  if (label.includes('кульков')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <circle cx="50" cy="50" r="44" fill="none" stroke="#0f172a" stroke-width="4"/>
      <circle cx="50" cy="50" r="24" fill="none" stroke="#0f172a" stroke-width="4"/>
      <circle cx="50" cy="13" r="6" fill="#475569"/>
      <circle cx="50" cy="87" r="6" fill="#475569"/>
      <circle cx="13" cy="50" r="6" fill="#475569"/>
      <circle cx="87" cy="50" r="6" fill="#475569"/>
      <circle cx="24" cy="24" r="6" fill="#475569"/>
      <circle cx="76" cy="24" r="6" fill="#475569"/>
      <circle cx="24" cy="76" r="6" fill="#475569"/>
      <circle cx="76" cy="76" r="6" fill="#475569"/>
    </svg>`;
  }

  if (label.includes('роликов') && label.includes('циліндр')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <circle cx="50" cy="50" r="44" fill="none" stroke="#0f172a" stroke-width="4"/>
      <circle cx="50" cy="50" r="24" fill="none" stroke="#0f172a" stroke-width="4"/>
      <rect x="45" y="7" width="10" height="12" rx="1" fill="#475569"/>
      <rect x="45" y="81" width="10" height="12" rx="1" fill="#475569"/>
      <rect x="7" y="45" width="12" height="10" rx="1" fill="#475569"/>
      <rect x="81" y="45" width="12" height="10" rx="1" fill="#475569"/>
      <g transform="rotate(45 50 50)">
        <rect x="45" y="7" width="10" height="12" rx="1" fill="#475569"/>
        <rect x="45" y="81" width="10" height="12" rx="1" fill="#475569"/>
        <rect x="7" y="45" width="12" height="10" rx="1" fill="#475569"/>
        <rect x="81" y="45" width="12" height="10" rx="1" fill="#475569"/>
      </g>
    </svg>`;
  }

  if (label.includes('роликов') && label.includes('коніч')) {
    return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <circle cx="50" cy="50" r="44" fill="none" stroke="#0f172a" stroke-width="4"/>
      <circle cx="50" cy="50" r="22" fill="none" stroke="#0f172a" stroke-width="3"/>
      <polygon points="44,7 56,7 53,17 47,17" fill="#475569"/>
      <g transform="rotate(45 50 50)"><polygon points="44,7 56,7 53,17 47,17" fill="#475569"/></g>
      <g transform="rotate(90 50 50)"><polygon points="44,7 56,7 53,17 47,17" fill="#475569"/></g>
      <g transform="rotate(135 50 50)"><polygon points="44,7 56,7 53,17 47,17" fill="#475569"/></g>
      <g transform="rotate(180 50 50)"><polygon points="44,7 56,7 53,17 47,17" fill="#475569"/></g>
      <g transform="rotate(225 50 50)"><polygon points="44,7 56,7 53,17 47,17" fill="#475569"/></g>
      <g transform="rotate(270 50 50)"><polygon points="44,7 56,7 53,17 47,17" fill="#475569"/></g>
      <g transform="rotate(315 50 50)"><polygon points="44,7 56,7 53,17 47,17" fill="#475569"/></g>
    </svg>`;
  }

  return `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" style="opacity: 0.3;">
    <circle cx="50" cy="50" r="40" fill="none" stroke="#64748b" stroke-width="3" stroke-dasharray="4 4"/>
    <path d="M30 50 H70 M50 30 V70" stroke="#64748b" stroke-width="3" stroke-linecap="round"/>
  </svg>`;
}

function sortValues(values) {
  return [...values].sort((a, b) => {
    const first = Number(String(a).replace(',', '.'));
    const second = Number(String(b).replace(',', '.'));

    if (!Number.isNaN(first) && !Number.isNaN(second)) {
      return first - second;
    }
    return String(a).localeCompare(String(b), 'uk');
  });
}

function buildBearingCatalog(rows) {
  const header = rows[0] || [];
  allHeaders = header.map((h, i) => normalizeCell(h) || String.fromCharCode(65 + i));
  while (allHeaders.length <= COLUMN.l) allHeaders.push(String.fromCharCode(65 + allHeaders.length));
  const displayHeaders = DISPLAY_COLUMNS.map(index => allHeaders[index]);
  const filterHeaders  = FILTER_COLUMNS.map(index => allHeaders[index]);

  const categoryRows = {};
  const details = {};
  const categoryOrder = [];

  rows.slice(1).forEach(row => {
    const category = normalizeCell(row[COLUMN.m]);

    if (!category) {
      return;
    }

    if (!categoryRows[category]) {
      categoryRows[category] = [];
      categoryOrder.push(category);
    }
    categoryRows[category].push(row);
  });

  categoryOrder.forEach(category => {
    const rowsForCategory = categoryRows[category];
    const displayValueSets = {};
    const filterValueSets = {};

    displayHeaders.forEach(headerName => {
      displayValueSets[headerName] = new Set();
    });

    filterHeaders.forEach(headerName => {
      filterValueSets[headerName] = new Set();
    });

    rowsForCategory.forEach(row => {
      DISPLAY_COLUMNS.forEach((colIndex, idx) => {
        const value = normalizeCell(row[colIndex]);
        if (value) displayValueSets[displayHeaders[idx]].add(value);
      });

      FILTER_COLUMNS.forEach((colIndex, idx) => {
        const value = normalizeCell(row[colIndex]);
        if (value) filterValueSets[filterHeaders[idx]].add(value);
      });
    });

    details[category] = {
      rows: rowsForCategory,
      displayHeaders,
      filterHeaders,
      displayValues: Object.fromEntries(
        displayHeaders.map(name => [name, sortValues(displayValueSets[name])])
      ),
      filterValues: Object.fromEntries(
        filterHeaders.map(name => [name, sortValues(filterValueSets[name])])
      )
    };
  });

  return {
    tiles: categoryOrder.map(label => ({ label, items: [] })),
    details
  };
}

async function loadBearingCatalog() {
  if (!window.XLSX) {
    throw new Error('SheetJS XLSX library is not loaded.');
  }

  const fileUrl = new URL(XLSX_PATH, window.location.href);
  fileUrl.searchParams.set('v', Date.now());
  const response = await fetch(fileUrl.toString(), { cache: 'no-store' });

  if (!response.ok) {
    throw new Error(`Cannot load ${XLSX_PATH}: ${response.status} ${response.statusText}`);
  }

  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });

  const allRows = [];
  workbook.SheetNames.forEach(name => {
    const sheet = workbook.Sheets[name];
    const sheetRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
    if (sheetRows.length > 1) {
      console.debug('Read sheet:', name, 'rows:', sheetRows.length);
      const normalized = sheetRows.slice(1).map(r => {
        const row = Array.isArray(r) ? r.slice() : [];
        while (row.length <= COLUMN.v) row.push('');
        return row;
      });
      allRows.push(...normalized);
    }
  });

  if (allRows.length === 0 && workbook.SheetNames.length > 0) {
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const sheetRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
    console.debug('Fallback read first sheet:', workbook.SheetNames[0], 'rows:', sheetRows.length);
    const normalized = sheetRows.slice(1).map(r => {
      const row = Array.isArray(r) ? r.slice() : [];
      while (row.length <= COLUMN.v) row.push('');
      return row;
    });
    allRows.push(...normalized);
  }

  const headerRow = workbook.SheetNames.length > 0
    ? (() => {
        const firstSheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '', blankrows: false });
        const row = Array.isArray(firstSheetRows[0]) ? firstSheetRows[0].slice() : [];
        while (row.length <= COLUMN.v) row.push('');
        return row;
      })()
    : [];

  return buildBearingCatalog([headerRow, ...allRows]);
}

function renderDisplayValues(resultContainer, filteredData) {
  resultContainer.innerHTML = '';
  resultContainer.className = 'results-cards-grid';

  if (!filteredData.rows || filteredData.rows.length === 0) {
    const empty = document.createElement('div');
    empty.textContent = 'Немає значень для цього набору фільтрів.';
    empty.style.color = '#666';
    empty.style.padding = '20px';
    resultContainer.appendChild(empty);
    return;
  }

  filteredData.rows.forEach(row => {
    const card = document.createElement('div');
    card.className = 'product-card';

    const kartka = normalizeCell(row[DISPLAY_COLUMNS[0]]);
    const kod = normalizeCell(row[DISPLAY_COLUMNS[1]]);
    const name = normalizeCell(row[DISPLAY_COLUMNS[2]]);
    
    const header = document.createElement('div');
    header.className = 'card-header';
    
    const nameElem = document.createElement('div');
    nameElem.className = 'card-name';
    nameElem.innerHTML = `<span class="product-title-text">${name || 'Без назви'}</span>`;
    header.appendChild(nameElem);
    
    const specsBox = document.createElement('div');
    specsBox.className = 'card-specs-box';
    
    const kartkaElem = document.createElement('div');
    kartkaElem.className = 'card-line';
    kartkaElem.innerHTML = `<span class="card-line-label">Картка:</span> <span class="card-line-value">${kartka || 'N/A'}</span>`;
    specsBox.appendChild(kartkaElem);
    
    const kodElem = document.createElement('div');
    kodElem.className = 'card-line';
    kodElem.innerHTML = `<span class="card-line-label">Код за каталогом:</span> <span class="card-line-value highlight">${kod || 'N/A'}</span>`;
    specsBox.appendChild(kodElem);
    
    header.appendChild(specsBox);
    card.appendChild(header);

    card.addEventListener('click', () => openModal(row));
    resultContainer.appendChild(card);
  });
}

// ---- MODAL ----
const modalOverlay = document.getElementById('modalOverlay');
const modalContent = document.getElementById('modalContent');
const modalClose   = document.getElementById('modalClose');

function openModal(row) {
  const v = (idx) => normalizeCell(row[idx]);
  const imageName = v(COLUMN.v);

  let html = '';

  if (imageName) {
    html += `
      <div class="modal-header-container">
        <img src="images/${imageName}" class="modal-image" alt="${v(COLUMN.c) || 'Підшипник'}" onerror="this.style.display='none'; this.parentElement.classList.add('no-image');">
        <div class="modal-header-text">
          <p class="modal-product-name">${v(COLUMN.c) || 'Без назви'}</p>
        </div>
      </div>
    `;
  } else {
    html += `<p class="modal-product-name" style="margin-bottom: 16px;">${v(COLUMN.c) || 'Без назви'}</p>`;
  }

  html += `<div class="modal-section">
    <p class="modal-section-title">Ідентифікація</p>`;
  [[COLUMN.b, allHeaders[COLUMN.b]], [COLUMN.a, allHeaders[COLUMN.a]]].forEach(([idx, lbl]) => {
    const val = v(idx);
    html += `<div class="modal-row">
      <span class="modal-label">${lbl}</span>
      <span class="modal-value${val ? '' : ' empty'}">${val || '—'}</span>
    </div>`;
  });
  html += `</div><hr class="modal-divider">`;

  const hasB2 = FILTER_COLUMNS.some(i => v(i) !== '');
  if (hasB2) {
    html += `<div class="modal-section">
      <p class="modal-section-title">Розміри</p>`;
    FILTER_COLUMNS.forEach(idx => {
      const val = v(idx);
      html += `<div class="modal-row">
        <span class="modal-label">${allHeaders[idx]}</span>
        <span class="modal-value${val ? '' : ' empty'}">${val || '—'}</span>
      </div>`;
    });
    html += `</div><hr class="modal-divider">`;
  }

  const hasB3 = EXTRA_COLUMNS.some(i => v(i) !== '');
  if (hasB3) {
    html += `<div class="modal-section">
      <p class="modal-section-title">Технічні характеристики</p>`;
    EXTRA_COLUMNS.forEach(idx => {
      const val = v(idx);
      html += `<div class="modal-row">
        <span class="modal-label">${allHeaders[idx]}</span>
        <span class="modal-value${val ? '' : ' empty'}">${val || '—'}</span>
      </div>`;
    });
    html += `</div>`;
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

function getFilteredDisplayValues(categoryDetails, filters) {
  const active = Object.entries(filters).filter(([, v]) => v.length > 0);
  const filteredRows = categoryDetails.rows.filter(row =>
    active.every(([headerName, values]) => {
      const idx = categoryDetails.filterHeaders.indexOf(headerName);
      if (idx === -1) return true;
      return values.includes(normalizeCell(row[FILTER_COLUMNS[idx]]));
    })
  );
  return {
    rows: filteredRows,
    filterHeaders: categoryDetails.filterHeaders,
    displayHeaders: categoryDetails.displayHeaders
  };
}

function getAvailableValuesForGroup(categoryDetails, filters, targetGroup) {
  const otherActive = Object.entries(filters).filter(([g, v]) => g !== targetGroup && v.length > 0);
  const subset = categoryDetails.rows.filter(row =>
    otherActive.every(([headerName, values]) => {
      const idx = categoryDetails.filterHeaders.indexOf(headerName);
      if (idx === -1) return true;
      return values.includes(normalizeCell(row[FILTER_COLUMNS[idx]]));
    })
  );
  const idx = categoryDetails.filterHeaders.indexOf(targetGroup);
  const available = new Set();
  subset.forEach(row => {
    const v = normalizeCell(row[FILTER_COLUMNS[idx]]);
    if (v) available.add(v);
  });
  return available;
}

function renderActiveTags(tagsBar, filters, filterPanel, categoryDetails, onUpdate) {
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

function updateFilterAvailability(filterPanel, categoryDetails, filters) {
  categoryDetails.filterHeaders.forEach(group => {
    const available = getAvailableValuesForGroup(categoryDetails, filters, group);

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

function createFilterSection(categoryDetails, onFilterChange) {
  const filterSection = document.createElement('div');
  filterSection.style.marginBottom = '20px';

  const title = document.createElement('h4');
  title.textContent = 'Фільтри';
  title.style.margin = '0 0 10px 0';
  filterSection.appendChild(title);

  categoryDetails.filterHeaders.forEach((headerName, headerIdx) => {
    const values = categoryDetails.filterValues[headerName];
    if (!values || values.length === 0) return;

    const colIndex = FILTER_COLUMNS[headerIdx];

    const details = document.createElement('details');
    details.className = 'details-filter-panel';
    details.open = true;

    const summary = document.createElement('summary');
    // Загортаємо в span з фіксованим line-height
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

      const count = categoryDetails.rows.filter(row => normalizeCell(row[colIndex]) === value).length;

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

function renderCategoryDetails(categoryDetails) {
  let itemsToShow = 12;
  let searchQuery = '';

  const filterPanel = document.createElement('div');
  filterPanel.style.marginBottom = '22px';

  const searchContainer = document.createElement('div');
  searchContainer.className = 'search-container';
  const searchInput = document.createElement('input');
  searchInput.type = 'text';
  searchInput.className = 'search-input';
  searchInput.placeholder = 'Швидкий пошук за кодом, карткою або найменуванням...';
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
    const filtered = getFilteredDisplayValues(categoryDetails, filters);

    if (searchQuery) {
      filtered.rows = filtered.rows.filter(row => {
        const kartka = normalizeCell(row[DISPLAY_COLUMNS[0]]).toLowerCase();
        const kod = normalizeCell(row[DISPLAY_COLUMNS[1]]).toLowerCase();
        const name = normalizeCell(row[DISPLAY_COLUMNS[2]]).toLowerCase();
        return kartka.includes(searchQuery) || kod.includes(searchQuery) || name.includes(searchQuery);
      });
    }

    const totalFilteredCount = filtered.rows.length;
    const paginatedRows = filtered.rows.slice(0, itemsToShow);

    renderActiveTags(tagsBar, filters, filterPanel, categoryDetails, () => {
      itemsToShow = 12;
      updateResults();
    });
    updateFilterAvailability(filterPanel, categoryDetails, filters);

    renderDisplayValues(displayValuesContainer, {
      rows: paginatedRows,
      filterHeaders: filtered.filterHeaders,
      displayHeaders: filtered.displayHeaders
    });

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
      btn.onclick = () => {
        itemsToShow += 12;
        updateResults();
      };
      loadMoreContainer.appendChild(btn);
      resultsContainer.appendChild(loadMoreContainer);
    }
  };

  const handleFilterChange = () => {
    itemsToShow = 12;
    updateResults();
  };

  const filtersEl = createFilterSection(categoryDetails, handleFilterChange);
  filterPanel.appendChild(filtersEl);

  // Спочатку рендеримо статичну структуру в DOM
  resultsContainer.appendChild(searchContainer);
  resultsContainer.appendChild(tagsBar);
  resultsContainer.appendChild(countLabel);
  resultsContainer.appendChild(displayValuesContainer);

  // Лише після цього викликаємо апдейт, щоб кнопка впала на самий низ
  updateResults();

  sideList.appendChild(filterPanel);
}

function showTileDetails(tile, activeContainer) {
  if (activeContainer && activeContainer.style) activeContainer.style.display = 'none';
  if (resultsContainer) {
    resultsContainer.style.display = 'block';
    resultsContainer.innerHTML = '';
    resultsContainer.className = 'results-panel';
  }

  sideTitle.textContent = tile.label;
  sideList.innerHTML = '';

  if (tile.items && tile.items.length > 0) {
    const itemsUl = document.createElement('ul');
    itemsUl.style.paddingLeft = '20px';
    itemsUl.style.margin = '0 0 15px 0';
    tile.items.forEach(item => {
      const li = document.createElement('li');
      li.textContent = item;
      itemsUl.appendChild(li);
    });
    sideList.appendChild(itemsUl);
  }

  if (bearingDetails[tile.label]) {
    renderCategoryDetails(bearingDetails[tile.label]);
  }
  sidePanel.style.display = 'block';
  if (catalogToolbar) catalogToolbar.style.display = 'flex';
}

function createTileElement(tile, activeContainer) {
  const div = document.createElement('div');
  div.className = 'tile';
  
  const iconSvg = getCategoryIcon(tile.label);

  div.innerHTML = `
    ${iconSvg}
    <span>${tile.label}</span>
  `;
  div.onclick = () => showTileDetails(tile, activeContainer);

  return div;
}

function renderTiles(target) {
  target.innerHTML = '';
  tiles.forEach(tile => {
    target.appendChild(createTileElement(tile, target));
  });
}

function showCatalogMessage(target, text, type = '') {
  const message = document.createElement('div');
  message.className = `catalog-message ${type}`.trim();
  message.textContent = text;
  target.appendChild(message);
}

async function initializeCatalog() {
  try {
    const bearingCatalog = await loadBearingCatalog();
    bearingDetails = bearingCatalog.details;
    tiles = [...staticTiles, ...bearingCatalog.tiles];

    renderTiles(container);

    if (bearingCatalog.tiles.length === 0) {
      showCatalogMessage(container, 'У файлі source/bearings.xlsx не знайдено значень у стовпці M.', 'error');
    }
  } catch (error) {
    console.error(error);
    tiles = [...staticTiles];
    bearingDetails = {};

    renderTiles(container);
    showCatalogMessage(
      container,
      'Не вдалося завантажити source/bearings.xlsx. Відкрийте сторінку через локальний сервер і перевірте, що файл існує в папці source.',
      'error'
    );
  }
}

backButton.onclick = () => {
  sidePanel.style.display = 'none';
  if (catalogToolbar) catalogToolbar.style.display = 'none';
  if (resultsContainer) resultsContainer.style.display = 'none';

  if (!container || !document.getElementById('tileSection')) {
    container = document.querySelector('.catalog-content .tile-container') || (function() {
      const newContainer = document.createElement('section');
      newContainer.className = 'tile-container';
      newContainer.id = 'tileSection';
      document.querySelector('.catalog-content').appendChild(newContainer);
      return newContainer;
    })();
  }

  container.style.display = 'grid';
  renderTiles(container);
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