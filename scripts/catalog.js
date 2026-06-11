const staticTiles = [
  {
    img: 'img1.jpg',
    label: 'Ліхтарі габаритні та декоративного освітлення',
    items: ['Світлодіодні стрічки', 'Габаритні ліхтарі', 'Декоративні елементи']
  },
  {
    img: 'img2.jpg',
    label: 'Лампи побутові та промислові',
    items: ['LED лампи', 'Галогенні', 'Люмінесцентні']
  },
  {
    img: 'img3.jpg',
    label: 'Сигнальні вогні та підсвічування номерного знаку',
    items: ['Лампи номерного знаку', 'Сигнальні вогні']
  },
  {
    img: 'img4.jpg',
    label: 'Лампи автомобільні',
    items: ['Ближнє світло', 'Дальнє світло', 'Поворотники']
  },
  {
    img: 'img5.jpg',
    label: 'Фари головного освітлення та передні протитуманні',
    items: ['Головні фари', 'Протитуманні фари']
  },
  {
    img: 'img6.jpg',
    label: 'Додаткові фари та ліхтарі',
    items: ['Робочі фари', 'Бокові ліхтарі']
  },
  {
    img: 'img7.jpg',
    label: 'Проблискові маяки, шашки таксі',
    items: ['Маяки', 'Таксі-шашки']
  },
  {
    img: 'img8.jpg',
    label: 'Патрони та роз\'єми для ламп, фар, ліхтарів',
    items: ['Патрони', 'Роз\'єми', 'Перехідники']
  },
  {
    img: 'img9.jpg',
    label: 'Освітлення внутрішнього простору ТЗ',
    items: ['Плафони', 'Підсвічування салону']
  },
  {
    img: 'img10.jpg',
    label: 'Інтер\'єрні світильники',
    items: ['Настінні', 'Декоративні', 'Точкові']
  },
  {
    img: 'img11.jpg',
    label: 'Ліхтарі, лампи портативні та переносні',
    items: ['Акумуляторні', 'Переноски', 'Ручні ліхтарі']
  },
  {
    img: 'img12.jpg',
    label: 'Ліхтарі екстер\'єрні',
    items: ['Фасадні', 'Прожектори']
  },
  {
    img: 'img13.jpg',
    label: 'Захисні решітки для фар та ліхтарів',
    items: ['Решітки', 'Захисні накладки']
  }
];

const lampDetails = {
  'Тип цоколя': ['A60', 'A65', 'A70', 'C37', 'E40'],
  'Номінальна потужність, Вт': ['5', '7', '8', '10', '15', '18'],
  'Температура світла, К': ['3000k', '4100k', '5000k', '6500k'],
  'Світловий потік, Lm': ['400', '520', '600', '800', '1000']
};

const XLSX_PATH = './source/bearings.xlsx';
const COLUMN = {
  d: 3,
  e: 4,
  f: 5,
  m: 12
};

let tiles = [...staticTiles];
let bearingDetails = {};

const container = document.getElementById('tileSection');
const sidePanel = document.getElementById('sidePanel');
const sideTitle = document.getElementById('sideTitle');
const sideList = document.getElementById('sideList');
const backButton = document.getElementById('backButton');

function normalizeCell(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
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
  const propertyColumns = [COLUMN.d, COLUMN.e, COLUMN.f];
  const propertyNames = propertyColumns.map(index => normalizeCell(header[index]) || `Стовпець ${index + 1}`);
  const details = {};
  const categoryOrder = [];

  rows.slice(1).forEach(row => {
    const category = normalizeCell(row[COLUMN.m]);

    if (!category) {
      return;
    }

    if (!details[category]) {
      details[category] = {};
      categoryOrder.push(category);

      propertyNames.forEach(name => {
        details[category][name] = new Set();
      });
    }

    propertyColumns.forEach((columnIndex, propertyIndex) => {
      const value = normalizeCell(row[columnIndex]);

      if (value) {
        details[category][propertyNames[propertyIndex]].add(value);
      }
    });
  });

  const normalizedDetails = {};

  categoryOrder.forEach(category => {
    normalizedDetails[category] = {};

    propertyNames.forEach(name => {
      normalizedDetails[category][name] = sortValues(details[category][name]);
    });
  });

  return {
    tiles: categoryOrder.map(label => ({ label, items: [] })),
    details: normalizedDetails
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
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  return buildBearingCatalog(rows);
}

function appendDetailGroups(target, detailGroups, withCheckboxes = true) {
  for (const [group, values] of Object.entries(detailGroups)) {
    const details = document.createElement('details');
    const summary = document.createElement('summary');
    summary.textContent = group;
    details.appendChild(summary);

    const ul = document.createElement('ul');
    ul.style.paddingLeft = '20px';
    ul.style.listStyle = 'none';

    values.forEach(val => {
      const li = document.createElement('li');

      if (withCheckboxes) {
        const label = document.createElement('label');
        label.style.cursor = 'pointer';
        label.style.display = 'block';
        label.style.marginBottom = '5px';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = val;
        checkbox.style.marginRight = '8px';

        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(val));
        li.appendChild(label);
      } else {
        li.textContent = val;
      }

      ul.appendChild(li);
    });

    details.appendChild(ul);
    target.appendChild(details);
  }
}

function showTileDetails(tile, activeContainer) {
  activeContainer.remove();

  sideTitle.textContent = tile.label;
  sideList.innerHTML = '';

  tile.items.forEach(item => {
    const li = document.createElement('li');
    li.textContent = item;
    sideList.appendChild(li);
  });

  if (tile.label === 'Лампи побутові та промислові') {
    appendDetailGroups(sideList, lampDetails);
  }

  if (bearingDetails[tile.label]) {
    appendDetailGroups(sideList, bearingDetails[tile.label]);
  }

  sidePanel.style.display = 'block';
}

function createTileElement(tile, activeContainer) {
  const div = document.createElement('div');
  div.className = tile.img ? 'tile' : 'tile no-image';
  div.innerHTML = `
    ${tile.img ? `<img src="images/${tile.img}" alt="${tile.label}" />` : ''}
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

  const newContainer = document.createElement('section');
  newContainer.className = 'tile-container';
  newContainer.id = 'tileSection';

  renderTiles(newContainer);
  document.querySelector('.catalog-content').appendChild(newContainer);
};

initializeCatalog();
