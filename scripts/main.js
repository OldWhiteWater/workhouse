document.addEventListener("DOMContentLoaded", () => {
  const menuBtn = document.getElementById('catalogBtn');
  const megaMenu = document.getElementById('megaMenu');
  const level1List = document.getElementById('menuLevel1');
  const level2List = document.getElementById('menuLevel2');
  const level3Box = document.getElementById('menuLevel3');
  const menuCloseBtn = document.getElementById('menuCloseBtn');
  const mainContent = document.getElementById('mainContent');

  let treeData = {};
  let isMenuOpen = false;

  // Загрузка данных
  async function loadData() {
    try {
      const response = await fetch("source/Каталог_товарів.xlsx");
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      
      rows.shift(); 

      rows.forEach(row => {
        const [l1, l2, l3] = row.map(cell => (cell || '').toString().trim());
        if (!l1) return;

        if (!treeData[l1]) treeData[l1] = {};
        if (l2) {
          if (!treeData[l1][l2]) treeData[l1][l2] = new Set();
          if (l3) treeData[l1][l2].add(l3);
        }
      });

      for (const l1 in treeData) {
        for (const l2 in treeData[l1]) {
          treeData[l1][l2] = Array.from(treeData[l1][l2]);
        }
      }

      initMenu();
    } catch (e) {
      console.error("Ошибка загрузки каталога:", e);
    }
  }

  // Рендер меню в шапке
  function initMenu() {
    const level1Keys = Object.keys(treeData);
    if (level1Keys.length === 0) return;

    level1List.innerHTML = ''; 

    level1Keys.forEach((l1) => {
      const li = document.createElement('li');
      li.className = 'menu-item level1-item';
      li.textContent = l1;
      
      // Ховер для ПК
      li.addEventListener('mouseenter', () => selectLevel1(l1, li));
      
      // Клик для выбора категории и отрисовки в main
      li.addEventListener('click', () => {
        selectLevel1(l1, li);
        renderMainContent(l1); // Отрисовка в main
        closeMenu(); 
      });
      
      level1List.appendChild(li);
    });
  }

  function selectLevel1(l1Name, liElement) {
    document.querySelectorAll('.level1-item').forEach(el => el.classList.remove('active'));
    liElement.classList.add('active');

    level2List.innerHTML = '';
    level3Box.innerHTML = '';

    const level2Keys = Object.keys(treeData[l1Name] || {});
    if (level2Keys.length === 0) return;

    level2Keys.forEach((l2, index) => {
      const li = document.createElement('li');
      li.className = 'menu-item level2-item';
      li.textContent = l2;
      
      li.addEventListener('mouseenter', () => selectLevel2(l1Name, l2, li));
      
      level2List.appendChild(li);

      if (index === 0 && window.innerWidth > 600) {
        setTimeout(() => selectLevel2(l1Name, l2, li), 0);
      }
    });
  }

  function selectLevel2(l1Name, l2Name, liElement) {
    document.querySelectorAll('.level2-item').forEach(el => el.classList.remove('active'));
    liElement.classList.add('active');

    level3Box.innerHTML = '';

    const level3Items = treeData[l1Name][l2Name] || [];
    if (level3Items.length === 0) {
       level3Box.innerHTML = '<div class="no-items">Немає товарів</div>';
       return;
    }

    const grid = document.createElement('div');
    grid.className = 'level3-grid';

    level3Items.forEach(l3 => {
      const item = document.createElement('div');
      item.className = 'level3-item';
      item.innerHTML = `<span class="level3-text">${l3}</span>`;
      grid.appendChild(item);
    });

    level3Box.appendChild(grid);
  }

  // Отрисовка контента в MAIN
  function renderMainContent(l1Name) {
    if (!mainContent) return;

    // Очищаем main перед новой отрисовкой
    mainContent.innerHTML = '';

    // Заголовок H1
    const title = document.createElement('h1');
    title.className = 'category-title';
    title.textContent = l1Name;
    mainContent.appendChild(title);

    // Контейнер для колонок
    const columnsContainer = document.createElement('div');
    columnsContainer.className = 'category-columns';

    const level2Keys = Object.keys(treeData[l1Name] || []);

    level2Keys.forEach((l2Name) => {
      const column = document.createElement('div');
      column.className = 'category-column';

      // Шапка колонки
      const colHeader = document.createElement('div');
      colHeader.className = 'column-header';
      colHeader.innerHTML = `<h3>${l2Name}</h3>`;
      column.appendChild(colHeader);

      // Список 3-го уровня
      const list = document.createElement('ul');
      list.className = 'column-list';
      
      const level3Items = treeData[l1Name][l2Name] || [];
      const maxItems = 8; // Показываем 8, остальные скрываем

      // Рендерим все элементы, но скрытым добавляем класс 'hidden-item'
      level3Items.forEach((l3, index) => {
        const li = document.createElement('li');
        if (index >= maxItems) {
          li.classList.add('hidden-item'); // Скрываем лишние
        }
        li.innerHTML = `<a href="#">${l3}</a> <span class="count">(${Math.floor(Math.random() * 100) + 1})</span>`;
        list.appendChild(li);
      });

      column.appendChild(list);

      // Кнопка "Показати всі категорії"
      if (level3Items.length > maxItems) {
        const showAll = document.createElement('button');
        showAll.className = 'show-all-link';
        showAll.textContent = 'Показати всі категорії';
        
        // Логика раскрытия
        showAll.addEventListener('click', function(e) {
          e.preventDefault();
          // Показываем скрытые элементы в этой колонке
          list.querySelectorAll('.hidden-item').forEach(el => {
            el.classList.remove('hidden-item');
          });
          // Удаляем кнопку после раскрытия
          this.remove();
        });
        
        column.appendChild(showAll);
      }

      columnsContainer.appendChild(column);
    });

    mainContent.appendChild(columnsContainer);
  }

  // Управление меню
  function closeMenu() {
    megaMenu.classList.remove('show');
    menuBtn.classList.remove('active');
    isMenuOpen = false;
  }

  menuBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    isMenuOpen = !isMenuOpen;
    if (isMenuOpen) {
      megaMenu.classList.add('show');
      menuBtn.classList.add('active');
    } else {
      closeMenu();
    }
  });

  if (menuCloseBtn) {
    menuCloseBtn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      closeMenu();
    });
  }

  document.addEventListener('click', (e) => {
    if (isMenuOpen && !megaMenu.contains(e.target) && !menuBtn.contains(e.target)) {
      closeMenu();
    }
  });

  loadData();
});