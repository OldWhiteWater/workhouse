document.addEventListener("DOMContentLoaded", () => {
  const menuBtn = document.getElementById('catalogBtn');
  const megaMenu = document.getElementById('megaMenu');
  const level1List = document.getElementById('menuLevel1');
  const level2List = document.getElementById('menuLevel2');
  const menuCloseBtn = document.getElementById('menuCloseBtn');
  const mainContent = document.getElementById('mainContent');

  let treeData = {};
  let isMenuOpen = false;

  // Загрузка данных
  async function loadData() {
    try {
      const response = await fetch("source/Каталог_товарів_меню.xlsx");
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
          // Мы сохраняем l3, чтобы потом вывести их в main
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

    const level2Keys = Object.keys(treeData[l1Name] || {});
    if (level2Keys.length === 0) return;

    level2Keys.forEach((l2) => {
      const li = document.createElement('li');
      li.className = 'menu-item level2-item';
      li.textContent = l2;
      
      // Клик по 2-му уровню тоже можем отправлять в main (опционально)
      li.addEventListener('click', () => {
         renderMainContent(l1Name);
         closeMenu();
      });
      
      level2List.appendChild(li);
    });
  }

  // Отрисовка контента в MAIN
  function renderMainContent(l1Name) {
    if (!mainContent) return;

    // Очищаем main
    mainContent.innerHTML = '';

    // Создаем контейнер для двух колонок
    const columnsContainer = document.createElement('div');
    columnsContainer.className = 'main-two-columns';

    // --- ЛЕВАЯ КОЛОНКА (Только выбранный Рівень 1) ---
    const leftCol = document.createElement('div');
    leftCol.className = 'main-col-left';
    
    const title = document.createElement('h1');
    title.className = 'main-title';
    title.textContent = l1Name;
    leftCol.appendChild(title);
    
    columnsContainer.appendChild(leftCol);

    // --- ПРАВАЯ КОЛОНКА (Список Рівень 2) ---
    const rightCol = document.createElement('div');
    rightCol.className = 'main-col-right';

    const level2Keys = Object.keys(treeData[l1Name] || {});
    
    if (level2Keys.length > 0) {
      level2Keys.forEach(l2 => {
        const div = document.createElement('div');
        div.className = 'main-subtitle';
        div.textContent = l2;
        // Если нужно, чтобы по клику на 2-й уровень происходил переход:
        // div.onclick = () => { ... };
        rightCol.appendChild(div);
      });
    } else {
      rightCol.innerHTML = '<div style="color:#94a3b8;">Немає підкатегорій</div>';
    }
    
    columnsContainer.appendChild(rightCol);

    // Добавляем всё в main
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