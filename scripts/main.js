// ------------------------------
// ЛОГІКА РОЗКРИТТЯ ДЕРЕВА
// ------------------------------
function activateCarets() {
  document.querySelectorAll(".caret").forEach(caret => {
    caret.addEventListener("click", function () {
      const nested = this.nextElementSibling;
      if (nested) {
        nested.classList.toggle("active");
        this.classList.toggle("caret-down");
      }
    });
  });
}

document.addEventListener("DOMContentLoaded", activateCarets);


// ------------------------------
// ЕКСПОРТ ДЕРЕВА В EXCEL
// ------------------------------

// Знаходимо максимальну глибину дерева
function getMaxDepth(node, depth = 1) {
  let max = depth;
  const children = Array.from(node.children);

  for (const li of children) {
    if (li.tagName !== "LI") continue;

    const sub = li.querySelector(":scope > ul");
    if (sub) {
      max = Math.max(max, getMaxDepth(sub, depth + 1));
    }
  }
  return max;
}

// Рекурсивний обхід дерева
function traverse(node, level = [], rows) {
  const items = Array.from(node.children);

  for (const li of items) {
    if (li.tagName !== "LI") continue;

    const span = li.firstElementChild?.tagName === "SPAN"
      ? li.firstElementChild
      : null;

    const label = span ? span.textContent.trim() : null;
    const sublist = li.querySelector(":scope > ul");

    if (label && sublist) {
      traverse(sublist, [...level, label], rows);
    } 
    else if (label && !sublist) {
      rows.push([...level, label]);
    } 
    else if (!label && !sublist) {
      const text = li.textContent.trim();
      if (text) rows.push([...level, text]);
    } 
    else if (!label && sublist) {
      traverse(sublist, level, rows);
    }
  }
}

// Основна функція експорту
function exportTreeToExcel() {
  const treeRoot = document.querySelector("ul");

  const maxDepth = getMaxDepth(treeRoot);

  const header = [];
  for (let i = 1; i <= maxDepth; i++) {
    header.push(`Рівень ${i}`);
  }

  const rows = [header];

  traverse(treeRoot, [], rows);

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Каталог");
  XLSX.writeFile(workbook, "Каталог_товарів.xlsx");
}


// ------------------------------
// ІМПОРТ EXCEL → HTML ДЕРЕВА
// ------------------------------
async function importTreeFromExcel() {
  const filePath = "source/Каталог_товарів.xls";

  const response = await fetch(filePath);
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  rows.shift(); // прибираємо заголовки

  const tree = {};

  for (const row of rows) {
    let current = tree;

    for (const cell of row) {
      if (!cell) break;

      if (!current[cell]) {
        current[cell] = {};
      }
      current = current[cell];
    }
  }

  const root = document.getElementById("importedTree");
  root.innerHTML = renderTree(tree);

  activateCarets();
}

// Рекурсивний рендер дерева
function renderTree(node, isRoot = true) {
  let html = isRoot ? "<ul>" : '<ul class="nested">';

  for (const key in node) {
    const hasChildren = Object.keys(node[key]).length > 0;

    if (hasChildren) {
      html += `
        <li>
          <span class="caret">${key}</span>
          ${renderTree(node[key], false)}
        </li>`;
    } else {
      html += `<li>${key}</li>`;
    }
  }

  html += "</ul>";
  return html;
}

function activateItemHighlighting() {
  const tree = document.getElementById("importedTree");

  tree.addEventListener("click", (event) => {
    const li = event.target.closest("li");
    if (!li) return;

    // Убираем подсветку со всех элементов
    tree.querySelectorAll(".active-item").forEach(el => {
      el.classList.remove("active-item");
    });

    // Добавляем подсветку выбранному
    li.classList.add("active-item");
  });
}

// в самом низу main.js
window.exportTreeToExcel = exportTreeToExcel;
window.importTreeFromExcel = importTreeFromExcel;


