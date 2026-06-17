function renderTree(node, isRoot = true) {
  let html = isRoot ? "<ul>" : '<ul class="nested">';
  for (const key in node) {
    const hasChildren = Object.keys(node[key]).length > 0;
    html += hasChildren 
      ? `<li class="tree-node"><div class="caret">${key}</div>${renderTree(node[key], false)}</li>`
      : `<li class="tree-node leaf">${key}</li>`;
  }
  html += "</ul>";
  return html;
}

async function importTreeFromExcel() {
  try {
    const response = await fetch("source/Каталог_товарів.xls");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    rows.shift();

    const tree = {};
    for (const row of rows) {
      let current = tree;
      for (const cell of row) {
        if (!cell) break;
        if (!current[cell]) current[cell] = {};
        current = current[cell];
      }
    }
    document.getElementById("importedTree").innerHTML = renderTree(tree);
  } catch (e) { console.error("Помилка імпорту:", e); }
}

document.addEventListener("DOMContentLoaded", () => {
  const treeRoot = document.getElementById("importedTree");
  treeRoot.addEventListener("click", (e) => {
    const caret = e.target.closest(".caret");
    if (caret) {
      caret.nextElementSibling?.classList.toggle("active");
      caret.classList.toggle("caret-down");
    }
    const li = e.target.closest("li");
    if (li) {
      document.querySelectorAll(".active-item").forEach(el => el.classList.remove("active-item"));
      li.classList.add("active-item");
    }
  });
});
