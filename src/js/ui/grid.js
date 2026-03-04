import {
  setSelection,
  getSelection,
  addSelection,
  getSelections,
  getSheetSelections,
  clearSelections,
  clearSheetSelections,
  getWorkbooks,
  getActiveWorkbook,
  setActiveWorkbook,
  addWorkbook,
  deleteWorkbook,
  getSheets,
  getActiveSheet,
  setActiveSheet,
  addSheet,
  deleteSheet,
  deleteRows,
  deleteColumns,
  getCell,
  setCell,
  appendRows,
  appendCols,
  initState,
} from "../services/state.js";

// Excel默认行列数
const EXCEL_ROWS = 1048576;
const EXCEL_COLS = 16384;

// 可见区域配置
const VISIBLE_ROWS = 50;
const VISIBLE_COLS = 20;

// 单元格尺寸
const COL_WIDTH = 80;
const ROW_HEIGHT = 22;
const HEADER_WIDTH = 50;

export function createGrid(containerEl, tabsEl, state) {
  createWorkbookTabs(state);
  createSheetTabs(tabsEl, containerEl, state);
  
  // 添加滚动事件监听
  const wrapper = containerEl.closest(".sheet-grid");
  if (wrapper) {
    wrapper.addEventListener("scroll", () => {
      if (!isScrolling) {
        isScrolling = true;
        requestAnimationFrame(() => {
          onGridScroll(wrapper, containerEl, state);
          isScrolling = false;
        });
      }
    });
  }
  
  // 初始渲染
  renderGrid(containerEl, state);
}

function createWorkbookTabs(state) {
  const tabsEl = document.getElementById("workbook-tabs");
  if (!tabsEl) return;
  tabsEl.innerHTML = "";

  const workbooks = getWorkbooks(state);
  const activeWb = getActiveWorkbook(state);

  workbooks.forEach((wb) => {
    const tab = document.createElement("button");
    tab.className = "workbook-tab";
    if (activeWb && wb.id === activeWb.id) {
      tab.classList.add("workbook-tab--active");
    }
    const tabContent = document.createElement("span");
    tabContent.textContent = wb.name || "工作簿";
    tab.appendChild(tabContent);

    const deleteBtn = document.createElement("span");
    deleteBtn.className = "tab-delete-btn";
    deleteBtn.textContent = "×";
    deleteBtn.title = "删除工作簿";
    deleteBtn.addEventListener("click", (e) => {
      e.stopPropagation(); // 防止触发父元素的点击事件
      if (confirm(`确定要删除工作簿 “${wb.name}” 吗？此操作不可撤销。`)) {
        deleteWorkbook(state, wb.id);
        // 重新渲染整个UI
        const gridRoot = document.getElementById("sheet-grid");
        const sheetTabs = document.getElementById("sheet-tabs");
        if(gridRoot && sheetTabs) {
            createGrid(gridRoot, sheetTabs, state);
        }
      }
    });
    tab.appendChild(deleteBtn);

    tab.addEventListener("click", () => {
      setActiveWorkbook(state, wb.id);
      createWorkbookTabs(state);
      const gridRoot = document.getElementById("sheet-grid");
      const sheetTabs = document.getElementById("sheet-tabs");
      if (gridRoot && sheetTabs) {
        createSheetTabs(sheetTabs, gridRoot, state);
        renderGrid(gridRoot, state);
      }
    });
    tabsEl.appendChild(tab);
  });

  const addBtn = document.createElement("button");
  addBtn.className = "workbook-tab workbook-tab--add";
  addBtn.textContent = "+";
  addBtn.title = "新建工作簿";
  addBtn.addEventListener("click", () => {
    const index = workbooks.length + 1;
    const sheet = {
      id: `sheet-1`,
      name: "Sheet1",
      rows: EXCEL_ROWS,
      cols: EXCEL_COLS,
      data: [],
      merges: [],
      headers: [],
    };
    const wb = {
      id: `workbook-${index}`,
      name: `工作簿${index}`,
      sheets: [sheet],
      activeSheetId: sheet.id,
    };
    addWorkbook(state, wb);
    createWorkbookTabs(state);
    const gridRoot = document.getElementById("sheet-grid");
    const sheetTabs = document.getElementById("sheet-tabs");
    if (gridRoot && sheetTabs) {
      createSheetTabs(sheetTabs, gridRoot, state);
      renderGrid(gridRoot, state);
    }
  });

  tabsEl.appendChild(addBtn);
}

function createSheetTabs(tabsEl, containerEl, state) {
  tabsEl.innerHTML = "";

  const sheets = getSheets(state);
  const activeSheet = getActiveSheet(state);

  const tabsWrapper = document.createElement("div");
  tabsWrapper.className = "sheet-tabs-list";

  sheets.forEach((sheet) => {
    const tab = document.createElement("button");
    tab.className = "sheet-tab";
    if (activeSheet && sheet.id === activeSheet.id) {
      tab.classList.add("sheet-tab--active");
    }
    const tabContent = document.createElement("span");
    tabContent.textContent = sheet.name;
    tab.appendChild(tabContent);

    const deleteBtn = document.createElement("span");
    deleteBtn.className = "tab-delete-btn";
    deleteBtn.textContent = "×";
    deleteBtn.title = "删除工作表";
    deleteBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (confirm(`确定要删除工作表 “${sheet.name}” 吗？`)) {
        deleteSheet(state, sheet.id);
        createGrid(containerEl, tabsEl, state);
      }
    });
    tab.appendChild(deleteBtn);

    tab.addEventListener("click", () => {
      // 保存当前选择
      const currentSelection = getSelection(state);
      if (currentSelection) {
        addSelection(state, currentSelection);
      }
      
      setActiveSheet(state, sheet.id);
      createSheetTabs(tabsEl, containerEl, state);
      renderGrid(containerEl, state);
      
      // 恢复该工作表的选择
      const selections = getSelections(state);
      const sheetSelection = selections.find(s => s.sheetId === sheet.id);
      if (sheetSelection) {
        setSelection(state, sheetSelection);
        const table = document.querySelector(".grid-table");
        if (table) {
          updateSelectionStyles(table, sheetSelection, state);
          updateStatusSelection(sheetSelection, state);
        }
      }
    });
    tabsWrapper.appendChild(tab);
  });

  const addBtn = document.createElement("button");
  addBtn.className = "sheet-tab sheet-tab--add";
  addBtn.textContent = "+";
  addBtn.title = "新建工作表";
  addBtn.addEventListener("click", () => {
    const index = sheets.length + 1;
    const newSheet = {
      id: `sheet-${index}`,
      name: `Sheet${index}`,
      rows: EXCEL_ROWS,
      cols: EXCEL_COLS,
      data: [],
    };
    addSheet(state, newSheet);
    createSheetTabs(tabsEl, containerEl, state);
    renderGrid(containerEl, state);
  });

  tabsEl.appendChild(tabsWrapper);
  tabsEl.appendChild(addBtn);
}

// 防止滚动事件触发过于频繁
let isScrolling = false;

function renderGrid(containerEl, state, isScrolling = false) {
  const activeSheet = getActiveSheet(state);
  if (!activeSheet) {
    containerEl.innerHTML = "";
    return;
  }

  // 确保 activeSheet.headers 是一个数组
  if (!Array.isArray(activeSheet.headers)) {
    activeSheet.headers = [];
  }

  const mergeInfo = buildMergeMaps(activeSheet.merges || []);

  // 计算总宽度和高度
  const totalWidth = HEADER_WIDTH + (activeSheet.cols * COL_WIDTH);
  const totalHeight = activeSheet.rows * ROW_HEIGHT;

  // 获取滚动位置，计算可见区域
  const wrapper = containerEl.closest(".sheet-grid");
  const scrollLeft = wrapper ? wrapper.scrollLeft : 0;
  const scrollTop = wrapper ? wrapper.scrollTop : 0;

  // 计算可见区域的起始和结束索引
  const startCol = Math.max(0, Math.floor(scrollLeft / COL_WIDTH) - 1);
  const endCol = Math.min(activeSheet.cols, startCol + Math.ceil(wrapper?.clientWidth / COL_WIDTH) + 2);
  const startRow = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - 1);
  const endRow = Math.min(activeSheet.rows, startRow + Math.ceil(wrapper?.clientHeight / ROW_HEIGHT) + 2);

  let table;
  let thead;
  let tbody;

  if (isScrolling) {
    // 滚动时，找到现有的table, thead, tbody并清空内容
    table = containerEl.querySelector(".grid-table");
    if (!table) {
      // 如果table不存在，回退到完整渲染
      renderGrid(containerEl, state, false);
      return;
    }
    thead = table.querySelector("thead");
    tbody = table.querySelector("tbody");
    if (thead) thead.innerHTML = "";
    if (tbody) tbody.innerHTML = "";
    else {
        tbody = document.createElement("tbody");
        table.appendChild(tbody);
    }

  } else {
    // 非滚动时（初次渲染），创建所有元素
    containerEl.innerHTML = "";
    const tableContainer = document.createElement("div");
    tableContainer.style.width = `${totalWidth}px`;
    tableContainer.style.height = `${totalHeight}px`;
    tableContainer.style.position = "relative";
    
    table = document.createElement("table");
    table.className = "grid-table";
    table.style.tableLayout = "fixed";
    table.style.position = "absolute";

    thead = document.createElement("thead");
    tbody = document.createElement("tbody");

    table.appendChild(thead);
    table.appendChild(tbody);
    tableContainer.appendChild(table);
    containerEl.appendChild(tableContainer);
  }

  // --- 重新渲染表头和内容 ---

  // 创建表头
  const headerRow = document.createElement("tr");

  // 左上角空白
  const corner = document.createElement("th");
  corner.className = "grid-header";
  corner.style.width = `${HEADER_WIDTH}px`;
  headerRow.appendChild(corner);

  // 渲染列标题（根据滚动位置）
  for (let c = startCol; c < endCol; c++) {
    const th = document.createElement("th");
    th.className = "grid-header";
    th.textContent = columnLabel(c);
    th.style.width = `${COL_WIDTH}px`;
    addColResizeHandle(th, c, table);
    // 添加列头点击事件，用于选择整列
    th.addEventListener("mousedown", (e) =>
      onHeaderMouseDown(e, "col", c, table, state),
    );
    th.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        showContextMenu(e, "col", c, state);
    });
    headerRow.appendChild(th);
  }

  thead.appendChild(headerRow);

  // 渲染行（根据滚动位置）
  for (let r = startRow; r < endRow; r++) {
    const tr = document.createElement("tr");

    // 行标题
    const rowHeader = document.createElement("th");
    rowHeader.className = "grid-header";
    rowHeader.textContent = r + 1;
    rowHeader.style.width = `${HEADER_WIDTH}px`;
    addRowResizeHandle(rowHeader, r, table);
    // 添加行头点击事件，用于选择整行
    rowHeader.addEventListener("mousedown", (e) =>
      onHeaderMouseDown(e, "row", r, table, state),
    );
    rowHeader.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        showContextMenu(e, "row", r, state);
    });
    tr.appendChild(rowHeader);

    // 渲染单元格
    for (let c = startCol; c < endCol; c++) {
      const coordKey = `${r},${c}`;
      if (mergeInfo.covered.has(coordKey)) {
        continue;
      }
      const td = document.createElement("td");
      td.className = "cell";
      // 检查是否需要高亮
      const header = activeSheet.headers[c];
      if (header && header.style && header.style.highlight) {
        td.classList.add("highlight-new-column");
      }
      td.dataset.row = r;
      td.dataset.col = c;
      td.style.width = `${COL_WIDTH}px`;
      
      // 处理合并单元格
      const mergeMeta = mergeInfo.tops.get(coordKey);
      if (mergeMeta) {
        if (mergeMeta.rowspan > 1) td.rowSpan = mergeMeta.rowspan;
        if (mergeMeta.colspan > 1) td.colSpan = mergeMeta.colspan;
      }
      
      const value = getCell(state, activeSheet.id, r, c);
      if (value != null) {
        td.textContent = value;
      }
      
      td.addEventListener("dblclick", () => {
        td.contentEditable = "true";
        td.focus();
      });

      td.addEventListener("blur", () => {
        td.contentEditable = "false";
        const text = td.textContent || "";
        setCell(state, activeSheet.id, r, c, text);
      });

      td.addEventListener("mousedown", (e) =>
        onCellMouseDown(e, td, table, state),
      );
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  // 恢复选择样式
  const selection = getSelection(state);
  if (selection) {
    updateSelectionStyles(table, selection, state);
  }

  if (table) {
    table.style.top = `${startRow * ROW_HEIGHT}px`;
  }

  // 更新状态栏
  updateStatusBar(startRow, endRow, startCol, endCol, activeSheet);

  // 滚动事件监听已经在初始化时添加，这里不需要重复添加
}

function onGridScroll(wrapperEl, containerEl, state) {
  // 滚动时重新渲染可见区域
  renderGrid(containerEl, state, true);
}

function addColResizeHandle(th, colIndex, table) {
  const handle = document.createElement("div");
  handle.className = "col-resize-handle";
  let startX = 0;
  let startWidth = 0;

  const onMouseMove = (e) => {
    const delta = e.clientX - startX;
    const newWidth = Math.max(40, startWidth + delta);
    const headerCells = table.querySelectorAll(
      `.grid-header:nth-child(${colIndex + 2})`,
    );
    headerCells.forEach((cell) => {
      cell.style.width = `${newWidth}px`;
    });
    const bodyCells = table.querySelectorAll(
      `td[data-col="${colIndex}"]`,
    );
    bodyCells.forEach((cell) => {
      cell.style.width = `${newWidth}px`;
    });
  };

  const onMouseUp = () => {
    document.removeEventListener("mousemove", onMouseMove);
    document.removeEventListener("mouseup", onMouseUp);
  };

  handle.addEventListener("mousedown", (e) => {
    e.preventDefault();
    startX = e.clientX;
    startWidth = th.getBoundingClientRect().width;
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
  });

  th.appendChild(handle);
}

function addRowResizeHandle(th, rowIndex, table) {
  const handle = document.createElement("div");
  handle.className = "row-resize-handle";
  let startY = 0;
  let startHeight = 0;

  const onMouseMove = (e) => {
    const delta = e.clientY - startY;
    const newHeight = Math.max(18, startHeight + delta);
    const headerCell = th;
    headerCell.style.height = `${newHeight}px`;
    headerCell.style.lineHeight = `${newHeight}px`;
    const row = table.querySelector(`tr:nth-child(${rowIndex + 2})`);
    if (row) {
      row.querySelectorAll("td").forEach((cell) => {
        cell.style.height = `${newHeight}px`;
        cell.style.lineHeight = `${newHeight}px`;
      });
    }
  };

  const onMouseUp = () => {
    document.removeEventListener("mousemove", onMouseMove);
    document.removeEventListener("mouseup", onMouseUp);
  };

  handle.addEventListener("mousedown", (e) => {
    e.preventDefault();
    startY = e.clientY;
    startHeight = th.getBoundingClientRect().height;
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
  });

  th.appendChild(handle);
}

function buildMergeMaps(merges) {
  const tops = new Map();
  const covered = new Set();
  if (!Array.isArray(merges)) {
    return { tops, covered };
  }
  merges.forEach((m) => {
    if (!m || !m.s || !m.e) return;
    const startRow = m.s.r;
    const startCol = m.s.c;
    const endRow = m.e.r;
    const endCol = m.e.c;
    const rowspan = endRow - startRow + 1;
    const colspan = endCol - startCol + 1;
    const topKey = `${startRow},${startCol}`;
    tops.set(topKey, { rowspan, colspan });
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const key = `${r},${c}`;
        if (key === topKey) continue;
        covered.add(key);
      }
    }
  });
  return { tops, covered };
}

function columnLabel(index) {
  let label = "";
  let n = index;
  while (n >= 0) {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  }
  return label;
}

function onCellMouseDown(event, cellEl, table, state) {
  // 如果单元格已经在编辑模式，则不执行任何操作，让浏览器处理默认行为（如文本选择）
  if (cellEl.isContentEditable) {
    return;
  }
  event.preventDefault();
  const startRow = Number(cellEl.dataset.row);
  const startCol = Number(cellEl.dataset.col);
  const activeSheet = getActiveSheet(state);
  
  // 检查是否按下了shift键或ctrl键
  const isShiftPressed = event.shiftKey;
  const isCtrlPressed = event.ctrlKey || event.metaKey;
  const currentSelection = getSelection(state);
  
  let selection;
  let shouldSaveOnMouseUp = false;
  
  if (isShiftPressed && currentSelection) {
    // 如果按下了shift键，从当前选择的起点到新的终点
    selection = {
      startRow: currentSelection.startRow,
      startCol: currentSelection.startCol,
      endRow: startRow,
      endCol: startCol,
      sheetId: activeSheet.id
    };
    shouldSaveOnMouseUp = true;
  } else if (isCtrlPressed && currentSelection) {
    // 如果按下了ctrl键，保存当前选择，然后开始新的选择
    addSelection(state, currentSelection);
    selection = {
      startRow,
      startCol,
      endRow: startRow,
      endCol: startCol,
      sheetId: activeSheet.id
    };
    shouldSaveOnMouseUp = true;
  } else {
    // 普通选择，清除当前工作表的所有选择
    clearSheetSelections(state, activeSheet.id);
    selection = {
      startRow,
      startCol,
      endRow: startRow,
      endCol: startCol,
      sheetId: activeSheet.id
    };
    shouldSaveOnMouseUp = true;
  }

  setSelection(state, selection);
  console.log("单元格选择已设置:", selection, "state.selection:", state.selection);
  updateSelectionStyles(table, selection, state);
  updateStatusSelection(selection, state);

  const onMouseMove = (moveEvent) => {
    const target = moveEvent.target;
    if (!target || !target.classList.contains("cell")) return;
    const endRow = Number(target.dataset.row);
    const endCol = Number(target.dataset.col);
    selection.endRow = endRow;
    selection.endCol = endCol;
    setSelection(state, selection);
    updateSelectionStyles(table, selection, state);
    updateStatusSelection(selection, state);
  };

  const onMouseUp = () => {
    // 鼠标释放时，保存当前选择区域
    if (shouldSaveOnMouseUp) {
      addSelection(state, selection);
    }
    document.removeEventListener("mousemove", onMouseMove);
    document.removeEventListener("mouseup", onMouseUp);
  };

  document.addEventListener("mousemove", onMouseMove);
  document.addEventListener("mouseup", onMouseUp);
}

function onHeaderMouseDown(event, type, index, table, state) {
  event.preventDefault();
  const activeSheet = getActiveSheet(state);
  
  // 检查是否按下了ctrl键
  const isCtrlPressed = event.ctrlKey || event.metaKey;
  const currentSelection = getSelection(state);
  
  let selection;
  if (type === "col") {
    // 选择整列
    selection = {
      startRow: 0,
      startCol: index,
      endRow: activeSheet.rows - 1,
      endCol: index,
      sheetId: activeSheet.id
    };
  } else {
    // 选择整行
    selection = {
      startRow: index,
      startCol: 0,
      endRow: index,
      endCol: activeSheet.cols - 1,
      sheetId: activeSheet.id
    };
  }
  
  // 如果按下了ctrl键，保存当前选择
  if (isCtrlPressed && currentSelection) {
    addSelection(state, currentSelection);
  } else {
    // 普通选择，清除当前工作表的所有选择
    clearSheetSelections(state, activeSheet.id);
  }
  
  setSelection(state, selection);
  updateSelectionStyles(table, selection, state);
  updateStatusSelection(selection, state);
  
  // 鼠标释放时，保存当前选择区域
  const onMouseUp = () => {
    addSelection(state, selection);
    document.removeEventListener("mouseup", onMouseUp);
  };
  document.addEventListener("mouseup", onMouseUp);
}

export function clearSelectionUI() {
  const table = document.querySelector(".grid-table");
  if (table) {
    updateSelectionStyles(table, null);
  }
  const state = initState();
  updateStatusSelection(null, state);
  // 清除所有选择
  clearSelections(state);
}

function updateSelectionStyles(table, selection, state) {
  const cells = table.querySelectorAll(".cell");
  cells.forEach((cell) => {
    cell.classList.remove("cell--selected", "cell--in-range");
  });

  if (!selection) return;

  const { startRow, startCol, endRow, endCol, sheetId } = selection;
  const minRow = Math.min(startRow, endRow);
  const maxRow = Math.max(startRow, endRow);
  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);

  cells.forEach((cell) => {
    const r = Number(cell.dataset.row);
    const c = Number(cell.dataset.col);
    if (r >= minRow && r <= maxRow && c >= minCol && c <= maxCol) {
      cell.classList.add("cell--in-range");
    }
  });

  // 获取当前工作表的所有选择区域
  const sheetSelections = getSheetSelections(state, sheetId);
  sheetSelections.forEach((sel, index) => {
    const { startRow: sRow, startCol: sCol, endRow: eRow, endCol: eCol } = sel;
    const minR = Math.min(sRow, eRow);
    const maxR = Math.max(sRow, eRow);
    const minC = Math.min(sCol, eCol);
    const maxC = Math.max(sCol, eCol);
    
    cells.forEach((cell) => {
      const r = Number(cell.dataset.row);
      const c = Number(cell.dataset.col);
      if (r === sRow && c === sCol && index === sheetSelections.length - 1) {
        cell.classList.add("cell--selected");
      } else if (r >= minR && r <= maxR && c >= minC && c <= maxC) {
        cell.classList.add("cell--in-range");
      }
    });
  });
}

function updateStatusSelection(selection, state) {
  const statusEl = document.getElementById("status-selection");
  if (!statusEl) return;

  // 获取所有选择区域
  const allSelections = getSelections(state);
  
  if (allSelections.length === 0) {
    statusEl.textContent = "未选择单元格";
    return;
  }
  
  // 按工作表分组
  const selectionsBySheet = {};
  allSelections.forEach(sel => {
    if (!selectionsBySheet[sel.sheetId]) {
      selectionsBySheet[sel.sheetId] = [];
    }
    selectionsBySheet[sel.sheetId].push(sel);
  });
  
  // 获取工作表名称映射
  const sheets = getSheets(state);
  const sheetNameMap = {};
  sheets.forEach(sheet => {
    sheetNameMap[sheet.id] = sheet.name;
  });
  
  // 生成状态栏文本
  const statusTexts = [];
  for (const sheetId in selectionsBySheet) {
    const sheetSelections = selectionsBySheet[sheetId];
    const sheetName = sheetNameMap[sheetId] || `表${Object.keys(selectionsBySheet).indexOf(sheetId) + 1}`;
    
    sheetSelections.forEach(sel => {
      const minRow = Math.min(sel.startRow, sel.endRow) + 1;
      const maxRow = Math.max(sel.startRow, sel.endRow) + 1;
      const minCol = columnLabel(Math.min(sel.startCol, sel.endCol));
      const maxCol = columnLabel(Math.max(sel.startCol, sel.endCol));
      
      if (minRow === maxRow && minCol === maxCol) {
        statusTexts.push(`${sheetName}：${minCol}${minRow}`);
      } else {
        statusTexts.push(`${sheetName}：${minCol}${minRow}-${maxCol}${maxRow}`);
      }
    });
  }
  
  statusEl.textContent = statusTexts.join("，");
}

function showContextMenu(event, type, index, state) {
    event.preventDefault();
    const menu = document.createElement("div");
    menu.className = "context-menu";
    menu.style.position = "absolute";
    menu.style.left = `${event.clientX}px`;
    menu.style.top = `${event.clientY}px`;

    const deleteOption = document.createElement("div");
    deleteOption.className = "context-menu-item";
    deleteOption.textContent = type === "row" ? "删除行" : "删除列";
    
    deleteOption.addEventListener("click", () => {
        const activeSheet = getActiveSheet(state);
        if (!activeSheet) return;

        if (type === "row") {
            if (confirm(`确定要删除第 ${index + 1} 行吗？`)) {
                deleteRows(state, activeSheet.id, index, 1);
            }
        } else {
            if (confirm(`确定要删除 ${columnLabel(index)} 列吗？`)) {
                deleteColumns(state, activeSheet.id, index, 1);
            }
        }
        
        const gridRoot = document.getElementById("sheet-grid");
        const sheetTabs = document.getElementById("sheet-tabs");
        if(gridRoot && sheetTabs) {
            createGrid(gridRoot, sheetTabs, state);
        }

        document.body.removeChild(menu);
    });

    menu.appendChild(deleteOption);
    document.body.appendChild(menu);

    const clickOutsideHandler = (e) => {
        if (!menu.contains(e.target)) {
            document.body.removeChild(menu);
            document.removeEventListener("click", clickOutsideHandler);
        }
    };
    document.addEventListener("click", clickOutsideHandler);
}

function updateStatusBar(startRow, endRow, startCol, endCol, activeSheet) {
  let statusBar = document.getElementById("status-bar");
  if (!statusBar) {
    statusBar = document.createElement("div");
    statusBar.id = "status-bar";
    statusBar.className = "status-bar";
    const sheetContainer = document.querySelector(".sheet-container");
    if (sheetContainer) {
        sheetContainer.appendChild(statusBar);
    }
  }

  const visibleRows = `当前显示: ${startRow + 1} - ${endRow} 行`;
  const totalRows = `总计: ${activeSheet.rows} 行`;
  const visibleCols = `${columnLabel(startCol)} - ${columnLabel(endCol - 1)} 列`;

  statusBar.innerHTML = `
    <div class="status-bar-item">${totalRows}</div>
    <div class="status-bar-item">${visibleRows}</div>
    <div class="status-bar-item">${visibleCols}</div>
    <div id="status-selection" class="status-bar-item status-bar-item--selection">未选择单元格</div>
  `;
}