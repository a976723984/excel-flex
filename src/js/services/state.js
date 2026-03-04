const LOCAL_STORAGE_KEY = "smart-excel-ai-config";

const globalState = {
  // 当前选中区域
  selection: null,
  // 跨表选择（每个工作表可以有多个选择区域）
  selections: {},
  // 配置（API 等）
  config: loadConfigFromStorage(),
  // 多工作簿数据
  workbooks: [],
  activeWorkbookId: null,
};

// Excel默认行列数
const EXCEL_ROWS = 1048576;
const EXCEL_COLS = 16384;

export function initState() {
  if (!globalState.workbooks.length) {
    const defaultSheet = createEmptySheet(1, "Sheet1", EXCEL_ROWS, EXCEL_COLS);
    const defaultWorkbook = {
      id: "workbook-1",
      name: "工作簿1",
      sheets: [defaultSheet],
      activeSheetId: defaultSheet.id,
    };
    globalState.workbooks = [defaultWorkbook];
    globalState.activeWorkbookId = defaultWorkbook.id;
  }
  return globalState;
}

// -------- 选择区域 --------

export function setSelection(state, selection) {
  state.selection = selection ? { ...selection } : null;
}

export function getSelection(state) {
  return state.selection;
}

// 添加跨表选择支持
export function addSelection(state, selection) {
  if (!state.selections) {
    state.selections = {};
  }
  const sheetId = selection.sheetId;
  if (!state.selections[sheetId]) {
    state.selections[sheetId] = [];
  }
  // 添加新选择区域
  state.selections[sheetId].push({ ...selection });
}

export function getSelections(state) {
  const selections = state.selections || {};
  const allSelections = [];
  for (const sheetId in selections) {
    allSelections.push(...selections[sheetId]);
  }
  return allSelections;
}

export function getSheetSelections(state, sheetId) {
  const selections = state.selections || {};
  return selections[sheetId] || [];
}

export function clearSelections(state) {
  state.selections = {};
  state.selection = null;
}

export function clearSheetSelections(state, sheetId) {
  const selections = state.selections || {};
  delete selections[sheetId];
}

// -------- 配置 --------

export function getConfig(state) {
  return state.config;
}

export function setConfig(state, partialConfig) {
  state.config = {
    ...state.config,
    ...partialConfig,
  };
  saveConfigToStorage(state.config);
}

// -------- Workbook & Sheet 管理 --------

export function getWorkbooks(state) {
  return state.workbooks;
}

export function getActiveWorkbook(state) {
  const id = state.activeWorkbookId;
  return state.workbooks.find((w) => w.id === id) || null;
}

export function setActiveWorkbook(state, workbookId) {
  if (!state.workbooks.find((w) => w.id === workbookId)) return;
  state.activeWorkbookId = workbookId;
}

export function addWorkbook(state, workbook) {
  state.workbooks.push(workbook);
  state.activeWorkbookId = workbook.id;
}

export function getSheets(state) {
  const wb = getActiveWorkbook(state);
  return wb ? wb.sheets : [];
}

export function getActiveSheet(state) {
  const wb = getActiveWorkbook(state);
  if (!wb) return null;
  const id = wb.activeSheetId;
  return wb.sheets.find((s) => s.id === id) || null;
}

export function setActiveSheet(state, sheetId) {
  const wb = getActiveWorkbook(state);
  if (!wb) return;
  if (!wb.sheets.find((s) => s.id === sheetId)) return;
  wb.activeSheetId = sheetId;
}

export function addSheet(state, sheet) {
  const wb = getActiveWorkbook(state);
  if (!wb) return;
  wb.sheets.push(sheet);
  wb.activeSheetId = sheet.id;
}

export function deleteWorkbook(state, workbookId) {
  const index = state.workbooks.findIndex(wb => wb.id === workbookId);
  if (index === -1) return;

  state.workbooks.splice(index, 1);

  if (state.activeWorkbookId === workbookId) {
    state.activeWorkbookId = state.workbooks[0]?.id || null;
  }
}

export function deleteSheet(state, sheetId) {
  const wb = getActiveWorkbook(state);
  if (!wb) return;

  const index = wb.sheets.findIndex(s => s.id === sheetId);
  if (index === -1) return;

  wb.sheets.splice(index, 1);

  if (wb.activeSheetId === sheetId) {
    wb.activeSheetId = wb.sheets[0]?.id || null;
  }
}

export function replaceSheetsForActiveWorkbook(state, sheets, activeSheetId) {
  const wb = getActiveWorkbook(state);
  if (!wb) return;
  wb.sheets = sheets.slice();
  wb.activeSheetId =
    activeSheetId && sheets.find((s) => s.id === activeSheetId)
      ? activeSheetId
      : sheets[0]?.id || null;
}

export function getAllSheetHeaders(state) {
  const workbooks = state.workbooks || [];
  
  const result = [];
  workbooks.forEach(workbook => {
    const sheets = workbook.sheets || [];
    sheets.forEach(sheet => {
      const headers = sheet.headers || [];
      result.push({
        workbookId: workbook.id,
        workbookName: workbook.name,
        sheetId: sheet.id,
        sheetName: sheet.name,
        columns: headers
      });
    });
  });
  return result;
}

export function getWorkbooksWithSheets(state) {
  const workbooks = state.workbooks || [];
  
  return workbooks.map(workbook => {
    const sheets = workbook.sheets || [];
    return {
      id: workbook.id,
      name: workbook.name,
      sheets: sheets.map(sheet => ({
        id: sheet.id,
        name: sheet.name,
        headers: sheet.headers || []
      }))
    };
  });
}

export function getSheetById(state, sheetId) {
  const workbooks = state.workbooks || [];
  
  for (const workbook of workbooks) {
    const sheets = workbook.sheets || [];
    const sheet = sheets.find(s => s.id === sheetId);
    if (sheet) {
      return sheet;
    }
  }
  
  return null;
}

// -------- 单元格读写 --------

export function getCell(state, sheetId, row, col) {
  const wb = getActiveWorkbook(state);
  if (!wb) return undefined;
  const sheet =
    wb.sheets.find((s) => s.id === sheetId) ||
    wb.sheets.find((s) => s.id === wb.activeSheetId);
  if (!sheet) return undefined;
  if (!sheet.data || !sheet.data[row]) return undefined;
  return sheet.data[row][col];
}

export function setCell(state, sheetId, row, col, value) {
  const wb = getActiveWorkbook(state);
  if (!wb) return;
  const sheet =
    wb.sheets.find((s) => s.id === sheetId) ||
    wb.sheets.find((s) => s.id === wb.activeSheetId);
  if (!sheet) return;

  if (!sheet.data) {
    sheet.data = [];
  }
  // 按需扩容二维数组
  while (sheet.data.length <= row) {
    sheet.data.push([]);
  }
  const rowArr = sheet.data[row];
  while (rowArr.length <= col) {
    rowArr.push("");
  }
  rowArr[col] = value;
}

export function appendRows(state, sheetId, count) {
  const wb = getActiveWorkbook(state);
  if (!wb) return;
  const sheet =
    wb.sheets.find((s) => s.id === sheetId) ||
    wb.sheets.find((s) => s.id === wb.activeSheetId);
  if (!sheet) return;
  sheet.rows += count;
}

export function appendCols(state, sheetId, count) {
  const wb = getActiveWorkbook(state);
  if (!wb) return;
  const sheet =
    wb.sheets.find((s) => s.id === sheetId) ||
    wb.sheets.find((s) => s.id === wb.activeSheetId);
  if (!sheet) return;
  sheet.cols += count;
}

export function ensureCols(state, sheetId, cols) {
  const wb = getActiveWorkbook(state);
  if (!wb) return;
  const sheet =
    wb.sheets.find((s) => s.id === sheetId) ||
    wb.sheets.find((s) => s.id === wb.activeSheetId);
  if (!sheet) return;
  if (cols > sheet.cols) {
    sheet.cols = cols;
  }
}

// -------- 内部工具 --------

function createEmptySheet(index, name, rows, cols) {
  return {
    id: `sheet-${index}`,
    name,
    rows,
    cols,
    data: [],
    merges: [],
  };
}

function loadConfigFromStorage() {
  try {
    const raw = window.localStorage.getItem(LOCAL_STORAGE_KEY);
    if (!raw) return {};
    return JSON.parse(raw) || {};
  } catch (e) {
    console.warn("Failed to load config from storage", e);
    return {};
  }
}

function saveConfigToStorage(config) {
  try {
    window.localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(config));
  } catch (e) {
    console.warn("Failed to save config to storage", e);
  }
}

export function deleteRows(state, sheetId, startRow, count = 1) {
  const sheet = getSheetById(state, sheetId);
  if (!sheet || !sheet.data) return;

  sheet.data.splice(startRow, count);
  sheet.rows -= count;
  // TODO: Adjust merges if necessary
}

export function deleteColumns(state, sheetId, startCol, count = 1) {
  const sheet = getSheetById(state, sheetId);
  if (!sheet || !sheet.data) return;

  for (let row = 0; row < sheet.data.length; row++) {
    if (sheet.data[row]) {
      sheet.data[row].splice(startCol, count);
    }
  }

  sheet.cols -= count;
  // TODO: Adjust merges if necessary
}
