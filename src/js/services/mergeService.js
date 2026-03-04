
import { getWorkbooksWithSheets, getSheetById } from './state.js';

/**
 * 获取所有工作簿及其工作表和列信息
 * @param {object} state 
 * @returns {Array}
 */
export function getSheetsWithColumns(state) {
  const workbooks = getWorkbooksWithSheets(state);
  return workbooks.map(wb => ({
    ...wb,
    sheets: wb.sheets.map(sh => {
      const sheet = getSheetById(state, sh.id);
      return {
        ...sh,
        columns: sheet?.headers || [],
      };
    }),
  }));
}

/**
 * 合并多个工作表
 * @param {object} state - 全局状态
 * @param {Array<{sheetId: string, columnIndex: number}>} selections - 要合并的表和列
 * @returns {object} - 包含合并后的数据和表头的新 sheet 对象
 */
export function mergeSheets(state, selections) {
  if (!selections || selections.length < 2) {
    throw new Error("请至少选择两个要合并的工作表。");
  }

  const [baseSelection, ...restSelections] = selections;
  const baseSheet = getSheetById(state, baseSelection.sheetId);
  if (!baseSheet || !baseSheet.data) {
    throw new Error(`无法找到或读取基础工作表: ${baseSelection.sheetId}`);
  }

  const baseData = baseSheet.data.slice(1); // 假设第一行是表头
  const baseHeaders = baseSheet.headers.map(h => h.name);
  const baseColumnIndex = baseSelection.columnIndex;

  let mergedData = baseData.map(row => [...row]);
  let mergedHeaders = [...baseHeaders];

  for (const selection of restSelections) {
    const sheet = getSheetById(state, selection.sheetId);
    if (!sheet || !sheet.data) {
      throw new Error(`无法找到或读取工作表: ${selection.sheetId}`);
    }

    const dataToMerge = sheet.data.slice(1);
    const headersToMerge = sheet.headers.map(h => h.name);
    const columnIndexToMerge = selection.columnIndex;

    const newMergedData = [];
    const newHeaders = [...mergedHeaders];

    // 添加新的表头，排除合并列
    headersToMerge.forEach((header, i) => {
      if (i !== columnIndexToMerge) {
        newHeaders.push(header);
      }
    });

    for (const baseRow of mergedData) {
      const baseValue = baseRow[baseColumnIndex];
      const matchingRow = dataToMerge.find(row => row[columnIndexToMerge] === baseValue);

      if (matchingRow) {
        const newRow = [...baseRow];
        matchingRow.forEach((cell, i) => {
          if (i !== columnIndexToMerge) {
            newRow.push(cell);
          }
        });
        newMergedData.push(newRow);
      }
    }
    mergedData = newMergedData;
    mergedHeaders = newHeaders;
  }

  // 将表头添加到数据的第一行
  mergedData.unshift(mergedHeaders);

  return {
    name: "合并结果",
    data: mergedData,
    rows: mergedData.length,
    cols: mergedHeaders.length,
    headers: mergedHeaders.map((name, i) => ({
      colIndex: i,
      name,
      colLabel: columnLabel(i),
    }))
  };
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
