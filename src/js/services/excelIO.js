const XLSX_GLOBAL = window.XLSX;

function columnLabel(index) {
  let label = "";
  let n = index;
  while (n >= 0) {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  }
  return label;
}

export function parseXlsxToSheets(file) {
  return new Promise((resolve, reject) => {
    if (!XLSX_GLOBAL) {
      reject(new Error("未找到 XLSX 库，请检查页面是否已正确引入。"));
      return;
    }
    const reader = new FileReader();
    reader.onerror = () => {
      reject(new Error("读取文件失败"));
    };
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX_GLOBAL.read(data, { type: "array" });
        const sheets = workbook.SheetNames.map((sheetName, index) => {
          const worksheet = workbook.Sheets[sheetName];
          const aoa = XLSX_GLOBAL.utils.sheet_to_json(worksheet, {
            header: 1,
            raw: false,
          });
          const rows = aoa.length;
          let cols = 0;
          for (let r = 0; r < rows; r++) {
            const row = aoa[r] || [];
            let lastNonEmpty = -1;
            for (let c = 0; c < row.length; c++) {
              if (row[c] != null && row[c] !== "") {
                lastNonEmpty = c;
              }
            }
            if (lastNonEmpty + 1 > cols) {
              cols = lastNonEmpty + 1;
            }
          }
          const dataMatrix = [];
          for (let r = 0; r < rows; r++) {
            const srcRow = aoa[r] || [];
            const destRow = [];
            for (let c = 0; c < cols; c++) {
              destRow[c] = srcRow[c] != null ? String(srcRow[c]) : "";
            }
            dataMatrix.push(destRow);
          }
          const merges = worksheet["!merges"] || [];
          
          const headers = [];
          if (dataMatrix.length > 0) {
            const firstRow = dataMatrix[0];
            for (let c = 0; c < cols; c++) {
              const headerValue = firstRow[c] || `列${c + 1}`;
              headers.push({
                colIndex: c,
                name: headerValue,
                colLabel: columnLabel(c)
              });
            }
          }
          
          return {
            id: `sheet-import-${Date.now()}-${index}`,
            name: sheetName || `Sheet${index + 1}`,
            rows: rows || 1,
            cols: cols || 1,
            data: dataMatrix,
            merges,
            headers,
          };
        });
        
        // 使用文件名作为工作簿名称
        const workbookName = file.name.replace(/\.(xlsx|xls|csv)$/i, "");
        
        resolve({ workbookName, sheets });
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

export function exportSheetsToXlsx(sheets, filename = "excel-flex.xlsx") {
  if (!XLSX_GLOBAL) {
    alert("未找到 XLSX 库，请检查页面是否已正确引入。");
    return;
  }
  const wb = XLSX_GLOBAL.utils.book_new();
  sheets.forEach((sheet) => {
    const rows = sheet.rows || 0;
    const cols = sheet.cols || 0;
    const aoa = [];
    for (let r = 0; r < rows; r++) {
      const rowArr = [];
      const srcRow = sheet.data && sheet.data[r] ? sheet.data[r] : [];
      for (let c = 0; c < cols; c++) {
        rowArr[c] = srcRow[c] != null ? srcRow[c] : "";
      }
      aoa.push(rowArr);
    }
    const ws = XLSX_GLOBAL.utils.aoa_to_sheet(aoa);
    if (sheet.merges && sheet.merges.length) {
      ws["!merges"] = sheet.merges;
    }
    XLSX_GLOBAL.utils.book_append_sheet(wb, ws, sheet.name || "Sheet");
  });
  XLSX_GLOBAL.writeFile(wb, filename);
}

