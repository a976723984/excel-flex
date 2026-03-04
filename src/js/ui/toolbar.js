import {
  getSelection,
  getActiveWorkbook,
  addWorkbook,
  getWorkbooks,
  setActiveWorkbook,
} from "../services/state.js";
import { runAiOnSelection } from "../services/aiClient.js";
import { clearSelectionUI, createGrid } from "./grid.js";
import { parseXlsxToSheets, exportSheetsToXlsx } from "../services/excelIO.js";
import { initPromptInput, getPromptText } from "./promptInput.js";
import { showMergeModal } from "./mergeModal.js";
import { updateColumnPanel } from "./columnPanel.js";

export function initToolbar(state) {
  const btnRunAi = document.getElementById("btn-run-ai");
  const btnClearSelection = document.getElementById("btn-clear-selection");
  const btnOpenConfig = document.getElementById("btn-open-config");
  const btnImport = document.getElementById("btn-import-excel");
  const btnExport = document.getElementById("btn-export-excel");
  const btnMerge = document.getElementById("btn-merge-sheets");
  const hiddenFileInput = document.getElementById("input-import-excel");
  const btnCloseAiProcess = document.getElementById("btn-close-ai-process");

  initPromptInput(state);

  if (btnRunAi) {
    btnRunAi.addEventListener("click", async () => {
      const selection = getSelection(state);
      
      // 调试信息
      console.log("AI按钮点击 - state:", state);
      console.log("AI按钮点击 - selection:", selection);
      
      const prompt = getPromptText();
      if (!prompt) {
        alert("请输入要对选中单元格执行的 AI 指令。");
        return;
      }
      await runAiOnSelection(state, selection || {}, prompt);
    });
  }

  if (btnClearSelection) {
    btnClearSelection.addEventListener("click", () => {
      clearSelectionUI();
      state.selection = null;
    });
  }

  if (btnOpenConfig) {
    btnOpenConfig.addEventListener("click", () => {
      const panel = document.getElementById("config-panel");
      if (!panel) return;
      panel.classList.toggle("config-panel--hidden");
    });
  }

  if (btnImport && hiddenFileInput) {
    btnImport.addEventListener("click", () => {
      hiddenFileInput.value = "";
      hiddenFileInput.click();
    });

    hiddenFileInput.addEventListener("change", async () => {
      const file = hiddenFileInput.files?.[0];
      if (!file) return;
      try {
        const result = await parseXlsxToSheets(file);
        if (!result || !result.sheets || !result.sheets.length) {
          alert("未解析到任何工作表。");
          return;
        }

        // 创建一个新的工作簿对象
        const newWorkbook = {
          id: `workbook-${Date.now()}`,
          name: result.workbookName,
          sheets: result.sheets,
          activeSheetId: result.sheets[0].id,
        };

        // 添加新工作簿并将其设为活动状态
        addWorkbook(state, newWorkbook);

        // 重新渲染整个UI以反映新工作簿
        const gridRoot = document.getElementById("sheet-grid");
        const tabsRoot = document.getElementById("sheet-tabs");
        if (gridRoot && tabsRoot) {
          createGrid(gridRoot, tabsRoot, state);
        }
        
        // 更新列头面板
        updateColumnPanel(state);

      } catch (e) {
        console.error(e);
        alert("导入 Excel 失败，请确认文件是否为合法的 .xlsx。");
      }
    });
  }

  if (btnExport) {
    btnExport.addEventListener("click", () => {
      const wb = getActiveWorkbook(state);
      if (!wb || !wb.sheets || !wb.sheets.length) {
        alert("当前没有可导出的工作表。");
        return;
      }
      exportSheetsToXlsx(wb.sheets, `${wb.name || "excel-flex"}.xlsx`);
    });
  }

  if (btnCloseAiProcess) {
    btnCloseAiProcess.addEventListener("click", () => {
      const panel = document.getElementById("ai-process-panel");
      if (!panel) return;
      panel.classList.add("ai-process-panel--hidden");
    });
  }

  if (btnMerge) {
    btnMerge.addEventListener("click", () => {
      showMergeModal();
    });
  }
}

