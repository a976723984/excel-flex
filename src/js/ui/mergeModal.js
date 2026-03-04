
import { getSheetsWithColumns, mergeSheets } from '../services/mergeService.js';
import { addSheet, getSheets } from '../services/state.js';
import { createGrid } from './grid.js';

const modal = document.getElementById('merge-modal');
const btnClose = document.getElementById('btn-close-merge-modal');
const btnCancel = document.getElementById('btn-cancel-merge');
const btnConfirm = document.getElementById('btn-confirm-merge');
const sheetsContainer = document.getElementById('merge-sheets-container');

let state = null;

export function initMergeModal(appState) {
  state = appState;

  btnClose.addEventListener('click', () => hideMergeModal());
  btnCancel.addEventListener('click', () => hideMergeModal());
  btnConfirm.addEventListener('click', () => {
    const selections = [];
    const sheetGroups = sheetsContainer.querySelectorAll('.sheet-group');
    sheetGroups.forEach(group => {
      const sheetSelect = group.querySelector('.sheet-select');
      const columnSelect = group.querySelector('.column-select');
      if (sheetSelect.value && columnSelect.value) {
        selections.push({
          sheetId: sheetSelect.value,
          columnIndex: parseInt(columnSelect.value, 10),
        });
      }
    });

    try {
      const newSheet = mergeSheets(state, selections);
      const allSheets = getSheets(state);
      newSheet.id = `sheet-${allSheets.length + 1}`;
      addSheet(state, newSheet);

      const gridRoot = document.getElementById("sheet-grid");
      const tabsRoot = document.getElementById("sheet-tabs");
      createGrid(gridRoot, tabsRoot, state);

      hideMergeModal();
    } catch (e) {
      alert(`合并失败: ${e.message}`);
    }
  });
}

export function showMergeModal() {
  if (!modal) return;
  populateSheets();
  modal.classList.remove('modal-wrapper--hidden');
}

function hideMergeModal() {
  if (!modal) return;
  modal.classList.add('modal-wrapper--hidden');
}

function populateSheets() {
  if (!sheetsContainer || !state) return;

  const sheetsWithColumns = getSheetsWithColumns(state);
  sheetsContainer.innerHTML = ''; // Clear previous content

  // For simplicity, we'll start with two sheet selectors
  for (let i = 0; i < 2; i++) {
    const group = document.createElement('div');
    group.className = 'sheet-group';

    const selector = document.createElement('div');
    selector.className = 'sheet-selector';

    const sheetSelect = document.createElement('select');
    sheetSelect.className = 'sheet-select';
    
    const columnSelect = document.createElement('select');
    columnSelect.className = 'column-select';

    sheetsWithColumns.forEach(workbook => {
      const optgroup = document.createElement('optgroup');
      optgroup.label = workbook.name;
      workbook.sheets.forEach(sheet => {
        const option = document.createElement('option');
        option.value = sheet.id;
        option.textContent = sheet.name;
        optgroup.appendChild(option);
      });
      sheetSelect.appendChild(optgroup);
    });

    sheetSelect.addEventListener('change', () => {
      const selectedSheetId = sheetSelect.value;
      const selectedSheet = findSheet(sheetsWithColumns, selectedSheetId);
      populateColumns(columnSelect, selectedSheet?.columns || []);
    });

    selector.appendChild(document.createTextNode(`工作表 ${i + 1}:`));
    selector.appendChild(sheetSelect);
    selector.appendChild(document.createTextNode(`合并列:`));
    selector.appendChild(columnSelect);
    
    group.appendChild(selector);
    sheetsContainer.appendChild(group);

    // Trigger change to populate columns for the default selection
    sheetSelect.dispatchEvent(new Event('change'));
  }
}

function populateColumns(selectElement, columns) {
  selectElement.innerHTML = '';
  columns.forEach(col => {
    const option = document.createElement('option');
    option.value = col.colIndex;
    option.textContent = col.name;
    selectElement.appendChild(option);
  });
}

function findSheet(workbooks, sheetId) {
  for (const workbook of workbooks) {
    const sheet = workbook.sheets.find(s => s.id === sheetId);
    if (sheet) return sheet;
  }
  return null;
}
