import { getWorkbooksWithSheets } from "../services/state.js";

let columnPanelVisible = false;

export function initColumnPanel(state) {
  const btnOpenColumnPanel = document.getElementById("btn-open-column-panel");
  const columnPanel = document.getElementById("column-panel");

  if (!columnPanel) return;

  // 创建打开列头面板的按钮
  if (!btnOpenColumnPanel) {
    const toolbarLeft = document.querySelector(".toolbar-left");
    if (toolbarLeft) {
      const openBtn = document.createElement("button");
      openBtn.className = "btn";
      openBtn.id = "btn-open-column-panel";
      openBtn.textContent = "列头";
      openBtn.addEventListener("click", () => {
        toggleColumnPanel(state);
      });
      toolbarLeft.appendChild(openBtn);
    }
  }

  // 初始化面板内容
  updateColumnPanel(state);
}

export function toggleColumnPanel(state) {
  const columnPanel = document.getElementById("column-panel");
  if (!columnPanel) return;

  columnPanelVisible = !columnPanelVisible;
  if (columnPanelVisible) {
    columnPanel.classList.remove("column-panel--hidden");
    updateColumnPanel(state);
  } else {
    columnPanel.classList.add("column-panel--hidden");
  }
}

export function updateColumnPanel(state) {
  const columnPanel = document.getElementById("column-panel");
  if (!columnPanel) return;

  const workbooks = getWorkbooksWithSheets(state);
  
  let html = '<div class="column-panel-title">表结构</div>';
  html += '<div class="column-panel-scroll-content">'; // 滚动容器开始
  
  if (workbooks.length === 0) {
    html += '<div class="panel-placeholder">暂无数据，请导入Excel文件</div>';
  } else {
    workbooks.forEach(workbook => {
      html += `
        <div class="column-workbook-group" data-workbook-id="${workbook.id}" data-expanded="true">
          <div class="column-workbook-header">
            <span class="toggle-icon">▼</span>
            <span class="column-workbook-name">${workbook.name}</span>
          </div>
          <div class="column-workbook-content">
      `;
      
      if (workbook.sheets && workbook.sheets.length > 0) {
        workbook.sheets.forEach(sheet => {
          html += `
            <div class="column-sheet-group" data-sheet-id="${sheet.id}" data-expanded="true">
              <div class="column-sheet-header">
                <span class="toggle-icon">▼</span>
                <span class="column-sheet-name">${sheet.name}</span>
              </div>
              <div class="column-sheet-content">
          `;
          
          if (sheet.headers && sheet.headers.length > 0) {
            sheet.headers.forEach(col => {
              html += `
                <div class="column-item" data-workbook="${workbook.id}" data-sheet="${sheet.id}" data-col="${col.colIndex}">
                  <span class="column-item-tag">${col.colLabel}</span>
                  <span>${col.name}</span>
                </div>
              `;
            });
          } else {
            html += '<div class="panel-placeholder small">暂无列</div>';
          }
          
          html += `</div></div>`; // aclose column-sheet-content and column-sheet-group
        });
      } else {
        html += '<div class="panel-placeholder small">暂无工作表</div>';
      }
      
      html += `</div></div>`; // close column-workbook-content and column-workbook-group
    });
  }
  
  html += '</div>'; // 滚动容器结束
  columnPanel.innerHTML = html;

  // 添加事件监听器
  addEventListenersToPanel(columnPanel);
}

function addEventListenersToPanel(panel) {
  // 为工作簿和工作表标题添加点击事件
  panel.querySelectorAll('.column-workbook-header, .column-sheet-header').forEach(header => {
    header.addEventListener('click', (e) => {
      const group = header.parentElement;
      const content = header.nextElementSibling;
      const icon = header.querySelector('.toggle-icon');
      const isExpanded = group.dataset.expanded === 'true';

      group.dataset.expanded = !isExpanded;
      content.style.display = isExpanded ? 'none' : '';
      icon.textContent = isExpanded ? '▶' : '▼';
      
      e.stopPropagation();
    });
  });

  // 为列项目添加点击事件
  panel.querySelectorAll(".column-item").forEach(item => {
    item.addEventListener("click", () => {
      const workbookId = item.dataset.workbook;
      const sheetId = item.dataset.sheet;
      const colIndex = parseInt(item.dataset.col);
      const workbookName = item.closest(".column-workbook-group").querySelector(".column-workbook-name").textContent;
      const sheetName = item.closest(".column-sheet-group").querySelector(".column-sheet-name").textContent;
      const colLabel = item.querySelector(".column-item-tag").textContent;
      const colName = item.querySelector("span:last-child").textContent;
      
      insertColumnReference(workbookName, sheetName, colLabel, colName);
    });
  });
}

function insertColumnReference(workbookName, sheetName, colLabel, colName) {
  const promptInput = document.getElementById("prompt-input");
  if (!promptInput) return;

  // 插入完整的列引用作为标签
  const tag = document.createElement('span');
  tag.className = 'reference-tag';
  tag.contentEditable = false;
  
  // 创建标签内容
  const tagContent = document.createElement('span');
  tagContent.textContent = `${workbookName}/${sheetName}/${colLabel}`;
  
  // 创建删除按钮
  const deleteBtn = document.createElement('span');
  deleteBtn.className = 'reference-tag-delete';
  deleteBtn.textContent = '×';
  deleteBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    tag.remove();
  });
  
  // 组装标签
  tag.appendChild(tagContent);
  tag.appendChild(deleteBtn);
  tag.dataset.reference = `${workbookName}/${sheetName}/${colLabel}`;
  
  // 在输入框末尾插入标签
  const space = document.createTextNode(' ');
  promptInput.appendChild(tag);
  promptInput.appendChild(space);
  
  // 聚焦到输入框并设置光标位置到末尾
  promptInput.focus();
  const selection = window.getSelection();
  const range = document.createRange();
  range.setStartAfter(space);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
}