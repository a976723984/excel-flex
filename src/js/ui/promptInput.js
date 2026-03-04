import { getWorkbooksWithSheets } from '../services/state.js';

let currentDropdownState = {
  visible: false,
  mode: 'workbook', // workbook, sheet, column
  items: [],
  selectedIndex: 0,
  context: {}
};

let currentReferencePath = [];
let isReferenceMode = false;

export function initPromptInput(state) {
  const promptInput = document.getElementById("prompt-input");
  const dropdown = document.getElementById("reference-dropdown");
  
  if (!promptInput || !dropdown) return;
  
  // 将 state 存储在 window 全局对象中
  window.appState = state;
  
  promptInput.addEventListener("input", (e) => handleInput(e, state, dropdown));
  promptInput.addEventListener("keydown", (e) => handleKeyDown(e, dropdown));
  promptInput.addEventListener("blur", () => {
    // 失去焦点时，如果在引用模式中，完成引用
    if (isReferenceMode) {
      completeReference(promptInput);
    }
  });
  
  // 点击下拉列表项
  dropdown.addEventListener("click", (e) => {
    const target = e.target.closest(".dropdown-item");
    if (target) {
      const index = parseInt(target.dataset.index);
      selectItem(index, state, promptInput, dropdown);
    }
  });
}

function handleInput(e, state, dropdown) {
  const promptInput = e.target;
  const selection = window.getSelection();
  
  if (!selection.rangeCount) return;
  
  const range = selection.getRangeAt(0);
  
  // 获取光标位置的文本内容
  let textBeforeCursor = "";
  let foundCursor = false;
  
  const collectTextBeforeCursor = (currentNode, currentRange) => {
    if (foundCursor) return;
    
    if (currentNode.nodeType === Node.TEXT_NODE) {
      const nodeLength = currentNode.textContent.length;
      for (let i = 0; i < nodeLength; i++) {
        if (currentRange.startContainer === currentNode && currentRange.startOffset === i) {
          foundCursor = true;
          return;
        }
        textBeforeCursor += currentNode.textContent[i];
      }
    } else if (currentNode.nodeType === Node.ELEMENT_NODE) {
      if (currentNode.classList && currentNode.classList.contains('reference-tag')) {
        // 引用标签作为整体处理
        const tagText = currentNode.textContent;
        textBeforeCursor += tagText;
        // 检查光标是否在标签内部
        if (currentRange.startContainer === currentNode || 
            currentNode.contains(currentRange.startContainer)) {
          foundCursor = true;
          return;
        }
      } else {
        // 递归处理子节点
        let child = currentNode.firstChild;
        while (child && !foundCursor) {
          collectTextBeforeCursor(child, currentRange);
          child = child.nextSibling;
        }
      }
    }
  };
  
  collectTextBeforeCursor(promptInput, range);
  
  console.log('textBeforeCursor:', textBeforeCursor);
  
  // 检查是否输入了 /
  const lastChar = textBeforeCursor[textBeforeCursor.length - 1];
  if (lastChar === '/') {
    // 进入引用模式
    isReferenceMode = true;
    currentReferencePath = [];
    showWorkbookDropdown(buildHierarchyTree(state), dropdown, '');
  } else if (isReferenceMode) {
    // 在引用模式中，处理用户输入
    // 找到最后一个 / 之后的文本
    const lastSlashIndex = textBeforeCursor.lastIndexOf('/');
    if (lastSlashIndex !== -1) {
      const inputText = textBeforeCursor.substring(lastSlashIndex + 1);
      console.log('inputText:', inputText);
      handleReferenceInput(inputText, state, dropdown);
    }
  }
}

function handleReferenceInput(inputText, state, dropdown) {
  const hierarchyTree = buildHierarchyTree(state);
  
  switch (currentDropdownState.mode) {
    case 'workbook':
      showWorkbookDropdown(hierarchyTree, dropdown, inputText);
      break;
    case 'sheet':
      if (currentReferencePath.length === 1) {
        const workbookName = currentReferencePath[0];
        showSheetDropdown(hierarchyTree, dropdown, workbookName, inputText);
      }
      break;
    case 'column':
      if (currentReferencePath.length === 2) {
        const workbookName = currentReferencePath[0];
        const sheetName = currentReferencePath[1];
        showColumnDropdown(hierarchyTree, dropdown, workbookName, sheetName, inputText);
      }
      break;
  }
}

function handleKeyDown(e, dropdown) {
  if (!currentDropdownState.visible) return;
  
  switch (e.key) {
    case 'ArrowDown':
      e.preventDefault();
      currentDropdownState.selectedIndex = (currentDropdownState.selectedIndex + 1) % currentDropdownState.items.length;
      renderDropdown(dropdown, currentDropdownState.items, currentDropdownState.mode);
      break;
    case 'ArrowUp':
      e.preventDefault();
      currentDropdownState.selectedIndex = (currentDropdownState.selectedIndex - 1 + currentDropdownState.items.length) % currentDropdownState.items.length;
      renderDropdown(dropdown, currentDropdownState.items, currentDropdownState.mode);
      break;
    case 'Enter':
      e.preventDefault();
      // 从 window 全局对象获取 state
      const state = window.appState;
      selectItem(currentDropdownState.selectedIndex, state, e.target, dropdown);
      break;
    case 'Escape':
      hideDropdown(dropdown);
      isReferenceMode = false;
      currentReferencePath = [];
      break;
  }
}

function selectItem(index, state, promptInput, dropdown) {
  const item = currentDropdownState.items[index];
  if (!item) return;
  
  // 获取当前的输入文本
  const selection = window.getSelection();
  let textNode = null;
  let text = '';
  if (selection.rangeCount) {
    const range = selection.getRangeAt(0);
    textNode = range.startContainer;
    if (textNode.nodeType === Node.TEXT_NODE) {
      text = textNode.textContent;
    }
  }
  
  switch (item.type) {
    case 'workbook':
      // 选择工作簿，进入工作表选择
      currentReferencePath = [item.name];
      currentDropdownState.mode = 'sheet';
      
      // 更新输入框中的文本
      if (textNode && textNode.nodeType === Node.TEXT_NODE) {
        const slashIndex = text.lastIndexOf('/');
        if (slashIndex !== -1) {
          textNode.textContent = text.substring(0, slashIndex + 1) + item.name + '/';
        }
      }
      
      // 显示工作表下拉列表
      if (state) {
        const hierarchyTree = buildHierarchyTree(state);
        showSheetDropdown(hierarchyTree, dropdown, item.name, '');
      }
      break;
    case 'sheet':
      // 选择工作表，进入列选择
      currentReferencePath = [item.workbook, item.name];
      currentDropdownState.mode = 'column';
      
      // 更新输入框中的文本
      if (textNode && textNode.nodeType === Node.TEXT_NODE) {
        const slashIndex = text.lastIndexOf('/');
        if (slashIndex !== -1) {
          const secondLastSlashIndex = text.lastIndexOf('/', slashIndex - 1);
          if (secondLastSlashIndex !== -1) {
            textNode.textContent = text.substring(0, secondLastSlashIndex + 1) + item.workbook + '/' + item.name + '/';
          }
        }
      }
      
      // 显示列下拉列表
      if (state) {
        const hierarchyTree = buildHierarchyTree(state);
        showColumnDropdown(hierarchyTree, dropdown, item.workbook, item.name, '');
      }
      break;
    case 'column':
      // 选择列，完成引用
      currentReferencePath = [item.workbook, item.sheet, item.colLabel];
      completeReference(promptInput, item);
      hideDropdown(dropdown);
      break;
  }
}

function completeReference(promptInput, columnItem = null) {
  if (currentReferencePath.length !== 3 || !columnItem) return;
  
  const [workbook, sheet, column] = currentReferencePath;
  
  // 创建引用标签
  const tag = document.createElement('span');
  tag.className = 'reference-tag';
  tag.contentEditable = false;
  
  // 创建标签内容
  const tagContent = document.createElement('span');
  tagContent.textContent = `${workbook}/${sheet}/${column}`;
  
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
  tag.dataset.reference = JSON.stringify({
    workbook,
    sheet,
    column: columnItem.name,
    colLabel: column
  });
  
  // 找到最后一个 / 的位置并替换
  const selection = window.getSelection();
  if (selection.rangeCount) {
    const range = selection.getRangeAt(0);
    const textNode = range.startContainer;
    if (textNode.nodeType === Node.TEXT_NODE) {
      const text = textNode.textContent;
      const slashIndex = text.lastIndexOf('/');
      if (slashIndex !== -1) {
        // 删除从 / 开始的文本
        const newRange = document.createRange();
        newRange.setStart(textNode, slashIndex);
        newRange.setEnd(textNode, text.length);
        newRange.deleteContents();
        
        // 在光标位置插入标签
        range.setStartAfter(newRange.startContainer, newRange.startOffset);
        range.collapse(true);
        range.insertNode(tag);
        
        // 在标签后插入空格
        const space = document.createTextNode(' ');
        range.insertNode(space);
        
        // 移动光标到空格后
        range.setStartAfter(space);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
      }
    }
  }
  
  // 退出引用模式
  isReferenceMode = false;
  currentReferencePath = [];
}

function buildHierarchyTree(state) {
  const workbooks = getWorkbooksWithSheets(state);
  const tree = {
    workbooks: {}
  };
  
  workbooks.forEach(workbook => {
    tree.workbooks[workbook.name] = {
      name: workbook.name,
      sheets: {}
    };
    
    workbook.sheets.forEach(sheet => {
      tree.workbooks[workbook.name].sheets[sheet.name] = {
        name: sheet.name,
        columns: {}
      };
      
      sheet.headers.forEach(col => {
        tree.workbooks[workbook.name].sheets[sheet.name].columns[col.colLabel] = {
          name: col.name,
          colLabel: col.colLabel
        };
      });
    });
  });
  
  return tree;
}

function showWorkbookDropdown(hierarchyTree, dropdown, filter = "") {
  const workbooks = Object.values(hierarchyTree.workbooks);
  const filteredWorkbooks = workbooks.filter(wb => 
    wb.name.toLowerCase().includes(filter.toLowerCase())
  );
  
  if (filteredWorkbooks.length === 0) {
    hideDropdown(dropdown);
    return;
  }
  
  const items = filteredWorkbooks.map(wb => ({
    type: 'workbook',
    name: wb.name
  }));
  
  currentDropdownState = {
    visible: true,
    mode: 'workbook',
    items,
    selectedIndex: 0
  };
  
  renderDropdown(dropdown, items, 'workbook');
}

function showSheetDropdown(hierarchyTree, dropdown, workbookFilter, sheetFilter = "") {
  // 查找匹配的工作簿
  let matchedWorkbook = null;
  const workbooks = Object.values(hierarchyTree.workbooks);
  for (const workbook of workbooks) {
    if (workbook.name.toLowerCase().includes(workbookFilter.toLowerCase())) {
      matchedWorkbook = workbook;
      break;
    }
  }
  
  if (!matchedWorkbook) {
    hideDropdown(dropdown);
    return;
  }
  
  const sheets = Object.values(matchedWorkbook.sheets);
  const filteredSheets = sheets.filter(sheet => 
    sheet.name.toLowerCase().includes(sheetFilter.toLowerCase())
  );
  
  if (filteredSheets.length === 0) {
    hideDropdown(dropdown);
    return;
  }
  
  const items = filteredSheets.map(sheet => ({
    type: 'sheet',
    name: sheet.name,
    workbook: matchedWorkbook.name
  }));
  
  currentDropdownState = {
    visible: true,
    mode: 'sheet',
    items,
    selectedIndex: 0,
    context: {
      workbook: matchedWorkbook.name
    }
  };
  
  renderDropdown(dropdown, items, 'sheet');
}

function showColumnDropdown(hierarchyTree, dropdown, workbookFilter, sheetFilter, columnFilter = "") {
  // 查找匹配的工作簿
  let matchedWorkbook = null;
  const workbooks = Object.values(hierarchyTree.workbooks);
  for (const workbook of workbooks) {
    if (workbook.name.toLowerCase().includes(workbookFilter.toLowerCase())) {
      matchedWorkbook = workbook;
      break;
    }
  }
  
  if (!matchedWorkbook) {
    hideDropdown(dropdown);
    return;
  }
  
  // 查找匹配的工作表
  let matchedSheet = null;
  const sheets = Object.values(matchedWorkbook.sheets);
  for (const sheet of sheets) {
    if (sheet.name.toLowerCase().includes(sheetFilter.toLowerCase())) {
      matchedSheet = sheet;
      break;
    }
  }
  
  if (!matchedSheet) {
    hideDropdown(dropdown);
    return;
  }
  
  const columns = Object.values(matchedSheet.columns);
  const filteredColumns = columns.filter(col => 
    col.name.toLowerCase().includes(columnFilter.toLowerCase()) ||
    col.colLabel.toLowerCase().includes(columnFilter.toLowerCase())
  );
  
  if (filteredColumns.length === 0) {
    hideDropdown(dropdown);
    return;
  }
  
  const items = filteredColumns.map(col => ({
    type: 'column',
    name: col.name,
    colLabel: col.colLabel,
    workbook: matchedWorkbook.name,
    sheet: matchedSheet.name
  }));
  
  currentDropdownState = {
    visible: true,
    mode: 'column',
    items,
    selectedIndex: 0,
    context: {
      workbook: matchedWorkbook.name,
      sheet: matchedSheet.name
    }
  };
  
  renderDropdown(dropdown, items, 'column');
}

function renderDropdown(dropdown, items, type) {
  let html = '';
  
  items.forEach((item, index) => {
    const isSelected = index === currentDropdownState.selectedIndex;
    let displayText = '';
    
    switch (type) {
      case 'workbook':
        displayText = item.name;
        break;
      case 'sheet':
        displayText = item.name;
        break;
      case 'column':
        displayText = `${item.colLabel} (${item.name})`;
        break;
    }
    
    html += `
      <div class="dropdown-item ${isSelected ? 'selected' : ''}" data-index="${index}">
        ${displayText}
      </div>
    `;
  });
  
  dropdown.innerHTML = html;
  dropdown.style.display = 'block';
  
  // 定位下拉列表
  const promptInput = document.getElementById("prompt-input");
  if (promptInput) {
    const rect = promptInput.getBoundingClientRect();
    dropdown.style.position = 'absolute';
    dropdown.style.left = `${rect.left}px`;
    dropdown.style.top = `${rect.bottom + 5}px`;
    dropdown.style.width = `${rect.width}px`;
  }
}

function hideDropdown(dropdown) {
  dropdown.style.display = 'none';
  currentDropdownState.visible = false;
}

export function getPromptText() {
  const promptInput = document.getElementById("prompt-input");
  if (!promptInput) return "";
  
  // 收集所有文本节点和引用标签的内容
  let text = "";
  
  const collectText = (node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      text += node.textContent;
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      if (node.classList && node.classList.contains('reference-tag')) {
        text += node.textContent;
      } else {
        let child = node.firstChild;
        while (child) {
          collectText(child);
          child = child.nextSibling;
        }
      }
    }
  };
  
  collectText(promptInput);
  return text.trim();
}
