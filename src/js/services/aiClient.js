import { getConfig, getActiveWorkbook, getSheetById, addSheet, getSheets, getWorkbooksWithSheets } from "./state.js";
 import { mergeSheets } from "./mergeService.js";
 import { createGrid } from "../ui/grid.js";
 import { updateColumnPanel } from "../ui/columnPanel.js";

export async function runAiOnSelection(state, selection, prompt) {
  const statusEl = document.getElementById("status-ai");
  const config = getConfig(state);

  if (!config.apiUrl || !config.apiKey) {
    alert("请先在配置面板中填写 API 地址和 API Key。");
    return;
  }

  if (statusEl) {
    statusEl.textContent = "AI 执行中...";
    statusEl.classList.add("status-ai--busy");
  }

  const panel = document.getElementById("ai-process-panel");
  const contentEl = document.getElementById("ai-process-content");

  if (panel && contentEl) {
    panel.classList.remove("ai-process-panel--hidden");
    contentEl.innerHTML = "<div>AI开始处理...</div>";
  }

  try {
    if (contentEl) {
      contentEl.innerHTML += "<div>AI 思考：正在分析您的指令并生成执行方案...</div>";
    }

    // 第一步：让 AI 生成执行方案
    const plan = await getAiExecutionPlan(state, selection, prompt, config, contentEl);

    if (contentEl) {
      contentEl.innerHTML += `<div>AI 决策：识别到操作“${plan.action}”。</div>`;
      contentEl.innerHTML += `<div>方案详情：<pre>${JSON.stringify(plan, null, 2)}</pre></div>`;
    }

    // 第二步：根据方案执行操作
    switch (plan.action) {
      case "filter":
        await executeSemanticFilter(state, selection, plan, config, contentEl);
        break;
      case "merge":
        await executeMerge(state, plan, config, contentEl);
        break;
      case "analysis":
        await executeAnalysis(state, selection, prompt, config, contentEl);
        break;
      case "classification":
        await executeClassification(state, selection, prompt, config, contentEl);
        break;
      case "generate_column":
        await executeGenerateColumn(state, plan, config, contentEl);
        break;  
      default:
        // 对于AI无法制定方案的简单请求，执行常规流程
        await executeRegularAiRequest(state, selection, prompt, config, contentEl);
        break;
    }

    if (contentEl) {
      contentEl.innerHTML += "<div>操作完成！</div>";
    }
  } catch (e) {
    console.error("AI 请求失败", e);
    if (contentEl) {
      contentEl.innerHTML += `<div style="color: red;">错误: ${e.message}</div>`;
    }
    alert("AI 请求失败，请检查配置与网络，或优化您的指令。");
  } finally {
    if (statusEl) {
      statusEl.textContent = "AI 空闲";
      statusEl.classList.remove("status-ai--busy");
    }
  }
}

// 让 AI 生成执行方案
async function getAiExecutionPlan(state, selection, prompt, config, contentEl) {
  const references = parseReferences(prompt);
  const { data: refData, columns: refCols } = getDataByReferences(state, references);

  // 获取所有工作簿的完整结构
  const allWorkbooks = getWorkbooksWithSheets(state).map(wb => ({
    name: wb.name,
    sheets: wb.sheets.map(s => ({
      id: s.id,
      name: s.name,
      columns: s.headers.map((h, index) => ({ index: index, name: h.name }))
    }))
  }));

  const systemPrompt = `
    你是一个强大的 Excel 操作助手。你的任务是分析用户的请求，并生成一个 JSON 格式的执行方案。
    
    可用的操作 (action) 包括: "filter", "merge", "sort", "analysis", "classification", "generate_column", "regular"。

    - 如果是 filter: 方案需要包含 target_sheet_id, target_column_index 和 filter_criteria。filter_criteria 应该是从用户请求中提炼出的核心语义条件。
    - 如果是 merge: 方案需要包含 'merge_pairs' (一个包含两个对象的数组，每个对象都有 'sheet_id' 和 'column_index') 和 'match_type' ('fuzzy' 或 'exact')。
    - 如果是 sort: 方案需要包含 target_sheet_id, target_column_index 和 sort_order ('asc' 或 'desc')。
    - 如果是 generate_column: 方案需要包含 target_sheet_id, new_column_name, reference_columns (一个包含列索引的数组) 和 generation_prompt (具体的生成要求)。
    - 对于其他操作，或者当你不确定时，将 action 设置为 "regular"。

    用户的请求是: "${prompt}"

    当前所有工作簿和工作表的结构是: 
    ${JSON.stringify(allWorkbooks, null, 2)}

    用户在提示词中引用了 ${refCols.length} 个列，其数据如下 (最多显示前5行):
    ${JSON.stringify(refCols.map((c, i) => ({ column: c.name, data: refData.map(r => r[i]).slice(0, 5) })), null, 2)}

    请根据以上所有信息，生成最合理的 JSON 执行方案。
  `;

  const payload = {
    prompt: systemPrompt,
    context: "请严格按照要求返回 JSON 格式的执行方案。"
  };

  const response = await callAiApi(config, payload, contentEl);
  const planJson = parseAiResponse(response, config.apiUrl);
  
  try {
    return JSON.parse(planJson);
  } catch (e) {
    console.error("解析AI执行方案失败", e, "原始JSON:", planJson);
    throw new Error("AI返回的执行方案格式不正确，无法解析。");
  }
}

// 执行语义筛选
async function executeSemanticFilter(state, selection, plan, config, contentEl) {
  if (contentEl) {
    contentEl.innerHTML += `<div>操作：正在基于方案执行“筛选”...</div>`;
  }

  const { target_sheet_id: targetSheetId, target_column_index: targetColumnIndex, filter_criteria: filterCriteria } = plan;

  const targetSheet = getSheetById(state, targetSheetId);
  if (!targetSheet) {
    throw new Error(`执行方案无效：无法找到 ID 为“${targetSheetId}”的工作表。`);
  }

  // 1. 提取数据
  const { data, columns } = extractDataWithColumns(state, selection, targetSheet);
  if (!data.length || !columns.length) {
    throw new Error("没有可供筛选的数据。");
  }

  const targetColumn = columns.find(c => c.index === targetColumnIndex);
  if (!targetColumn) {
    const availableColumns = JSON.stringify(columns.map(c => ({ index: c.index, name: c.name })));
    throw new Error(`执行方案无效：无法找到索引为 ${targetColumnIndex} 的列。可用的列为: ${availableColumns}`);
  }

  if (contentEl) {
    contentEl.innerHTML += `<div>操作：将在“${targetColumn.name}”列上，根据条件“${filterCriteria}”进行筛选。</div>`;
  }

  // 2. 提取目标列的数据，并分批次让 AI 筛选出匹配的行
  const columnData = data.map((row, index) => ({ rowIndex: index, value: row[targetColumnIndex] }));
  const batchSize = 100; // 每批处理100行
  const allMatchingRowIndexes = [];

  for (let i = 0; i < columnData.length; i += batchSize) {
    const batch = columnData.slice(i, i + batchSize);
    if (contentEl) {
      contentEl.innerHTML += `<div>AI 思考：正在分析“${targetColumn.name}”列的数据 (第 ${i + 1} - ${i + batch.length} 行)...</div>`;
    }

    const filterPayload = {
      prompt: `我正在处理名为“${targetColumn.name}”的列。请从以下数据中，筛选出所有“${filterCriteria}”的行。`, 
      context: `这是“${targetColumn.name}”列的数据: ${JSON.stringify(batch)}. 请分析每一行的 'value'，并返回一个包含所有符合条件的行的 'rowIndex' 的JSON数组。`,
    };
    const filterResponse = await callAiApi(config, filterPayload, contentEl);
    const batchMatchingIndexes = JSON.parse(parseAiResponse(filterResponse, config.apiUrl));
    
    if (batchMatchingIndexes && batchMatchingIndexes.length > 0) {
      allMatchingRowIndexes.push(...batchMatchingIndexes);
    }
  }

  if (allMatchingRowIndexes.length === 0) {
    throw new Error("未找到任何符合条件的数据。");
  }

  if (contentEl) {
    contentEl.innerHTML += `<div>AI 识别结果：共找到 ${allMatchingRowIndexes.length} 行符合条件的数据。</div>`;
  }

  // 3. 构建新表
  const headerRow = columns.map(c => c.name);
  const filteredData = [headerRow];
  allMatchingRowIndexes.forEach(rowIndex => {
    if (data[rowIndex]) {
      filteredData.push(data[rowIndex]);
    }
  });

  const newSheetData = {
    name: "筛选结果",
    data: filteredData,
    rows: filteredData.length,
    cols: headerRow.length,
    headers: headerRow.map((name, i) => ({
      colIndex: i,
      name,
      colLabel: String.fromCharCode(65 + i),
    }))
  };

  const newSheet = {
    ...newSheetData,
    id: `sheet-${Date.now()}`,
  };

  // 4. 添加新表并刷新视图
  addSheet(state, newSheet);
  const gridRoot = document.getElementById("sheet-grid");
  const tabsRoot = document.getElementById("sheet-tabs");
  createGrid(gridRoot, tabsRoot, state);

  if (contentEl) {
    contentEl.innerHTML += `<div>操作：筛选完成，已生成新工作表“筛选结果”。</div>`;
  }
}



// 解析引用标签，提取工作簿、工作表和列信息
function parseReferences(prompt) {
  const references = [];
  // 更新正则表达式以精确匹配列标签（例如 A, B, AA），并忽略后面的'×'
  const referencePattern = /([^/\s]+)\/([^/\s]+)\/([A-Z]+)/g;
  let match;
  
  while ((match = referencePattern.exec(prompt)) !== null) {
    references.push({
      workbook: match[1],
      sheet: match[2],
      column: match[3] // 这将是列标签, 例如 'A', 'B'
    });
  }
  
  return references;
}

// 根据引用获取数据
function getDataByReferences(state, references) {
  const data = [];
  const columns = [];
  
  references.forEach(ref => {
    // 查找工作簿
    const workbook = state.workbooks?.find(wb => wb.name === ref.workbook);
    if (!workbook) return;
    
    // 查找工作表
    const sheet = workbook.sheets?.find(s => s.name === ref.sheet);
    if (!sheet) return;
    
    // 查找列
    const column = sheet.headers?.find(h => h.name === ref.column || h.colLabel === ref.column);
    if (!column) return;
    
    // 提取列信息
    columns.push({
      index: column.colIndex,
      name: column.name || `列${String.fromCharCode(65 + column.colIndex)}`,
      label: column.colLabel || String.fromCharCode(65 + column.colIndex)
    });
    
    // 提取列数据
    const columnData = [];
    if (sheet.data) {
      for (let r = 0; r < sheet.data.length; r++) {
        const row = sheet.data[r] || [];
        columnData.push(row[column.colIndex] || "");
      }
    }
    data.push(columnData);
  });
  
  // 转置数据，确保每行包含所有列的数据
  const transposedData = [];
  const maxRows = Math.max(...data.map(col => col.length), 0);
  for (let r = 0; r < maxRows; r++) {
    const row = [];
    for (let c = 0; c < data.length; c++) {
      row.push(data[c][r] || "");
    }
    transposedData.push(row);
  }
  
  return { data: transposedData, columns };
}

// 执行分析操作
async function executeAnalysis(state, selection, prompt, config, contentEl) {
  if (contentEl) {
    contentEl.innerHTML += "<div>思考：执行数据分析...</div>";
  }

  // 提取数据
  const { data, columns } = extractDataWithColumns(state, selection);
  if (data.length === 0) {
    throw new Error("没有可供分析的数据。");
  }

  // 分批处理
  const batchSize = 100;
  let analysisResults = "";

  for (let i = 0; i < data.length; i += batchSize) {
    const batch = data.slice(i, i + batchSize);
    if (contentEl) {
      contentEl.innerHTML += `<div>AI 思考：正在分析数据 (第 ${i + 1} - ${i + batch.length} 行)...</div>`;
    }

    const analysisPayload = {
      prompt: `请基于用户的请求“${prompt}”，分析以下这批数据。`,
      context: `这是整个数据集的一部分。\n\n列信息: ${JSON.stringify(columns)}\n\n数据:\n${JSON.stringify(batch)}`
    };

    const response = await callAiApi(config, analysisPayload, contentEl);
    const parsedContent = parseAiResponse(response, config.apiUrl);
    analysisResults += parsedContent + "\n\n"; // 将各批次结果合并
  }

  if (contentEl) {
    contentEl.innerHTML += "<div>操作：处理AI分析结果...</div>";
  }

  // 将完整的分析结果添加到新工作表
  addAnalysisResultToSheet(state, analysisResults);
  
  if (contentEl) {
    contentEl.innerHTML += "<div>操作：将分析结果添加到新工作表...</div>";
  }
}

// 执行分类匹配操作
async function executeClassification(state, selection, prompt, config, contentEl) {
  if (contentEl) {
    contentEl.innerHTML += "<div>思考：执行分类匹配...</div>";
  }

  try {
    // 提取数据
    const { dataWithRowNumbers, columns } = extractDataWithColumnsAndRowNumbers(state, selection);
    if (dataWithRowNumbers.length === 0) {
      throw new Error("没有可供分类的数据。");
    }

    // 分批处理
    const batchSize = 100;
    const allResults = [];

    for (let i = 0; i < dataWithRowNumbers.length; i += batchSize) {
      const batch = dataWithRowNumbers.slice(i, i + batchSize);
      if (contentEl) {
        contentEl.innerHTML += `<div>AI 思考：正在处理数据 (第 ${i + 1} - ${i + batch.length} 行)...</div>`;
      }

      const classificationPayload = {
      prompt: `请根据用户的请求“${prompt}”，对以下数据进行分类匹配。`,
      context: `这是整个数据集的一部分。\n\n列信息: ${JSON.stringify(columns)}\n\n数据:\n${JSON.stringify(batch)}\n\n请返回一个包含每个输入行的 'rowNumber' 和 'match' 结果的JSON数组。`
    };

    const batchResponse = await callAiApi(config, classificationPayload, contentEl);
      const parsedContent = parseAiResponse(batchResponse, config.apiUrl);
      const results = (typeof parsedContent === 'string' ? JSON.parse(parsedContent) : parsedContent).results || [];
      allResults.push(...results);
    }

    if (contentEl) {
      contentEl.innerHTML += `<div>操作：组合所有匹配结果，共 ${allResults.length} 条...</div>`;
    }

    // 检查是否有匹配结果
    if (allResults.length === 0) {
      if (contentEl) {
        contentEl.innerHTML += `<div style="color: yellow;">警告: 没有匹配结果</div>`;
      }
      // 生成默认结果
      const defaultResults = dataWithRowNumbers.map(item => ({
        rowNumber: item.rowNumber,
        match: "无匹配结果"
      }));
      allResults.push(...defaultResults);
    }

    // 根据行号组合数据
    const combinedData = combineDataByRowNumber(dataWithRowNumbers, allResults);
    
    // 检查组合后的数据
    if (combinedData.length === 0) {
      throw new Error("组合数据失败");
    }

    // 将组合结果添加到新工作簿
    addNewWorkbookWithData(state, "分类匹配结果", combinedData);
    
    if (contentEl) {
      contentEl.innerHTML += "<div>操作：将组合结果添加到新工作簿...</div>";
    }
  } catch (error) {
    console.error("分类匹配失败", error);
    if (contentEl) {
      contentEl.innerHTML += `<div style="color: red;">错误: ${error.message}</div>`;
    }
    throw error;
  }
}

// 执行常规AI请求
async function executeRegularAiRequest(state, selection, prompt, config, contentEl) {
  if (contentEl) {
    contentEl.innerHTML += "<div>操作：调用远程AI API...</div>";
    await new Promise(resolve => setTimeout(resolve, 500));
  }

  // 解析引用标签
  const references = parseReferences(prompt);
  let data, columns;
  
  if (references.length > 0) {
    // 如果有引用，根据引用获取数据
    if (contentEl) {
      contentEl.innerHTML += `<div>操作：解析引用标签，找到 ${references.length} 个引用...</div>`;
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    
    const result = getDataByReferences(state, references);
    data = result.data;
    columns = result.columns;
    
    if (contentEl) {
      contentEl.innerHTML += `<div>操作：根据引用获取数据，共 ${data.length} 行，${columns.length} 列...</div>`;
      await new Promise(resolve => setTimeout(resolve, 500));
    }
  } else {
    // 如果没有引用，使用选中的数据
    const result = extractDataWithColumns(state, selection);
    data = result.data;
    columns = result.columns;
    
    if (contentEl) {
      contentEl.innerHTML += `<div>操作：分析选中数据，共 ${data.length} 行，${columns.length} 列...</div>`;
      await new Promise(resolve => setTimeout(resolve, 500));
    }
  }

  // 构建请求数据
  const requestData = {
    prompt,
    data: data,
    columns: columns,
    context: "请根据用户的请求处理这些数据。"
  };

  // 调用远程API
  await callAiApi(config, requestData, contentEl);

  if (contentEl) {
    contentEl.innerHTML += "<div>操作：处理AI响应...</div>";
    await new Promise(resolve => setTimeout(resolve, 500));
  }
}

// 调用AI API
async function callAiApi(config, payload, contentEl) {
  if (contentEl) {
    contentEl.innerHTML += `<div>准备调用API: ${config.apiUrl}</div>`;
    contentEl.innerHTML += `<div>模型: ${config.model}</div>`;
  }

  // 将 prompt 和 context 组合成最终发送给 AI 的内容
  const finalContent = `${payload.prompt}\n\n${payload.context || ''}`;

  try {
    const requestData = {
      model: config.model,
      messages: [
        {
          role: "user",
          content: finalContent
        }
      ],
      temperature: 0.7,
      max_tokens: 4096, // 增加最大 token 限制以处理更大的数据批次
    };

    if (contentEl) {
      contentEl.innerHTML += `<div>请求入参: <pre>${JSON.stringify(requestData, null, 2)}</pre></div>`;
    }

    const resp = await fetch(config.apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify(requestData),
    });

    if (contentEl) {
      contentEl.innerHTML += "<div>正在等待AI响应...</div>";
    }

    if (!resp.ok) {
      throw new Error(`API请求失败: ${resp.status} ${resp.statusText}`);
    }

    const response = await resp.json();

    if (contentEl) {
      contentEl.innerHTML += "<div>收到AI响应</div>";
      // 使用 <pre> 标签来格式化和完整显示JSON响应
      contentEl.innerHTML += `<div>响应内容: <pre>${JSON.stringify(response, null, 2)}</pre></div>`;
      await new Promise(resolve => setTimeout(resolve, 500));
    }

    return response;
  } catch (error) {
    console.error("API请求失败", error);
    if (contentEl) {
      contentEl.innerHTML += `<div style="color: red;">API请求失败: ${error.message}</div>`;
    }
    throw error;
  }
}

// 提取选中的数据
function extractSelectedData(state, selection) {
  const data = [];
  
  // 遍历所有选中的区域
  Object.keys(selection).forEach(sheetId => {
    const sheet = getSheetById(state, sheetId);
    if (!sheet) return;
    
    const sheetSelections = selection[sheetId];
    sheetSelections.forEach(sel => {
      // 提取选中区域的数据
      for (let r = sel.startRow; r <= sel.endRow; r++) {
        const row = [];
        for (let c = sel.startCol; c <= sel.endCol; c++) {
          row.push(sheet.cells[r]?.[c] || "");
        }
        data.push(row);
      }
    });
  });
  
  return data;
}

// 提取选中的数据并包含行号
function extractSelectedDataWithRowNumbers(state, selection) {
  const data = [];
  
  // 遍历所有选中的区域
  Object.keys(selection).forEach(sheetId => {
    const sheet = getSheetById(state, sheetId);
    if (!sheet) return;
    
    const sheetSelections = selection[sheetId];
    sheetSelections.forEach(sel => {
      // 提取选中区域的数据，包含行号
      for (let r = sel.startRow; r <= sel.endRow; r++) {
        const row = {
          rowNumber: r,
          data: []
        };
        for (let c = sel.startCol; c <= sel.endCol; c++) {
          row.data.push(sheet.cells[r]?.[c] || "");
        }
        data.push(row);
      }
    });
  });
  
  return data;
}

// 提取指定工作表的完整数据和列信息
function extractDataWithColumns(state, selection, targetSheet) {
  // 优先使用传入的 targetSheet，否则回退到当前活动的工作表
  const sheetToUse = targetSheet || (() => {
    const activeWorkbook = getActiveWorkbook(state);
    if (!activeWorkbook || activeWorkbook.sheets.length === 0) return null;
    return activeWorkbook.sheets.find(s => s.id === activeWorkbook.activeSheetId) || activeWorkbook.sheets[0];
  })();

  if (!sheetToUse || !sheetToUse.data || sheetToUse.data.length === 0) {
    return { data: [], columns: [] };
  }

  // 确保从正确的工作表提取列头，并为每一列添加索引
  const columns = (sheetToUse.headers || []).map((h, index) => ({ ...h, index }));
  const data = [];

  // 确保从正确的工作表提取数据
  const rowCount = sheetToUse.data.length;
  // 使用列头的数量作为 colCount，如果列头不存在，则使用第一行数据的长度
  const colCount = sheetToUse.headers ? sheetToUse.headers.length : (sheetToUse.data[0] ? sheetToUse.data[0].length : 0);

  for (let r = 0; r < rowCount; r++) {
    const rowData = [];
    const sheetRow = sheetToUse.data[r] || [];
    for (let c = 0; c < colCount; c++) {
      rowData.push(sheetRow[c] || "");
    }
    data.push(rowData);
  }

  return { data, columns };
}

// 提取选中的数据、列信息和行号
function extractDataWithColumnsAndRowNumbers(state, selection) {
  const { data, columns } = extractDataWithColumns(state, selection);
  
  const dataWithRowNumbers = data.map((row, index) => {
    // 如果有 selection，行号需要加上起始行号
    const rowNumber = selection && selection.startRow != null ? selection.startRow + index : index;
    return {
      rowNumber,
      data: row
    };
  });

  return { dataWithRowNumbers, columns };
}

// 根据行号组合数据
function combineDataByRowNumber(originalData, matchResults) {
  // 创建行号到数据的映射
  const rowMap = new Map();
  originalData.forEach(item => {
    rowMap.set(item.rowNumber, item.data);
  });
  
  // 组合数据
  const combinedData = [];
  
  // 添加表头
  if (originalData.length > 0) {
    const headers = [...originalData[0].data, "匹配结果"];
    combinedData.push(headers);
  }
  
  // 添加数据行
  matchResults.forEach(result => {
    const rowData = rowMap.get(result.rowNumber);
    if (rowData) {
      combinedData.push([...rowData, result.match]);
    }
  });
  
  return combinedData;
}

// 将分析结果添加到新工作表
function addAnalysisResultToSheet(state, analysisResult) {
  const activeWorkbook = getActiveWorkbook(state);
  if (!activeWorkbook) return;
  
  // 创建新工作表
  const newSheet = {
    id: `sheet-${Date.now()}`,
    name: "分析结果",
    cells: {},
    headers: [
      { colIndex: 0, name: "分析结果", colLabel: "A" }
    ]
  };
  
  // 解析AI响应
  let analysisResultText = "";
  if (typeof analysisResult === 'string') {
    analysisResultText = analysisResult;
  } else {
    analysisResultText = JSON.stringify(analysisResult, null, 2);
  }
  
  // 添加分析结果到工作表
  const lines = analysisResultText.split("\n");
  lines.forEach((line, index) => {
    newSheet.cells[index] = {
      0: line
    };
  });
  
  // 添加到工作簿
  activeWorkbook.sheets.push(newSheet);
  
  // 切换到新工作表
  state.activeSheetId = newSheet.id;
}

// 将数据添加到新工作簿
function addNewWorkbookWithData(state, workbookName, data) {
  // 创建新工作簿
  const newWorkbook = {
    id: `workbook-${Date.now()}`,
    name: workbookName,
    sheets: [],
    activeSheetId: null
  };
  
  // 创建新工作表
  const newSheet = {
    id: `sheet-${Date.now()}`,
    name: "匹配结果",
    cells: {},
    headers: []
  };
  
  // 添加表头
  if (data.length > 0) {
    const headers = data[0];
    headers.forEach((header, index) => {
      newSheet.headers.push({
        colIndex: index,
        name: header,
        colLabel: String.fromCharCode(65 + index) // A, B, C, ...
      });
    });
  }
  
  // 添加数据
  data.forEach((row, rowIndex) => {
    newSheet.cells[rowIndex] = {};
    row.forEach((cell, colIndex) => {
      newSheet.cells[rowIndex][colIndex] = cell;
    });
  });
  
  // 添加工作表到工作簿
  newWorkbook.sheets.push(newSheet);
  newWorkbook.activeSheetId = newSheet.id;
  
  // 添加工作簿到状态
  state.workbooks.push(newWorkbook);
  state.activeWorkbookId = newWorkbook.id;
  state.activeSheetId = newSheet.id;
}

// 根据列名查找工作表
function getSheetByColumn(state, columnName) {
  for (const workbook of state.workbooks) {
    for (const sheet of workbook.sheets) {
      if (sheet.headers && sheet.headers.some(h => h.name === columnName)) {
        return sheet;
      }
    }
  }
  return null;
}

// 解析AI响应
function parseAiResponse(response, apiUrl) {
  const lowerApiUrl = apiUrl.toLowerCase();
  let content = "";

  // 兼容 OpenAI / DeepSeek 等
  if (lowerApiUrl.includes("openai") || lowerApiUrl.includes("deepseek")) {
    if (response.choices && response.choices[0] && response.choices[0].message) {
      content = response.choices[0].message.content;
    }
  } 
  // 添加其他平台的解析逻辑, 例如 Anthropic
  else if (lowerApiUrl.includes("anthropic")) {
    if (response.content && response.content[0] && response.content[0].text) {
      content = response.content[0].text;
    }
  } 
  // 默认或通用解析
  else if (response.content) {
    content = response.content;
  } else if (response.data) {
    content = response.data;
  } else {
    // 如果都匹配不上，返回原始响应字符串
    content = JSON.stringify(response);
  }

  // 清理从AI响应中提取的JSON字符串
  // 移除markdown代码块标记
  const jsonRegex = /```json\s*([\s\S]*?)\s*```/;
  const match = content.match(jsonRegex);
  if (match && match[1]) {
    return match[1].trim();
  }

  // 如果没有找到markdown块，直接返回清理过的字符串
  return content.trim();
}

// 执行合并操作
async function executeMerge(state, plan, config, contentEl) {
  if (contentEl) {
    contentEl.innerHTML += `<div>操作：正在基于方案执行“合并”...</div>`;
  }

  const { merge_pairs, match_type } = plan;

  if (!merge_pairs || merge_pairs.length !== 2) {
    throw new Error("执行方案无效：合并操作需要正好两个选择项。");
  }

  // 验证选择项
  const selections = merge_pairs.map(p => {
    const sheet = getSheetById(state, p.sheet_id);
    if (!sheet) {
      throw new Error(`执行方案无效：无法找到 ID 为“${p.sheet_id}”的工作表。`);
    }
    if (!sheet.headers || p.column_index >= sheet.headers.length) {
      throw new Error(`执行方案无效：在工作表“${sheet.name}”中无法找到索引为 ${p.column_index} 的列。`);
    }
    return { sheetId: p.sheet_id, columnIndex: p.column_index };
  });

  if (match_type === 'exact') {
    // 执行现有的精确匹配逻辑
    if (contentEl) {
      contentEl.innerHTML += `<div>操作：正在执行精确匹配合并...</div>`;
    }
    const newSheetData = mergeSheets(state, selections);
    addSheet(state, { ...newSheetData, id: `sheet-${Date.now()}` });
  } else if (match_type === 'fuzzy') {
    // 执行新的模糊匹配逻辑
    if (contentEl) {
      contentEl.innerHTML += `<div>操作：正在执行 AI 模糊匹配合并...</div>`;
    }
    await executeFuzzyMerge(state, selections, config, contentEl);
  } else {
    throw new Error(`未知的合并匹配类型: ${match_type}`);
  }

  // 刷新视图
  const gridRoot = document.getElementById("sheet-grid");
  const tabsRoot = document.getElementById("sheet-tabs");
  createGrid(gridRoot, tabsRoot, state);

  if (contentEl) {
    contentEl.innerHTML += `<div>操作：合并完成，已生成新工作表。</div>`;
  }
}

async function executeFuzzyMerge(state, selections, config, contentEl) {
  const [selectionA, selectionB] = selections;

  const sheetA = getSheetById(state, selectionA.sheetId);
  const sheetB = getSheetById(state, selectionB.sheetId);

  // 提取完整数据和列数据
  const dataA = sheetA.data.slice(1); // all rows except header
  const dataB = sheetB.data.slice(1);
  const columnAData = dataA.map((row, index) => ({ originalIndex: index, value: row[selectionA.columnIndex] }));
  const columnBData = dataB.map((row, index) => ({ originalIndex: index, value: row[selectionB.columnIndex] }));

  // 分批进行模糊匹配
  const batchSize = 50;
  const allMatchedPairs = [];

  for (let i = 0; i < columnAData.length; i += batchSize) {
    const batchA = columnAData.slice(i, i + batchSize);
    if (contentEl) {
      contentEl.innerHTML += `<div>AI 思考：正在为 ${sheetA.name} 的 ${i + 1}-${i + batchA.length} 行数据寻找匹配项...</div>`;
    }

    const fuzzyMatchPayload = {
      prompt: `你是一个数据匹配专家。我有一个主列表 A 和一个用于查找的列表 B。你的任务是为列表 A 中的每一项，从列表 B 中找到语义最相近的一项。请返回一个 JSON 数组，其中每个对象包含 A 和 B 中匹配项的索引。例如: [{"index_a": 0, "index_b": 42}, {"index_a": 1, "index_b": 15}]。如果找不到好的匹配项，请不要包含该项。`,
      context: `列表 A (需要为其查找匹配):\n${JSON.stringify(batchA.map(item => ({ index: item.originalIndex, value: item.value })))}\n\n列表 B (用于查找匹配):\n${JSON.stringify(columnBData.map(item => ({ index: item.originalIndex, value: item.value })))}`,
    };

    const response = await callAiApi(config, fuzzyMatchPayload, contentEl);
    const matchedPairs = JSON.parse(parseAiResponse(response, config.apiUrl));
    if (matchedPairs && matchedPairs.length > 0) {
      allMatchedPairs.push(...matchedPairs);
    }
  }

  if (allMatchedPairs.length === 0) {
    throw new Error("AI 未能找到任何模糊匹配项。");
  }

  // 构建新表
  const headersA = sheetA.headers.map(h => h.name);
  const headersB = sheetB.headers.map(h => h.name);
  const mergedHeaders = [...headersA, ...headersB.filter((h, i) => i !== selectionB.columnIndex)];

  const mergedData = [mergedHeaders];
  allMatchedPairs.forEach(pair => {
    const rowA = dataA[pair.index_a];
    const rowB = dataB[pair.index_b];
    if (rowA && rowB) {
      const mergedRow = [...rowA, ...rowB.filter((cell, i) => i !== selectionB.columnIndex)];
      mergedData.push(mergedRow);
    }
  });

  const newSheetData = {
    name: "模糊合并结果",
    data: mergedData,
    rows: mergedData.length,
    cols: mergedHeaders.length,
    headers: mergedHeaders.map((name, i) => ({
      colIndex: i,
      name,
      colLabel: String.fromCharCode(65 + i),
    }))
  };

  addSheet(state, { ...newSheetData, id: `sheet-${Date.now()}` });
}

// 执行“生成新列”操作
async function executeGenerateColumn(state, plan, config, contentEl) {
  if (contentEl) {
    contentEl.innerHTML += `<div>操作：正在基于方案执行“生成新列”...</div>`;
  }

  const {
    target_sheet_id: targetSheetId,
    reference_columns: refColumnIndexes,
    generation_prompt: generationPrompt,
    new_column_name: newColumnName,
  } = plan;

  const targetSheet = getSheetById(state, targetSheetId);
  if (!targetSheet) {
    throw new Error(`执行方案无效：无法找到 ID 为“${targetSheetId}”的工作表。`);
  }

  // 1. 提取参考列的数据
  const refColumnNames = refColumnIndexes.map(index => targetSheet.headers[index]?.name || `列 ${index + 1}`).join(", ");
  if (contentEl) {
    contentEl.innerHTML += `<div>操作：提取参考列 (${refColumnNames}) 的数据...</div>`;
  }

  const refData = targetSheet.data.slice(1).map((row, rowIndex) => ({
    rowIndex,
    values: refColumnIndexes.map(index => row[index]),
  }));

  if (refData.length === 0) {
    if (contentEl) {
      contentEl.innerHTML += `<div>警告：没有可用于生成新列的数据行。</div>`;
    }
    return; // 如果没有数据行，则提前退出
  }

  // 2. 分批处理并调用 AI 生成新数据
  const batchSize = 100;
  const generatedValues = [];

  for (let i = 0; i < refData.length; i += batchSize) {
    const batch = refData.slice(i, i + batchSize);
    if (contentEl) {
      contentEl.innerHTML += `<div>AI 思考：正在处理 ${i + 1} - ${i + batch.length} 行数据，以生成新列“${newColumnName}”...</div>`;
    }

    const generationPayload = {
      prompt: `你是一个数据处理专家。根据以下参考数据和用户要求，为每一行生成一个新的值。
      用户要求: "${generationPrompt}"
      参考列: ${refColumnNames}`,
      context: `请为下面的每一项数据生成一个新值。返回一个JSON数组，其中每个元素是对应输入行的新值。
      例如，如果输入是 [{"rowIndex":0,"values":["a","b"]},{"rowIndex":1,"values":["c","d"]}]，你应该返回 ["新值1", "新值2"]。
      当前批次的数据:
      ${JSON.stringify(batch)}`,
    };

    const response = await callAiApi(config, generationPayload, contentEl);
    const parsedResp = parseAiResponse(response, config.apiUrl);
    console.log(`批次 ${i / batchSize + 1} - AI 解析后的响应:`, parsedResp);
    const batchGeneratedValues = JSON.parse(parsedResp);
    console.log(`批次 ${i / batchSize + 1} - JSON 解析后的值:`, batchGeneratedValues);
    
    if (batchGeneratedValues && batchGeneratedValues.length > 0) {
      generatedValues.push(...batchGeneratedValues);
    } else {
      // 如果AI未能为批次返回任何值，则用空字符串填充
      generatedValues.push(...Array(batch.length).fill(""));
    }
  }

  console.log("AI 处理完成，所有批次生成的总值为:", generatedValues);
  if (contentEl) {
    contentEl.innerHTML += `<div>AI 处理完成：共生成 ${generatedValues.length} 个新值。</div>`;
    contentEl.innerHTML += `<pre>${JSON.stringify(generatedValues, null, 2)}</pre>`;
  }

  // 3. 将新列添加到工作表
  if (contentEl) {
    contentEl.innerHTML += `<div>操作：正在将新列“${newColumnName}”添加到工作表“${targetSheet.name}”...</div>`;
  }

  // 添加新表头
  const newHeader = {
    colIndex: targetSheet.headers.length,
    name: newColumnName,
    colLabel: String.fromCharCode(65 + targetSheet.headers.length),
    style: { highlight: true }, // 添加高亮标记
  };
  targetSheet.headers.push(newHeader);

  // 更新表头行
  if (targetSheet.data.length > 0) {
    targetSheet.data[0].push(newColumnName);
  } else {
    // 如果工作表是空的, 创建一个只包含新列名的表头
    targetSheet.data.push([newColumnName]);
  }

  // 添加新列数据到数据行
  const dataRows = targetSheet.data.slice(1);
  dataRows.forEach((row, index) => {
    // 确保即使AI返回的数据少于预期，也不会出错
    row.push(generatedValues[index] || "");
  });

  // 更新工作表元数据
  targetSheet.cols = targetSheet.headers.length;

  // 4. 刷新UI并强制重新计算布局
  const gridRoot = document.getElementById("sheet-grid");
  const tabsRoot = document.getElementById("sheet-tabs");
  if (gridRoot && tabsRoot) {
    // 强制重新渲染，这将触发宽度的重新计算
    createGrid(gridRoot, tabsRoot, state);
  }
  updateColumnPanel(state);

  if (contentEl) {
    contentEl.innerHTML += `<div>操作完成！新列已添加。</div>`;
  }
}


