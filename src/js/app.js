import { createGrid } from "./ui/grid.js";
import { initToolbar } from "./ui/toolbar.js";
import { initConfigPanel } from "./ui/configPanel.js";
import { initColumnPanel } from "./ui/columnPanel.js";
import { initMergeModal } from "./ui/mergeModal.js";
import { initState } from "./services/state.js";

// 初始化全局状态
const state = initState();

// 初始化 Grid
const gridRoot = document.getElementById("sheet-grid");
const tabsRoot = document.getElementById("sheet-tabs");
createGrid(gridRoot, tabsRoot, state);

// 初始化工具栏（含 AI 操作入口）
initToolbar(state);

// 初始化配置面板（API 地址、Secret 等）
initConfigPanel(state);

// 初始化列头面板
initColumnPanel(state);

// 初始化合并模态框
initMergeModal(state);

