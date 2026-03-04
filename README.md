## 智能 Excel 操作工具（H5）

这是一个基于 H5 的智能 Excel 操作工具项目骨架，包含类似 Excel 的表格界面以及 AI 操作入口和配置界面。

### 功能规划（后续可实现）

- **单元格选择**：支持单选和框选多个单元格。
- **AI 操作面板**：对选中单元格发送提示词（Prompt），执行如：
  - 根据选中数据合并到另一张“工作表”
  - 根据描述批量修改单元格内容
- **配置界面**：
  - 配置大模型 API 地址
  - 配置 Secret 等安全信息（仅保存在本地浏览器，示例实现）

### 项目结构

- `index.html`：入口页面（Excel 风格布局）
- `src/`
  - `styles/excel.css`：主要样式（表格 + 工具栏 + 配置面板）
  - `js/app.js`：入口脚本，初始化 UI
  - `js/ui/grid.js`：Excel 风格网格组件
  - `js/ui/toolbar.js`：顶部工具栏 + Prompt 输入区域
  - `js/ui/configPanel.js`：右侧/弹出配置面板
  - `js/services/aiClient.js`：调用大模型 API 的封装
  - `js/services/state.js`：全局状态（当前选择、配置等）
- `config/ai-config.sample.json`：API 配置示例文件
- `package.json`：前端依赖与脚本（可按需扩展）

### 本地开发（示例）

```bash
npm install
# 使用一个静态服务器（例如 serve、http-server）或 VSCode / Cursor 插件打开本目录
# 之后浏览器访问 http://localhost:PORT/index.html
```

后续你可以在当前骨架基础上逐步实现选择单元格、调用 AI 和配置管理等具体逻辑。

