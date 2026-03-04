# AI生成列功能修复与优化计划

## 目标

1.  **修复Bug**：解决AI生成的新列无法正确显示数据的问题。
2.  **功能增强**：为新生成的列添加明黄色的背景高亮，以作区分。

## 实现步骤

### 1. 修复新列生成逻辑

问题根源在于 `aiClient.js` 的 `executeGenerateColumn` 函数中，处理数据时将表头（Header）也一并发送给了AI，并错误地处理了返回的数据。我将按以下步骤修复它：

*   **调整数据提取逻辑**:
    *   在调用AI之前，我会修改 `executeGenerateColumn` 函数，使其仅提取和处理工作表中的数据行，而排除表头行。
    *   我将使用 `targetSheet.data.slice(1)` 来获取所有数据行，并确保传递给AI的 `refData` 不包含表头信息。

*   **修正数据合并逻辑**:
    *   在从AI获取到生成的 `generatedValues` 后，我会更新数据合并的逻辑。
    *   首先，将新的列名添加到表头行 `targetSheet.data[0]`。
    *   然后，遍历数据行（从 `targetSheet.data` 的第二行开始），将 `generatedValues` 中的值一一对应地追加到每一行的末尾。

### 2. 实现新列高亮功能

为了在视觉上突出显示新生成的列，我将实现以下功能：

*   **在 `aiClient.js` 中定义样式**:
    *   在 `executeGenerateColumn` 函数中创建新列的表头对象（`newHeader`）时，我会为其添加一个 `style` 属性，用于定义背景颜色。例如：`style: { highlight: true }`。

*   **在 `grid.js` 中应用样式**:
    *   我将修改 `src/js/ui/grid.js` 文件中的 `createGrid` 函数。
    *   在 `Handsontable` 的配置对象中，我会使用 `cells` 渲染器函数来动态地为单元格添加CSS类。
    *   该函数会检查当前单元格所在列的表头（header）是否包含 `style.highlight` 属性。如果包含，就为该单元格添加一个名为 `highlight-new-column` 的CSS类。

*   **在 `excel.css` 中添加样式**:
    *   最后，我会在 `src/styles/excel.css` 文件中添加 `highlight-new-column` 类的样式定义，将其背景色设置为明黄色。

```css
.highlight-new-column {
  background-color: #FFFF00 !important; /* 明黄色 */
}
```
