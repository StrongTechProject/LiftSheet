这里是为你准备的 `choice.md` 文档。它包含了脚本的完整代码、配置逻辑说明以及维护指南，非常适合保存在项目仓库或本地文件夹中，供后续迭代查阅。

------

# Choice Template Automation Script

## 1. 项目简介

此脚本用于 Google Sheets 的训练计划模板（LiftSheet）。它监听特定列（C列）的编辑事件，根据用户选择的动作名称（如 `Main Squat`），自动从 `Choice` 模板表中拉取对应的预设格式和数据，并覆盖粘贴到当前行。

**主要功能：**

- **自动监听**：无需手动运行，修改 C 列即触发。
- **模板注入**：根据关键词将 `Choice` 表中预定义的 B:N 区域块复制到当前位置。
- **重置功能**：支持通过选择 `Choice of 3Lift` 将区域还原为初始状态。

------

## 2. 核心代码

**文件名**：`Code.gs`

JavaScript

```
/**
 * ============================================================
 * LIFTSHEET TEMPLATE AUTOMATION
 * Version: 1.1
 * Last Updated: 2025-12
 * ============================================================
 */

/**
 * [配置区域]
 * 定义关键词与 "Choice" 表中源区域的映射关系。
 * 维护时只需修改此处的 MAPPINGS 对象。
 */
const CONFIG = {
  SOURCE_SHEET_NAME: "Choice", // 存放模板的源工作表名称
  TRIGGER_COLUMNS: [3, 18, 31, 45],           // 监听的列号 (C=3, R=18, AE=31, AS=45)
  
  // 格式: "下拉菜单选项值" : "Choice表中对应的源区域(B:N)"
 // 格式: "下拉菜单选项值" : "Choice表中对应的源区域(B:N)"
  // 格式: "下拉菜单选项值" : "Choice表中对应的源区域(B:N)"
  MAPPINGS: {
    // Main Lifts
    "Main Squat":      "B2:N7",
    "Main Bench":      "B9:N14",
    "Main Deadlift":   "B16:N21",

    // Secondary Lifts
    "Secondary Squat":    "B23:N28",
    "Secondary Bench":    "B30:N35",
    "Secondary Deadlift": "B37:N42",

    // Tertiary Lifts
    "Tertiary Bench":     "B44:N49", // Updated

    // Reset / Default
    "Choice of Mainlift":    "B51:N56"  // Updated
  }
};

/**
 * [触发器] 主动监听函数
 * 当表格被用户编辑时自动运行
 */
function onEdit(e) {
  // 1. 安全检查：确保事件对象存在
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const range = e.range;
  
  // 2. 过滤逻辑：
  // - 忽略源模板表本身 (防止在编辑模板时意外触发)
  // - 仅监听指定的列 (C, R, AE, AS)
  if (sheet.getName() === CONFIG.SOURCE_SHEET_NAME) return;
  if (range.getColumn() !== CONFIG.TRIGGER_COLUMN) return;

  // 3. 获取输入值
  const value = e.value; 
  
  // 4. 匹配逻辑：如果值为空或不在配置表中，则退出
  if (!value || !CONFIG.MAPPINGS.hasOwnProperty(value)) return;

  // 5. 执行替换
  replaceTemplate(sheet, range, value);
}

/**
 * [核心逻辑] 执行复制粘贴
 */
function replaceTemplate(activeSheet, activeRange, keyword) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
  
  // 错误处理：源表不存在
  if (!sourceSheet) {
    console.error(`Error: Source sheet "${CONFIG.SOURCE_SHEET_NAME}" not found.`);
    return;
  }

  // 获取源区域
  const sourceRangeNotation = CONFIG.MAPPINGS[keyword];
  const sourceRange = sourceSheet.getRange(sourceRangeNotation);
  
  // 定位目标区域：
  // 监听的是 C 列，但模板是从 B 列开始的，所以向左偏移 1 列 (-1)
  const targetRange = activeRange.offset(0, -1);
  
  // 执行复制 (包含值、公式、格式、数据验证等所有属性)
  sourceRange.copyTo(targetRange);
}
```

------

## 3. 维护指南

### 如何新增一个动作模板？

如果未来需要增加新的动作类型（例如 "Main Overhead Press"）：

1. **在 Sheet 中操作**：

   - 进入 `Choice` 表。
   - 找一块空白区域（例如 `B56:N61`），设计好你的模板样式、公式和格式。

2. **在代码中操作**：

   - 打开 `Extensions` > `Apps Script`。

   - 在 `CONFIG.MAPPINGS` 对象中添加一行：

     JavaScript

     ```
     "Main Overhead Press": "B56:N61",
     ```

   - 保存脚本 (`Cmd/Ctrl + S`)。

### 关于 "Choice of Mainlift" 的重置逻辑

- **用途**：当用户选错了动作，想退回到未选择的状态时使用。
- **注意**：`Choice` 表中的 `B49:N54` 区域应当被设计为“初始状态”（例如：灰色背景、仅 C 列有文字，其余为空或占位符）。
- **安全性**：脚本执行的粘贴操作**不会**再次触发 `onEdit`，因此即使 `B49:N54` 模板里的 C 列包含 "Choice of Mainlift" 文字，也不会导致无限循环。

## 4. 常见问题 (FAQ)

- **Q: 为什么选择了菜单脚本没反应？**
  - A1: 检查 `Choice` 表是否存在。
  - A2: 检查输入的文字是否与代码中的 `Key` 完全一致（**区分大小写**，注意单词间的空格）。
  - A3: 只有**手动编辑**会触发。如果是通过其他脚本修改的值，不会触发此脚本。
- **Q: 脚本可以粘贴到 C 列右边的区域吗？**
  - A: 可以。目前代码设置为 `offset(0, -1)` (向左偏移1列到 B 列)。如果模板需要从 C 列开始粘贴，将代码修改为 `offset(0, 0)` 即可。