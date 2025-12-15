/**
 * ============================================================
 * LIFTSHEET TEMPLATE AUTOMATION
 * Version: 1.3
 * Last Updated: 2025-12
 * ============================================================
 */

const CONFIG = {
  SOURCE_SHEET_NAME: "Choice", 
  // 监听的列号数组: C=3, R=18, AE=31, AS=45
  TRIGGER_COLUMNS: [3, 18, 31, 45], 
  
  MAPPINGS: {
    // Main Lifts
    "Main Squat":      "B2:N7",
    "Main Bench":      "B9:N14",
    "Main Deadlift":   "B16:N21",

    // Secondary Lifts
    "Secondary Squat":    "B23:N28",
    "Secondary Bench":    "B30:N35",
    "Secondary Deadlift": "B37:N42",

    // Tertiary Lifts (坐标已更新)
    "Tertiary Bench":     "B44:N49", 

    // Reset / Default (坐标已更新)
    "Choice of Mainlift":    "B51:N56",
    // Accessories
    "Shoulder":           "C58:K58",
    "Push":               "C59:K59",
    "Horizontal Row":     "C61:K61",
    "Vertical Row":       "C62:K62",
    "Anterior chain":     "C64:K64",
    "Posterior chain":    "C65:K65",
    "Choice of Accessories":"C67:K67"
  }
};

/**
 * 主动监听函数
 */
function onEdit(e) {
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const range = e.range;
  
  if (sheet.getName() === CONFIG.SOURCE_SHEET_NAME) return;
  
  // 检查编辑的列号是否在指定的监听列数组中
  if (!CONFIG.TRIGGER_COLUMNS.includes(range.getColumn())) return; 

  const value = e.value; 
  
  if (!value || !CONFIG.MAPPINGS.hasOwnProperty(value)) return;

  replaceTemplate(sheet, range, value);
}

/**
 * 执行复制粘贴逻辑
 */
function replaceTemplate(activeSheet, activeRange, keyword) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`错误：找不到名为 "${CONFIG.SOURCE_SHEET_NAME}" 的工作表。`);
    return;
  }

  const sourceRangeNotation = CONFIG.MAPPINGS[keyword];
  const sourceRange = sourceSheet.getRange(sourceRangeNotation);
  
  const targetRange = activeRange.offset(0, -1);
  
  // Define accessory keywords that should paste directly to the active cell
  const accessoryKeywords = [
    "Shoulder", "Push", "Horizontal Row", "Vertical Row", 
    "Anterior chain", "Posterior chain", "Choice of Accessories"
  ];

  let pasteTargetRange;
  if (accessoryKeywords.includes(keyword)) {
    pasteTargetRange = activeRange; // Paste directly to the active cell
  } else {
    pasteTargetRange = activeRange.offset(0, -1); // Paste one cell to the left
  }
  
  sourceRange.copyTo(pasteTargetRange);
}