/**
 * ============================================================
 * LIFTSHEET TEMPLATE AUTOMATION
 * Version: 1.6 (Silent & Lightweight)
 * ============================================================
 */

const CONFIG = {
  SOURCE_SHEET_NAME: "Choice", 
  // 监听列: C=3, R=18, AE=31, AS=45
  TRIGGER_COLUMNS: [3, 18, 31, 45], 
  
  // 原地粘贴列表 (Offset 0)
  DIRECT_PASTE_KEYS: [
    "Shoulder", "Push", "Horizontal Row", "Vertical Row", 
    "Anterior chain", "Posterior chain", "Choice of Accessories"
  ],

  MAPPINGS: {
    // Main Lifts (Offset -1)
    "Main Squat":      "B2:N7",
    "Main Bench":      "B9:N14",
    "Main Deadlift":   "B16:N21",

    // Secondary Lifts (Offset -1)
    "Secondary Squat":    "B23:N28",
    "Secondary Bench":    "B30:N35",
    "Secondary Deadlift": "B37:N42",

    // Tertiary Lifts (Offset -1)
    "Tertiary Bench":     "B44:N49", 

    // Reset (Offset -1)
    "Choice of Mainlift": "B51:N56",

    // Accessories (Offset 0)
    "Shoulder":           "C58:K58",
    "Push":               "C59:K59",
    "Horizontal Row":     "C61:K61",
    "Vertical Row":       "C62:K62",
    "Anterior chain":     "C64:K64",
    "Posterior chain":    "C65:K65",
    "Choice of Accessories":"C67:K67"
  }
};

function onEdit(e) {
  // 极简过滤：无值、无范围、不在监听列则直接退出
  if (!e || !e.range || !e.value) return;
  if (!CONFIG.TRIGGER_COLUMNS.includes(e.range.getColumn())) return;
  
  // 检查是否为 Choice 表
  const sheet = e.range.getSheet();
  if (sheet.getName() === CONFIG.SOURCE_SHEET_NAME) return;

  // 检查值是否有效
  if (!CONFIG.MAPPINGS.hasOwnProperty(e.value)) return;

  // 执行核心逻辑
  runSilentPaste(e.source, sheet, e.range, e.value);
}

function runSilentPaste(spreadsheet, activeSheet, activeRange, keyword) {
  try {
    const sourceSheet = spreadsheet.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
    if (!sourceSheet) return;

    const sourceRange = sourceSheet.getRange(CONFIG.MAPPINGS[keyword]);
    
    // 判断粘贴位置：Accessories 原地粘贴(0)，Main Lifts 向左一格(-1)
    const offsetCol = CONFIG.DIRECT_PASTE_KEYS.includes(keyword) ? 0 : -1;
    const targetRange = activeRange.offset(0, offsetCol);
    
    // 执行粘贴
    sourceRange.copyTo(targetRange);
    
  } catch (error) {
    // 静默失败：不弹窗，只在后台记录日志（用户不可见，除非打开编辑器查看）
    console.error("AutoPaste Error:", error.message);
  }
}