/**
 * 统一菜单管理器
 * 文件名: MenuManager.gs
 * 
 * 这个文件统一管理所有自定义菜单
 * 所有其他代码文件中的 onOpen() 函数都可以删除或注释掉
 */

/**
 * 工作簿打开时自动执行
 * 这是唯一的 onOpen() 函数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // ============================================
  // 主菜单: MyScript
  // ============================================
  ui.createMenu('⚙️ MyScript')
    // 📊 数据同步功能组
    .addSubMenu(ui.createMenu('📊 Data Sync')
      .addItem('🔄 Sync Performance Data', 'syncPerformanceData')
      .addSeparator()
      .addItem('📋 View All Sheets', 'listAllSheets'))
    
    .addSeparator()
    
// 🤖 AI 助手功能组
    .addSubMenu(ui.createMenu('🤖 AI Assistant')
      .addItem('💬 Open Chat Window', 'openAISidebar')
      .addItem('❌ Close Chat Window', 'closeAISidebar')
      .addSeparator()
      .addItem('🔑 Configure API Key', 'showAPIKeyConfigDialog')
      .addItem('👀 View Current Key', 'showCurrentAPIKey')
      .addSeparator()
      .addItem('🔌 Test API Connection', 'testAPIConnection'))
    
    .addSeparator()
    
    // 🛠️ 工具箱
    .addSubMenu(ui.createMenu('🛠️ Toolbox')
      .addItem('🧹 Clear Cache', 'clearCache')
      .addItem('📝 View Logs', 'viewLogs')
      .addItem('ℹ️ About', 'showAbout'))
    
    .addToUi();
  
  Logger.log('✅ Menu initialized');
}

/**
 * 手动重新加载菜单
 * 如果菜单没有正常显示，可以运行这个函数
 */
function reloadMenu() {
  onOpen();
  SpreadsheetApp.getUi().alert('Menu reloaded', 'Please check the top menu bar', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================
// 工具箱功能实现
// ============================================

/**
 * 清理缓存（预留功能）
 */
function clearCache() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear Cache',
    'Are you sure you want to remove all cached data?',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    // 这里添加清理缓存的代码
    // 例如：清理脚本属性、临时数据等
    const properties = PropertiesService.getScriptProperties();
    properties.deleteAllProperties();
    
    ui.alert('✅ Done', 'Cache cleared', ui.ButtonSet.OK);
    Logger.log('Cache cleared');
  }
}

/**
 * 查看日志
 */
function viewLogs() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'View Logs',
    'Open the Apps Script editor and click\n"View" → "Logs"\n\nor use the shortcut Ctrl+Enter (Windows)\nCmd+Enter (Mac)',
    ui.ButtonSet.OK
  );
}

/**
 * 关于信息
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  const version = '1.0.0';
  const lastUpdate = '2024-11-14';
  
  const aboutText = 
    '📊 Google Sheets Toolkit\n\n' +
    'Version: ' + version + '\n' +
    'Last update: ' + lastUpdate + '\n\n' +
    'Modules:\n' +
    '• Data sync utilities\n' +
    '• AI chat assistant (DeepSeek)\n' +
    '• Toolbox utilities\n\n' +
    '💡 Tip: Explore more features in the menu!';
  
  ui.alert('About MyScript', aboutText, ui.ButtonSet.OK);
}

// ============================================
// 菜单配置说明
// ============================================

/**
 * 如何添加新菜单项：
 * 
 * 1. 在 onOpen() 函数中找到对应的位置
 * 2. 使用 .addItem('显示名称', '函数名') 添加
 * 3. 使用 .addSeparator() 添加分隔线
 * 4. 使用 .addSubMenu() 添加子菜单
 * 
 * 示例：
 * .addItem('🔥 新功能', 'newFeatureFunction')
 * 
 * 注意事项：
 * - 确保函数名存在且可访问
 * - 使用 emoji 让菜单更直观
 * - 相关功能归类到同一子菜单
 * - 使用分隔线区分不同功能组
 */

/**
 * 菜单结构说明：
 * 
 * ⚙️ MyScript (主菜单)
 *   ├── 📊 数据同步 (子菜单)
 *   │   ├── 🔄 同步表现数据
 *   │   ├── ---
 *   │   ├── 📋 查看所有工作表
 *   │   └── ⚡ 快速同步
 *   ├── ---
 *   ├── 🤖 AI 助手 (子菜单)
 *   │   ├── 💬 打开对话窗口
 *   │   ├── ❌ 关闭对话窗口
 *   │   ├── ---
 *   │   └── 🔌 测试 API 连接
 *   ├── ---
 *   └── 🛠️ 工具箱 (子菜单)
 *       ├── 🧹 清理缓存
 *       ├── 📝 查看日志
 *       └── ℹ️ 关于
 */
