// 监听单元格修改事件
function onEdit(e) {
  // 基础检查，防止事件对象为空
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  // 定义需要监听的列号
  // J=10, X=24, AL=38, AZ=52
  const triggerColumns = [10, 24, 38, 52];
  
  // 检查修改的列是否在监听列表中
  if (!triggerColumns.includes(col)) {
    return;
  }
  
  // 定义有效的触发值和对应的修饰符
  const triggerValues = {
    'Optimal': 0,
    'Easy[-1]': -1,
    'Comfortable[-0.5]': -0.5,
    'Challenging[+0.5]': 0.5,
    'Hard[+1]': 1
  };
  
  // 优化：优先使用 e.value，如果不存在（移动端延迟/粘贴），则主动读取单元格值
  let value = e.value;
  
  if (value === undefined || value === null) {
    // 仅当编辑的是单个单元格时才处理，避免多行粘贴导致性能问题
    if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
      value = range.getValue();
    } else {
      return;
    }
  }
  
  // 检查输入的值是否在有效值列表中
  if (!triggerValues.hasOwnProperty(value)) {
    return;
  }
  
  // 执行修改逻辑
  executeModifications(sheet, row, col, triggerValues[value]);
}

// 执行修改操作
function executeModifications(sheet, row, col, modifier) {
  try {
    // 对应关系：触发列 -> [源列, 左边格源列, 下一行左两格源列]
    const colMappings = {
      10: { source: 6, left: 5, nextRowLeft: 7 },  // J列触发: 修改F, I改E, H12改G
      24: { source: 20, left: 19, nextRowLeft: 21 }, // X列触发: 修改T, W改S, 下一行改U
      38: { source: 34, left: 33, nextRowLeft: 35 }, // AL列触发: 修改AH, AK改AG, 下一行改AI
      52: { source: 48, left: 47, nextRowLeft: 49 }  // AZ列触发: 修改AV, AY改AU, 下一行改AW
    };
    
    const mapping = colMappings[col];
    if (!mapping) return;
    
    // ============ 第一步：获取源列F的值 ============
    const sourceValue = sheet.getRange(row, mapping.source).getValue();
    
    // 检查值是否为数字
    if (typeof sourceValue !== 'number') {
      Logger.log('源列值不是数字，操作中止');
      return;
    }
    
    // ============ 第二步：对J11本身进行更改（J11 = F11 + modifier） ============
    const newValue = sourceValue + modifier;
    sheet.getRange(row, col).setValue(newValue);
    
    // ============ 第三步：将左边一格改为E11的值 ============
    const leftCellValue = sheet.getRange(row, mapping.left).getValue();
    sheet.getRange(row, col - 1).setValue(leftCellValue);
    
    // ============ 第四步：将下一行左两格先改为Processing ============
    sheet.getRange(row + 1, col - 2).setValue('Processing');
    
    // ============ 第五步：等待1秒后改为G12的值 ============
    Utilities.sleep(1000);
    const nextRowValue = sheet.getRange(row + 1, mapping.nextRowLeft).getValue();
    sheet.getRange(row + 1, col - 2).setValue(nextRowValue);
    
    Logger.log(`操作完成: 行${row}, 列${col}, 修饰符${modifier}, 新值${newValue}`);
    
  } catch (error) {
    Logger.log('错误: ' + error.toString());
  }
}

// 手动测试函数（可选）
function testScript() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // 测试用例：在F11输入触发值
  const testCases = [
    { cell: 'J11', value: 'Optimal' },
    { cell: 'J12', value: 'Easy[-1]' },
    { cell: 'J13', value: 'Comfortable[-0.5]' },
    { cell: 'J14', value: 'Challenging[+0.5]' },
    { cell: 'J15', value: 'Hard[+1]' }
  ];
  
  testCases.forEach((testCase) => {
    sheet.getRange(testCase.cell).setValue(testCase.value);
    
    const range = sheet.getRange(testCase.cell);
    const e = {
      range: range,
      value: testCase.value,
      source: SpreadsheetApp.getActiveSpreadsheet()
    };
    
    onEdit(e);
    Logger.log(`测试: ${testCase.cell} = ${testCase.value} 完成`);
  });
}