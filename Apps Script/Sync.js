function syncPerformanceData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Prompt user to input period identifier (e.g., M1B1)
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Enter Training Period',
    'Please enter format like: M1B1 (Month-Block)',
    ui.ButtonSet.OK_CANCEL
  );
  
  // Check if user clicked cancel
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  var rawInput = response.getResponseText().trim();
  if (!rawInput) {
    ui.alert('Error', 'Input cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  var upperInput = rawInput.toUpperCase();
  var compactInput = upperInput.replace(/\s+/g, '');
  
  // Validate input format (e.g., M1B1) even if extra text exists
  var regex = /M(\d+)B(\d+)/i;
  var match = upperInput.match(regex);
  
  if (!match) {
    ui.alert('Error', 'Invalid format! Please include pattern like M1B1', ui.ButtonSet.OK);
    return;
  }
  
  var month = match[1];
  var block = match[2];
  var baseCode = 'M' + month + 'B' + block;
  
  // Try exact sheet match first
  var sourceSheet = ss.getSheetByName(rawInput);
  var sourceSheetName = sourceSheet ? sourceSheet.getName() : null;
  
  if (!sourceSheet) {
    var sheets = ss.getSheets();
    var candidates = [];
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      var sheetUpper = sheetName.toUpperCase();
      var sheetCompact = sheetUpper.replace(/\s+/g, '');
      if (sheetUpper.indexOf(upperInput) !== -1 ||
          sheetCompact.indexOf(compactInput) !== -1 ||
          sheetUpper.indexOf(baseCode) !== -1) {
        candidates.push(sheetName);
      }
    }
    
    if (candidates.length === 0) {
      ui.alert('Error', 'No sheet matched input or code: ' + rawInput, ui.ButtonSet.OK);
      return;
    }
    
    if (candidates.length === 1) {
      sourceSheetName = candidates[0];
    } else {
      var optionsText = 'Multiple sheets matched:\n\n';
      for (var j = 0; j < candidates.length; j++) {
        optionsText += (j + 1) + '. ' + candidates[j] + '\n';
      }
      optionsText += '\nEnter the number to use.';
      var chooseResponse = ui.prompt('Select Sheet', optionsText, ui.ButtonSet.OK_CANCEL);
      if (chooseResponse.getSelectedButton() != ui.Button.OK) {
        return;
      }
      var choice = parseInt(chooseResponse.getResponseText(), 10);
      if (isNaN(choice) || choice < 1 || choice > candidates.length) {
        ui.alert('Error', 'Invalid selection.', ui.ButtonSet.OK);
        return;
      }
      sourceSheetName = candidates[choice - 1];
    }
    
    sourceSheet = ss.getSheetByName(sourceSheetName);
  }
  
  // Detect training frequency from sheet name
  var frequency = detectFrequency(sourceSheetName);
  
  // If frequency not detected, ask user
  if (!frequency) {
    var freqResponse = ui.prompt(
      'Select Training Frequency',
      'Training frequency not detected in sheet name.\nPlease enter: 3, 4, or 5 (Days per week)',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (freqResponse.getSelectedButton() != ui.Button.OK) {
      return;
    }
    
    var freqInput = freqResponse.getResponseText().trim();
    if (freqInput === '3' || freqInput === '4' || freqInput === '5') {
      frequency = parseInt(freqInput);
    } else {
      ui.alert('Error', 'Invalid frequency! Please enter 3, 4, or 5.', ui.ButtonSet.OK);
      return;
    }
  }
  
  var confirm = ui.alert(
    'Confirm Source Sheet',
    'Use sheet "' + sourceSheetName + '" for syncing?\n' +
    '(Detected code: ' + baseCode + ')\n' +
    '(Training frequency: ' + frequency + ' Days)',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm != ui.Button.OK) {
    return;
  }
  
  var sheetNameForFormula = sourceSheetName.replace(/'/g, "''");
  
  // Get Performance Tracker sheet
  var performanceSheet = ss.getSheetByName('📈Performance Tracker');
  if (!performanceSheet) {
    ui.alert('Error', 'Sheet not found: 📈Performance Tracker', ui.ButtonSet.OK);
    return;
  }
  
  // Get Volume Tracker sheet
  var volumeSheet = ss.getSheetByName('📊Volume Tracker');
  if (!volumeSheet) {
    ui.alert('Error', 'Sheet not found: 📊Volume Tracker', ui.ButtonSet.OK);
    return;
  }
  
  var m = parseInt(month);
  var b = parseInt(block);
  
  // Get volume tracker row numbers based on frequency
  var volumeRows = getVolumeRows(frequency);
  
  // Sync all 4 weeks (W1-W4)
  var syncedWeeks = [];
  
  for (var week = 1; week <= 4; week++) {
    // Calculate source column index based on week
    // Week1: C (col 3), Week2: Q (col 17), Week3: AE (col 31), Week4: AS (col 45)
    // Interval is 14 columns
    var sourceColIndex = 3 + (week - 1) * 14;
    var sourceCol = columnToLetter(sourceColIndex);
    
    // Calculate D column (col 4) for Week series
    var sourceColIndex_D = 4 + (week - 1) * 14;
    var sourceCol_D = columnToLetter(sourceColIndex_D);
    
    // Calculate E column series for E1RM
    // Week1: E (col 5), Week2: S (col 19), Week3: AG (col 33), Week4: AU (col 47)
    var sourceColIndex2 = 5 + (week - 1) * 14;
    var sourceCol2 = columnToLetter(sourceColIndex2);
    
    // Calculate I column (col 9) for Volume Tracker
    var sourceColIndex_I = 9 + (week - 1) * 14;
    var sourceCol_I = columnToLetter(sourceColIndex_I);
    
    // Calculate target row for Performance Tracker
    // M1B1W1 -> row 9, M1B1W2 -> row 10, M1B1W3 -> row 11, M1B1W4 -> row 12
    // M1B2W1 -> row 13, M1B2W2 -> row 14...
    // M1B3W1 -> row 17...
    // M2B1W1 -> row 21...
    
    // Base row: 9 (M1B1W1 starting row)
    // Each Block has 4 rows (4 weeks)
    // Each Month has 3 Blocks, so each Month has 12 rows
    var performanceRow = 9 + (m - 1) * 12 + (b - 1) * 4 + (week - 1);
    
    // Calculate target row for Volume Tracker
    // M1B1W1 -> row 4, M1B1W2 -> row 5, M1B1W3 -> row 6, M1B1W4 -> row 7
    // Base row: 4
    var volumeRow = 4 + (m - 1) * 12 + (b - 1) * 4 + (week - 1);
    
    // ========== Performance Tracker ==========
    // Write formulas to D:F (Weekly Top Load: Squat, Bench, Deadlift)
    performanceSheet.getRange(performanceRow, 4).setFormula("='" + sheetNameForFormula + "'!" + sourceCol + "4");  // D: Squat
    performanceSheet.getRange(performanceRow, 5).setFormula("='" + sheetNameForFormula + "'!" + sourceCol + "5");  // E: Bench
    performanceSheet.getRange(performanceRow, 6).setFormula("='" + sheetNameForFormula + "'!" + sourceCol + "6");  // F: Deadlift
    
    // Write formulas to H:J (Weekly E1RM: Squat, Bench, Deadlift)
    performanceSheet.getRange(performanceRow, 8).setFormula("='" + sheetNameForFormula + "'!" + sourceCol2 + "4");  // H: Squat E1RM
    performanceSheet.getRange(performanceRow, 9).setFormula("='" + sheetNameForFormula + "'!" + sourceCol2 + "5");  // I: Bench E1RM
    performanceSheet.getRange(performanceRow, 10).setFormula("='" + sheetNameForFormula + "'!" + sourceCol2 + "6"); // J: Deadlift E1RM
    
    // ========== Volume Tracker ==========
    // Write formulas to D:F (from I row range based on frequency)
    volumeSheet.getRange(volumeRow, 4).setFormula("='" + sheetNameForFormula + "'!" + sourceCol_I + volumeRows.group1[0]);  // D
    volumeSheet.getRange(volumeRow, 5).setFormula("='" + sheetNameForFormula + "'!" + sourceCol_I + volumeRows.group1[1]);  // E
    volumeSheet.getRange(volumeRow, 6).setFormula("='" + sheetNameForFormula + "'!" + sourceCol_I + volumeRows.group1[2]);  // F
    
    // Write formulas to G:I (from C row range based on frequency)
    volumeSheet.getRange(volumeRow, 7).setFormula("='" + sheetNameForFormula + "'!" + sourceCol + volumeRows.group2[0]);   // G
    volumeSheet.getRange(volumeRow, 8).setFormula("='" + sheetNameForFormula + "'!" + sourceCol + volumeRows.group2[1]);   // H
    volumeSheet.getRange(volumeRow, 9).setFormula("='" + sheetNameForFormula + "'!" + sourceCol + volumeRows.group2[2]);   // I
    
    // Write formulas to J:L (from C row range based on frequency)
    volumeSheet.getRange(volumeRow, 10).setFormula("='" + sheetNameForFormula + "'!" + sourceCol + volumeRows.group3[0]);   // J
    volumeSheet.getRange(volumeRow, 11).setFormula("='" + sheetNameForFormula + "'!" + sourceCol + volumeRows.group3[1]);   // K
    volumeSheet.getRange(volumeRow, 12).setFormula("='" + sheetNameForFormula + "'!" + sourceCol + volumeRows.group3[2]);   // L
    
    syncedWeeks.push('Week' + week + ' (Performance: Row ' + performanceRow + ', Volume: Row ' + volumeRow + ')');
  }
  
  ui.alert(
    '✅ Sync Completed',
    'Successfully linked data from ' + sourceSheetName + ' to:\n' +
    '• 📈Performance Tracker\n' +
    '• 📊Volume Tracker\n\n' +
    '📊 Training Frequency: ' + frequency + ' Days\n' +
    '📊 Synced Weeks:\n' +
    syncedWeeks.join('\n') + '\n\n' +
    '🔗 All formulas are now linked and will auto-update!',
    ui.ButtonSet.OK
  );
}

// Detect training frequency from sheet name
function detectFrequency(sheetName) {
  var upper = sheetName.toUpperCase();
  
  // Look for patterns like (3Days), (4Days), (5Days) or 3DAYS, 4DAYS, 5DAYS
  if (upper.indexOf('3DAY') !== -1 || upper.indexOf('3 DAY') !== -1) {
    return 3;
  }
  if (upper.indexOf('4DAY') !== -1 || upper.indexOf('4 DAY') !== -1) {
    return 4;
  }
  if (upper.indexOf('5DAY') !== -1 || upper.indexOf('5 DAY') !== -1) {
    return 5;
  }
  
  return null; // Not detected
}

// Get volume tracker row numbers based on frequency
function getVolumeRows(frequency) {
  switch(frequency) {
    case 3:
      return {
        group1: [64, 65, 66],  // I64-66
        group2: [64, 65, 66],  // C64-66
        group3: [68, 69, 70]   // C68-70
      };
    case 4:
      return {
        group1: [86, 87, 88],  // I86-88
        group2: [86, 87, 88],  // C86-88
        group3: [90, 91, 92]   // C90-92
      };
    case 5:
      return {
        group1: [94, 95, 96],  // I94-96
        group2: [94, 95, 96],  // C94-96
        group3: [98, 99, 100]  // C98-100
      };
    default:
      // Default to 4 days if something goes wrong
      return {
        group1: [86, 87, 88],
        group2: [86, 87, 88],
        group3: [90, 91, 92]
      };
  }
}

// Convert column index to column letter (e.g., 1->A, 27->AA, 52->AZ)
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// List all sheet names (for debugging)
function listAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ui = SpreadsheetApp.getUi();
  
  var sheetNames = 'All sheets in current workbook:\n\n';
  for (var i = 0; i < sheets.length; i++) {
    sheetNames += (i + 1) + '. "' + sheets[i].getName() + '"\n';
  }
  
  ui.alert('Sheet List', sheetNames, ui.ButtonSet.OK);
}

// Create custom menu
//function onOpen() {
//  var ui = SpreadsheetApp.getUi();
//  ui.createMenu('📊 Synchronize Statistics')
//    .addItem('🔄 Sync Performance Data', 'syncPerformanceData')
//    .addSeparator()
//    .addItem('📋 View All Sheet Names', 'listAllSheets')
//    .addToUi();
//}