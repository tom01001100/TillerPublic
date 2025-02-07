function investmentBalancesToBudgetSheet() {
  // ========================================
  // Must Review and Update These Variables
  // ========================================
  var sourceSpreadsheetId = "aY2Lrvvq8BYrxjyD8EhltKpoQB6bO1NMRkgmAGaTjHRW"; // Change this to your source sheet's (Investment Sheet) ID
  var accountColumnIndex = 3; // Column D (3) is "Account" column on my sheet. Update so the number is for the Account column on your sheet. (zero-based index: A=0, B=1, C=2, D=3, etc.)
  
  // ==============================================
  // Less likely to need changes to these variables
  // ==============================================
  var sourceSheetName = "Balance History"; // Source spreadsheet (Investment Sheet)
  var targetSheetName = "Balance History"; // Destination sheet (Budget Sheet)

  // ==============================================
  // Script Logic - No changes needed below this point
  // ==============================================
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);

  if (!sourceSheet || !targetSheet) {
    Logger.log("One or both sheets not found!");
    return;
  }

  var data = sourceSheet.getRange(2, 1, 500, sourceSheet.getLastColumn()).getValues(); // Read first 500 rows (starting from row 2)
  var uniqueAccounts = new Set();
  var filteredRows = [];

  // Iterate through rows and store unique accounts
  for (var i = 0; i < data.length; i++) {
    var accountName = data[i][accountColumnIndex]; // Use configurable column for "Account"

    if (!uniqueAccounts.has(accountName)) {
      uniqueAccounts.add(accountName);
      filteredRows.push(data[i]);
    }
  }

  var numNewRows = filteredRows.length;
  if (numNewRows === 0) {
    Logger.log("No new unique rows to insert.");
    return;
  }

  // Step 1: Insert new rows between row 2 and 3 (some conditional formatting or array formulas may not like row 2)
  targetSheet.insertRowsBefore(3, numNewRows);

  // Step 2: Copy row 2 and paste it to the last new empty row
  var row2Values = targetSheet.getRange(2, 1, 1, targetSheet.getLastColumn()).getValues();
  targetSheet.getRange(2 + numNewRows, 1, 1, row2Values[0].length).setValues(row2Values);

  // Step 3: Paste new unique account data into row 2
  targetSheet.getRange(2, 1, numNewRows, filteredRows[0].length).setValues(filteredRows);

  Logger.log("Successfully inserted " + numNewRows);
}
