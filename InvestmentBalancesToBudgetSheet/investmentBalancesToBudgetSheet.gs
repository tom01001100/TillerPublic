function investmentBalancesToBudgetSheet() {
    var sourceSpreadsheetId = "aY2Lrvvq8BYrxjyD8EhltKpoQB6bO1NMRkgmAGaTjHRW"; // Change this to your source sheet's (Investment Sheet) ID
    var sourceSheetName = "Balance History"; // Source spreadsheet (Investment Sheet)
    var targetSheetName = "Balance History"; // Destination sheet (Budget Sheet)
  
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
      var accountName = data[i][3]; // Column D is 'Account' on my sheet (zero-based index is 3) - Change this to your column (A is 1, B is 2, etc.)
  
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
  
    // Step 1: Insert new rows between row 2 and 3 (some conditional formatting or arrayformulas may not like row 2)
    targetSheet.insertRowsBefore(3, numNewRows);
  
    // Step 2: Copy row 2 and paste it to the last new empty row
    var row2Values = targetSheet.getRange(2, 1, 1, targetSheet.getLastColumn()).getValues();
    targetSheet.getRange(2 + numNewRows, 1, 1, row2Values[0].length).setValues(row2Values);
  
    // Step 3: Paste new unique account data into row 2
    targetSheet.getRange(2, 1, numNewRows, filteredRows[0].length).setValues(filteredRows);
  
    Logger.log("Successfully inserted " + numNewRows + " new rows and copied formatting.");
  }