function paycheckDeductionTransactionAutoInsert() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentSheet = ss.getActiveSheet();
    const sheetName = currentSheet.getName();

    // Check if the sheet name is one of my "Paycheck Deduction Transaction Generator" sheets.
    if (sheetName !== "Paycheck-Adam" && sheetName !== "Paycheck-Eve" && sheetName !== "Paycheck Deductions Generator") {
      SpreadsheetApp.getUi().alert(
        `Error: This script can only be run on Paychecks sheets.`
      );
      return; // Exit the script without doing anything
    }

    const transactionsSheet = ss.getSheetByName("Transactions");
    if (!transactionsSheet) {
      SpreadsheetApp.getUi().alert("Error: Sheet 'Transactions' does not exist.");
      return;
    }

    // Dynamically concatenate values from AQ10 and AQ12 to form the range address
    // These are the two ranges in the "Select $AT$4 Through $BO$16" section.
    const rangeStart = currentSheet.getRange("AQ10").getValue();
    const rangeEnd = currentSheet.getRange("AQ12").getValue();

    // Make a range with full colon
    const rangeAddress = `${rangeStart}:${rangeEnd}`;

    if (!rangeStart || !rangeEnd) {
      SpreadsheetApp.getUi().alert("Error: Cells AQ10 and AQ12 must not be empty.");
      return;
    }

    // Get the number of rows to insert from AQ5
    // This is the number in the "Insert nn Rows in Transactions Sheet" section.
    const rowsToInsert = currentSheet.getRange("AQ5").getValue();
    if (!Number.isInteger(rowsToInsert) || rowsToInsert < 0) {
      SpreadsheetApp.getUi().alert("Error: Cell AQ5 must contain a non-negative integer.");
      return;
    }

    try {
      // Insert rows between row 2 and 3 in the Transactions sheet
      // I have had issues with conditional formatting or arrayformulas if I try to insert at the very top.
      // Feel free to do this differently if you like to live dangerously.
      if (rowsToInsert > 0) {
        transactionsSheet.insertRowsAfter(2, rowsToInsert);
      }

      // Get the data from the specified concatenated range
      const range = currentSheet.getRange(rangeAddress);
      const data = range.getValues();

      // Paste the data into the Transactions sheet starting at B3
      // Your Transactions sheet may be different, but I use the first column for indicators, hence column B
      transactionsSheet
        .getRange(3, 2, data.length, data[0].length) // Row 3 and Column 2 (B3)
        .setValues(data);

    } catch (e) {
      SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
    }
  }
