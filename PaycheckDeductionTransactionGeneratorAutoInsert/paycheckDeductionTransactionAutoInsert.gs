// Add or remove sheet names here to control which sheets this script is allowed to run on.
const ALLOWED_SHEET_NAMES = [
  "Paycheck-Adam",
  "Paycheck-Eve",
  "Paycheck Deductions Generator",
];

function paycheckDeductionTransactionAutoInsert() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentSheet = ss.getActiveSheet();
    const sheetName = currentSheet.getName();

    // Check if the sheet name is one of my "Paycheck Deduction Transaction Generator" sheets.
    if (!ALLOWED_SHEET_NAMES.includes(sheetName)) {
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

    // Finding the "Transaction ID" column index, so we can insert unique IDs if that column exists.
    // Strip any $ signs (e.g. "$AT$4" -> "AT4") before parsing.
    // Then build a header range that is 1 row above rangeStart, spanning through rangeEnd's column.
    const startMatch = rangeStart.replace(/\$/g, "").match(/^([A-Za-z]+)(\d+)$/);
    const endMatch = rangeEnd.replace(/\$/g, "").match(/^([A-Za-z]+)(\d+)$/);
    if (!startMatch || !endMatch) {
      SpreadsheetApp.getUi().alert("Error: AQ10 and AQ12 must contain valid cell addresses (e.g. AT4 or $AT$4).");
      return;
    }
    const startCol = startMatch[1];
    const headerRow = parseInt(startMatch[2], 10) - 1;
    const endCol = endMatch[1];
    const headerRangeAddress = `${startCol}${headerRow}:${endCol}${headerRow}`;
    const headers = currentSheet.getRange(headerRangeAddress).getValues()[0];
    const txnIdColIndex = headers.findIndex(h => h === "Transaction ID");

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

      // If a "Transaction ID" column was found in the headers, populate each row with a unique UUID.
      if (txnIdColIndex !== -1) {
        for (let i = 0; i < data.length; i++) {
          if (!data[i][txnIdColIndex]) {
            data[i][txnIdColIndex] = Utilities.getUuid();
          }
        }
      }

      // Paste the data into the Transactions sheet starting at B3
      // Your Transactions sheet may be different, but I use the first column for indicators, hence column B
      transactionsSheet
        .getRange(3, 2, data.length, data[0].length) // Row 3 and Column 2 (B3)
        .setValues(data);

    } catch (e) {
      SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
    }
  }