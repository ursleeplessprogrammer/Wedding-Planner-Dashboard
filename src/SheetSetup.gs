/**
 * Sets up the Google Sheet for wedding planning by creating
 * and populating essential sheets with their respective column headers.
 *
 * The 'Home' sheet is ensured to be the first sheet.
 * Other sheets are created/updated with predefined headers and basic formatting.
 */
function setupWeddingSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let message = "Sheet creation and header setup complete.\n\n";

  // Define all sheets and their headers
  // The 'Home' sheet is handled separately as it's a dashboard and doesn't have traditional column headers.
  const sheetDefinitions = {
    'Tasks': [
      'Task Name', 'Category', 'Due Date', 'Assigned To', 'Status', 'Priority', 'Notes', 'Completion Date'
    ],
    'Vendors': [
      'Vendor Category', 'Vendor Name', 'Contact Person', 'Phone', 'Email', 'Website',
      'Contract Signed?', 'Deposit Paid?', 'Total Cost', 'Amount Paid', 'Balance Due',
      'Payment Due Date', 'Notes', 'Link to Contract/Invoice'
    ],
    'Guest List': [
      'Guest Name(s)', 'Address', 'Email', 'Phone', 'Invited?', 'RSVP Status',
      'Number Attending', 'Dietary Restrictions/Allergies', 'Table Number',
      'Gift Received?', 'Thank You Sent?', 'Relationship'
    ],
    'Budget': [
      'Category', 'Item/Description', 'Estimated Cost', 'Actual Cost', 'Difference',
      'Paid To', 'Payment Date', 'Payment Status', 'Notes'
    ],
    'Itinerary': [
      'Time', 'Activity', 'Location', 'Responsible Person/Vendor', 'Notes'
    ],
    'Inspiration': [
      'Category', 'Idea/Description', 'Link', 'Notes'
    ]
  };

  // --- Handle the 'Home' sheet first ---
  // Ensure 'Home' exists and is the first sheet
  let homeSheet = ss.getSheetByName('Home');
  if (!homeSheet) {
    homeSheet = ss.insertSheet('Home', 0); // Create as the first sheet
    message += "Created 'Home' sheet.\n";
  } else {
    // If 'Home' exists but isn't the first sheet, move it
    if (homeSheet.getIndex() !== 1) { // getIndex() is 1-based
      ss.setActiveSheet(homeSheet);
      ss.moveActiveSheet(1); // Move to the first position
      message += "Moved 'Home' sheet to the first position.\n";
    } else {
      message += "'Home' sheet already exists and is in the correct position.\n";
    }
  }

  // --- Create/Update other sheets and set headers ---
  for (const sheetName in sheetDefinitions) {
    let sheet = ss.getSheetByName(sheetName);
    const headers = sheetDefinitions[sheetName];

    if (!sheet) {
      // Create new sheet if it doesn't exist
      sheet = ss.insertSheet(sheetName);
      message += `Created '${sheetName}' sheet.\n`;
    } else {
      // If sheet exists, clear the first row to ensure clean header placement on re-run
      if (sheet.getLastColumn() > 0) {
          sheet.getRange(1, 1, 1, sheet.getLastColumn()).clearContent();
      }
      message += `Updated '${sheetName}' sheet with new headers.\n`;
    }

    // --- Check if there are headers before trying to set a range ---
    if (headers.length > 0) {
      // Set headers in the first row
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // Apply basic formatting to headers for readability
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setWrap(true)
           .setVerticalAlignment('middle')
           .setBackground('#cfe2f3'); // Light blue background for headers

      // Auto-resize columns to fit content
      sheet.autoResizeColumns(1, headers.length);
    } else {
      console.warn(`Sheet '${sheetName}' was created but has no headers defined in the script.`);
      message += `Warning: '${sheetName}' sheet created but has no headers defined in the script.\n`;
    }

    // Freeze the first row
    sheet.setFrozenRows(1);
  }
}
