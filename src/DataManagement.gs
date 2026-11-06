/**
 * Applies named ranges to all relevant columns in the detailed sheets
 * and sets up data validation dropdowns for specific columns.
 */
function applyNamedRangesAndValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let message = "Named ranges and data validation applied.\n\n";

  // --- Named Ranges Definitions ---
  // Format: { sheetName: { headerText: namedRangeSuffix } }
  // Named ranges make it easier to reference data in formulas (e.g., Tasks_Status instead of Tasks!E:E)
  const namedRangeDefinitions = {
    'Tasks': {
      'Task Name': 'TaskName', 'Category': 'Category', 'Due Date': 'DueDate',
      'Assigned To': 'AssignedTo', 'Status': 'Status', 'Priority': 'Priority',
      'Notes': 'Notes', 'Completion Date': 'CompletionDate'
    },
    'Vendors': {
      'Vendor Category': 'Category', 'Vendor Name': 'Name', 'Contact Person': 'ContactPerson',
      'Phone': 'Phone', 'Email': 'Email', 'Website': 'Website',
      'Contract Signed?': 'ContractSigned', 'Deposit Paid?': 'DepositPaid',
      'Total Cost': 'TotalCost', 'Amount Paid': 'AmountPaid', 'Balance Due': 'BalanceDue',
      'Payment Due Date': 'PaymentDueDate', 'Notes': 'Notes',
      'Link to Contract/Invoice': 'ContractLink'
    },
    'Guest List': {
      'Guest Name(s)': 'GuestNames', 'Address': 'Address', 'Email': 'Email',
      'Phone': 'Phone', 'Invited?': 'Invited', 'RSVP Status': 'RSVPStatus',
      'Number Attending': 'NumberAttending', 'Dietary Restrictions/Allergies': 'DietaryRestrictions',
      'Table Number': 'TableNumber', 'Gift Received?': 'GiftReceived',
      'Thank You Sent?': 'ThankYouSent', 'Relationship': 'Relationship'
    },
    'Budget': {
      'Category': 'Category', 'Item/Description': 'ItemDescription', 'Estimated Cost': 'EstimatedCost',
      'Actual Cost': 'ActualCost', 'Difference': 'Difference', 'Paid To': 'PaidTo',
      'Payment Date': 'PaymentDate', 'Payment Status': 'PaymentStatus', 'Notes': 'Notes'
    },
    'Itinerary': {
      'Time': 'Time', 'Activity': 'Activity', 'Location': 'Location',
      'Responsible Person/Vendor': 'ResponsiblePerson', 'Notes': 'Notes'
    },
    'Inspiration': {
      'Category': 'Category', 'Idea/Description': 'IdeaDescription', 'Link': 'Link', 'Notes': 'Notes'
    }
  };

  // --- Data Validation Definitions ---
  // Format: { sheetName: { headerText: [option1, option2, ...] } }
  const validationDefinitions = {
    'Tasks': {
      'Category': ['Venue', 'Food', 'Dress', 'Decor', 'Guests', 'Ceremony', 'Reception', 'Photography', 'Music', 'Stationery', 'Rings', 'Honeymoon', 'Other'],
      'Status': ['Not Started', 'In Progress', 'Completed', 'On Hold', 'Overdue'],
      'Priority': ['High', 'Medium', 'Low']
    },
    'Vendors': {
      'Vendor Category': ['Venue', 'Caterer', 'Photographer', 'Videographer', 'DJ', 'Band', 'Florist', 'Dress Shop', 'Suit Rental', 'Hair Stylist', 'Makeup Artist', 'Officiant', 'Stationery', 'Cake', 'Transportation', 'Favors', 'Rings', 'Planner', 'Other']
    },
    'Guest List': {
      'RSVP Status': ['Yes', 'No', 'Pending', 'Declined']
    },
    'Budget': {
      'Category': ['Venue', 'Catering', 'Attire', 'Photography', 'Videography', 'Music', 'Flowers', 'Stationery', 'Rings', 'Favors', 'Transportation', 'Hair & Makeup', 'Officiant', 'Decor', 'Honeymoon', 'Miscellaneous'],
      'Payment Status': ['Paid', 'Deposit Paid', 'Due', 'Overdue']
    }
  };

  // Iterate through each sheet to apply named ranges and data validation
  for (const sheetName in namedRangeDefinitions) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet '${sheetName}' not found for named range/validation setup.`);
      continue;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Apply Named Ranges
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      const namedRangeSuffix = namedRangeDefinitions[sheetName][header];
      if (namedRangeSuffix) {
        // Remove spaces from sheetName for cleaner named range names
        const rangeName = `${sheetName.replace(/\s/g, '')}_${namedRangeSuffix}`; // e.g., Tasks_Status
        // Range from row 2 to the maximum number of rows in the sheet, for the current column
        const range = sheet.getRange(2, i + 1, sheet.getMaxRows() - 1, 1);
        ss.setNamedRange(rangeName, range);
        // message += `  - Named range '${rangeName}' created for ${sheetName} > ${header}.\n`; // Uncomment for detailed logging
      }
    }

    // Apply Data Validation
    const sheetValidationDefs = validationDefinitions[sheetName];
    if (sheetValidationDefs) {
      for (let i = 0; i < headers.length; i++) {
        const header = headers[i];
        const validationOptions = sheetValidationDefs[header];
        if (validationOptions) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(validationOptions)
            .setAllowInvalid(false) // Don't allow values not in the list
            .setHelpText(`Select a valid ${header}.`)
            .build();
          // Apply to entire column from row 2 to the maximum number of rows
          const range = sheet.getRange(2, i + 1, sheet.getMaxRows() - 1, 1);
          range.setDataValidation(rule);
          // message += `  - Data validation applied to ${sheetName} > ${header}.\n`; // Uncomment for detailed logging
        }
      }
    }
  }
}
