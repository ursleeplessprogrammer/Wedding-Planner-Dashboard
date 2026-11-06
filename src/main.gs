/**
 * Main function to set up the entire wedding planning spreadsheet.
 * Runs all necessary setup functions in order.
 *
 * To run: Select 'initializeWeddingPlanner' from the dropdown in the Apps Script editor
 * and click the 'Run' button.
 */
function initializeWeddingPlanner() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Initializing Wedding Planner', 'This script will set up your sheets, headers, named ranges, data validation, and explanation page. Please allow permissions if prompted.', ui.ButtonSet.OK);

  setupWeddingSheets(); // Creates sheets and sets basic headers
  setupDashboardLayout(); // Configures the Home sheet layout
  applyNamedRangesAndValidation(); // Creates named ranges and dropdowns
  setupExplanationSheet(); // <<< ADD THIS LINE

  ui.alert('Setup Complete!', 'Your wedding planner dashboard is now ready! Please review the "Home" sheet and fill in sample data in other sheets to see the dashboard come alive. Check the "Explanation" sheet for guidance!', ui.ButtonSet.OK);
}
