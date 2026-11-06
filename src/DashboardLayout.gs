/**
 * Helper function to apply multiple styles to a range.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range to style.
 * @param {Object} styles An object containing style properties.
 */
function applyStyle(range, styles) {
  if (styles.fontSize) range.setFontSize(styles.fontSize);
  if (styles.fontWeight) range.setFontWeight(styles.fontWeight);
  if (styles.background) range.setBackground(styles.background);
  if (styles.fontColor) range.setFontColor(styles.fontColor);
  if (styles.horizontalAlignment) range.setHorizontalAlignment(styles.horizontalAlignment);
  if (styles.verticalAlignment) range.setVerticalAlignment(styles.verticalAlignment);
  if (styles.wrap) range.setWrap(styles.wrap);
  return range;
}

/**
 * Sets up the layout and initial static content for the 'Home' (Dashboard) sheet,
 * including inserting all dynamic formulas and hyperlinks, with an optimized, clean UI.
 */
function setupDashboardLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const homeSheet = ss.getSheetByName('Home');
  if (!homeSheet) {
    Logger.log("Home sheet not found. Cannot set up dashboard layout.");
    return;
  }

  // --- Define an optimized color palette and styles ---
  const PALETTE = {
    BACKGROUND: '#fcfcfc',     // Very light, almost white background for the whole sheet
    PRIMARY_LIGHT: '#e0f2f1',   // Soft mint/sage green
    PRIMARY_MEDIUM: '#a7d9d6',  // Medium mint/sage green
    PRIMARY_DARK: '#00796b',    // Dark teal/green for text
    ACCENT_LIGHT: '#f0f4f7',    // Light grey-blue for subtle highlights
    ACCENT_MEDIUM: '#d0d9e0',   // Medium grey-blue for subtle borders
    TEXT_PRIMARY: '#333333',    // Dark grey for main text
    TEXT_SECONDARY: '#777777',  // Medium grey for labels/secondary text
    LINK: '#1976d2',            // Standard blue for links
    WARNING: '#fff3e0',         // Light orange for warnings
    ERROR: '#ffebee',           // Light red for errors
    SUCCESS: '#e8f5e9'          // Light green for success
  };

  const FONT_SETTINGS = {
    TITLE: { fontSize: 36, fontWeight: 'bold', fontColor: PALETTE.PRIMARY_DARK, horizontalAlignment: 'center', verticalAlignment: 'middle' },
    SUBTITLE: { fontSize: 18, fontWeight: 'normal', fontColor: PALETTE.TEXT_SECONDARY, horizontalAlignment: 'center', verticalAlignment: 'middle' },
    SECTION_HEADER: { fontSize: 18, fontWeight: 'bold', fontColor: PALETTE.TEXT_PRIMARY, horizontalAlignment: 'left', verticalAlignment: 'middle' },
    KEY_METRIC_LABEL: { fontSize: 14, fontWeight: 'normal', fontColor: PALETTE.TEXT_SECONDARY, horizontalAlignment: 'left', verticalAlignment: 'middle' },
    KEY_METRIC_VALUE_LARGE: { fontSize: 24, fontWeight: 'bold', fontColor: PALETTE.TEXT_PRIMARY, horizontalAlignment: 'right', verticalAlignment: 'middle' },
    KEY_METRIC_VALUE_MEDIUM: { fontSize: 18, fontWeight: 'bold', fontColor: PALETTE.TEXT_PRIMARY, horizontalAlignment: 'right', verticalAlignment: 'middle' },
    TASK_LIST: { fontSize: 12, fontColor: PALETTE.TEXT_PRIMARY, wrap: true, verticalAlignment: 'top' },
    NAVIGATION_LINK: { fontSize: 14, fontColor: PALETTE.LINK, horizontalAlignment: 'left', verticalAlignment: 'middle' }
  };

  // Clear existing content and formatting
  homeSheet.clear();
  homeSheet.clearConditionalFormatRules();
  // Explicitly remove all embedded charts on the sheet
  const charts = homeSheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    homeSheet.removeChart(charts[i]);
  }

  // Set overall sheet background
  homeSheet.getRange('A:Z').setBackground(PALETTE.BACKGROUND);

  // Set default column widths and row heights for better spacing
  homeSheet.setColumnWidth(1, 40);  // Left Spacer (A)
  homeSheet.setColumnWidth(2, 250); // Content Column 1 (labels) (B)
  homeSheet.setColumnWidth(3, 150); // Content Column 1 (values/progress) (C)
  homeSheet.setColumnWidth(4, 60);  // Middle Spacer (D) - Wider for separation
  homeSheet.setColumnWidth(5, 250); // Content Column 2 (labels) (E)
  homeSheet.setColumnWidth(6, 150); // Content Column 2 (values/progress) (F)
  homeSheet.setColumnWidth(7, 40);  // Right Spacer (G)

  homeSheet.setRowHeight(1, 30);  // Top padding
  homeSheet.setRowHeight(4, 25);  // Spacer below main title
  homeSheet.setRowHeight(7, 25);  // Spacer below countdown
  homeSheet.setRowHeight(10, 25); // Spacer below progress
  homeSheet.setRowHeight(17, 25); // Spacer below summaries
  homeSheet.setRowHeight(24, 25); // Spacer below tasks/deadlines
  homeSheet.setRowHeight(29, 30); // Bottom padding

  // --- Main Title ---
  let range = homeSheet.getRange('B2:F2').merge().setValue('IVAN & [PARTNER\'S NAME] WEDDING PLANNER');
  applyStyle(range, FONT_SETTINGS.TITLE).setBackground(PALETTE.BACKGROUND);
  range = homeSheet.getRange('B3:F3').merge().setValue('Your Central Hub for Elegant Wedding Preparations');
  applyStyle(range, FONT_SETTINGS.SUBTITLE).setBackground(PALETTE.BACKGROUND);


  // --- Wedding Date & Countdown ---
  homeSheet.getRange('B5').setValue('Wedding Date:');
  applyStyle(homeSheet.getRange('B5'), FONT_SETTINGS.KEY_METRIC_LABEL);
  range = homeSheet.getRange('C5:D5').merge().setValue('YYYY-MM-DD'); // Placeholder for manual date input
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_LARGE).setHorizontalAlignment('left').setBackground(PALETTE.ACCENT_LIGHT);

  homeSheet.getRange('E5').setValue('Days Until Wedding:');
  applyStyle(homeSheet.getRange('E5'), FONT_SETTINGS.KEY_METRIC_LABEL);
  range = homeSheet.getRange('F5');
  range.setFormula('=IF(ISBLANK(C5),"",DATEDIF(TODAY(),C5,"D")&" days")'); // Formula for countdown
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_LARGE).setBackground(PALETTE.ACCENT_LIGHT);
  homeSheet.getRange('B5:F5').setBorder(false, false, true, false, false, false, PALETTE.ACCENT_MEDIUM, SpreadsheetApp.BorderStyle.SOLID); // Subtle bottom border


  // --- Overall Planning Progress ---
  range = homeSheet.getRange('B8:F8').merge().setValue('Overall Planning Progress');
  applyStyle(range, FONT_SETTINGS.SECTION_HEADER).setBackground(PALETTE.BACKGROUND);

  // Cell for percentage (B9) and sparkline (C9)
  range = homeSheet.getRange('B9');
  range.setFormula('=IFERROR(COUNTIF(Tasks_Status,"Completed")/COUNTA(Tasks_Status),"0%")'); // Formula for percentage
  range.setNumberFormat('0%'); // Format as percentage
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_LARGE).setHorizontalAlignment('left').setBackground(PALETTE.ACCENT_LIGHT);

  range = homeSheet.getRange('C9:F9').merge();
  range.setFormula('=SPARKLINE(B9,{"charttype","bar";"max",1;"color1","' + PALETTE.PRIMARY_DARK + '"})'); // Formula for sparkline
  applyStyle(range, {background: PALETTE.ACCENT_LIGHT}); // Apply background to the merged range
  homeSheet.getRange('B8:F9').setBorder(false, false, true, false, false, false, PALETTE.ACCENT_MEDIUM, SpreadsheetApp.BorderStyle.SOLID); // Subtle bottom border


  // --- Budget Overview ---
  range = homeSheet.getRange('B11:C11').merge().setValue('💰 Budget Overview');
  applyStyle(range, FONT_SETTINGS.SECTION_HEADER).setBackground(PALETTE.BACKGROUND);

  homeSheet.getRange('B12').setValue('Total Estimated:');
  applyStyle(homeSheet.getRange('B12'), FONT_SETTINGS.KEY_METRIC_LABEL);
  range = homeSheet.getRange('C12');
  range.setFormula('=IFERROR(SUM(Budget_EstimatedCost),0)'); // Formula
  range.setNumberFormat('Rp #,##0'); // Currency format
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_MEDIUM).setBackground(PALETTE.ACCENT_LIGHT);

  homeSheet.getRange('B13').setValue('Total Actual Spent:');
  applyStyle(homeSheet.getRange('B13'), FONT_SETTINGS.KEY_METRIC_LABEL);
  range = homeSheet.getRange('C13');
  range.setFormula('=IFERROR(SUM(Budget_ActualCost),0)'); // Formula
  range.setNumberFormat('Rp #,##0'); // Currency format
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_MEDIUM).setBackground(PALETTE.ACCENT_LIGHT);

  homeSheet.getRange('B14').setValue('Remaining Budget:');
  applyStyle(homeSheet.getRange('B14'), FONT_SETTINGS.KEY_METRIC_LABEL);
  range = homeSheet.getRange('C14');
  range.setFormula('=IFERROR(C12-C13,0)'); // Formula
  range.setNumberFormat('Rp #,##0'); // Currency format
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_MEDIUM).setBackground(PALETTE.ACCENT_LIGHT);
  homeSheet.getRange('B11:C14').setBorder(true, true, true, true, true, true, PALETTE.ACCENT_MEDIUM, SpreadsheetApp.BorderStyle.SOLID); // Box border


  // --- Guest List Summary ---
  range = homeSheet.getRange('E11:F11').merge().setValue('👥 Guest List Summary');
  applyStyle(range, FONT_SETTINGS.SECTION_HEADER).setBackground(PALETTE.BACKGROUND);

  homeSheet.getRange('E12').setValue('Total Invited:');
  applyStyle(homeSheet.getRange('E12'), FONT_SETTINGS.KEY_METRIC_LABEL);
  range = homeSheet.getRange('F12');
  range.setFormula('=IFERROR(COUNTA(GuestList_GuestNames),0)'); // Formula
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_MEDIUM).setBackground(PALETTE.ACCENT_LIGHT);

  homeSheet.getRange('E13').setValue('Total Confirmed:');
  applyStyle(homeSheet.getRange('E13'), FONT_SETTINGS.KEY_METRIC_LABEL);
  range = homeSheet.getRange('F13');
  range.setFormula('=IFERROR(COUNTIF(GuestList_RSVPStatus,"Yes"),0)'); // Formula
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_MEDIUM).setBackground(PALETTE.ACCENT_LIGHT);

  homeSheet.getRange('E14').setValue('Total Declined:');
  applyStyle(homeSheet.getRange('E14'), FONT_SETTINGS.KEY_METRIC_LABEL);
  range = homeSheet.getRange('F14');
  range.setFormula('=IFERROR(COUNTIF(GuestList_RSVPStatus,"No"),0)'); // Formula
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_MEDIUM).setBackground(PALETTE.ACCENT_LIGHT);

  homeSheet.getRange('E15').setValue('Dietary Restrictions:');
  applyStyle(homeSheet.getRange('E15'), FONT_SETTINGS.KEY_METRIC_LABEL);
  range = homeSheet.getRange('F15');
  range.setFormula('=IFERROR(COUNTIF(GuestList_DietaryRestrictions,"<>"),0)'); // Formula
  applyStyle(range, FONT_SETTINGS.KEY_METRIC_VALUE_MEDIUM).setBackground(PALETTE.ACCENT_LIGHT);
  homeSheet.getRange('E11:F15').setBorder(true, true, true, true, true, true, PALETTE.ACCENT_MEDIUM, SpreadsheetApp.BorderStyle.SOLID); // Box border


  // --- Urgent Tasks ---
  range = homeSheet.getRange('B19:D19').merge().setValue('🚨 Urgent Tasks');
  applyStyle(range, FONT_SETTINGS.SECTION_HEADER).setBackground(PALETTE.BACKGROUND);
  range = homeSheet.getRange('B20:D24').merge();
  const urgentTasksFormula = '=IFERROR(TEXTJOIN(CHAR(10), TRUE, ARRAY_CONSTRAIN(SORT(FILTER(Tasks_TaskName & " (Due: " & TEXT(Tasks_DueDate, "MM/DD") & ")", (Tasks_Status<>"Completed") * (Tasks_Status<>"On Hold") * (Tasks_DueDate<=TODAY()+60), (Tasks_Priority="High") + (Tasks_DueDate<TODAY())), 2, TRUE), 10, 1)), "No urgent tasks! Enjoy the calm.")';
  range.setFormula(urgentTasksFormula);
  applyStyle(range, FONT_SETTINGS.TASK_LIST).setHorizontalAlignment('left').setBackground(PALETTE.WARNING);
  homeSheet.getRange('B19:D24').setBorder(true, true, true, true, true, true, PALETTE.ACCENT_MEDIUM, SpreadsheetApp.BorderStyle.SOLID); // Box border


  // --- Upcoming Deadlines ---
  range = homeSheet.getRange('E19:F19').merge().setValue('🗓️ Upcoming Deadlines');
  applyStyle(range, FONT_SETTINGS.SECTION_HEADER).setBackground(PALETTE.BACKGROUND);
  range = homeSheet.getRange('E20:F24').merge();
  const upcomingDeadlinesFormula = '=IFERROR(TEXTJOIN(CHAR(10), TRUE, ARRAY_CONSTRAIN(SORT(FILTER({Tasks_TaskName & " (Task - " & TEXT(Tasks_DueDate, "MM/DD") & ")"; Vendors_Name & " (Payment - " & TEXT(Vendors_PaymentDueDate, "MM/DD") & ")"}, {Tasks_Status<>"Completed"; TRUE} * {Tasks_DueDate>=TODAY(); Vendors_PaymentDueDate>=TODAY()} * {Vendors_PaymentDueDate<=TODAY()+90; Vendors_PaymentDueDate<=TODAY()+90}), 2, TRUE), 10, 1)), "No major deadlines in the next 90 days!")';
  range.setFormula(upcomingDeadlinesFormula);
  applyStyle(range, FONT_SETTINGS.TASK_LIST).setHorizontalAlignment('left').setBackground(PALETTE.ACCENT_LIGHT);
  homeSheet.getRange('E19:F24').setBorder(true, true, true, true, true, true, PALETTE.ACCENT_MEDIUM, SpreadsheetApp.BorderStyle.SOLID); // Box border


  // --- Quick Navigation ---
  range = homeSheet.getRange('B27:F27').merge().setValue('🚀 Quick Navigation');
  applyStyle(range, FONT_SETTINGS.SECTION_HEADER).setBackground(PALETTE.BACKGROUND);

  // Navigation Links
  const navLinks = [
    { cell: 'B28', text: '📋 Tasks Checklist', sheet: 'Tasks' },
    { cell: 'C28', text: '🤝 Vendor Management', sheet: 'Vendors' },
    { cell: 'D28', text: '👥 Guest List & RSVP', sheet: 'Guest List' },
    { cell: 'E28', text: '💰 Budget Tracker', sheet: 'Budget' },
    { cell: 'F28', text: '🗓️ Wedding Day Itinerary', sheet: 'Itinerary' },
    { cell: 'B29', text: '💡 Inspiration & Notes', sheet: 'Inspiration' },
    { cell: 'C29', text: '🔗 Wedding Website', sheet: 'https://yourweddingwebsite.com' } // Added website link here
  ];

  navLinks.forEach(link => {
    range = homeSheet.getRange(link.cell);
    if (link.sheet.startsWith('http')) { // External link
      range.setFormula('=HYPERLINK("' + link.sheet + '","' + link.text + '")');
    } else { // Internal sheet link
      range.setFormula('=HYPERLINK("#gid="&SHEETID("' + link.sheet + '"),"' + link.text + '")');
    }
    applyStyle(range, FONT_SETTINGS.NAVIGATION_LINK).setBackground(PALETTE.PRIMARY_LIGHT);
    range.setBorder(true, true, true, true, true, true, PALETTE.ACCENT_MEDIUM, SpreadsheetApp.BorderStyle.SOLID);
  });
  homeSheet.getRange('B27:F29').setBorder(true, true, true, true, true, true, PALETTE.ACCENT_MEDIUM, SpreadsheetApp.BorderStyle.SOLID); // Box border

  // Hide gridlines for a cleaner dashboard look
  homeSheet.setHiddenGridlines(true);
}
