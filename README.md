# 💍 Wedding Planner Dashboard

> **A Google Sheets workspace for organizing wedding plans, tracking progress, and keeping important details in one place.**

## 📌 Project Overview

This project uses Google Apps Script to create and format a wedding-planning spreadsheet. It sets up sheets for tasks, vendors, guests, budget, itinerary, and inspiration, then builds a styled **Home** dashboard with progress, budget and guest summaries, urgent tasks, deadlines, and navigation links.

Run `initializeWeddingPlanner()` from the Apps Script editor to set up the spreadsheet.

## 🗺️ Contents

- [Features](#-features)
- [Spreadsheet Structure](#-spreadsheet-structure)
- [Setup](#-setup)
- [Run the Project](#-run-the-project)
- [Apps Script Files](#-apps-script-files)
- [Dashboard Formulas and Named Ranges](#-dashboard-formulas-and-named-ranges)
- [Customization](#️-customization)
- [Important Notes](#️-important-notes)
- [Troubleshooting](#-troubleshooting)

## ✨ Features

- Creates a Home dashboard and planning sheets.
- Shows planning progress, budget totals, and guest RSVP counts.
- Lists urgent tasks and upcoming deadlines.
- Adds navigation links to planning sheets and an optional wedding website.
- Applies formatting, column widths, and frozen header rows.
- Includes a `SHEETID()` custom spreadsheet function for internal sheet links.
- Provides an Explanation sheet when `setupExplanationSheet()` is implemented.

## 📑 Spreadsheet Structure

The setup script creates these sheets:

| Sheet | Purpose |
|---|---|
| **Home** | Dashboard with summaries, tasks, deadlines, and navigation |
| **Tasks** | Checklist, due dates, assignees, priorities, and status |
| **Vendors** | Vendor contacts, contracts, costs, payments, and notes |
| **Guest List** | Guest details, invitations, RSVPs, dietary notes, and table numbers |
| **Budget** | Estimated and actual costs, payments, and balances |
| **Itinerary** | Wedding-day schedule, locations, and responsible people |
| **Inspiration** | Ideas, descriptions, links, and notes |
| **Explanation** | Guidance for using the planner; requires `setupExplanationSheet()` |

### Sheet columns

**Tasks**

```text
Task Name, Category, Due Date, Assigned To, Status, Priority, Notes, Completion Date
```

**Vendors**

```text
Vendor Category, Vendor Name, Contact Person, Phone, Email, Website,
Contract Signed?, Deposit Paid?, Total Cost, Amount Paid, Balance Due,
Payment Due Date, Notes, Link to Contract/Invoice
```

**Guest List**

```text
Guest Name(s), Address, Email, Phone, Invited?, RSVP Status,
Number Attending, Dietary Restrictions/Allergies, Table Number,
Gift Received?, Thank You Sent?, Relationship
```

**Budget**

```text
Category, Item/Description, Estimated Cost, Actual Cost, Difference,
Paid To, Payment Date, Payment Status, Notes
```

**Itinerary**

```text
Time, Activity, Location, Responsible Person/Vendor, Notes
```

**Inspiration**

```text
Category, Idea/Description, Link, Notes
```

## ⚙️ Setup

1. Create or open a Google Spreadsheet.
2. Open **Extensions → Apps Script**.
3. Add the supplied `.gs` files to the Apps Script project.
4. Review and customize the placeholder names, date, currency, and website link.
5. Save the project.
6. In the Apps Script editor, select `initializeWeddingPlanner` from the function list.
7. Click **Run** and authorize the requested spreadsheet permissions.
8. Return to the spreadsheet and review the created sheets.

Google Apps Script uses a spreadsheet’s bound project to access the active spreadsheet and its UI.

## ▶️ Run the Project

Select and run:

```javascript
initializeWeddingPlanner()
```

The function runs the setup operations in sequence:

```text
Create planning sheets and headers
              ↓
Format the Home dashboard
              ↓
Create named ranges and validations
              ↓
Set up the Explanation sheet
```

Some referenced functions must exist in the project. In the supplied `main.gs`, these are called:

```javascript
setupWeddingSheets();
setupDashboardLayout();
applyNamedRangesAndValidation();
setupExplanationSheet();
```

The supplied files include `setupWeddingSheets()` and `setupDashboardLayout()`, but **do not include** `applyNamedRangesAndValidation()` or `setupExplanationSheet()`. Add those functions or remove their calls before running `initializeWeddingPlanner()`. Otherwise, Apps Script will stop with a “function is not defined” error.

## 🧩 Apps Script Files

### `main.gs`

Contains the main setup function, `initializeWeddingPlanner()`. It displays setup messages and calls the other setup functions.

### `SheetSetup.gs`

Creates or finds the planning sheets, writes their header rows, applies basic header formatting, resizes columns, and freezes the first row.

### `DashboardLayout.gs`

Defines the dashboard style helper and configures the Home sheet’s layout, formulas, summary sections, task panels, and navigation links.

### `DataManagement.gs`

The supplied version duplicates `applyStyle()` and `setupDashboardLayout()` from `DashboardLayout.gs`. Retain these functions in **one file only**. If `DataManagement.gs` is intended for named ranges and validation, put those functions there instead.

### `Utilities.gs`

Defines the custom `SHEETID(sheetName)` spreadsheet function, which returns a sheet’s numeric ID for dashboard navigation formulas.

## 📊 Dashboard Formulas and Named Ranges

The dashboard references named ranges including:

```text
Tasks_Status
Tasks_TaskName
Tasks_DueDate
Tasks_Priority
Vendors_Name
Vendors_PaymentDueDate
Budget_EstimatedCost
Budget_ActualCost
GuestList_GuestNames
GuestList_RSVPStatus
GuestList_DietaryRestrictions
```

The supplied code calls `applyNamedRangesAndValidation()`, but its implementation is not included. These named ranges must be created and mapped to the appropriate sheet columns for dashboard formulas to work.

The dashboard uses formulas for:

- Countdown to the wedding date
- Overall task completion
- Estimated and actual budget
- Guest RSVP totals
- Urgent tasks
- Upcoming payment and task deadlines

A formula may show an error or a fallback message if its named ranges are missing or if the underlying data does not match the expected values.

## 🎨 Customization

### Couple names and dashboard text

In `DashboardLayout.gs`, replace:

```javascript
'[YOUR'S NAME] & [PARTNER\'S NAME] WEDDING PLANNER'
```

with your preferred title.

You can also customize the subtitle:

```javascript
'Your Central Hub for Elegant Wedding Preparations'
```

### Wedding date

The dashboard places a `YYYY-MM-DD` placeholder in the wedding-date area. Replace it with a valid spreadsheet date. The countdown formula uses that date.

### Currency

Budget values use the format:

```javascript
'Rp #,##0'
```

Change this to match your currency and preferred number format.

### Wedding website

Replace:

```javascript
'https://yourweddingwebsite.com'
```

with your actual website URL, or remove that navigation item.

### Categories and dropdowns

Task, vendor, guest, payment, and priority dropdowns depend on the missing `applyNamedRangesAndValidation()` function. Define the desired lists there and ensure the named ranges match the dashboard formulas.

## ⚠️ Important Notes

- **Duplicate functions:** `DashboardLayout.gs` and `DataManagement.gs` both define `applyStyle()` and `setupDashboardLayout()`. Remove one copy of each.
- **Missing setup functions:** `applyNamedRangesAndValidation()` and `setupExplanationSheet()` are called by `main.gs` but are not included in the supplied code.
- **Setup can clear data:** `setupDashboardLayout()` calls `homeSheet.clear()`, and `setupWeddingSheets()` clears the first row of existing planning sheets before rewriting headers. Running initialization again may remove existing dashboard content or replace sheet headers. Back up the spreadsheet first, especially before rerunning setup.
- **Custom formula support:** The `SHEETID()` custom function is used inside a hyperlink formula. Test that formula in your spreadsheet locale; formula separators and custom-function behavior can vary.
- **Guest counts:** The dashboard’s guest formulas count nonempty guest-list rows and RSVP values. If one row represents a household rather than one person, “Total Invited” may not equal the total number of attendees.
- **Deadline formulas:** The urgent-task and deadline sections depend on valid dates, statuses, named ranges, and formula behavior in the spreadsheet.
- **Privacy:** The guest list may contain personal information. Restrict spreadsheet access to trusted collaborators and avoid sharing it publicly.

## 🛠️ Troubleshooting

### “Function not found”

Check that every function called by `initializeWeddingPlanner()` exists in the project. In particular, provide `applyNamedRangesAndValidation()` and `setupExplanationSheet()`, or remove those calls.

### “Identifier has already been declared” or unexpected duplicate behavior

Search the Apps Script project for repeated function definitions. Keep only one definition of `applyStyle()` and `setupDashboardLayout()`.

### Dashboard formulas show errors or zero values

Confirm that the required named ranges exist and refer to the correct columns. Check that status values use the expected text, such as `Completed`, `On Hold`, `Yes`, and `No`.

### Navigation links do not work

Confirm the target sheets exist, test `SHEETID()` in a spreadsheet cell, and check whether your spreadsheet locale requires different formula separators.

### Initialization changed existing content

Restore a backup or version history if needed. The setup functions clear dashboard content and may clear existing header rows.

---
## 📄 License

Please refer to the repository for applicable licensing information.

## 👨‍💻 Author

**ursleeplessprogrammer**

*This project provides a customizable starting point for wedding planning in Google Sheets. Review the setup functions before running them on a spreadsheet that already contains important data.*
