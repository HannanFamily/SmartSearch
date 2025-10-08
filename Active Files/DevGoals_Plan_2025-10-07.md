# DevGoals: Ordered by Complexity with Sub-Steps (2025-10-07)

## 1. Easiest → Hardest

### 1.1. Have SearchBox search Tag ID for matches as well always and favor Tag ID results if there is a 4 to 5 digit number string in the search. (Search function)
- Sub-steps:
  1. Update search logic to always include Tag ID in search.
  2. Add logic to detect 4-5 digit number strings in input.
  3. Prioritize Tag ID matches in results display.
  4. Test with various search inputs.

### 1.2. If no results are found in a search, clear slicers and keep the search. (Search function)
- Sub-steps:
  1. Add check for zero search results.
  2. Trigger slicer clear routine if no results.
  3. Re-run search with slicers cleared.
  4. Validate user experience and edge cases.

### 1.3. Have a auto copy function when the "SAP ID" is clicked in the results. With a message in Status cell stating that "SAP ID for " & Equipment Description & " has been copied to your clipboard." (Dashboard Function)
- Sub-steps:
  1. Add click event handler for SAP ID cell in results.
  2. Implement clipboard copy for SAP ID.
  3. Update Status cell with confirmation message.
  4. Test with multiple result rows.

### 1.4. Move from the only whole work search to a fuzzy search. Maybe start by making this a search mode until refined, or as a auto feature if no search results are found from a search. (Search function/mode)
- Sub-steps:
  1. Research and select fuzzy search algorithm (e.g., Levenshtein).
  2. Implement as a new search mode.
  3. Add fallback to fuzzy search if no exact results.
  4. Tune threshold and test accuracy.

### 1.5. Match Equipment to HVAC filters (From "HVAC Filters" sheet) in DataTable to Equipment in DataTable and add data fields for Stock Number, Quantity and Filter size for all filters. Step 2 create search mode for HVAC filter lookup and display results in a clear way. (Look up)
- Sub-steps:
  1. Import/validate HVAC Filters sheet data.
  2. Add new fields to DataTable.
  3. Write matching logic between Equipment and HVAC Filters.
  4. Populate Stock Number, Quantity, Filter size.
  5. Create new search mode for HVAC filter lookup.
  6. Design results display for clarity.

### 1.6. Have a module to look up Power source Location from manually entered "Power Source" column in DataTable and fill the "Power Source Location" column. Find match in "Equipment Description" and return "Physical Location" of Power Source to "Power Source Location" Column. (Data type set up)
- Sub-steps:
  1. Add/validate "Power Source" and "Power Source Location" columns.
  2. Write lookup logic to match Equipment Description to Power Source.
  3. Retrieve and fill Physical Location.
  4. Make routine manually triggerable.
  5. Test with sample data.

### 1.7. Have a way for users to make corrections to or add missing data fields. Maybe with a button "Edit or Add data to selected equipment". (User Function)
- Sub-steps:
  1. Design user interface for editing/adding data.
  2. Implement dropdown for data type selection.
  3. Add logic for power source/linked data selection.
  4. Write change log to hidden sheet (SAP ID, changes, user name).
  5. Implement user name auto-fill (optional, secondary).
  6. Test and validate workflow.

### 1.8. Create a standard way to capture or transfer changes from one version to another. Maybe keep things like the code and configuration table in a separate workbook as some sort of addin function and the DataTable and Dashboard in another. (Project Management)
- Sub-steps:
  1. Design separation of code/config vs. data/dashboard.
  2. Implement export/import routines for config/code.
  3. Test transfer between versions.
  4. Document process for users.

### 1.9. Find patterns in equipment data based on Equipment Description, System, object type, etc. to find groupings of equipment and common sub-parts of that equipment. (Database indexing)
- Sub-steps:
  1. Gather sample data and known associations.
  2. Research clustering/grouping algorithms.
  3. Implement pattern detection and grouping logic.
  4. Validate with real data and refine.
  5. Document findings and update database structure as needed.

---

This plan is ready for review and action in your next session. Each goal is broken down for stepwise progress.
