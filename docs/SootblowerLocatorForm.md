# Sootblower Locator Form Documentation

## Overview

The Sootblower Locator form provides a user-friendly interface to search, filter, and display sootblower equipment in the dashboard. It supports filtering by number and by group type (Retracts or Wall blowers).

## Form Components

### 1. Main Form
![Sootblower Locator Form](form_image.png)

The form includes:
- Sootblower number input field (numeric only)
- Group filter options (All, Retracts, Wall)
- Action buttons (Search, Show All, Show Associated, Close)
- Results display area
- Status bar

### 2. Search Options

- **Sootblower Number**: Enter digits only to find a specific sootblower number
- **Filter by Group**:
  - **All Types**: Show all sootblower types
  - **Retracts (IK/EL)**: Only show SBIK and SBEL type sootblowers
  - **Wall (IR/WB)**: Only show SBIR and SBWB type sootblowers

### 3. Action Buttons

- **Search**: Search for sootblowers matching the current number and filter
- **Show All**: Display all sootblowers matching the current filter
- **Show Associated**: View equipment associated with the current results
- **Close**: Close the form

## Usage Instructions

1. Open the form by selecting the Sootblower Locator option
2. Enter a sootblower number or leave blank to search all
3. Select a group filter if needed
4. Click Search or Show All to display results
5. Results will appear in the Dashboard worksheet
6. Use Show Associated to find related equipment

## Implementation Notes

The form can be created in two ways:

1. **Static form** created in the VBA editor:
   - Import the provided `.frm` file or create manually
   - Use the `.txt` file as a reference for control layout and code

2. **Dynamic form** created at runtime:
   - Uses the SootblowerFormCreator module
   - Creates all controls programmatically
   - Used as a fallback when the static form isn't available

3. **Runtime binder (no VBIDE)** for any plain form:
   - Use `ActiveModules/mod_SSB_RuntimeBinder.bas` with `ActiveModules/C_SSB_BtnHandler.cls`
   - Drop in any simple UserForm with: 1 TextBox, 3 OptionButtons, 4 CommandButtons
   - Optionally set control Tags to guide binding:
     - TextBox.Tag = `role:sb_number`
     - OptionButtons.Tag = `role:opt_all`, `role:opt_retracts`, `role:opt_wall`
     - Buttons.Tag = `role:btn_search`, `role:btn_showall`, `role:btn_assoc`, `role:btn_close`
   - Then run: `SSB_BindAndShow "UserForm1"` (replace with your form name)

## Adding the Form to Your Project

### Option 1: Import the Form File

1. In the VBA Editor, right-click on "Forms" in the Project Explorer
2. Select "Import File..."
3. Navigate to and select "frmSootblowerLocator.frm"
4. The form will be added to your project

### Option 2: Create the Form Manually

1. In the VBA Editor, select Insert > UserForm
2. Name the new form "frmSootblowerLocator"
3. Add all the controls as described in "frmSootblowerLocator.txt"
4. Copy the code from "frmSootblowerLocator.txt" to the form's code module

### Option 3: Use the Dynamic Creator

The SootblowerFormCreator module will automatically create the form at runtime if the static form isn't available. This requires no additional setup beyond ensuring the module is in your project.

### Option 4: Use the Runtime Binder (recommended during dev)

1. Insert a basic UserForm (e.g., `UserForm1`) with controls as listed above
2. Optionally set the control Tags as shown to make binding explicit
3. Run `SSB_BindAndShow "UserForm1"`
4. Click the buttons to trigger search/display logic in `mod_SootblowerLocator`

### Option 5: Use the Design-time Builder (auto-create frmSootblowerLocator)

This option programmatically creates the compiled form using the VBIDE, with all controls and Tags set. Then it uses the runtime binder to wire events.

Requirements:
- Trust Center: enable "Trust access to the VBA project object model"

Steps:
1. Ensure `ActiveModules/SootblowerFormBuilder.bas` is imported
2. In the VBA editor, run `Dev_BuildAndShow_SootblowerForm`
3. The builder will create `frmSootblowerLocator` (or rebuild it), then call the binder to show it
4. If the binder is unavailable, it will simply show the form (buttons will be inert until wired)

## Troubleshooting

If the form fails to appear or buttons do nothing:
1. Check if the "frmSootblowerLocator" form exists in your project
2. Verify the SootblowerFormCreator module is present
3. If using the runtime binder, ensure `mod_SSB_RuntimeBinder` and `C_SSB_BtnHandler` are present
4. Confirm control Tags or order match expectations (see above)
3. Look for diagnostic log files in the logs/Diagnostic_Notes folder
4. Check the immediate window for any error messages

## Additional Resources

- See mod_SootblowerLocator.bas for the backend search functionality
- Examine SootblowerFormCreator.bas for the dynamic form creation code
- Check logs/Diagnostic_Notes for detailed logs of form operations