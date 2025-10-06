# SEARCH ENGINE INTEGRATION GUIDE

## Overview
The new search engine provides a consolidated, extensible system for implementing search functionality in your Excel dashboard with **full slicer integration** and **comprehensive ConfigTable support**. It consists of three main modules:

1. **mod_SearchEngine_Fixed.bas** - Core search engine with slicer integration
2. **mod_SearchModeBuilder.bas** - Easy mode creation and management  
3. **mod_DashboardInterface.bas** - Dashboard integration functions

## Key Features

### ✅ **Slicer Integration**
- **Preserves slicer state**: Search results respect active slicer selections
- **Temporary filter column**: Uses hidden helper column to combine search + slicer filtering
- **Slicer pulse detection**: Automatically detects when slicers change
- **Smart refresh**: Shows slicer-filtered data when no search is active

### ✅ **ConfigTable Integration**
- **Dynamic configuration**: All settings pulled from ConfigTable
- **Input discovery**: Automatically finds search inputs marked as "Input Named Range"
- **Output columns**: Uses Out_Column1..Out_Column8 from ConfigTable
- **Flexible setup**: Easy to reconfigure without code changes

## Quick Start

### 1. Basic Setup
```vb
' Initialize the search system
Call InitializeDashboard

' This will:
' - Initialize the search engine
' - Create default modes for existing tables
' - Set up dashboard interface
```

### 2. Perform a Search
```vb
' Quick text search
resultCount = QuickSearch("pump")

' Mode-specific search
resultCount = ExecuteSearch("Equipment Search")

' Search with criteria
Dim criteria As SearchCriteria
criteria.SearchText = "valve"
criteria.ExactMatch = False
resultCount = ExecuteSearch("Equipment Search", criteria)
```

### 3. Create New Search Modes

#### Simple Text Search Mode
```vb
' Create a mode that searches text in specific columns
Call CreateTextSearchMode("Equipment Search", "EquipmentData", "Description,Type,Tag")
```

#### Exact Match Mode (for IDs, codes)
```vb
' Create a mode for exact tag lookups
Call CreateExactMatchMode("Tag Lookup", "EquipmentData", "Tag")
```

#### Location-Based Search
```vb
' Create a mode for location filtering
Call CreateLocationSearchMode("By Location", "EquipmentData", "Location")
```

#### Custom Filtered View
```vb
' Create a mode with custom filter criteria
Call CreateFilteredViewMode("Pumps Only", "EquipmentData", "[@Type]='Pump'", "Tag,Description,Location")
```

## Dashboard Integration

### Button Event Handlers
```vb
' Assign these to your dashboard buttons
Public Sub btnSearch_Click()
    Call PerformDashboardSearch
End Sub

Public Sub btnClear_Click()  
    Call ClearDashboardSearch
End Sub

Public Sub btnShowAll_Click()
    Call ShowAllData
End Sub
```

### Named Ranges Required
Ensure these named ranges exist on your dashboard:

- **SearchBox** - Main search input
- **ModeSelector** - Dropdown for search mode selection
- **SearchStatus** - Status message display
- **ResultsStartCell** - Where to output results (e.g., "A10")

Optional named ranges:
- **ExactMatch** - Checkbox for exact match
- **CaseSensitive** - Checkbox for case sensitivity
- **LastUpdated** - Timestamp display

## Setting Up Predefined Mode Templates

### Equipment Modes
```vb
Call SetupEquipmentModes
' Creates:
' - "Equipment Search" - Text search in Description,Tag,Type
' - "Tag Lookup" - Exact tag matching
' - "By Type" - Filter by equipment type
' - "By Location" - Filter by location
```

### Sootblower Modes  
```vb
Call SetupSootblowerModes
' Creates:
' - "Sootblower by Number" - Exact number matching
' - "Sootblower by Location" - Filter by floor/location
' - "By Cabinet" - Filter by cabinet
```

### Valve Modes
```vb
Call SetupValveModes  
' Creates:
' - "Valve by Number" - Exact valve number matching
' - "By System" - Filter by valve system
```

## Configuration

## Slicer Integration Details

### How It Works
The search engine maintains full compatibility with Excel slicers:

1. **Slicer State Preservation**: When you perform a search, slicer selections remain active
2. **Combined Filtering**: Search results show only items that match BOTH:
   - Active slicer selections
   - Search criteria
3. **Temporary Filter Column**: Uses a hidden helper column to combine filters without disrupting slicers
4. **Automatic Cleanup**: Temporary filters are removed when search is cleared

### Slicer Behavior Examples

#### Example 1: Search with Active Slicers
```
1. User selects "Pump" in Equipment Type slicer
2. User searches for "cooling" 
3. Results show only cooling pumps (both filters applied)
4. Slicer still shows "Pump" selected
```

#### Example 2: Clear Search, Keep Slicers
```
1. User has slicers active and search results showing
2. User clicks "Clear Search"
3. Search box clears, but slicers remain active
4. Results show all items matching current slicer selections
```

#### Example 3: Refresh After Slicer Change
```
1. User changes slicer selection
2. System automatically detects change via pulse cell
3. Results refresh to show new slicer-filtered data
4. Active search (if any) is reapplied to new slicer subset
```

### ConfigTable Keys for Slicer Integration
The system uses comprehensive ConfigTable integration with these keys:

| Key | Description | Example |
|-----|-------------|---------|
| **DASHBOARD_SHEET** | Name of dashboard worksheet | "Dashboard" |
| **DATA_TABLE_NAME** | Main data table name | "EquipmentData" |
| **SPLICER_THRESHOLD** | Slicer change detection threshold | "250" |
| **TEMP_FILTER_COL_NAME** | Name for temporary search filter column | "Temp_SearchInclude" |
| **SLICER_PULSE_ANCHOR** | Named range for slicer pulse detection | "SlicerPulseAnchor" |
| **InputCell_DescripSearch** | Description search input range | "SearchBox" |
| **InputCell_ValveNumSearch** | Valve/Tag search input range | "ValveSearchBox" |
| **ResultsStartCell** | Where to output search results | "ResultsStart" |
| **StatusCell** | Status message display cell | "SearchStatus" |
| **Out_Column1..Out_Column8** | Output column definitions | "Tag", "Description", etc. |

### Input Named Ranges
In your ConfigTable, mark search inputs with TYPE = "Input Named Range":

| ConfigKey | ConfigValue | TYPE |
|-----------|-------------|------|
| SearchInput1 | SearchBox | Input Named Range |
| SearchInput2 | ValveSearchBox | Input Named Range |
| SearchInput3 | LocationFilter | Input Named Range |

### Slicer Integration Setup

1. **Create Slicer Pulse Cell**: Add a named range "SlicerPulseAnchor" (hidden cell)
2. **Configure Threshold**: Set SPLICER_THRESHOLD in ConfigTable (e.g., "250")
3. **Temp Filter Column**: Set TEMP_FILTER_COL_NAME (e.g., "Temp_SearchInclude")

The system automatically:
- Detects slicer changes via SUBTOTAL formula
- Preserves slicer filtering when searching
- Shows slicer-filtered data when search is cleared

## Migration from Old System

### Replace Old Function Calls
```vb
' Old way:
Call PerformSearch

' New way:
Call PerformDashboardSearch
```

```vb
' Old way:
Call Safe_OutputAllVisible

' New way:
Call ShowAllData
```

### Update Mode Definitions
If you have existing modes in `mod_ModeDrivenSearch`, recreate them using the new builder:

```vb
' Instead of manually coding modes, use:
Call CreateTextSearchMode("ModeName", "TableName", "Columns")
```

## Error Handling

The new system provides comprehensive error handling:

```vb
' Check for errors after search operations
resultCount = ExecuteSearch("MyMode")
If resultCount = -1 Then
    MsgBox "Search error: " & GetLastError()
End If
```

## Testing and Validation

### Validate All Modes
```vb
Call ValidateAllModes
' Checks that all registered modes work correctly
```

### Test Dashboard Functions
```vb
Call TestDashboard
' Runs through basic dashboard operations
```

### View Current State
```vb
Call ShowDashboardState
' Displays current mode, search text, etc.
```

## Examples

### Example 1: Adding Equipment Search by Manufacturer
```vb
Sub AddManufacturerSearch()
    ' Create a filtered view for specific manufacturer
    Call CreateFilteredViewMode("Manufacturer_ABC", "EquipmentData", "[@Manufacturer]='ABC Corp'", "Tag,Description,Manufacturer,Type")
End Sub
```

### Example 2: Multi-Column Text Search
```vb
Sub AddMultiColumnSearch()
    ' Search across multiple text fields
    Call CreateTextSearchMode("Comprehensive Search", "EquipmentData", "Description,Tag,Type,Location,Notes")
End Sub
```

### Example 3: Custom Search with VBA Function
```vb
Sub PerformCustomSearch()
    Dim criteria As SearchCriteria
    criteria.SearchText = "motor"
    criteria.ExactMatch = False
    criteria.CaseSensitive = False
    criteria.SearchColumns = "Description,Type"
    
    Dim results As Long
    results = ExecuteSearch("Equipment Search", criteria)
    
    MsgBox "Found " & results & " motor-related items"
End Sub
```

## Benefits

1. **✅ Slicer Integration**: Search results respect and preserve slicer selections
2. **✅ ConfigTable Driven**: All configuration through tables, no hardcoded values
3. **Easy Extension**: Add new search modes with single function calls
4. **Consistent Interface**: All search operations use the same API
5. **Error Recovery**: Robust error handling prevents crashes
6. **Performance**: Optimized for large datasets with slicer filtering
7. **Flexibility**: Works with existing Excel structure and slicer setup
8. **Maintainable**: Clear separation of concerns
9. **Dynamic Inputs**: Automatically discovers search inputs from ConfigTable
10. **Backward Compatible**: Works with existing dashboard structure

## Next Steps

1. Import the three new modules into your Excel workbook
2. Run `InitializeDashboard` to set up the system
3. Test with `TestDashboard` 
4. Create your custom modes using the builder functions
5. Update your dashboard buttons to use the new interface functions

For troubleshooting, check the VBA Immediate Window for detailed logging messages.