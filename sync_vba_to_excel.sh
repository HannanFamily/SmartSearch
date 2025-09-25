#!/bin/bash

# VBA to Excel Synchronization Script (Bash)
# ==========================================
# This script prepares VBA files for import into Excel workbooks
# Usage: ./sync_vba_to_excel.sh [workbook_name]

PROJECT_DIR="$(pwd)"
BACKUP_DIR="$PROJECT_DIR/Old_Code"
WORKBOOK_NAME="$1"

echo "=== VBA to Excel Synchronization ==="

# Find Excel workbooks
XLSM_FILES=(*.xlsm)
if [ ! -e "${XLSM_FILES[0]}" ]; then
    echo "ERROR: No .xlsm files found in current directory"
    exit 1
fi

# Filter out backup files
WORKBOOK_FILES=()
for file in "${XLSM_FILES[@]}"; do
    if [[ ! "$file" =~ backup ]]; then
        WORKBOOK_FILES+=("$file")
    fi
done

# Select workbook
if [ -n "$WORKBOOK_NAME" ]; then
    TARGET_WORKBOOK="$WORKBOOK_NAME"
    if [ ! -f "$TARGET_WORKBOOK" ]; then
        echo "ERROR: Workbook '$WORKBOOK_NAME' not found"
        exit 1
    fi
elif [ ${#WORKBOOK_FILES[@]} -eq 1 ]; then
    TARGET_WORKBOOK="${WORKBOOK_FILES[0]}"
else
    echo "Multiple workbooks found:"
    for i in "${!WORKBOOK_FILES[@]}"; do
        echo "  $((i+1)). ${WORKBOOK_FILES[$i]}"
    done
    read -p "Select workbook (1-${#WORKBOOK_FILES[@]}): " choice
    if [[ "$choice" =~ ^[0-9]+$ ]] && [ "$choice" -ge 1 ] && [ "$choice" -le ${#WORKBOOK_FILES[@]} ]; then
        TARGET_WORKBOOK="${WORKBOOK_FILES[$((choice-1))]}"
    else
        echo "Invalid selection"
        exit 1
    fi
fi

echo "Target workbook: $TARGET_WORKBOOK"

# Find VBA files
VBA_FILES=(*.bas *.cls)
ACTUAL_VBA_FILES=()
for file in "${VBA_FILES[@]}"; do
    if [ -f "$file" ]; then
        ACTUAL_VBA_FILES+=("$file")
    fi
done

if [ ${#ACTUAL_VBA_FILES[@]} -eq 0 ]; then
    echo "No VBA files found in project directory"
    exit 0
fi

echo "Found ${#ACTUAL_VBA_FILES[@]} VBA files:"
for file in "${ACTUAL_VBA_FILES[@]}"; do
    if [[ "$file" == *.bas ]]; then
        echo "  - $file (Module)"
    else
        echo "  - $file (Class)"
    fi
done

# Create backup directory
mkdir -p "$BACKUP_DIR"

# Create backup
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
WORKBOOK_BASE="${TARGET_WORKBOOK%.*}"
WORKBOOK_EXT="${TARGET_WORKBOOK##*.}"
BACKUP_NAME="${WORKBOOK_BASE}_backup_${TIMESTAMP}.${WORKBOOK_EXT}"
BACKUP_PATH="$BACKUP_DIR/$BACKUP_NAME"

cp "$TARGET_WORKBOOK" "$BACKUP_PATH"
echo "Backup created: $BACKUP_PATH"

# Generate VBA import script
IMPORT_SCRIPT="import_vba_modules.bas"

cat > "$IMPORT_SCRIPT" << 'EOF'
Sub ImportVBAModules()
    ' Auto-generated VBA import script
    ' Run this macro in Excel to import updated modules
    
    Dim fso As Object
    Dim projectPath As String
    Dim vbcomp As Object
    Dim moduleName As String
    Dim filePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    projectPath = ThisWorkbook.Path
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Debug.Print "Starting VBA module import at " & Now()
    Debug.Print "Project path: " & projectPath
    Debug.Print String(50, "=")
    
EOF

for file in "${ACTUAL_VBA_FILES[@]}"; do
    MODULE_NAME="${file%.*}"
    cat >> "$IMPORT_SCRIPT" << EOF
    
    ' Process $file
    moduleName = "$MODULE_NAME"
    filePath = projectPath & "\\$file"
    
    If fso.FileExists(filePath) Then
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(moduleName)
        On Error GoTo 0
        
        Set vbcomp = ThisWorkbook.VBProject.VBComponents.Import(filePath)
        Debug.Print "✓ Imported: $file"
    Else
        Debug.Print "✗ File not found: $file"
    End If
EOF
done

cat >> "$IMPORT_SCRIPT" << 'EOF'

    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Debug.Print String(50, "=")
    Debug.Print "VBA module import completed at " & Now()
    MsgBox "VBA module import complete!" & vbCrLf & vbCrLf & _
           "Check the Immediate Window (Ctrl+G) for detailed results.", vbInformation, "Import Complete"
End Sub
EOF

# Generate instructions
INSTRUCTIONS_FILE="VBA_SYNC_INSTRUCTIONS.txt"

cat > "$INSTRUCTIONS_FILE" << EOF
VBA MODULE SYNCHRONIZATION INSTRUCTIONS
=======================================
Workbook: $TARGET_WORKBOOK
Generated: $(date '+%Y-%m-%d %H:%M:%S')
Backup: $BACKUP_NAME

AUTOMATIC METHOD (Recommended):
1. Open $TARGET_WORKBOOK in Excel
2. Enable macros if prompted
3. Press Alt+F11 to open VBA Editor
4. Import the file 'import_vba_modules.bas'
5. Run the 'ImportVBAModules' macro
6. Check Immediate Window (Ctrl+G) for results
7. Save the workbook
8. Delete the temporary import module when done

MODULES TO UPDATE:
EOF

for file in "${ACTUAL_VBA_FILES[@]}"; do
    MODULE_NAME="${file%.*}"
    if [[ "$file" == *.bas ]]; then
        TYPE="Standard Module"
    else
        TYPE="Class Module"
    fi
    
    FILE_SIZE=$(stat -c%s "$file" 2>/dev/null || wc -c < "$file")
    MODIFIED=$(stat -c%y "$file" 2>/dev/null || date -r "$file" '+%Y-%m-%d %H:%M:%S')
    
    cat >> "$INSTRUCTIONS_FILE" << EOF

- $file ($TYPE)
  Module Name: $MODULE_NAME
  File Size: $FILE_SIZE bytes
  Modified: $MODIFIED
  Path: $PROJECT_DIR/$file
EOF
done

cat >> "$INSTRUCTIONS_FILE" << EOF

TROUBLESHOOTING:
- Enable "Trust access to the VBA project object model" in Excel
- Check file permissions and paths
- Always test functionality after importing

WORKFLOW:
- VBA files are the source of truth
- Use Git for version control
- Run this script after updating VBA files
EOF

echo
echo "=== SYNCHRONIZATION READY ==="
echo "Created files:"
echo "  - $BACKUP_NAME (backup)"
echo "  - import_vba_modules.bas (import script)"
echo "  - VBA_SYNC_INSTRUCTIONS.txt (instructions)"
echo
echo "=== NEXT STEPS ==="
echo "1. Open $TARGET_WORKBOOK in Excel"
echo "2. Follow instructions in VBA_SYNC_INSTRUCTIONS.txt"
echo "3. Use import_vba_modules.bas for automatic import"
echo
echo "Synchronization preparation complete!"