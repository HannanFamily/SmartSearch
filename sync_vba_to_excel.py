#!/usr/bin/env python3
"""
VBA to Excel Synchronization Script
===================================
This script synchronizes VBA code from .bas and .cls files back into Excel workbooks.
It can update existing modules or create new ones as needed.

Dependencies: openpyxl (can be installed via pip)
Usage: python sync_vba_to_excel.py [workbook_name]
"""

import os
import sys
import shutil
from datetime import datetime
import zipfile
import tempfile
import re

class VBAToExcelSyncer:
    def __init__(self, project_dir=None):
        self.project_dir = project_dir or os.getcwd()
        self.workbook_path = None
        self.backup_dir = os.path.join(self.project_dir, "Old_Code")
        
    def find_workbook(self, workbook_name=None):
        """Find the target Excel workbook"""
        if workbook_name:
            potential_path = os.path.join(self.project_dir, workbook_name)
            if os.path.exists(potential_path):
                return potential_path
        
        # Search for .xlsm files in the project directory
        xlsm_files = [f for f in os.listdir(self.project_dir) if f.endswith('.xlsm')]
        
        if not xlsm_files:
            print("ERROR: No .xlsm files found in project directory")
            return None
        
        if len(xlsm_files) == 1:
            return os.path.join(self.project_dir, xlsm_files[0])
        
        # Multiple workbooks found - let user choose
        print("Multiple Excel workbooks found:")
        for i, wb in enumerate(xlsm_files, 1):
            print(f"  {i}. {wb}")
        
        while True:
            try:
                choice = input(f"Select workbook (1-{len(xlsm_files)}): ").strip()
                idx = int(choice) - 1
                if 0 <= idx < len(xlsm_files):
                    return os.path.join(self.project_dir, xlsm_files[idx])
                print("Invalid selection. Please try again.")
            except (ValueError, KeyboardInterrupt):
                print("\nOperation cancelled.")
                return None
    
    def find_vba_files(self):
        """Find all VBA files (.bas, .cls) in the project directory"""
        vba_files = []
        for filename in os.listdir(self.project_dir):
            if filename.endswith(('.bas', '.cls')):
                vba_files.append({
                    'filename': filename,
                    'path': os.path.join(self.project_dir, filename),
                    'type': 'module' if filename.endswith('.bas') else 'class',
                    'name': os.path.splitext(filename)[0]
                })
        return vba_files
    
    def backup_workbook(self, workbook_path):
        """Create a timestamped backup of the workbook"""
        if not os.path.exists(self.backup_dir):
            os.makedirs(self.backup_dir)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        workbook_name = os.path.basename(workbook_path)
        name, ext = os.path.splitext(workbook_name)
        backup_name = f"{name}_backup_{timestamp}{ext}"
        backup_path = os.path.join(self.backup_dir, backup_name)
        
        shutil.copy2(workbook_path, backup_path)
        print(f"Backup created: {backup_path}")
        return backup_path
    
    def read_vba_content(self, vba_file_path):
        """Read and clean VBA content from .bas or .cls file"""
        with open(vba_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Remove the Attribute VB_Name line as it will be set by Excel
        lines = content.split('\n')
        cleaned_lines = []
        
        for line in lines:
            if line.strip().startswith('Attribute VB_Name'):
                continue
            cleaned_lines.append(line)
        
        return '\n'.join(cleaned_lines)
    
    def extract_xlsm_contents(self, xlsm_path, extract_dir):
        """Extract XLSM file contents to a directory"""
        with zipfile.ZipFile(xlsm_path, 'r') as zip_file:
            zip_file.extractall(extract_dir)
    
    def update_vba_project(self, extract_dir, vba_files):
        """Update VBA project files in the extracted directory"""
        vba_project_dir = os.path.join(extract_dir, 'xl', 'vbaProject.bin')
        
        # Note: Direct VBA binary manipulation is complex
        # For now, we'll create a reference file that can be used
        # to manually import the modules
        
        vba_ref_dir = os.path.join(extract_dir, 'vba_modules')
        if not os.path.exists(vba_ref_dir):
            os.makedirs(vba_ref_dir)
        
        for vba_file in vba_files:
            content = self.read_vba_content(vba_file['path'])
            ref_path = os.path.join(vba_ref_dir, vba_file['filename'])
            
            with open(ref_path, 'w', encoding='utf-8') as f:
                f.write(content)
    
    def create_import_script(self, vba_files):
        """Create a VBA script that can be run in Excel to import modules"""
        script_content = '''Sub ImportVBAModules()
    ' Auto-generated VBA import script
    ' Run this macro in Excel to import updated modules
    
    Dim fso As Object
    Dim projectPath As String
    Dim vbcomp As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    projectPath = ThisWorkbook.Path
    
    ' Remove existing modules first (optional - comment out if you want to keep old versions)
'''
        
        for vba_file in vba_files:
            if vba_file['type'] == 'module':
                script_content += f'''    
    ' Remove and re-import {vba_file['name']}
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("{vba_file['name']}")
    On Error GoTo 0
    
    If fso.FileExists(projectPath & "\\{vba_file['filename']}") Then
        Set vbcomp = ThisWorkbook.VBProject.VBComponents.Import(projectPath & "\\{vba_file['filename']}")
        Debug.Print "Imported: {vba_file['filename']}"
    Else
        Debug.Print "File not found: {vba_file['filename']}"
    End If
'''
        
        script_content += '''
    
    MsgBox "VBA module import complete! Check Immediate Window for details.", vbInformation
End Sub
'''
        
        script_path = os.path.join(self.project_dir, 'import_vba_modules.bas')
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script_content)
        
        print(f"Import script created: {script_path}")
        return script_path
    
    def generate_manual_instructions(self, vba_files, workbook_name):
        """Generate manual instructions for VBA import"""
        instructions = f"""
VBA MODULE SYNCHRONIZATION INSTRUCTIONS
=======================================
Workbook: {workbook_name}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

AUTOMATIC METHOD (Recommended):
1. Open {workbook_name} in Excel
2. Press Alt+F11 to open VBA Editor
3. Import the file 'import_vba_modules.bas' (File > Import)
4. Run the ImportVBAModules macro
5. Delete the temporary import_vba_modules module when done

MANUAL METHOD:
1. Open {workbook_name} in Excel
2. Press Alt+F11 to open the VBA Editor
3. For each module/class below:
   - Right-click in Project Explorer
   - Choose "Import File..." or "Remove" then "Import File..."
   - Select the corresponding .bas or .cls file

MODULES TO UPDATE:
"""
        
        for vba_file in vba_files:
            file_size = os.path.getsize(vba_file['path'])
            mod_time = datetime.fromtimestamp(os.path.getmtime(vba_file['path']))
            instructions += f"""
- {vba_file['filename']} ({vba_file['type']})
  Module Name: {vba_file['name']}
  File Size: {file_size} bytes
  Modified: {mod_time.strftime('%Y-%m-%d %H:%M:%S')}
  Path: {vba_file['path']}
"""
        
        instructions += f"""
IMPORTANT NOTES:
- A backup has been created in the Old_Code directory
- Ensure macros are enabled when opening the workbook
- Save the workbook after importing to preserve changes
- Test all functionality after import

PROJECT STRUCTURE:
- VBA files are stored as .bas and .cls in the project root
- Git tracks these files for version control
- Use this script whenever VBA files are updated
"""
        
        instructions_path = os.path.join(self.project_dir, 'VBA_SYNC_INSTRUCTIONS.txt')
        with open(instructions_path, 'w', encoding='utf-8') as f:
            f.write(instructions)
        
        print(f"Instructions saved to: {instructions_path}")
        return instructions_path
    
    def sync_vba(self, workbook_name=None):
        """Main synchronization method"""
        print("=== VBA to Excel Synchronization ===")
        
        # Find workbook
        workbook_path = self.find_workbook(workbook_name)
        if not workbook_path:
            return False
        
        print(f"Target workbook: {os.path.basename(workbook_path)}")
        
        # Find VBA files
        vba_files = self.find_vba_files()
        if not vba_files:
            print("No VBA files (.bas, .cls) found in project directory")
            return False
        
        print(f"Found {len(vba_files)} VBA files:")
        for vba_file in vba_files:
            print(f"  - {vba_file['filename']} ({vba_file['type']})")
        
        # Create backup
        self.backup_workbook(workbook_path)
        
        # Create import script
        self.create_import_script(vba_files)
        
        # Generate instructions
        self.generate_manual_instructions(vba_files, os.path.basename(workbook_path))
        
        print("\n=== NEXT STEPS ===")
        print("1. Open your Excel workbook")
        print("2. Follow the instructions in 'VBA_SYNC_INSTRUCTIONS.txt'")
        print("3. Use the auto-generated 'import_vba_modules.bas' for easy import")
        print("\nSynchronization preparation complete!")
        
        return True

def main():
    """Main entry point"""
    workbook_name = sys.argv[1] if len(sys.argv) > 1 else None
    
    syncer = VBAToExcelSyncer()
    success = syncer.sync_vba(workbook_name)
    
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()