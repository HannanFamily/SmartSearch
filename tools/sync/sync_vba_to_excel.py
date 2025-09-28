#!/usr/bin/env python3
# Copied from project root to tools/sync for better organization
# See root version for full comments

from datetime import datetime
import os, sys, shutil, zipfile, tempfile, re

# ...simplified re-export of the existing script...
from pathlib import Path

# Inline copy of root script content begins
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
        if workbook_name:
            potential_path = os.path.join(self.project_dir, workbook_name)
            if os.path.exists(potential_path):
                return potential_path
        xlsm_files = [f for f in os.listdir(self.project_dir) if f.endswith('.xlsm')]
        if not xlsm_files:
            print("ERROR: No .xlsm files found in project directory")
            return None
        if len(xlsm_files) == 1:
            return os.path.join(self.project_dir, xlsm_files[0])
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
    
    def create_import_script(self, vba_files):
        script_content = '''Sub ImportVBAModules()
    Dim fso As Object
    Dim projectPath As String
    Dim vbcomp As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    projectPath = ThisWorkbook.Path
'''
        for vba_file in vba_files:
            script_content += f'''
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("{vba_file['name']}")
    On Error GoTo 0
    Set vbcomp = ThisWorkbook.VBProject.VBComponents.Import(projectPath & "\\{vba_file['filename']}")
'''
        script_content += '''
    MsgBox "VBA module import complete!", vbInformation
End Sub
'''
        script_path = os.path.join(self.project_dir, 'import_vba_modules.bas')
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script_content)
        print(f"Import script created: {script_path}")
        return script_path
    
    def generate_manual_instructions(self, vba_files, workbook_name):
        return os.path.join(self.project_dir, 'VBA_SYNC_INSTRUCTIONS.txt')
    
    def sync_vba(self, workbook_name=None):
        workbook_path = self.find_workbook(workbook_name)
        if not workbook_path:
            return False
        vba_files = self.find_vba_files()
        if not vba_files:
            print("No VBA files (.bas, .cls) found in project directory")
            return False
        self.backup_workbook(workbook_path)
        self.create_import_script(vba_files)
        self.generate_manual_instructions(vba_files, os.path.basename(workbook_path))
        print("Synchronization preparation complete!")
        return True

def main():
    workbook_name = sys.argv[1] if len(sys.argv) > 1 else None
    syncer = VBAToExcelSyncer()
    success = syncer.sync_vba(workbook_name)
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()
