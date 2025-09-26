#!/usr/bin/env python3
"""
Excel VBA Import Tool
====================
This script imports VBA modules directly into Excel workbooks
"""

import sys
import os
import shutil
from pathlib import Path
import zipfile
import tempfile
import xml.etree.ElementTree as ET
from datetime import datetime

def import_vba_module_to_excel(excel_file, vba_file):
    """Import a single VBA module into an Excel workbook"""
    
    print(f"🔄 Importing {vba_file} into {excel_file}")
    
    # Check if files exist
    if not os.path.exists(excel_file):
        print(f"❌ Excel file not found: {excel_file}")
        return False
        
    if not os.path.exists(vba_file):
        print(f"❌ VBA file not found: {vba_file}")
        return False
    
    # Create backup
    backup_path = f"{excel_file}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    try:
        shutil.copy2(excel_file, backup_path)
        print(f"✅ Backup created: {backup_path}")
    except Exception as e:
        print(f"⚠️ Could not create backup: {e}")
    
    # Work with temporary files
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_excel = os.path.join(temp_dir, "temp_workbook.xlsm")
        shutil.copy2(excel_file, temp_excel)
        
        try:
            # Extract the Excel file
            extract_dir = os.path.join(temp_dir, "xl")
            with zipfile.ZipFile(temp_excel, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # Copy VBA module
            vba_modules_dir = os.path.join(temp_dir, "xl", "vbaProject")
            if not os.path.exists(vba_modules_dir):
                os.makedirs(vba_modules_dir, exist_ok=True)
            
            # Copy the VBA file
            module_name = os.path.basename(vba_file)
            dest_vba = os.path.join(vba_modules_dir, module_name)
            shutil.copy2(vba_file, dest_vba)
            print(f"✅ VBA module copied to Excel structure")
            
            # Repackage the Excel file
            with zipfile.ZipFile(excel_file, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for root, dirs, files in os.walk(temp_dir):
                    if 'temp_workbook.xlsm' not in root:
                        for file in files:
                            file_path = os.path.join(root, file)
                            arc_name = os.path.relpath(file_path, temp_dir)
                            zip_out.write(file_path, arc_name)
            
            print(f"✅ Excel file updated successfully")
            return True
            
        except Exception as e:
            print(f"❌ Error during import: {e}")
            # Restore backup
            if os.path.exists(backup_path):
                shutil.copy2(backup_path, excel_file)
                print(f"🔄 Restored from backup")
            return False

def main():
    """Main function to import QuickDevAnalysis.bas into Excel"""
    
    # Get current directory
    project_dir = os.getcwd()
    
    # Define files
    vba_file = os.path.join(project_dir, "QuickDevAnalysis.bas")
    excel_files = [
        "Search Dashboard v1.1 STABLE.xlsm",
        "Search Dashboard v1.1 DevCOPY.xlsm"
    ]
    
    print("=" * 60)
    print("🚀 EXCEL VBA MODULE IMPORT TOOL")
    print("=" * 60)
    print(f"📂 Project Directory: {project_dir}")
    print(f"📋 VBA Module: {os.path.basename(vba_file)}")
    
    # Check if VBA file exists
    if not os.path.exists(vba_file):
        print(f"❌ VBA file not found: {vba_file}")
        print("Please ensure QuickDevAnalysis.bas exists in the current directory.")
        return 1
    
    print(f"✅ Found VBA module: {os.path.basename(vba_file)}")
    
    # Find Excel files
    available_excel_files = []
    for excel_file in excel_files:
        excel_path = os.path.join(project_dir, excel_file)
        if os.path.exists(excel_path):
            available_excel_files.append(excel_path)
            print(f"✅ Found Excel file: {excel_file}")
    
    if not available_excel_files:
        print("❌ No Excel files found!")
        print("Expected files:", ", ".join(excel_files))
        return 1
    
    # Import to all available Excel files
    success_count = 0
    for excel_file in available_excel_files:
        print(f"\n📥 Processing: {os.path.basename(excel_file)}")
        if import_vba_module_to_excel(excel_file, vba_file):
            success_count += 1
            print(f"✅ Successfully imported to {os.path.basename(excel_file)}")
        else:
            print(f"❌ Failed to import to {os.path.basename(excel_file)}")
    
    print("\n" + "=" * 60)
    print("📊 IMPORT SUMMARY")
    print("=" * 60)
    print(f"✅ Successful imports: {success_count}")
    print(f"❌ Failed imports: {len(available_excel_files) - success_count}")
    
    if success_count > 0:
        print(f"\n🎯 NEXT STEPS:")
        print(f"1. Open your Excel file")
        print(f"2. Press Alt+F8 to see macros")
        print(f"3. Run 'AnalyzeDevEnvironment'")
        print(f"4. Check the 'Dev_Analysis' worksheet")
        print(f"\n💡 Available macros:")
        print(f"   • AnalyzeDevEnvironment - Main analysis")
        print(f"   • RefreshAnalysis - Quick refresh")
        print(f"   • ExportAnalysisToText - Export results")
    
    return 0 if success_count > 0 else 1

if __name__ == "__main__":
    sys.exit(main())