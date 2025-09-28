#!/usr/bin/env python3
"""
Excel VBA Dashboard - Quick Operations Interface
===============================================
Provides a menu-driven interface for common Search Dashboard VBA operations
"""

import os
import sys
import subprocess
from pathlib import Path

# Configuration (dynamic)
# Prefer explicit env override, else current interpreter
PYTHON_PATH = os.environ.get("PYTHON_PATH", sys.executable)
BASE_DIR = Path(__file__).parent
CONTROLLER_PATH = str((BASE_DIR / "excel_vba_controller.py").resolve())
WORKBOOK = str((BASE_DIR.parent / "Search Dashboard v1.3.xlsm").resolve())

class VBADashboard:
    def __init__(self):
    self.base_path = BASE_DIR
        
    def run_command(self, args):
        """Run Excel VBA controller with given arguments"""
        cmd = [PYTHON_PATH, CONTROLLER_PATH, "--workbook", WORKBOOK] + args
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, cwd=self.base_path)
            if result.stdout:
                print(result.stdout)
            if result.stderr:
                print("ERROR:", result.stderr)
            return result.returncode == 0
        except Exception as e:
            print(f"ERROR: {e}")
            return False
    
    def show_menu(self):
        """Display the main menu"""
        print("\n" + "="*50)
        print("  Excel VBA Dashboard - Search Dashboard v1.3")
        print("="*50)
        print("1.  Show Workbook Info")
        print("2.  List VBA Modules")
        print("3.  Interactive Mode")
        print("4.  Run Diagnostics")
        print("5.  Export Modules")
        print("6.  Sync Modules")
        print("7.  Test Search")
        print("8.  Show Sootblower Form")
        print("9.  Get Search Input")
        print("10. Set Search Input")
        print("11. Clear Search Results")
        print("12. Run Config Diagnostics")
        print("13. Show Config Table")
        print("14. Custom VBA Command")
        print("0.  Exit")
        print("-"*50)
        
    def get_choice(self):
        """Get user menu choice"""
        try:
            choice = input("Enter your choice (0-14): ").strip()
            return int(choice)
        except ValueError:
            return -1
    
    def execute_choice(self, choice):
        """Execute the selected menu option"""
        if choice == 0:
            return False
            
        elif choice == 1:
            print("\n📊 Showing workbook info...")
            self.run_command(["--show-info"])
            
        elif choice == 2:
            print("\n📋 Listing VBA modules...")
            self.run_command(["--list-modules"])
            
        elif choice == 3:
            print("\n🔧 Starting interactive mode...")
            self.run_command(["--interactive"])
            
        elif choice == 4:
            print("\n🔍 Running diagnostics...")
            # Try different possible macro names
            macros_to_try = [
                "RunQuickSearchDiagnostics",
                "QuickSearchDiagnostics.RunQuickSearchDiagnostics", 
                "mod_PrimaryConsolidatedModule.RunConfigDiagnostics"
            ]
            for macro in macros_to_try:
                print(f"Trying macro: {macro}")
                if self.run_command(["--run-macro", macro]):
                    break
                    
        elif choice == 5:
            print("\n📤 Exporting modules...")
            macros_to_try = [
                "ExportModulesToActiveFolder",
                "Dev_Exports.ExportModulesToActiveFolder",
                "SyncManager.ExportModulesToActiveFolder"
            ]
            for macro in macros_to_try:
                print(f"Trying macro: {macro}")
                if self.run_command(["--run-macro", macro]):
                    break
                    
        elif choice == 6:
            print("\n🔄 Syncing modules...")
            macros_to_try = [
                "SyncModules_FromActiveFolder",
                "SyncManager.SyncModules_FromActiveFolder"
            ]
            for macro in macros_to_try:
                print(f"Trying macro: {macro}")
                if self.run_command(["--run-macro", macro]):
                    break
                    
        elif choice == 7:
            print("\n🔍 Testing search functionality...")
            search_term = input("Enter search term (or press Enter for 'pump'): ").strip()
            if not search_term:
                search_term = "pump"
            print(f"Setting search input to: {search_term}")
            self.run_command(["--set-name", "InputCell_DescripSearch", search_term])
            print("Running search...")
            macros_to_try = [
                "Safe_PerformSearch",
                "mod_PrimaryConsolidatedModule.Safe_PerformSearch",
                "PerformSearch"
            ]
            for macro in macros_to_try:
                print(f"Trying macro: {macro}")
                if self.run_command(["--run-macro", macro]):
                    break
                    
        elif choice == 8:
            print("\n🖥️  Creating and showing Sootblower form...")
            macros_to_try = [
                "CreateAndShowSootblowerForm",
                "SootblowerFormCreator.CreateAndShowSootblowerForm"
            ]
            for macro in macros_to_try:
                print(f"Trying macro: {macro}")
                if self.run_command(["--run-macro", macro]):
                    break
                    
        elif choice == 9:
            print("\n📥 Getting search input...")
            self.run_command(["--get-name", "InputCell_DescripSearch"])
            
        elif choice == 10:
            print("\n📤 Setting search input...")
            search_term = input("Enter search term: ").strip()
            if search_term:
                self.run_command(["--set-name", "InputCell_DescripSearch", search_term])
            else:
                print("No search term entered.")
                
        elif choice == 11:
            print("\n🧹 Clearing search results...")
            macros_to_try = [
                "ClearOldResults",
                "mod_PrimaryConsolidatedModule.ClearOldResults"
            ]
            for macro in macros_to_try:
                print(f"Trying macro: {macro}")
                if self.run_command(["--run-macro", macro]):
                    break
                    
        elif choice == 12:
            print("\n⚙️  Running config diagnostics...")
            macros_to_try = [
                "RunConfigDiagnostics",
                "mod_PrimaryConsolidatedModule.RunConfigDiagnostics"
            ]
            for macro in macros_to_try:
                print(f"Trying macro: {macro}")
                if self.run_command(["--run-macro", macro]):
                    break
                    
        elif choice == 13:
            print("\n📋 Showing config table...")
            # Get config values
            print("Getting key configuration values...")
            config_keys = [
                "InputCell_DescripSearch",
                "InputCell_ValveNumSearch", 
                "ResultsStartCell",
                "StatusCell",
                "MAX_OUTPUT_ROWS"
            ]
            for key in config_keys:
                print(f"  {key}:")
                self.run_command(["--get-name", key])
                
        elif choice == 14:
            print("\n⚡ Custom VBA command...")
            command = input("Enter VBA macro name: ").strip()
            if command:
                args = input("Enter arguments (space-separated, optional): ").strip()
                if args:
                    self.run_command(["--run-macro", command, "--macro-args"] + args.split())
                else:
                    self.run_command(["--run-macro", command])
            else:
                print("No command entered.")
                
        else:
            print("❌ Invalid choice. Please try again.")
            
        input("\nPress Enter to continue...")
        return True
    
    def run(self):
        """Main dashboard loop"""
        print("🚀 Starting Excel VBA Dashboard...")
        
        # Check if files exist
        if not os.path.exists(CONTROLLER_PATH):
            print(f"ERROR: Controller script not found: {CONTROLLER_PATH}")
            return
            
        if not os.path.exists(PYTHON_PATH):
            print(f"ERROR: Python not found: {PYTHON_PATH}")
            return
        
        while True:
            try:
                self.show_menu()
                choice = self.get_choice()
                
                if not self.execute_choice(choice):
                    break
                    
            except KeyboardInterrupt:
                print("\n\n👋 Goodbye!")
                break
            except Exception as e:
                print(f"\n❌ Unexpected error: {e}")
                input("Press Enter to continue...")

if __name__ == "__main__":
    dashboard = VBADashboard()
    dashboard.run()