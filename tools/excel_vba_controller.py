#!/usr/bin/env python3
"""
Excel VBA Controller - Terminal Interface
==========================================
Provides command-line interface to interact with Excel VBA projects including:
- Running VBA procedures and functions
- Accessing and manipulating UserForms
- Reading/writing workbook data
- Module management and code execution
- Real-time interaction with the Search Dashboard

Usage:
    python excel_vba_controller.py --help
    python excel_vba_controller.py --workbook "Search Dashboard v1.3.xlsm" --run-macro "QuickSearchDiagnostics.RunQuickSearchDiagnostics"
    python excel_vba_controller.py --interactive
"""

import argparse
import sys
import os
import time
import json
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 not installed. Install with: pip install pywin32")
    sys.exit(1)

class ExcelVBAController:
    """Controller for Excel VBA operations via COM automation"""
    
    def __init__(self, workbook_path: Optional[str] = None, visible: bool = True):
        self.app = None
        self.workbook = None
        self.workbook_path = workbook_path
        self.visible = visible
        self._connected = False
        
    def connect(self) -> bool:
        """Connect to Excel application"""
        try:
            # Try to connect to existing Excel instance
            try:
                self.app = win32com.client.GetActiveObject("Excel.Application")
                print("Connected to existing Excel instance")
            except:
                # Create new Excel instance
                self.app = win32com.client.Dispatch("Excel.Application")
                print("Created new Excel instance")
            
            self.app.Visible = self.visible
            self.app.DisplayAlerts = False
            
            # Open workbook if specified
            if self.workbook_path:
                full_path = os.path.abspath(self.workbook_path)
                if os.path.exists(full_path):
                    self.workbook = self.app.Workbooks.Open(full_path)
                    print(f"Opened workbook: {full_path}")
                else:
                    print(f"ERROR: Workbook not found: {full_path}")
                    return False
            else:
                # Try to use active workbook
                try:
                    self.workbook = self.app.ActiveWorkbook
                    print(f"Using active workbook: {self.workbook.Name}")
                except:
                    print("No active workbook found")
                    
            self._connected = True
            return True
            
        except Exception as e:
            print(f"ERROR connecting to Excel: {e}")
            return False
    
    def disconnect(self):
        """Disconnect from Excel"""
        if self.app:
            try:
                if not self.visible:
                    self.app.Quit()
                self.app = None
                self.workbook = None
                self._connected = False
                print("Disconnected from Excel")
            except:
                pass
    
    def run_macro(self, macro_name: str, *args) -> Any:
        """Run a VBA macro/procedure"""
        if not self._connected:
            print("ERROR: Not connected to Excel")
            return None
            
        try:
            print(f"Running macro: {macro_name}")
            if args:
                result = self.app.Run(macro_name, *args)
            else:
                result = self.app.Run(macro_name)
            print(f"Macro completed successfully")
            return result
        except Exception as e:
            print(f"ERROR running macro {macro_name}: {e}")
            return None
    
    def get_range_value(self, range_address: str, sheet_name: Optional[str] = None) -> Any:
        """Get value from Excel range"""
        try:
            if sheet_name:
                sheet = self.workbook.Worksheets(sheet_name)
                return sheet.Range(range_address).Value
            else:
                return self.workbook.ActiveSheet.Range(range_address).Value
        except Exception as e:
            print(f"ERROR getting range value: {e}")
            return None
    
    def set_range_value(self, range_address: str, value: Any, sheet_name: Optional[str] = None) -> bool:
        """Set value in Excel range"""
        try:
            if sheet_name:
                sheet = self.workbook.Worksheets(sheet_name)
                sheet.Range(range_address).Value = value
            else:
                self.workbook.ActiveSheet.Range(range_address).Value = value
            return True
        except Exception as e:
            print(f"ERROR setting range value: {e}")
            return False
    
    def get_named_range_value(self, name: str) -> Any:
        """Get value from named range"""
        try:
            return self.workbook.Names(name).RefersToRange.Value
        except Exception as e:
            print(f"ERROR getting named range {name}: {e}")
            return None
    
    def set_named_range_value(self, name: str, value: Any) -> bool:
        """Set value in named range"""
        try:
            self.workbook.Names(name).RefersToRange.Value = value
            return True
        except Exception as e:
            print(f"ERROR setting named range {name}: {e}")
            return False
    
    def list_modules(self) -> List[str]:
        """List all VBA modules in the workbook"""
        try:
            modules = []
            for component in self.workbook.VBProject.VBComponents:
                modules.append(f"{component.Name} ({component.Type})")
            return modules
        except Exception as e:
            print(f"ERROR listing modules: {e}")
            return []
    
    def get_module_code(self, module_name: str) -> str:
        """Get VBA code from a module"""
        try:
            component = self.workbook.VBProject.VBComponents(module_name)
            return component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
        except Exception as e:
            print(f"ERROR getting module code: {e}")
            return ""
    
    def execute_vba_statement(self, statement: str) -> Any:
        """Execute a single VBA statement"""
        try:
            # Create a temporary procedure to execute the statement
            temp_proc = f"""
Sub TempExecute()
    {statement}
End Sub
"""
            # Add temporary module
            temp_module = self.workbook.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
            temp_module.CodeModule.AddFromString(temp_proc)
            
            # Run the procedure
            result = self.app.Run("TempExecute")
            
            # Clean up
            self.workbook.VBProject.VBComponents.Remove(temp_module)
            
            return result
        except Exception as e:
            print(f"ERROR executing VBA statement: {e}")
            return None
    
    def show_userform(self, form_name: str) -> bool:
        """Show a UserForm"""
        try:
            self.execute_vba_statement(f"{form_name}.Show")
            return True
        except Exception as e:
            print(f"ERROR showing UserForm {form_name}: {e}")
            return False
    
    def hide_userform(self, form_name: str) -> bool:
        """Hide a UserForm"""
        try:
            self.execute_vba_statement(f"{form_name}.Hide")
            return True
        except Exception as e:
            print(f"ERROR hiding UserForm {form_name}: {e}")
            return False
    
    def get_workbook_info(self) -> Dict[str, Any]:
        """Get information about the current workbook"""
        if not self.workbook:
            return {}
        
        try:
            info = {
                'name': self.workbook.Name,
                'path': self.workbook.FullName,
                'sheets': [sheet.Name for sheet in self.workbook.Worksheets],
                'modules': self.list_modules(),
                'saved': self.workbook.Saved
            }
            return info
        except Exception as e:
            print(f"ERROR getting workbook info: {e}")
            return {}

def interactive_mode(controller: ExcelVBAController):
    """Interactive command mode"""
    print("\n=== Excel VBA Interactive Mode ===")
    print("Available commands:")
    print("  run <macro_name> [args...]     - Run VBA macro")
    print("  get <range_address>           - Get range value")
    print("  set <range_address> <value>   - Set range value")
    print("  getname <name>                - Get named range value")
    print("  setname <name> <value>        - Set named range value")
    print("  modules                       - List VBA modules")
    print("  code <module_name>            - Show module code")
    print("  exec <vba_statement>          - Execute VBA statement")
    print("  showform <form_name>          - Show UserForm")
    print("  hideform <form_name>          - Hide UserForm")
    print("  info                          - Show workbook info")
    print("  quit                          - Exit interactive mode")
    print()
    
    while True:
        try:
            command = input("VBA> ").strip()
            if not command:
                continue
                
            parts = command.split()
            cmd = parts[0].lower()
            
            if cmd == 'quit':
                break
            elif cmd == 'run' and len(parts) > 1:
                macro_name = parts[1]
                args = parts[2:] if len(parts) > 2 else []
                result = controller.run_macro(macro_name, *args)
                if result is not None:
                    print(f"Result: {result}")
            elif cmd == 'get' and len(parts) > 1:
                value = controller.get_range_value(parts[1])
                print(f"Value: {value}")
            elif cmd == 'set' and len(parts) > 2:
                success = controller.set_range_value(parts[1], parts[2])
                print(f"Set: {'Success' if success else 'Failed'}")
            elif cmd == 'getname' and len(parts) > 1:
                value = controller.get_named_range_value(parts[1])
                print(f"Value: {value}")
            elif cmd == 'setname' and len(parts) > 2:
                success = controller.set_named_range_value(parts[1], parts[2])
                print(f"Set: {'Success' if success else 'Failed'}")
            elif cmd == 'modules':
                modules = controller.list_modules()
                for module in modules:
                    print(f"  {module}")
            elif cmd == 'code' and len(parts) > 1:
                code = controller.get_module_code(parts[1])
                print(code[:1000] + "..." if len(code) > 1000 else code)
            elif cmd == 'exec' and len(parts) > 1:
                statement = ' '.join(parts[1:])
                result = controller.execute_vba_statement(statement)
                if result is not None:
                    print(f"Result: {result}")
            elif cmd == 'showform' and len(parts) > 1:
                success = controller.show_userform(parts[1])
                print(f"Show form: {'Success' if success else 'Failed'}")
            elif cmd == 'hideform' and len(parts) > 1:
                success = controller.hide_userform(parts[1])
                print(f"Hide form: {'Success' if success else 'Failed'}")
            elif cmd == 'info':
                info = controller.get_workbook_info()
                print(json.dumps(info, indent=2))
            else:
                print("Unknown command or missing arguments")
                
        except KeyboardInterrupt:
            print("\nUse 'quit' to exit")
        except Exception as e:
            print(f"ERROR: {e}")

def main():
    parser = argparse.ArgumentParser(description="Excel VBA Controller - Terminal Interface")
    parser.add_argument("--workbook", "-w", help="Path to Excel workbook")
    parser.add_argument("--visible", action="store_true", default=True, help="Make Excel visible")
    parser.add_argument("--hidden", action="store_true", help="Keep Excel hidden")
    parser.add_argument("--run-macro", "-r", help="Run VBA macro")
    parser.add_argument("--macro-args", nargs="*", help="Arguments for macro")
    parser.add_argument("--get-range", help="Get value from range")
    parser.add_argument("--set-range", nargs=2, metavar=("RANGE", "VALUE"), help="Set range value")
    parser.add_argument("--get-name", help="Get value from named range")
    parser.add_argument("--set-name", nargs=2, metavar=("NAME", "VALUE"), help="Set named range value")
    parser.add_argument("--interactive", "-i", action="store_true", help="Interactive mode")
    parser.add_argument("--list-modules", action="store_true", help="List VBA modules")
    parser.add_argument("--show-info", action="store_true", help="Show workbook info")
    
    args = parser.parse_args()
    
    # Determine visibility
    visible = args.visible and not args.hidden
    
    # Create controller
    controller = ExcelVBAController(args.workbook, visible=visible)
    
    try:
        # Connect to Excel
        if not controller.connect():
            sys.exit(1)
        
        # Execute commands
        if args.run_macro:
            macro_args = args.macro_args or []
            result = controller.run_macro(args.run_macro, *macro_args)
            if result is not None:
                print(f"Macro result: {result}")
                
        if args.get_range:
            value = controller.get_range_value(args.get_range)
            print(f"Range {args.get_range}: {value}")
            
        if args.set_range:
            range_addr, value = args.set_range
            success = controller.set_range_value(range_addr, value)
            print(f"Set range {range_addr}: {'Success' if success else 'Failed'}")
            
        if args.list_modules:
            modules = controller.list_modules()
            print("VBA Modules:")
            for module in modules:
                print(f"  {module}")
                
        if args.show_info:
            info = controller.get_workbook_info()
            print(json.dumps(info, indent=2))
            
        if args.interactive:
            interactive_mode(controller)
            
    finally:
        controller.disconnect()

if __name__ == "__main__":
    main()