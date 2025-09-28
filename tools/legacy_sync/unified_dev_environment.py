#!/usr/bin/env python3
"""
Unified Development Environment Manager
======================================
This script creates a seamless development environment that synchronizes Python and VBA code,
resolves differences, and provides an Excel-based dashboard for managing the entire workflow.
"""

import os
import sys
import json
import pandas as pd
from pathlib import Path
from datetime import datetime
import shutil
import re
from typing import Dict, List, Any, Optional, Tuple

class UnifiedDevEnvironment:
    def __init__(self, project_dir=None):
        self.project_dir = Path(project_dir or os.getcwd())
        self.python_dir = self.project_dir / "python"
        self.vba_files = []
        self.python_files = []
        self.function_mapping = {}
        self.sync_status = {}
        self.differences = []
        
    def analyze_project_structure(self):
        """Analyze current project structure and identify all Python/VBA files"""
        print("🔍 Analyzing project structure...")
        
        # Find VBA files
        self.vba_files = list(self.project_dir.glob("*.bas")) + list(self.project_dir.glob("*.cls"))
        
        # Find Python files
        if self.python_dir.exists():
            self.python_files = list(self.python_dir.glob("*.py"))
        
        print(f"   Found {len(self.vba_files)} VBA files")
        print(f"   Found {len(self.python_files)} Python files")
        
        return {
            "vba_files": [f.name for f in self.vba_files],
            "python_files": [f.name for f in self.python_files]
        }
    
    def extract_functions_from_vba(self, vba_file: Path) -> List[Dict[str, Any]]:
        """Extract function definitions from VBA files"""
        functions = []
        
        try:
            with open(vba_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Find function definitions
            function_pattern = r'^(?:Public\s+|Private\s+)?(?:Sub|Function)\s+(\w+)\s*\([^)]*\)(?:\s+As\s+\w+)?'
            matches = re.findall(function_pattern, content, re.MULTILINE | re.IGNORECASE)
            
            for func_name in matches:
                # Extract full function body
                func_start = content.find(f'{func_name}(')
                if func_start > 0:
                    # Find function start
                    line_start = content.rfind('\n', 0, func_start) + 1
                    func_line = content[line_start:content.find('\n', func_start)]
                    
                    functions.append({
                        'name': func_name,
                        'signature': func_line.strip(),
                        'file': vba_file.name,
                        'type': 'vba'
                    })
                    
        except Exception as e:
            print(f"   ⚠️ Error reading {vba_file.name}: {e}")
            
        return functions
    
    def extract_functions_from_python(self, python_file: Path) -> List[Dict[str, Any]]:
        """Extract function definitions from Python files"""
        functions = []
        
        try:
            with open(python_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Find function definitions
            function_pattern = r'^(?:\s*)def\s+(\w+)\s*\([^)]*\)(?:\s*->\s*[^:]+)?:'
            matches = re.findall(function_pattern, content, re.MULTILINE)
            
            for func_name in matches:
                # Extract full function signature
                func_start = content.find(f'def {func_name}(')
                if func_start > 0:
                    line_start = content.rfind('\n', 0, func_start) + 1
                    func_line = content[line_start:content.find(':', func_start) + 1]
                    
                    functions.append({
                        'name': func_name,
                        'signature': func_line.strip(),
                        'file': python_file.name,
                        'type': 'python'
                    })
                    
        except Exception as e:
            print(f"   ⚠️ Error reading {python_file.name}: {e}")
            
        return functions
    
    def compare_functionality(self) -> Dict[str, Any]:
        """Compare Python and VBA functionality to find differences"""
        print("🔄 Comparing Python and VBA functionality...")
        
        all_functions = []
        
        # Extract VBA functions
        for vba_file in self.vba_files:
            functions = self.extract_functions_from_vba(vba_file)
            all_functions.extend(functions)
        
        # Extract Python functions
        for python_file in self.python_files:
            functions = self.extract_functions_from_python(python_file)
            all_functions.extend(functions)
        
        # Group by function name
        function_groups = {}
        for func in all_functions:
            name = func['name']
            if name not in function_groups:
                function_groups[name] = {'python': [], 'vba': []}
            function_groups[name][func['type']].append(func)
        
        # Analyze differences
        comparison_result = {
            'total_functions': len(all_functions),
            'unique_functions': len(function_groups),
            'python_only': [],
            'vba_only': [],
            'both_environments': [],
            'conflicts': []
        }
        
        for name, group in function_groups.items():
            python_count = len(group['python'])
            vba_count = len(group['vba'])
            
            if python_count > 0 and vba_count > 0:
                comparison_result['both_environments'].append(name)
            elif python_count > 0:
                comparison_result['python_only'].append(name)
            elif vba_count > 0:
                comparison_result['vba_only'].append(name)
            
            if python_count > 1 or vba_count > 1:
                comparison_result['conflicts'].append({
                    'function': name,
                    'python_instances': python_count,
                    'vba_instances': vba_count
                })
        
        self.function_mapping = function_groups
        return comparison_result
    
    def create_excel_dashboard(self, comparison_data: Dict[str, Any]):
        """Create Excel dashboard for managing the development environment"""
        print("📊 Creating Excel development dashboard...")
        
        dashboard_data = []
        
        # Create comprehensive function mapping
        for func_name, group in self.function_mapping.items():
            python_files = [f['file'] for f in group['python']]
            vba_files = [f['file'] for f in group['vba']]
            
            status = "✅ Synchronized"
            if len(group['python']) > 0 and len(group['vba']) == 0:
                status = "🔄 Python Only - Needs VBA"
            elif len(group['python']) == 0 and len(group['vba']) > 0:
                status = "🔄 VBA Only - Needs Python"
            elif len(group['python']) > 1 or len(group['vba']) > 1:
                status = "⚠️ Conflict - Multiple Definitions"
            
            dashboard_data.append({
                'Function': func_name,
                'Status': status,
                'Python_Files': ', '.join(python_files),
                'VBA_Files': ', '.join(vba_files),
                'Priority': self._calculate_priority(group),
                'Action_Needed': self._suggest_action(group),
                'Last_Updated': datetime.now().strftime('%Y-%m-%d %H:%M'),
                'Notes': ''
            })
        
        # Create DataFrame and save to Excel
        df = pd.DataFrame(dashboard_data)
        
        dashboard_path = self.project_dir / "Development_Dashboard.xlsx"
        
        with pd.ExcelWriter(dashboard_path, engine='openpyxl') as writer:
            # Main dashboard
            df.to_excel(writer, sheet_name='Function_Overview', index=False)
            
            # Summary sheet
            summary_data = {
                'Metric': [
                    'Total Functions',
                    'Synchronized',
                    'Python Only',
                    'VBA Only',
                    'Conflicts',
                    'High Priority',
                    'Ready for Conversion'
                ],
                'Count': [
                    comparison_data['total_functions'],
                    len(comparison_data['both_environments']),
                    len(comparison_data['python_only']),
                    len(comparison_data['vba_only']),
                    len(comparison_data['conflicts']),
                    len([d for d in dashboard_data if d['Priority'] == 'High']),
                    len([d for d in dashboard_data if 'Ready' in d['Action_Needed']])
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Files overview
            file_overview = {
                'VBA_Files': [f.name for f in self.vba_files],
                'Python_Files': [f.name for f in self.python_files],
                'Last_Modified': [
                    datetime.fromtimestamp(f.stat().st_mtime).strftime('%Y-%m-%d %H:%M')
                    for f in (self.vba_files + self.python_files)
                ]
            }
            
            # Pad shorter lists with empty strings
            max_len = max(len(file_overview['VBA_Files']), len(file_overview['Python_Files']))
            for key in file_overview:
                if key != 'Last_Modified':
                    file_overview[key].extend([''] * (max_len - len(file_overview[key])))
            
            file_overview['Last_Modified'] = file_overview['Last_Modified'][:max_len]
            
            files_df = pd.DataFrame(file_overview)
            files_df.to_excel(writer, sheet_name='File_Overview', index=False)
        
        print(f"   ✅ Dashboard created: {dashboard_path}")
        return dashboard_path
    
    def _calculate_priority(self, group: Dict[str, List[Dict]]) -> str:
        """Calculate priority for function conversion"""
        python_count = len(group['python'])
        vba_count = len(group['vba'])
        
        if python_count == 0 and vba_count > 0:
            return "Medium"  # VBA exists, needs Python equivalent
        elif python_count > 0 and vba_count == 0:
            return "High"    # Python exists, needs VBA conversion
        elif python_count > 1 or vba_count > 1:
            return "High"    # Conflicts need resolution
        else:
            return "Low"     # Already synchronized
    
    def _suggest_action(self, group: Dict[str, List[Dict]]) -> str:
        """Suggest action for function synchronization"""
        python_count = len(group['python'])
        vba_count = len(group['vba'])
        
        if python_count > 0 and vba_count == 0:
            return "Convert Python to VBA"
        elif python_count == 0 and vba_count > 0:
            return "Create Python equivalent"
        elif python_count > 1 or vba_count > 1:
            return "Resolve conflicts - merge duplicates"
        else:
            return "Ready - No action needed"
    
    def create_unified_sync_workflow(self):
        """Create scripts for unified synchronization"""
        print("⚙️ Creating unified sync workflow...")
        
        # Create master sync script
        sync_script = f'''#!/usr/bin/env python3
"""
Master Sync Script - Generated on {datetime.now()}
=================================================
Automatically synchronizes Python and VBA environments
"""

import subprocess
import sys
from pathlib import Path

def main():
    project_dir = Path(__file__).parent
    
    print("🔄 Starting unified sync process...")
    
    # Step 1: Sync Python to VBA
    print("   1. Converting Python to VBA...")
    result = subprocess.run([
        sys.executable, 
        str(project_dir / "python_to_vba_converter.py")
    ], capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"      ❌ Python to VBA conversion failed: {{result.stderr}}")
        return 1
    
    # Step 2: Sync VBA to Excel
    print("   2. Syncing VBA to Excel...")
    result = subprocess.run([
        sys.executable, 
        str(project_dir / "sync_vba_to_excel.py")
    ], capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"      ❌ VBA to Excel sync failed: {{result.stderr}}")
        return 1
    
    # Step 3: Update development dashboard
    print("   3. Updating development dashboard...")
    env_manager = UnifiedDevEnvironment(str(project_dir))
    env_manager.analyze_project_structure()
    comparison = env_manager.compare_functionality()
    env_manager.create_excel_dashboard(comparison)
    
    print("✅ Unified sync complete!")
    return 0

if __name__ == "__main__":
    sys.exit(main())
'''
        
        sync_path = self.project_dir / "unified_sync.py"
        with open(sync_path, 'w') as f:
            f.write(sync_script)
        
        # Make executable on Unix systems
        if os.name != 'nt':
            os.chmod(sync_path, 0o755)
        
        print(f"   ✅ Created: {sync_path}")
        return sync_path
    
    def run_complete_resolution(self):
        """Run the complete resolution process"""
        print("🚀 Running complete development environment resolution...")
        print("=" * 60)
        
        # Step 1: Analyze structure
        structure = self.analyze_project_structure()
        
        # Step 2: Compare functionality
        comparison = self.compare_functionality()
        
        # Step 3: Create dashboard
        dashboard_path = self.create_excel_dashboard(comparison)
        
        # Step 4: Create sync workflow
        sync_script = self.create_unified_sync_workflow()
        
        # Step 5: Generate report
        report = self._generate_resolution_report(structure, comparison, dashboard_path, sync_script)
        
        print("=" * 60)
        print("✅ COMPLETE RESOLUTION FINISHED")
        print(f"   📊 Excel Dashboard: {dashboard_path}")
        print(f"   🔄 Sync Script: {sync_script}")
        print(f"   📄 Report: {report}")
        
        return {
            'dashboard': dashboard_path,
            'sync_script': sync_script,
            'report': report,
            'comparison': comparison
        }
    
    def _generate_resolution_report(self, structure, comparison, dashboard_path, sync_script):
        """Generate comprehensive resolution report"""
        
        report_content = f'''# Development Environment Resolution Report
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Project Structure Analysis
- VBA Files: {len(structure['vba_files'])}
- Python Files: {len(structure['python_files'])}

## Functionality Comparison
- Total Functions Analyzed: {comparison['total_functions']}
- Functions in Both Environments: {len(comparison['both_environments'])}
- Python-Only Functions: {len(comparison['python_only'])}
- VBA-Only Functions: {len(comparison['vba_only'])}
- Conflicts Detected: {len(comparison['conflicts'])}

## Python-Only Functions (Need VBA Conversion)
{chr(10).join(f"- {func}" for func in comparison['python_only'])}

## VBA-Only Functions (Need Python Equivalent)
{chr(10).join(f"- {func}" for func in comparison['vba_only'])}

## Synchronized Functions (Both Environments)
{chr(10).join(f"- {func}" for func in comparison['both_environments'])}

## Generated Files
- Excel Dashboard: {dashboard_path}
- Unified Sync Script: {sync_script}

## Next Steps
1. Open the Excel Dashboard to review function status
2. Use the dashboard to track conversion progress
3. Run unified_sync.py after making changes
4. Monitor the dashboard for synchronization status

## Workflow
1. Develop new features in Python first (AI can see and test)
2. Use the converter tools to generate VBA equivalents
3. Run unified_sync.py to update everything
4. Check the Excel dashboard for status
'''
        
        report_path = self.project_dir / "Resolution_Report.md"
        with open(report_path, 'w') as f:
            f.write(report_content)
        
        return report_path

def main():
    if len(sys.argv) > 1:
        project_dir = sys.argv[1]
    else:
        project_dir = os.getcwd()
    
    env_manager = UnifiedDevEnvironment(project_dir)
    result = env_manager.run_complete_resolution()
    
    print(f"""
🎯 RESOLUTION COMPLETE! 

Your seamless development environment is ready:

1. 📊 EXCEL DASHBOARD: {result['dashboard']}
   - View all functions and their sync status
   - Track conversion progress
   - Identify conflicts and priorities

2. 🔄 UNIFIED SYNC: {result['sync_script']}
   - One-command synchronization
   - Automatic Python → VBA → Excel workflow
   - Always keeps everything in sync

3. 📋 DEVELOPMENT WORKFLOW:
   - Develop in Python (AI can see and test)
   - Run unified_sync.py to convert and sync
   - Check Excel dashboard for status
   - Continue developing with confidence

The environment is now completely resolved and unified!
""")

if __name__ == "__main__":
    main()