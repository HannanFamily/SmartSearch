"""
Mode-Driven Search Engine (Python Prototype)
===========================================
This module demonstrates a mode-driven search selector, where the search logic and output columns change based on the selected mode.
"""
import pandas as pd
from typing import List, Dict, Any

class ModeConfig:
    def __init__(self, mode_name: str, filter_func, output_columns: List[str], description: str = ""):
        self.mode_name = mode_name
        self.filter_func = filter_func  # Callable: (row, params) -> bool
        self.output_columns = output_columns
        self.description = description

class ModeDrivenSearchEngine:
    def __init__(self, data: pd.DataFrame, mode_configs: List[ModeConfig]):
        self.data = data
        self.mode_configs = {mc.mode_name: mc for mc in mode_configs}
        self.active_mode = None
        self.search_params = {}

    def set_mode(self, mode_name: str, search_params: Dict[str, Any] = None):
        if mode_name not in self.mode_configs:
            raise ValueError(f"Mode '{mode_name}' not found.")
        self.active_mode = self.mode_configs[mode_name]
        self.search_params = search_params or {}

    def search(self) -> pd.DataFrame:
        if not self.active_mode:
            raise RuntimeError("No mode selected.")
        # Apply filter function row-wise
        filtered = self.data[self.data.apply(lambda row: self.active_mode.filter_func(row, self.search_params), axis=1)]
        # Select output columns
        return filtered[self.active_mode.output_columns]

# --- Sample Usage ---
def sample_data():
    return pd.DataFrame([
        {'ID': 'EQ001', 'Type': 'Pump', 'Desc': 'Main Water Pump', 'Location': 'A', 'Status': 'Active'},
        {'ID': 'EQ002', 'Type': 'Valve', 'Desc': 'Control Valve', 'Location': 'B', 'Status': 'Active'},
        {'ID': 'EQ003', 'Type': 'Pump', 'Desc': 'Backup Pump', 'Location': 'A', 'Status': 'Inactive'},
        {'ID': 'EQ004', 'Type': 'Motor', 'Desc': 'Pump Motor', 'Location': 'C', 'Status': 'Active'},
        {'ID': 'EQ005', 'Type': 'Valve', 'Desc': 'Relief Valve', 'Location': 'A', 'Status': 'Inactive'},
    ])

def pump_mode_filter(row, params):
    # Only show pumps, optionally filter by status
    if row['Type'] != 'Pump':
        return False
    if 'status' in params and params['status']:
        return row['Status'] == params['status']
    return True

def valve_mode_filter(row, params):
    # Only show valves, optionally filter by location
    if row['Type'] != 'Valve':
        return False
    if 'location' in params and params['location']:
        return row['Location'] == params['location']
    return True

def all_active_mode_filter(row, params):
    # Show all active equipment
    return row['Status'] == 'Active'

def main():
    data = sample_data()
    modes = [
        ModeConfig(
            mode_name='Pump Search',
            filter_func=pump_mode_filter,
            output_columns=['ID', 'Desc', 'Status'],
            description='Show only pumps, filterable by status.'
        ),
        ModeConfig(
            mode_name='Valve by Location',
            filter_func=valve_mode_filter,
            output_columns=['ID', 'Desc', 'Location'],
            description='Show only valves, filterable by location.'
        ),
        ModeConfig(
            mode_name='All Active',
            filter_func=all_active_mode_filter,
            output_columns=['ID', 'Type', 'Desc', 'Status'],
            description='Show all active equipment.'
        )
    ]
    engine = ModeDrivenSearchEngine(data, modes)

    print("\n--- Pump Search (Active Only) ---")
    engine.set_mode('Pump Search', {'status': 'Active'})
    print(engine.search())

    print("\n--- Valve by Location (A) ---")
    engine.set_mode('Valve by Location', {'location': 'A'})
    print(engine.search())

    print("\n--- All Active Equipment ---")
    engine.set_mode('All Active')
    print(engine.search())

if __name__ == "__main__":
    main()
