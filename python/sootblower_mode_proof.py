"""
Self-contained proof of Sootblower Location search mode
======================================================
Combines mode-driven search engine and sootblower mode test in one file for AI execution.
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
        filtered = self.data[self.data.apply(lambda row: self.active_mode.filter_func(row, self.search_params), axis=1)]
        return filtered[self.active_mode.output_columns]

def sootblower_sample_data():
    return pd.DataFrame([
        {'ID': 'SB001', 'Type': 'Sootblower', 'Location': 'Boiler 1', 'Floor': 1, 'Side': 'A', 'Status': 'Active'},
        {'ID': 'SB002', 'Type': 'Sootblower', 'Location': 'Boiler 2', 'Floor': 2, 'Side': 'B', 'Status': 'Active'},
        {'ID': 'SB003', 'Type': 'Sootblower', 'Location': 'Boiler 1', 'Floor': 2, 'Side': 'A', 'Status': 'Inactive'},
        {'ID': 'SB004', 'Type': 'Valve', 'Location': 'Boiler 1', 'Floor': 1, 'Side': 'A', 'Status': 'Active'},
        {'ID': 'SB005', 'Type': 'Sootblower', 'Location': 'Boiler 2', 'Floor': 1, 'Side': 'B', 'Status': 'Active'},
    ])

def sootblower_location_filter(row, params):
    if row['Type'] != 'Sootblower':
        return False
    if 'location' in params and params['location']:
        if row['Location'] != params['location']:
            return False
    if 'floor' in params and params['floor']:
        if row['Floor'] != params['floor']:
            return False
    if 'side' in params and params['side']:
        if row['Side'] != params['side']:
            return False
    return True

def main():
    data = sootblower_sample_data()
    sootblower_mode = ModeConfig(
        mode_name='Sootblower Location',
        filter_func=sootblower_location_filter,
        output_columns=['ID', 'Location', 'Floor', 'Side', 'Status'],
        description='Show sootblowers by location, floor, and side.'
    )
    engine = ModeDrivenSearchEngine(data, [sootblower_mode])

    print("\n--- All Sootblowers ---")
    engine.set_mode('Sootblower Location', {})
    print(engine.search())

    print("\n--- Sootblowers in Boiler 1 ---")
    engine.set_mode('Sootblower Location', {'location': 'Boiler 1'})
    print(engine.search())

    print("\n--- Sootblowers in Boiler 2, Floor 1, Side B ---")
    engine.set_mode('Sootblower Location', {'location': 'Boiler 2', 'floor': 1, 'side': 'B'})
    print(engine.search())

if __name__ == "__main__":
    main()
