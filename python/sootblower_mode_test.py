"""
Test for Sootblower Location Search Mode
========================================
This script adds and tests a 'Sootblower Location' mode in the mode-driven search engine.
"""
import pandas as pd
from mode_search_engine import ModeConfig, ModeDrivenSearchEngine

def sootblower_sample_data():
    return pd.DataFrame([
        {'ID': 'SB001', 'Type': 'Sootblower', 'Location': 'Boiler 1', 'Floor': 1, 'Side': 'A', 'Status': 'Active'},
        {'ID': 'SB002', 'Type': 'Sootblower', 'Location': 'Boiler 2', 'Floor': 2, 'Side': 'B', 'Status': 'Active'},
        {'ID': 'SB003', 'Type': 'Sootblower', 'Location': 'Boiler 1', 'Floor': 2, 'Side': 'A', 'Status': 'Inactive'},
        {'ID': 'SB004', 'Type': 'Valve', 'Location': 'Boiler 1', 'Floor': 1, 'Side': 'A', 'Status': 'Active'},
        {'ID': 'SB005', 'Type': 'Sootblower', 'Location': 'Boiler 2', 'Floor': 1, 'Side': 'B', 'Status': 'Active'},
    ])

def sootblower_location_filter(row, params):
    # Only show Sootblowers, filter by location, floor, and side if provided
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
