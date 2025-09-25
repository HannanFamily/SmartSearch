"""
Sootblower Number Search Mode Proof
===================================
Search for a Sootblower by its number and return all data, especially location and power supply.
"""
import pandas as pd

class ModeConfig:
    def __init__(self, mode_name, filter_func, output_columns, description=""):
        self.mode_name = mode_name
        self.filter_func = filter_func
        self.output_columns = output_columns
        self.description = description

class ModeDrivenSearchEngine:
    def __init__(self, data, mode_configs):
        self.data = data
        self.mode_configs = {mc.mode_name: mc for mc in mode_configs}
        self.active_mode = None
        self.search_params = {}

    def set_mode(self, mode_name, search_params=None):
        if mode_name not in self.mode_configs:
            raise ValueError(f"Mode '{mode_name}' not found.")
        self.active_mode = self.mode_configs[mode_name]
        self.search_params = search_params or {}

    def search(self):
        if not self.active_mode:
            raise RuntimeError("No mode selected.")
        filtered = self.data[self.data.apply(lambda row: self.active_mode.filter_func(row, self.search_params), axis=1)]
        return filtered[self.active_mode.output_columns]

def sootblower_sample_data():
    # Realistic Sootblower DataTable sample (Table1: Sootblower Data)
    # Headers: Type, Number, Floor, Side, SB Cabinet, Cabinet Floor, Cabinet side
    return pd.DataFrame([
        {'Type': 'Sootblower', 'Number': 101, 'Floor': 1, 'Side': 'A', 'SB Cabinet': 'SB-CAB1', 'Cabinet Floor': 1, 'Cabinet side': 'A'},
        {'Type': 'Sootblower', 'Number': 102, 'Floor': 1, 'Side': 'A', 'SB Cabinet': 'SB-CAB1', 'Cabinet Floor': 1, 'Cabinet side': 'A'},
        {'Type': 'Sootblower', 'Number': 103, 'Floor': 1, 'Side': 'A', 'SB Cabinet': 'SB-CAB2', 'Cabinet Floor': 1, 'Cabinet side': 'A'},
        {'Type': 'Sootblower', 'Number': 104, 'Floor': 1, 'Side': 'A', 'SB Cabinet': 'SB-CAB2', 'Cabinet Floor': 1, 'Cabinet side': 'A'},
        {'Type': 'Valve',      'Number': 201, 'Floor': 1, 'Side': 'A', 'SB Cabinet': '',        'Cabinet Floor': '', 'Cabinet side': ''},
    ])

def sootblower_number_filter(row, params):
    # Search for Sootblower by number (as int or str)
    if row['Type'] != 'Sootblower':
        return False
    number = params.get('number')
    if number is not None:
        # Accept both int and str input
        return str(row['Number']) == str(number)
    return False

def main():
    data = sootblower_sample_data()
    mode = ModeConfig(
        mode_name='Sootblower Number',
        filter_func=sootblower_number_filter,
        output_columns=['Type', 'Number', 'Floor', 'Side', 'SB Cabinet', 'Cabinet Floor', 'Cabinet side'],
        description='Search for a Sootblower by number and return all data.'
    )
    engine = ModeDrivenSearchEngine(data, [mode])

    print("\n--- Search for Sootblower Number 102 ---")
    engine.set_mode('Sootblower Number', {'number': 102})
    print(engine.search())

    print("\n--- Search for Sootblower Number 105 ---")
    engine.set_mode('Sootblower Number', {'number': 105})
    print(engine.search())

    print("\n--- Search for Sootblower Number 999 (not found) ---")
    engine.set_mode('Sootblower Number', {'number': 999})
    print(engine.search())

if __name__ == "__main__":
    main()
