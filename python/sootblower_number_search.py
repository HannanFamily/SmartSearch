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
    # Complete Sootblower Data sheet - All sections (IK, IR, WB, IKAH)
    # Columns: Type, Number, Floor, Side, SB Cabinet, Cabinet Floor, Cabinet side
    data = [
        # IK Series (Left column)
        {'Type': 'IK', 'Number': 7,   'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 8,   'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 9,   'Floor': 17,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 10,  'Floor': 17,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 15,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 16,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 17,  'Floor': 17,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 18,  'Floor': 17,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 19,  'Floor': 19,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 20,  'Floor': 19,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 21,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 22,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 23,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 24,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 25,  'Floor': 17,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 26,  'Floor': 17,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 27,  'Floor': 17,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 28,  'Floor': 17,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 31,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 32,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 33,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 34,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 35,  'Floor': 17,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 36,  'Floor': 17,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 37,  'Floor': 19,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 38,  'Floor': 19,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 39,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 40,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 41,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 42,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 43,  'Floor': 17,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 44,  'Floor': 17,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 45,  'Floor': 17,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 46,  'Floor': 17,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 49,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 50,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 51,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 52,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 53,  'Floor': 19,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 54,  'Floor': 19,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 55,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 56,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 57,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 58,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 59,  'Floor': 19,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 60,  'Floor': 19,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 61,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 62,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 63,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 64,  'Floor': 18,   'Side': 'E', 'SB Cabinet': 19,   'Cabinet Floor': 19, 'Cabinet side': 'E'},
        {'Type': 'IK', 'Number': 65,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 67,  'Floor': 18,   'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 69,  'Floor': 17,   'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 71,  'Floor': 17,   'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 73,  'Floor': 16.5, 'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 75,  'Floor': 16.5, 'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 77,  'Floor': 16.5, 'Side': 'E', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 79,  'Floor': 15.5, 'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 81,  'Floor': 15,   'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 83,  'Floor': 15,   'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 85,  'Floor': 14.5, 'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        {'Type': 'IK', 'Number': 87,  'Floor': 14.5, 'Side': 'W', 'SB Cabinet': 17,   'Cabinet Floor': 17, 'Cabinet side': 'W'},
        
        # IR Series (Middle column)
        {'Type': 'IR', 'Number': 27,  'Floor': 11, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 29,  'Floor': 11, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 31,  'Floor': 11, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 33,  'Floor': 11, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 35,  'Floor': 11, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 37,  'Floor': 11, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 39,  'Floor': 11, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 41,  'Floor': 11, 'Side': 'E', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 43,  'Floor': 11, 'Side': 'E', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 45,  'Floor': 11, 'Side': 'E', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 47,  'Floor': 11, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 49,  'Floor': 11, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 51,  'Floor': 11, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 54,  'Floor': 12, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 56,  'Floor': 12, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 58,  'Floor': 12, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 60,  'Floor': 12, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 62,  'Floor': 12, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 64,  'Floor': 12, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 66,  'Floor': 12, 'Side': 'E', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 68,  'Floor': 12, 'Side': 'E', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 70,  'Floor': 12, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 72,  'Floor': 12, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 74,  'Floor': 12, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 76,  'Floor': 12, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 78,  'Floor': 12, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 79,  'Floor': 13, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 81,  'Floor': 13, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 83,  'Floor': 13, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 85,  'Floor': 13, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 87,  'Floor': 13, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 89,  'Floor': 13, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 91,  'Floor': 13, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 93,  'Floor': 13, 'Side': 'E', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 95,  'Floor': 13, 'Side': 'E', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 97,  'Floor': 13, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 99,  'Floor': 13, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 101, 'Floor': 13, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 103, 'Floor': 13, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 106, 'Floor': 14, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 108, 'Floor': 14, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 110, 'Floor': 14, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 112, 'Floor': 14, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 114, 'Floor': 14, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 116, 'Floor': 14, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'IR', 'Number': 118, 'Floor': 14, 'Side': 'E', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 120, 'Floor': 14, 'Side': 'E', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 122, 'Floor': 14, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 124, 'Floor': 14, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 126, 'Floor': 14, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 128, 'Floor': 14, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        {'Type': 'IR', 'Number': 130, 'Floor': 14, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'W'},
        
        # WB Series (Right column)
        {'Type': 'WB', 'Number': 28,  'Floor': 11, 'Side': 'W', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 32,  'Floor': 11, 'Side': 'N', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 34,  'Floor': 11, 'Side': 'N', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 36,  'Floor': 11, 'Side': 'N', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 38,  'Floor': 11, 'Side': 'N', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 42,  'Floor': 11, 'Side': 'E', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 46,  'Floor': 11, 'Side': 'S', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 48,  'Floor': 11, 'Side': 'S', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 50,  'Floor': 11, 'Side': 'S', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 59,  'Floor': 12, 'Side': 'N', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 61,  'Floor': 12, 'Side': 'N', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 63,  'Floor': 12, 'Side': 'N', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 71,  'Floor': 12, 'Side': 'S', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 73,  'Floor': 12, 'Side': 'S', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 75,  'Floor': 12, 'Side': 'S', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 77,  'Floor': 12, 'Side': 'S', 'SB Cabinet': 12, 'Cabinet Floor': 12, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 80,  'Floor': 13, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 84,  'Floor': 13, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 86,  'Floor': 13, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 88,  'Floor': 13, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 90,  'Floor': 13, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 94,  'Floor': 13, 'Side': 'W', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 98,  'Floor': 13, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 100, 'Floor': 13, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 102, 'Floor': 13, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 111, 'Floor': 14, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 113, 'Floor': 14, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 115, 'Floor': 14, 'Side': 'N', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 123, 'Floor': 14, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 125, 'Floor': 14, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 127, 'Floor': 14, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        {'Type': 'WB', 'Number': 129, 'Floor': 14, 'Side': 'S', 'SB Cabinet': 13, 'Cabinet Floor': 13, 'Cabinet side': 'E'},
        
        # IKAH Series (Bottom right)
        {'Type': 'IKAH', 'Number': 1, 'Floor': 4, 'Side': 'N', 'SB Cabinet': '', 'Cabinet Floor': '', 'Cabinet side': ''},
        {'Type': 'IKAH', 'Number': 2, 'Floor': 4, 'Side': 'N', 'SB Cabinet': '', 'Cabinet Floor': '', 'Cabinet side': ''},
        {'Type': 'IKAH', 'Number': 3, 'Floor': 4, 'Side': 'N', 'SB Cabinet': '', 'Cabinet Floor': '', 'Cabinet side': ''},
        {'Type': 'IKAH', 'Number': 4, 'Floor': 5, 'Side': 'N', 'SB Cabinet': '', 'Cabinet Floor': '', 'Cabinet side': ''},
        {'Type': 'IKAH', 'Number': 5, 'Floor': 5, 'Side': 'N', 'SB Cabinet': '', 'Cabinet Floor': '', 'Cabinet side': ''},
        {'Type': 'IKAH', 'Number': 6, 'Floor': 5, 'Side': 'N', 'SB Cabinet': '', 'Cabinet Floor': '', 'Cabinet side': ''},
    ]
    return pd.DataFrame(data)

def sootblower_number_filter(row, params):
    # Search across all Sootblower types (IK, IR, WB, IKAH)
    number = params.get('number')
    if number is not None:
        return str(row['Number']) == str(number)
    return False

def main():
    data = sootblower_sample_data()
    mode = ModeConfig(
        mode_name='Sootblower Number',
        filter_func=sootblower_number_filter,
        output_columns=['Type', 'Number', 'Floor', 'Side', 'SB Cabinet', 'Cabinet Floor', 'Cabinet side'],
        description='Search for a Sootblower by number across all types (IK, IR, WB, IKAH).'
    )
    engine = ModeDrivenSearchEngine(data, [mode])

    print("\n--- Search for Sootblower Number 102 (all types) ---")
    engine.set_mode('Sootblower Number', {'number': 102})
    result_102 = engine.search()
    if not result_102.empty:
        print(result_102.to_string(index=False))
    else:
        print("No results found")

    print("\n--- Search for Sootblower Number 27 (should find IR27) ---")
    engine.set_mode('Sootblower Number', {'number': 27})
    result_27 = engine.search()
    if not result_27.empty:
        print(result_27.to_string(index=False))
    else:
        print("No results found")

    print("\n--- Search for Sootblower Number 75 (should find IK75 and WB75) ---")
    engine.set_mode('Sootblower Number', {'number': 75})
    result_75 = engine.search()
    if not result_75.empty:
        print(result_75.to_string(index=False))
    else:
        print("No results found")

if __name__ == "__main__":
    main()
