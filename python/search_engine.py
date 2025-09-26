"""
Equipment Search Engine - Python Version
========================================
This is a Python implementation of the VBA search engine for development and testing.
AI can execute this code directly to see results and validate logic before VBA conversion.
"""

import pandas as pd
import re
from typing import List, Dict, Any, Optional, Tuple
import json
from pathlib import Path

class EquipmentSearchEngine:
    def __init__(self, data_file: Optional[str] = None, config_file: Optional[str] = None):
        """Initialize the search engine with data and configuration."""
        self.data = pd.DataFrame()
        self.config = {}
        self.mapping = {}  # Synonym mapping
        
        if data_file:
            self.load_data(data_file)
        if config_file:
            self.load_config(config_file)
    
    def load_data(self, file_path: str):
        """Load equipment data from CSV file."""
        try:
            self.data = pd.read_csv(file_path)
            print(f"Loaded {len(self.data)} equipment records")
        except Exception as e:
            print(f"Error loading data: {e}")
    
    def load_config(self, file_path: str):
        """Load configuration from JSON file."""
        try:
            with open(file_path, 'r') as f:
                self.config = json.load(f)
            print("Configuration loaded successfully")
        except Exception as e:
            print(f"Error loading config: {e}")
    
    def build_synonym_index(self, mapping_data: List[Dict[str, str]]) -> Dict[str, List[str]]:
        """Build synonym index from mapping data (equivalent to VBA BuildSynonymIndex)."""
        synonym_index = {}
        
        for mapping in mapping_data:
            raw_term = mapping.get('RawTerm', '').strip().lower()
            standard_term = mapping.get('StandardTerm', '').strip().lower()
            
            if raw_term and standard_term:
                if standard_term not in synonym_index:
                    synonym_index[standard_term] = []
                if raw_term not in synonym_index[standard_term]:
                    synonym_index[standard_term].append(raw_term)
        
        return synonym_index
    
    def build_search_regexes(self, search_text: str, synonym_index: Dict[str, List[str]]) -> List[re.Pattern]:
        """Build regex patterns for search terms with synonym expansion."""
        if not search_text.strip():
            return []
        
        # Split search text into tokens
        tokens = [token.strip().lower() for token in search_text.split() if token.strip()]
        regex_patterns = []
        
        for token in tokens:
            # Build alternation pattern with synonyms
            alternatives = [re.escape(token)]
            
            # Add synonyms
            for standard_term, synonyms in synonym_index.items():
                if token in synonyms or token == standard_term:
                    alternatives.extend([re.escape(syn) for syn in synonyms])
                    alternatives.append(re.escape(standard_term))
            
            # Remove duplicates and create word boundary pattern
            unique_alternatives = list(set(alternatives))
            pattern = r'\b(?:' + '|'.join(unique_alternatives) + r')\b'
            
            try:
                regex_patterns.append(re.compile(pattern, re.IGNORECASE))
            except re.error as e:
                print(f"Regex error for token '{token}': {e}")
                # Fallback to simple word boundary search
                regex_patterns.append(re.compile(r'\b' + re.escape(token) + r'\b', re.IGNORECASE))
        
        return regex_patterns
    
    def search_equipment(self, 
                        description_search: str = "", 
                        valve_search: str = "",
                        max_results: int = 1000) -> pd.DataFrame:
        """
        Main search function - equivalent to VBA PerformSearch.
        
        Args:
            description_search: Description text to search for
            valve_search: Valve number to search for (exact match)
            max_results: Maximum number of results to return
            
        Returns:
            DataFrame with matching equipment records
        """
        if self.data.empty:
            print("No data loaded")
            return pd.DataFrame()
        
        # Start with all visible data (in VBA this would be filtered by slicers)
        results = self.data.copy()
        
        # Apply description search if provided
        if description_search.strip():
            # Build synonym mapping (in real implementation, load from data)
            sample_mapping = [
                {"RawTerm": "pump", "StandardTerm": "pumping equipment"},
                {"RawTerm": "motor", "StandardTerm": "electric motor"},
                {"RawTerm": "valve", "StandardTerm": "control valve"}
            ]
            
            synonym_index = self.build_synonym_index(sample_mapping)
            regex_patterns = self.build_search_regexes(description_search, synonym_index)
            
            # Apply description filter
            if regex_patterns:
                description_column = 'Equipment Description'  # Configurable
                if description_column in results.columns:
                    mask = results[description_column].str.contains(
                        '|'.join([pattern.pattern for pattern in regex_patterns]), 
                        case=False, na=False, regex=True
                    )
                    results = results[mask]
                else:
                    print(f"Warning: Description column '{description_column}' not found")
        
        # Apply valve number search if provided
        if valve_search.strip():
            valve_column = 'Valve Number'  # Configurable
            if valve_column in results.columns:
                # Exact match for valve number
                results = results[results[valve_column].astype(str).str.lower() == valve_search.lower()]
            else:
                print(f"Warning: Valve column '{valve_column}' not found")
        
        # Limit results
        if len(results) > max_results:
            results = results.head(max_results)
            print(f"Results limited to {max_results} records")
        
        # Sort by description (equivalent to VBA sorting)
        description_column = 'Equipment Description'
        if description_column in results.columns:
            results = results.sort_values(by=description_column)
        
        return results
    
    def output_no_results(self) -> pd.DataFrame:
        """Return empty DataFrame with column headers (equivalent to VBA OutputNoResults)."""
        if self.data.empty:
            return pd.DataFrame()
        
        # Return empty DataFrame with same columns
        empty_result = pd.DataFrame(columns=self.data.columns)
        print("Enter search criteria to display results.")
        return empty_result
    
    def output_all_visible(self, max_results: int = 1000) -> pd.DataFrame:
        """Return all visible data (equivalent to VBA OutputAllVisible)."""
        if self.data.empty:
            return pd.DataFrame()
        
        results = self.data.copy()
        
        if len(results) > max_results:
            results = results.head(max_results)
            print(f"Displayed {max_results} visible rows.")
        else:
            print(f"Displayed {len(results)} visible rows.")
        
        return results
    
    def refresh_results(self, description_search: str = "", valve_search: str = "") -> pd.DataFrame:
        """
        Main entry point - equivalent to VBA RefreshResults.
        Decides whether to search, show all, or show no results.
        """
        # Check if we have active search criteria
        desc_active = len(description_search.strip()) > 0
        valve_active = len(valve_search.strip()) >= 3  # Minimum length like VBA
        
        if desc_active or valve_active:
            print(f"Performing search: desc='{description_search}', valve='{valve_search}'")
            return self.search_equipment(description_search, valve_search)
        else:
            print("No search criteria provided - showing no results")
            return self.output_no_results()  # Changed from output_all_visible to match VBA update
    
    def get_column_info(self) -> Dict[str, Any]:
        """Get information about available columns."""
        if self.data.empty:
            return {}
        
        return {
            'columns': list(self.data.columns),
            'row_count': len(self.data),
            'data_types': self.data.dtypes.to_dict(),
            'sample_row': self.data.head(1).to_dict('records')[0] if len(self.data) > 0 else {}
        }

def create_sample_data():
    """Create sample equipment data for testing."""
    sample_data = [
        {
            'SAP Equipment ID': 'EQ001',
            'Equipment Description': 'Primary Cooling Water Pump',
            'Functional System': 'Cooling Water',
            'Work Area': 'Plant A',
            'Valve Number': 'V001',
            'Object Type': 'Pump',
            'Physical Location': 'Building 1'
        },
        {
            'SAP Equipment ID': 'EQ002', 
            'Equipment Description': 'Emergency Diesel Generator',
            'Functional System': 'Emergency Power',
            'Work Area': 'Plant B',
            'Valve Number': 'V002',
            'Object Type': 'Generator',
            'Physical Location': 'Building 2'
        },
        {
            'SAP Equipment ID': 'EQ003',
            'Equipment Description': 'Control Room Air Handler',
            'Functional System': 'HVAC',
            'Work Area': 'Control Room',
            'Valve Number': '',
            'Object Type': 'Air Handler',
            'Physical Location': 'Control Building'
        },
        {
            'SAP Equipment ID': 'EQ004',
            'Equipment Description': 'Main Steam Valve Assembly',
            'Functional System': 'Steam System',
            'Work Area': 'Plant A',
            'Valve Number': 'V004',
            'Object Type': 'Valve',
            'Physical Location': 'Steam Header'
        },
        {
            'SAP Equipment ID': 'EQ005',
            'Equipment Description': 'Backup Water Pump Motor',
            'Functional System': 'Cooling Water',
            'Work Area': 'Plant A', 
            'Valve Number': '',
            'Object Type': 'Motor',
            'Physical Location': 'Pump House'
        }
    ]
    
    return pd.DataFrame(sample_data)

# Demonstration and testing
if __name__ == "__main__":
    print("=== Equipment Search Engine - Python Version ===\n")
    
    # Create search engine instance
    engine = EquipmentSearchEngine()
    
    # Load sample data
    sample_df = create_sample_data()
    engine.data = sample_df
    
    print("Sample Data Loaded:")
    print(engine.data.to_string(index=False))
    print(f"\nTotal records: {len(engine.data)}")
    
    # Test column info
    print("\n=== Column Information ===")
    col_info = engine.get_column_info()
    for key, value in col_info.items():
        if key != 'sample_row':
            print(f"{key}: {value}")
    
    # Test different search scenarios
    print("\n=== Test 1: No search criteria (should show no results) ===")
    results1 = engine.refresh_results()
    print(f"Results: {len(results1)} rows")
    
    print("\n=== Test 2: Description search for 'pump' ===")
    results2 = engine.refresh_results(description_search="pump")
    print(f"Results: {len(results2)} rows")
    if not results2.empty:
        print(results2[['SAP Equipment ID', 'Equipment Description']].to_string(index=False))
    
    print("\n=== Test 3: Valve number search ===")
    results3 = engine.refresh_results(valve_search="V001")
    print(f"Results: {len(results3)} rows")
    if not results3.empty:
        print(results3[['SAP Equipment ID', 'Equipment Description', 'Valve Number']].to_string(index=False))
    
    print("\n=== Test 4: Combined search ===")
    results4 = engine.refresh_results(description_search="water", valve_search="V001")
    print(f"Results: {len(results4)} rows")
    if not results4.empty:
        print(results4[['SAP Equipment ID', 'Equipment Description', 'Valve Number']].to_string(index=False))
    
    print("\n=== Python Search Engine Demo Complete ===")