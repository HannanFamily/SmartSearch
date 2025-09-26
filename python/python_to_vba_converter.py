"""
Python to VBA Code Converter
============================
This tool helps convert Python search logic to VBA equivalents.
Useful for maintaining parallel development environments.
"""

import re
from typing import Dict, List, Any

class PythonToVBAConverter:
    def __init__(self):
        self.type_mapping = {
            'str': 'String',
            'int': 'Long', 
            'float': 'Double',
            'bool': 'Boolean',
            'list': 'Variant',
            'dict': 'Object',
            'pd.DataFrame': 'Variant'
        }
        
        self.function_mapping = {
            'len()': 'UBound() - LBound() + 1',
            'print()': 'Debug.Print',
            '.strip()': 'Trim$()',
            '.lower()': 'LCase$()',
            '.upper()': 'UCase$()',
            'range()': 'For i = 1 To n',
            'enumerate()': 'For i = LBound() To UBound()',
            '.split()': 'Split()',
            '.join()': 'Join()',
            '.append()': 'ReDim Preserve arr(UBound(arr) + 1): arr(UBound(arr)) = value'
        }
    
    def convert_function_signature(self, python_func: str):
        """Convert Python function definition to VBA."""
        # Extract function name and parameters
        match = self.new_method(python_func)
        if not match:
            return python_func
        
        func_name = match.group(1)
        params = match.group(2)
        
        # Convert to VBA function signature
        vba_func = f"Public Function {func_name}("
        
        if params.strip():
            # Process parameters
            param_parts = [p.strip() for p in params.split(',')]
            vba_params = []
            
            for param in param_parts:
                if '=' in param:  # Default value
                    name, default = param.split('=', 1)
                    name = name.strip()
                    default = default.strip().strip('"\'')
                    if default == '""' or default == "''":
                        vba_params.append(f"Optional {name} As String = \"\"")
                    elif default.isdigit():
                        vba_params.append(f"Optional {name} As Long = {default}")
                    else:
                        vba_params.append(f"Optional {name} As Variant = {default}")
                else:
                    # Required parameter
                    name = param.strip()
                    if ':' in name:  # Type hint
                        name, type_hint = name.split(':', 1)
                        name = name.strip()
                        type_hint = type_hint.strip()
                        vba_type = self.type_mapping.get(type_hint, 'Variant')
                        vba_params.append(f"{name} As {vba_type}")
                    else:
                        vba_params.append(f"{name} As Variant")
            
            vba_func += ', '.join(vba_params)
        
        vba_func += ") As Variant"
        return vba_func, func_name  # Return function name as well

    def new_method(self, python_func):
        match = re.match(r'def\s+(\w+)\s*\((.*?)\):', python_func)
        return match
    
    def convert_search_logic(self, python_code: str, function_name: str) -> str:
        """Convert Python search logic to VBA equivalent."""
        vba_code = []
        lines = python_code.split('\n')
        indent_level = 0

        for line in lines:
            stripped = line.strip()
            if not stripped or stripped.startswith('#'):
                continue

            # Convert common patterns
            vba_line = self.convert_line(stripped, indent_level, function_name)
            if vba_line:
                vba_code.append('    ' * indent_level + vba_line)
            
            # Track indentation for control structures
            if stripped.endswith(':'):
                indent_level += 1
            elif not line.startswith(' ') and indent_level > 0:
                indent_level = 0
        
        return '\n'.join(vba_code)
    
    def convert_line(self, line: str, indent: int, function_name: str) -> str:
        """Convert individual Python line to VBA."""
        # Function definitions
        if line.startswith('def '):
            vba_func, _ = self.convert_function_signature(line)
            return vba_func.replace(':', '')

        # Return statements
        if line.startswith('return '):
            value = line[7:]
            return f"Set {function_name} = {value}" if 'DataFrame' in value else f"{function_name} = {value}"
        
        # If statements
        if line.startswith('if '):
            condition = line[3:].rstrip(':')
            return f"If {self.convert_condition(condition)} Then"
        elif line.startswith('elif '):
            condition = line[5:].rstrip(':')
            return f"ElseIf {self.convert_condition(condition)} Then"
        elif line == 'else:':
            return "Else"
        
        # For loops
        if line.startswith('for '):
            return self.convert_for_loop(line)
        
        # Variable assignments
        if '=' in line and not any(op in line for op in ['==', '!=', '<=', '>=', '<', '>']):
            return self.convert_assignment(line)
        
        # Print statements
        if line.startswith('print('):
            content = line[6:-1]
            return f"Debug.Print {content}"
        
        # Method calls
        return self.convert_method_calls(line)
    
    def convert_condition(self, condition: str) -> str:
        """Convert Python conditions to VBA."""
        # Replace Python operators with VBA equivalents
        condition = condition.replace('==', '=')
        condition = condition.replace('!=', '<>')
        condition = condition.replace(' and ', ' And ')
        condition = condition.replace(' or ', ' Or ')
        condition = condition.replace(' not ', ' Not ')
        condition = condition.replace('len(', 'Len(')
        return condition
    
    def convert_assignment(self, line: str) -> str:
        """Convert Python assignments to VBA."""
        parts = line.split('=', 1)
        var_name = parts[0].strip()
        value = parts[1].strip()
        
        # Handle object assignments
        if 'DataFrame' in value or 'dict' in value or 'list' in value:
            return f"Set {var_name} = {value}"
        else:
            return f"Dim {var_name} As Variant: {var_name} = {value}"
    
    def generate_vba_equivalent(self, python_search_function: str) -> str:
        """Generate complete VBA equivalent of Python search function."""
        template = '''
Public Function {function_name}({parameters}) As Variant
    On Error GoTo EH

    Dim dataLo As ListObject
    Set dataLo = lo(DATA_TABLE_NAME)

    {function_body}

    Exit Function
EH:
    LogErrorLocal "{function_name}", Err.Number, Err.Description
    {function_name} = Array()
End Function
'''

        # Extract function details
        lines = python_search_function.strip().split('\n')
        first_line = lines[0]

        # Get function signature and function name
        vba_signature, func_name = self.convert_function_signature(first_line)
        match = re.match(r'Public Function (\w+)\((.*?)\)', vba_signature)
        if match:
            parameters = match.group(2)

            # Convert function body
            body_lines = lines[1:]
            body_code = self.convert_search_logic('\n'.join(body_lines), func_name)

            return template.format(
                function_name=func_name,
                parameters=parameters,
                function_body=body_code
            )

        return python_search_function

def demonstrate_conversion():
    """Demonstrate Python to VBA conversion."""
    converter = PythonToVBAConverter()
    
    python_code = '''
def search_equipment(description_search="", valve_search="", max_results=1000):
    if not description_search.strip() and not valve_search.strip():
        return output_no_results()
    
    results = []
    for row in data:
        if description_search:
            if not matches_description(row, description_search):
                continue
        
        if valve_search:
            if row['valve_number'].lower() != valve_search.lower():
                continue
        
        results.append(row)
        
        if len(results) >= max_results:
            break
    
    return results
'''
    
    print("=== Python to VBA Conversion Demo ===\n")
    print("PYTHON CODE:")
    print(python_code)
    print("\nCONVERTED VBA:")
    vba_equivalent = converter.generate_vba_equivalent(python_code)
    print(vba_equivalent)

if __name__ == "__main__":
    demonstrate_conversion()