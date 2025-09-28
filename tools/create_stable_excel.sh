#!/bin/bash
# Quick Excel Stabilization Script
# Adds DevEnvironmentAnalyzer to Excel workbook

echo "=== Creating Stable Excel Version ==="

# Copy current workbook to stable version
cp "Search Dashboard v1.1 DevCOPY.xlsm" "Search Dashboard v1.1 STABLE.xlsm"

# Try to sync the DevEnvironmentAnalyzer
echo "Adding DevEnvironmentAnalyzer module..."

python3 sync_vba_to_excel.py "Search Dashboard v1.1 STABLE.xlsm" 2>/dev/null

if [ $? -eq 0 ]; then
    echo "✅ Successfully created stable Excel version with DevEnvironmentAnalyzer!"
    echo "   File: Search Dashboard v1.1 STABLE.xlsm"
    echo ""
    echo "📋 To use the analyzer:"
    echo "   1. Open Search Dashboard v1.1 STABLE.xlsm"
    echo "   2. Press Alt+F8"
    echo "   3. Run 'AnalyzeDevEnvironment'"
    echo "   4. Check the new analysis worksheets"
else
    echo "⚠️  Sync failed, but stable copy created."
    echo "   Manual steps:"
    echo "   1. Open Search Dashboard v1.1 STABLE.xlsm"
    echo "   2. Press Alt+F11 (VBA Editor)"
    echo "   3. Insert → Module"
    echo "   4. Copy contents of DevEnvironmentAnalyzer.bas"
    echo "   5. Save and run AnalyzeDevEnvironment macro"
fi

echo ""
echo "🎯 Your stable Excel file is ready: Search Dashboard v1.1 STABLE.xlsm"