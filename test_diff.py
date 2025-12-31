import os
import shutil
from excel_diff.excel_parser import ExcelParser
from excel_diff.diff_engine import DiffEngine

def run_test_scenario(name, data_a, data_b):
    print(f"\n--- Testing Scenario: {name} ---")
    try:
        engine = DiffEngine(data_a, data_b)
        results = engine.compare()
        
        for i, sheet in enumerate(results):
            print(f"Sheet {i}: {sheet['name_a']} vs {sheet['name_b']}")
            print(f"  - Match: {sheet['is_match']}")
            print(f"  - Rows found: {len(sheet['data']['rows'])}")
            
            # Check for data presence
            if len(sheet['data']['rows']) > 0:
                print(f"  - Data Check: ✅ Data successfully paired")
            else:
                print(f"  - Data Check: ❌ DATA MISSING!")
                
    except Exception as e:
        print(f"  - ❌ CRASHED: {e}")

# Mock Data representing what ExcelParser outputs
# Scenario 1: Renamed Sheet (The issue we just fixed)
mock_a = {"Sheet1": [["Header"], ["Data A"]]}
mock_b = {"RENAMED_SHEET": [["Header"], ["Data B"]]}

# Scenario 2: Different Column Counts
mock_c = {"Sheet1": [["A", "B"], ["1", "2"]]}
mock_d = {"Sheet1": [["A", "B", "C"], ["1", "2", "3"]]}

if __name__ == "__main__":
    run_test_scenario("Renamed Sheet Pairing", mock_a, mock_b)
    run_test_scenario("Extra Columns Handling", mock_c, mock_d)