class DiffEngine:
    def __init__(self, excel_a: dict, excel_b: dict):
        self.excel_a = excel_a
        self.excel_b = excel_b

    def compare(self) -> list:
        result = [] 
        names_a = list(self.excel_a.keys())
        names_b = list(self.excel_b.keys())
        
        num_sheets = max(len(names_a), len(names_b))

        for i in range(num_sheets):
            real_key_a = names_a[i] if i < len(names_a) else None
            real_key_b = names_b[i] if i < len(names_b) else None
            
            # FIX: Only use "MISSING" for the display name, never for the dictionary lookup
            display_name_a = real_key_a if real_key_a else "MISSING"
            display_name_b = real_key_b if real_key_b else "MISSING"
            
            is_match = (real_key_a == real_key_b) if (real_key_a and real_key_b) else False
            
            # Extract rows; if a sheet is missing, pass an empty list []
            rows_a = self.excel_a.get(real_key_a, []) if real_key_a else []
            rows_b = self.excel_b.get(real_key_b, []) if real_key_b else []

            # Even if names don't match, we compare the rows by position
            sheet_diff = self._compare_sheet(rows_a, rows_b)
            
            result.append({
                "name_a": display_name_a,
                "name_b": display_name_b,
                "is_match": is_match,
                "data": sheet_diff
            })
        return result

    def _compare_sheet(self, rows_a, rows_b):
        max_rows = max(len(rows_a), len(rows_b))
        
        # Calculate max columns
        max_cols_a = max((len(r) for r in rows_a), default=0)
        max_cols_b = max((len(r) for r in rows_b), default=0)
        max_cols = max(max_cols_a, max_cols_b)

        rows = []
        for r in range(max_rows):
            row_a = rows_a[r] if r < len(rows_a) else []
            row_b = rows_b[r] if r < len(rows_b) else []

            cells = []
            for c in range(max_cols):
                val_a = row_a[c] if c < len(row_a) else None
                val_b = row_b[c] if c < len(row_b) else None

                v_a = str(val_a).strip() if val_a is not None else ""
                v_b = str(val_b).strip() if val_b is not None else ""

                if v_a == v_b:
                    status = "equal"
                elif v_a and not v_b:
                    status = "deleted"
                elif v_b and not v_a:
                    status = "added"
                else:
                    status = "modified"

                cells.append({
                    "col": c,
                    "a": val_a if val_a is not None else "",
                    "b": val_b if val_b is not None else "",
                    "status": status
                })

            rows.append({"row_index": r+1, "cells": cells})

        return {"max_cols": max_cols, "rows": rows}