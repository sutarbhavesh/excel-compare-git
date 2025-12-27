# excel_diff/diff_engine.py

class DiffEngine:
    def __init__(self, excel_a: dict, excel_b: dict):
        self.excel_a = excel_a
        self.excel_b = excel_b

    def compare(self) -> dict:
        result = {}
        all_sheets = set(self.excel_a.keys()) | set(self.excel_b.keys())

        for sheet in all_sheets:
            rows_a = self.excel_a.get(sheet, [])
            rows_b = self.excel_b.get(sheet, [])

            result[sheet] = self._compare_sheet(rows_a, rows_b)

        return result

    def _compare_sheet(self, rows_a, rows_b):
        max_rows = max(len(rows_a), len(rows_b))
        max_cols = max(
            max((len(r) for r in rows_a), default=0),
            max((len(r) for r in rows_b), default=0),
        )

        rows = []

        for r in range(max_rows):
            row_a = rows_a[r] if r < len(rows_a) else []
            row_b = rows_b[r] if r < len(rows_b) else []

            cells = []
            for c in range(max_cols):
                val_a = row_a[c] if c < len(row_a) else ""
                val_b = row_b[c] if c < len(row_b) else ""

                if val_a == val_b:
                    status = "same"
                elif val_a and not val_b:
                    status = "deleted"
                elif val_b and not val_a:
                    status = "added"
                else:
                    status = "modified"

                cells.append({
                    "col": c,
                    "a": val_a,
                    "b": val_b,
                    "status": status
                })

            rows.append({
                "row_index": r,
                "cells": cells
            })

        return {
            "max_cols": max_cols,
            "rows": rows
        }
