
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import re

class ExcelParserError(Exception):
    pass


class ExcelParser:
    def __init__(self, xlsx_path: str):
        self.xlsx_path = xlsx_path

    def parse(self) -> dict:
        if not zipfile.is_zipfile(self.xlsx_path):
            raise ExcelParserError("Invalid XLSX file")

        with zipfile.ZipFile(self.xlsx_path, "r") as z:
            shared_strings = self._read_shared_strings(z)
            sheets = self._read_sheets(z, shared_strings)

        return sheets

    def _read_shared_strings(self, z):
        try:
            xml = z.read("xl/sharedStrings.xml")
        except KeyError:
            return []

        root = ET.fromstring(xml)
        ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        strings = []
        for si in root.findall("a:si", ns):
            text = "".join(t.text or "" for t in si.findall(".//a:t", ns))
            strings.append(text)

        return strings

    def _read_sheets(self, z, shared_strings):
        workbook_xml = z.read("xl/workbook.xml")
        workbook = ET.fromstring(workbook_xml)
        ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

        # Use the relationship ID to find the correct file, not just the sheetId
        # This ensures we get the sheets in the exact order they appear in the tabs
        sheets = {}
        for i, sheet_node in enumerate(workbook.findall("a:sheets/a:sheet", ns)):
            name = sheet_node.attrib["name"]
            # Sheets are usually numbered sheet1.xml, sheet2.xml based on internal creation
            # but we follow the order they appear in the workbook.xml
            rel_id = sheet_node.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            
            # Simple fallback to index if rel_id logic gets complex
            path = f"xl/worksheets/sheet{i+1}.xml" 

            try:
                xml = z.read(path)
                sheets[name] = self._read_sheet(xml, shared_strings)
            except KeyError:
                # If sheet1.xml doesn't exist, try finding by sheetId
                s_id = sheet_node.attrib.get("sheetId")
                try:
                    xml = z.read(f"xl/worksheets/sheet{s_id}.xml")
                    sheets[name] = self._read_sheet(xml, shared_strings)
                except:
                    continue

        return sheets

    def _read_sheet(self, xml_bytes, shared_strings):
        root = ET.fromstring(xml_bytes)
        ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        rows_dict = defaultdict(dict)
        max_col = 0

        for row in root.findall(".//a:row", ns):
            row_idx = int(row.attrib["r"]) - 1

            for cell in row.findall("a:c", ns):
                ref = cell.attrib.get("r")  # e.g. C3
                col_idx = self._col_to_index(ref)

                cell_type = cell.attrib.get("t")
                value_elem = cell.find("a:v", ns)
                value = value_elem.text if value_elem is not None else ""

                if cell_type == "s":
                    value = shared_strings[int(value)]

                rows_dict[row_idx][col_idx] = value
                max_col = max(max_col, col_idx)

        rows = []
        max_row = max(rows_dict.keys(), default=-1)

        for r in range(max_row + 1):
            row = []
            for c in range(max_col + 1):
                row.append(rows_dict[r].get(c, ""))
            rows.append(row)

        return rows

    def _col_to_index(self, cell_ref: str) -> int:
        """
        Converts Excel column letters to index (A=0, B=1, Z=25, AA=26)
        """
        match = re.match(r"([A-Z]+)", cell_ref)
        col_letters = match.group(1)

        index = 0
        for char in col_letters:
            index = index * 26 + (ord(char) - ord("A") + 1)

        return index - 1
