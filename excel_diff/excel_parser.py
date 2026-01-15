import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import re
import xlrd  # For .xls files

class ExcelParserError(Exception):
    pass

class ExcelParser:
    def __init__(self, file_path: str):
        self.file_path = file_path

    def parse(self) -> dict:
        # Route based on file extension
        if self.file_path.lower().endswith('.xls'):
            return self._parse_xls()
        
        # Existing .xlsx logic
        if not zipfile.is_zipfile(self.file_path):
            raise ExcelParserError("Invalid XLSX file or unsupported format")

        with zipfile.ZipFile(self.file_path, "r") as z:
            shared_strings = self._read_shared_strings(z)
            sheets = self._read_sheets(z, shared_strings)

        return sheets

    def _parse_xls(self) -> dict:
        """Parses legacy .xls files and returns the same structure as XLSX parser."""
        try:
            workbook = xlrd.open_workbook(self.file_path)
            sheets = {}
            for sheet in workbook.sheets():
                rows = []
                for r in range(sheet.nrows):
                    row_data = []
                    for val in sheet.row_values(r):
                        # Clean up numbers: if it's 10.0, treat it as "10"
                        if isinstance(val, float) and val.is_integer():
                            row_data.append(str(int(val)))
                        else:
                            row_data.append(str(val) if val is not None else "")
                    rows.append(row_data)
                sheets[sheet.name] = rows
            return sheets
        except Exception as e:
            raise ExcelParserError(f"Error reading .xls file: {e}")

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
        sheets = {}
        for i, sheet_node in enumerate(workbook.findall("a:sheets/a:sheet", ns)):
            name = sheet_node.attrib["name"]
            path = f"xl/worksheets/sheet{i+1}.xml" 
            try:
                xml = z.read(path)
                sheets[name] = self._read_sheet(xml, shared_strings)
            except KeyError:
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
                ref = cell.attrib.get("r")
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
        match = re.match(r"([A-Z]+)", cell_ref)
        col_letters = match.group(1)
        index = 0
        for char in col_letters:
            index = index * 26 + (ord(char) - ord("A") + 1)
        return index - 1