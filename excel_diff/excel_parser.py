
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
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        sheets = {}
        for sheet in workbook.findall("a:sheets/a:sheet", ns):
            name = sheet.attrib["name"]
            sheet_id = sheet.attrib["sheetId"]
            path = f"xl/worksheets/sheet{sheet_id}.xml"

            try:
                xml = z.read(path)
            except KeyError:
                continue

            sheets[name] = self._read_sheet(xml, shared_strings)

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
