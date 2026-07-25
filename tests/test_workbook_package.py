from pathlib import Path
import unittest
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Projet_VBA_BS.xlsm"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


class WorkbookPackageTests(unittest.TestCase):
    def test_package_is_valid_and_keeps_vba(self) -> None:
        with zipfile.ZipFile(WORKBOOK) as package:
            self.assertIsNone(package.testzip())
            self.assertIn("xl/vbaProject.bin", package.namelist())
            vba_project = package.read("xl/vbaProject.bin")

        self.assertIn(b"totalMtmPnl = 0#", vba_project)
        self.assertIn(
            b'Range("E" & rowIndex & ":L" & rowIndex).Delete',
            vba_project,
        )
        self.assertNotIn(b"rawMat > 10", vba_project)
        self.assertNotIn(b"\\Users\\wahid\\", vba_project)
        self.assertNotIn("\\Users\\wahid\\".encode("utf-16le"), vba_project)

    def test_workbook_has_no_personal_path_and_opens_on_portfolio(self) -> None:
        with zipfile.ZipFile(WORKBOOK) as package:
            workbook_xml = package.read("xl/workbook.xml")

        self.assertNotIn(b"absPath", workbook_xml)
        self.assertNotIn(b"OneDrive", workbook_xml)

        root = ET.fromstring(workbook_xml)
        view = root.find(
            f".//{{{SHEET_NS}}}bookViews/{{{SHEET_NS}}}workbookView"
        )
        self.assertIsNotNone(view)
        self.assertEqual(view.attrib.get("activeTab"), "0")

        sheets = {
            sheet.attrib["name"]: sheet
            for sheet in root.findall(f".//{{{SHEET_NS}}}sheet")
        }
        self.assertEqual(sheets["GridPort"].attrib.get("state"), "veryHidden")

    def test_corrected_reference_values_are_stored(self) -> None:
        with zipfile.ZipFile(WORKBOOK) as package:
            grid_xml = package.read("xl/worksheets/sheet3.xml")
            portfolio_xml = package.read("xl/worksheets/sheet1.xml")

        grid = ET.fromstring(grid_xml)
        cells = {
            cell.attrib["r"]: cell
            for cell in grid.findall(f".//{{{SHEET_NS}}}c")
        }

        def number(reference: str) -> float:
            value = cells[reference].find(f"{{{SHEET_NS}}}v")
            self.assertIsNotNone(value)
            return float(value.text)

        self.assertAlmostEqual(number("B2"), 69.80290176895662, places=10)
        self.assertAlmostEqual(number("B12"), 21.63180112062848, places=10)
        self.assertAlmostEqual(number("B92"), 4.991787573706517, places=10)
        self.assertAlmostEqual(number("I2"), 63.86365626771473, places=10)
        self.assertAlmostEqual(number("I3"), 58.86365626771473, places=10)

        portfolio = ET.fromstring(portfolio_xml)
        portfolio_cells = {
            cell.attrib["r"]: cell
            for cell in portfolio.findall(f".//{{{SHEET_NS}}}c")
        }
        self.assertEqual(
            portfolio_cells["F6"].find(f"{{{SHEET_NS}}}v").text,
            "31",
        )
        dividend_label = portfolio_cells["B12"].find(
            f".//{{{SHEET_NS}}}t"
        )
        self.assertEqual(dividend_label.text, "Dividende (q)")

    def test_existing_buttons_are_preserved(self) -> None:
        with zipfile.ZipFile(WORKBOOK) as package:
            portfolio_xml = package.read("xl/worksheets/sheet1.xml")

        self.assertIn(b"macro=\"[0]!Ouvrir_Portfolio\"", portfolio_xml)
        self.assertIn(b"macro=\"[0]!LanceTout\"", portfolio_xml)


if __name__ == "__main__":
    unittest.main()
