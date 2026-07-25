from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"


class VbaSourceRegressionTests(unittest.TestCase):
    def test_sources_are_utf8(self) -> None:
        for path in SRC.iterdir():
            if path.suffix.lower() in {".bas", ".cls", ".frm"}:
                path.read_text(encoding="utf-8")

    def test_grid_totals_are_reset_for_every_spot(self) -> None:
        source = (SRC / "Portfolio.bas").read_text(encoding="utf-8")
        outer_loop = source.index("For pointIndex = 1 To pointCount")
        inner_loop = source.index("For positionIndex = 1 To positions.Count")
        reset_block = source[outer_loop:inner_loop]

        for variable in (
            "totalValue",
            "totalDelta",
            "totalGamma",
            "totalVega",
            "totalTheta",
            "totalRho",
            "totalExpiryPnl",
            "totalMtmPnl",
        ):
            self.assertIn(f"{variable} = 0#", reset_block)

    def test_maturity_is_always_converted_from_days(self) -> None:
        source = (SRC / "Portfolio.bas").read_text(encoding="utf-8")
        self.assertIn("maturityDays / DAYS_IN_YEAR", source)
        self.assertIn("Abs(maturityDays - Fix(maturityDays))", source)
        self.assertNotIn("rawMat > 10", source)

    def test_legacy_mixed_underlying_sample_is_migrated(self) -> None:
        source = (SRC / "DesignModule.bas").read_text(encoding="utf-8")
        self.assertIn("MigrateLegacySampleData ws", source)
        self.assertIn('ws.Range("F6").Value = "Asset"', source)

    def test_deletion_cannot_remove_the_dashboard_row(self) -> None:
        source = (SRC / "frmPortfolio.frm").read_text(encoding="utf-8")
        self.assertNotIn("Rows(i).Delete", source)
        self.assertNotIn("Rows(rowIndex).Delete", source)
        self.assertIn('Range("E" & rowIndex & ":L" & rowIndex).Delete', source)

    def test_no_full_sheet_or_full_column_formatting(self) -> None:
        portfolio = (SRC / "Portfolio.bas").read_text(encoding="utf-8")
        design = (SRC / "DesignModule.bas").read_text(encoding="utf-8")
        self.assertNotIn(".Cells.Interior.Color", portfolio)
        self.assertNotIn('Range("M:T").Clear', design)

    def test_message_boxes_use_vba_line_breaks(self) -> None:
        for name in ("Portfolio.bas", "DesignModule.bas"):
            source = (SRC / name).read_text(encoding="utf-8")
            self.assertNotIn(r"\n", source)

    def test_option_explicit_is_present(self) -> None:
        for name in ("Portfolio.bas", "DesignModule.bas", "Classe1.cls", "frmPortfolio.frm"):
            source = (SRC / name).read_text(encoding="utf-8")
            self.assertIn("Option Explicit", source)

    def test_vba_procedure_blocks_are_balanced(self) -> None:
        for name in ("Portfolio.bas", "DesignModule.bas", "Classe1.cls", "frmPortfolio.frm"):
            lines = [
                line.strip()
                for line in (SRC / name).read_text(encoding="utf-8").splitlines()
            ]
            starts = sum(
                line.startswith(("Public Sub ", "Private Sub ", "Sub "))
                for line in lines
            )
            functions = sum(
                line.startswith(
                    ("Public Function ", "Private Function ", "Function ")
                )
                for line in lines
            )
            properties = sum(
                line.startswith(("Public Property ", "Private Property "))
                for line in lines
            )
            self.assertEqual(starts, lines.count("End Sub"), name)
            self.assertEqual(functions, lines.count("End Function"), name)
            self.assertEqual(properties, lines.count("End Property"), name)


if __name__ == "__main__":
    unittest.main()
