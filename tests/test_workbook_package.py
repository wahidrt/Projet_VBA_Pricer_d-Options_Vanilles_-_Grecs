from hashlib import sha256
from pathlib import Path
import unittest
import zipfile


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Projet_VBA_BS.xlsm"
INSTALLER = ROOT / "Installer-Corrections.cmd"
SYNC_SCRIPT = ROOT / "scripts" / "Sync-VbaProject.ps1"

# Ce classeur a été enregistré par Microsoft Excel. Toute modification du
# binaire doit être faite par Excel, puis cette empreinte mise à jour
# volontairement. Une simple archive ZIP valide ne garantit pas qu'Excel
# acceptera le projet VBA embarqué.
TRUSTED_WORKBOOK_SHA256 = (
    "c7d2df22f575ae802ae1f3c4565690e2d29468ec2fac6318508ad9d7b798f073"
)
TRUSTED_VBA_PROJECT_SHA256 = (
    "964e02b17b7645136162950e4e35ce554b0ab8634704aa706cb3962a71175fb5"
)


class WorkbookPackageTests(unittest.TestCase):
    def test_excel_generated_template_is_unchanged(self) -> None:
        self.assertEqual(
            sha256(WORKBOOK.read_bytes()).hexdigest(),
            TRUSTED_WORKBOOK_SHA256,
        )

    def test_package_is_valid_and_keeps_trusted_vba(self) -> None:
        with zipfile.ZipFile(WORKBOOK) as package:
            self.assertIsNone(package.testzip())
            self.assertIn("xl/vbaProject.bin", package.namelist())
            vba_project = package.read("xl/vbaProject.bin")

        self.assertEqual(
            sha256(vba_project).hexdigest(),
            TRUSTED_VBA_PROJECT_SHA256,
        )

    def test_existing_buttons_are_preserved(self) -> None:
        with zipfile.ZipFile(WORKBOOK) as package:
            portfolio_xml = package.read("xl/worksheets/sheet1.xml")

        self.assertIn(b'macro="[0]!Ouvrir_Portfolio"', portfolio_xml)
        self.assertIn(b'macro="[0]!LanceTout"', portfolio_xml)

    def test_installer_creates_a_separate_corrected_workbook(self) -> None:
        installer = INSTALLER.read_text(encoding="utf-8")
        sync_script = SYNC_SCRIPT.read_text(encoding="utf-8")

        self.assertIn("Sync-VbaProject.ps1", installer)
        self.assertIn("Projet_VBA_BS_corrige.xlsm", installer)
        self.assertIn("Projet_VBA_BS_corrige.xlsm", sync_script)
        self.assertIn("$workbook.SaveAs($resolvedOutput, 52)", sync_script)
        self.assertIn(
            "Le fichier de sortie doit être différent du classeur modèle.",
            sync_script,
        )


if __name__ == "__main__":
    unittest.main()
