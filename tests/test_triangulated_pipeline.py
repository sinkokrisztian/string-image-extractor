import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import image_ocr_match_report as report


class TriangulatedPipelineTests(unittest.TestCase):
    def test_ocr_candidate_package_keeps_pass_metadata(self):
        line = report.OCRLine(
            pair_id="0001",
            side="source",
            image_filename="enis01ct009b.eps",
            ocr_pass="psm6_gray",
            preprocessing="gray",
            psm="6",
            line_no=1,
            raw_line="Volume",
            cleaned_line="Volume",
            avg_tesseract_confidence=91.0,
            bbox_left=1,
            bbox_top=2,
            bbox_right=30,
            bbox_bottom=10,
            kept_or_filtered="kept",
            selected_pass=True,
            line_quality_score=3.0,
        )
        evidence = report.ClassicOCRResult([], [line], ["Volume"], "Volume", [])
        package = report.ocr_candidates_package("enis01ct009b.eps", "en", "English", "source", evidence)
        self.assertEqual(package["ocr_passes"][0]["pass_name"], "psm6_gray")
        self.assertTrue(package["ocr_passes"][0]["selected_pass"])
        self.assertEqual(package["ocr_passes"][0]["lines"][0]["cleaned_line"], "Volume")

    def test_triangulation_selects_llm_ai_agreement(self):
        classic = report.ClassicOCRResult([], [], ["Hanger6"], "Hanger6", [])
        llm = report.OCRNormalizedResult(
            image_filename="huis01ct009b.eps",
            language_code="hu",
            language_name="Hungarian",
            side="target",
            source_engine="llm_ocr_normalization",
            model="gpt-4.1-mini",
            prompt_version=report.LLM_OCR_NORMALIZATION_PROMPT_VERSION,
            ui_objects=[
                report.NormalizedGUIObject(
                    object_id="llm_001",
                    normalized_text="Hangerő",
                    normalization_confidence=0.92,
                    ocr_evidence_confidence=0.80,
                    reading_order=1,
                )
            ],
        )
        ai = report.OCRNormalizedResult(
            image_filename="huis01ct009b.eps",
            language_code="hu",
            language_name="Hungarian",
            side="target",
            source_engine="ai_image_ocr",
            model="gpt-4.1-mini",
            prompt_version=report.AI_IMAGE_OCR_PROMPT_VERSION,
            ui_objects=[
                report.NormalizedGUIObject(
                    object_id="ai_001",
                    normalized_text="Hangerő",
                    normalization_confidence=0.95,
                    ocr_evidence_confidence=0.95,
                    reading_order=1,
                    correction_type="vision_read",
                )
            ],
        )
        rows = report.triangulate_side("0001", "target", "huis01ct009b.eps", classic, llm, ai)
        self.assertEqual(rows[0].agreement_status, "llm_ai_agree")
        self.assertEqual(rows[0].selected_final_text, "Hangerő")
        self.assertEqual(rows[0].selected_source, "merged")

    def test_multi_sheet_workbook_contains_expected_sheets(self):
        pair = report.ImagePair(
            pair_id="0001",
            source_image_filename="enis01ct009b.eps",
            source_image_path="src",
            target_language_code="hu",
            target_language_name="Hungarian",
            target_image_filename="huis01ct009b.eps",
            target_image_path="trg",
            image_match_method="exact_normalized",
            image_match_confidence=1.0,
            image_match_status="OK",
        )
        pipe_report = report.PipelineReport(
            image_pairs=[pair],
            classic_lines=[],
            classic_tokens=[],
            llm_results=[],
            ai_results=[],
            triangulation_rows=[],
            final_matches=[],
            summary={"ocr_engine": "classic"},
        )
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "report.xlsx"
            report.write_multi_sheet_report(pipe_report, out)
            wb = load_workbook(out)
            self.assertIn("Image pairs", wb.sheetnames)
            self.assertIn("Raw OCR - Classic Lines", wb.sheetnames)
            self.assertIn("Final Matches", wb.sheetnames)
            self.assertIn("Summary", wb.sheetnames)


if __name__ == "__main__":
    unittest.main()
