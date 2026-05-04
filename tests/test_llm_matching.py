import json
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import image_ocr_match_report as report


class LLMMatchingTests(unittest.TestCase):
    def test_noise_line_rejects_common_garbage(self):
        self.assertTrue(report.is_noise_line("2K OK 2K 2K OK OK", 70))

    def test_parse_llm_match_response_rejects_malformed_json(self):
        with self.assertRaises(ValueError):
            report.parse_llm_match_response('[{"source_string": "Date"')

    def test_validate_llm_matches_rejects_hallucinated_text(self):
        raw = [
            report.LLMGuiMatch("Date and time", "Datum es ido", 0.95, "matched", ""),
            report.LLMGuiMatch("Made up source", "Datum es ido", 0.95, "matched", ""),
            report.LLMGuiMatch("Auto", "Invented target", 0.95, "matched", ""),
        ]
        validated, rejected = report.validate_llm_matches(
            raw,
            "Date and time\nAuto",
            "Datum es ido",
            ["Date and time", "Auto"],
        )
        pairs = {(m.source_string, m.target_string, m.status) for m in validated}
        self.assertIn(("Date and time", "Datum es ido", "matched"), pairs)
        self.assertIn(("Auto", "", "uncertain"), pairs)
        self.assertEqual(len(rejected), 2)

    def test_make_report_without_llm_uses_positional_alignment(self):
        source_file = Path("enis01ct001a.eps")
        target_file = Path("huis01ct001a.eps")
        match = report.MatchResult(
            source_name=source_file.name,
            target_lang="",
            target_name=target_file.name,
            method="exact_normalized",
            confidence=1.0,
            alternatives=[],
            status="OK",
            notes="",
        )

        def fake_extract(path, *_args, **_kwargs):
            if Path(path).name.startswith("en"):
                return ["Date and time", "Auto"], None
            return ["Datum es ido", "Automatikus"], None

        with patch.object(report, "list_images", return_value=[target_file]), patch.object(
            report, "match_source_to_target", return_value=[match]
        ), patch.object(report, "extract_text_lines", side_effect=fake_extract), patch.object(
            report, "match_gui_strings_with_llm"
        ) as llm_mock:
            df = report.make_report(
                source_files=[source_file],
                lang_code="hu",
                target_dir=Path("."),
                tesseract_lang_target="hun",
                tesseract_lang_source="eng",
                tesseract_cmd="tesseract",
                ghostscript_cmd=None,
                eps_dpi=600,
                use_temp_local=False,
                fuzzy_threshold=0.72,
                progress_every=0,
                allow_fuzzy=False,
                reject_ambiguous_fuzzy=True,
                strict_text_only=False,
                use_llm_matching=False,
                openai_model="gpt-4.1-mini",
                openai_api_key=None,
                llm_timeout_sec=45,
                llm_max_retries=2,
            )

        llm_mock.assert_not_called()
        self.assertEqual(df["Source OCR text"].tolist(), ["Date and time", "Auto"])
        self.assertEqual(df["Target OCR text"].tolist(), ["Datum es ido", "Automatikus"])
        self.assertEqual(df["Match method"].tolist(), ["exact_normalized", "exact_normalized"])


if __name__ == "__main__":
    unittest.main()
