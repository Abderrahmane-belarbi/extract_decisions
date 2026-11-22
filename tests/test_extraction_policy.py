import unittest
from pathlib import Path
from unittest.mock import patch

import extract_decisions as app
from extraction_policy import mask_uncertain_fields
from evaluate_extraction import score_predictions


class PolicyTests(unittest.TestCase):
    def test_request_timeout_is_configurable_and_forwarded(self):
        self.assertEqual(app.arguments([]).request_timeout, 3600)
        self.assertEqual(app.arguments(["--request-timeout", "7200"]).request_timeout, 7200)
        with patch.object(app.urllib.request, "urlopen", side_effect=TimeoutError("timed out")) as request:
            with self.assertRaisesRegex(RuntimeError, "within 7200 seconds"):
                app.ollama_call("model", "image", app.SCHEMA, "prompt", 12288, 1600, 7200)
        self.assertEqual(request.call_args.kwargs["timeout"], 7200)

    def test_default_timeout_explains_cpu_failure(self):
        with patch.object(app.urllib.request, "urlopen", side_effect=TimeoutError("timed out")):
            with self.assertRaisesRegex(RuntimeError, "Increase --request-timeout"):
                app.ollama_call("model", "image", app.SCHEMA, "prompt", 12288, 1600)

    def test_ambiguous_name_split_preserves_full_name_and_candidates(self):
        fields = {"surname": "الأول", "given_name": "الثاني", "full_name_as_written": "الأول الثاني"}
        values, uncertain, candidates = mask_uncertain_fields(fields, ["surname", "not_a_field", "surname"])
        self.assertEqual(values, {"surname": "", "given_name": "", "full_name_as_written": "الأول الثاني"})
        self.assertEqual(uncertain, ["given_name", "surname"])
        self.assertEqual(candidates, {"surname": "الأول", "given_name": "الثاني"})

    def test_missing_grade_and_effective_date_are_not_inferred(self):
        result = app.clean_model_result({"decision_date": "04 JAN. 2026", "index_number": "98+653",
                                         "full_name_as_written": "اسم كامل", "category": "ب"})
        self.assertEqual(result["fields"]["grade"], "")
        self.assertEqual(result["fields"]["decision_effective_date"], "")
        self.assertEqual(result["fields"]["surname"], "")
        self.assertEqual(result["fields"]["category"], "ب")

    def test_uncertain_values_are_not_extraction_targets(self):
        result = app.clean_model_result({"visa_number": "123", "uncertain_fields": ["visa_number"]})
        self.assertEqual(result["fields"]["visa_number"], "")
        self.assertEqual(result["review_candidates"], {"visa_number": "123"})

    def test_cache_tracks_prompt_schema_and_actual_context(self):
        job = app.PageJob(Path("scan.jpg"), 1, 1, "sourcehash")
        def key(ctx=4096):
            return app.cache_key(job, "model", 200, 1700, False, 0, ctx)
        original = key()
        self.assertNotEqual(original, key(8192))
        with patch.object(app, "SYSTEM_PROMPT", app.SYSTEM_PROMPT + "changed"):
            self.assertNotEqual(original, key())
        with patch.object(app, "CROP_PROMPT", app.CROP_PROMPT + "changed"):
            self.assertNotEqual(original, key())
        with patch.object(app, "SCHEMA", {**app.SCHEMA, "description": "changed"}):
            self.assertNotEqual(original, key())
        self.assertEqual(original, key())

    def test_category_letters_and_invalid_categories(self):
        record = app.DecisionRecord(Path("scan.jpg"), [1], {key: "" for key in app.FIELDS})
        for value in ["01", "16", "١٤", "أ", "ب"]:
            record.fields["category"] = value
            self.assertFalse(any("الصنف" in warning for warning in app.validate_record(record)), value)
        for value in ["0", "17", "رئيس", "A"]:
            record.fields["category"] = value
            self.assertTrue(any("الصنف" in warning for warning in app.validate_record(record)), value)

    def test_crop_disagreement_is_blank_and_preserved_for_review(self):
        from PIL import Image
        page = {"readable": True, "upright": True, "page_role": "start", "decision_number": "15",
                "decision_date": "04 JAN. 2026", "surname": "لقب", "given_name": "اسم",
                "full_name_as_written": "لقب اسم", "uncertain_fields": []}
        with patch.object(app, "ollama_call", side_effect=[page, {"decision_number": "16"}, {}, {}]):
            result, _, _, _ = app.model_page(Image.new("RGB", (600, 800)), "model", 800, 12288, 0, True)
        self.assertEqual(result["fields"]["decision_number"], "")
        self.assertEqual(result["review_candidates"]["decision_number"], "15")
        self.assertIn("decision_number", result["uncertain_fields"])

    def test_evaluation_requires_verified_truth_and_counts_false_fills(self):
        expected = {key: "" for key in app.FIELDS}
        expected.update({"decision_number": "15", "uncertain_fields": ["visa_number"]})
        reference = {"record_id": "r:1", "output": expected, "human_verified": False}
        prediction = {"record_id": "r:1", "output": {"decision_number": "", "grade": "3", "visa_number": "9"}}
        with self.assertRaisesRegex(ValueError, "No human-verified"):
            score_predictions([reference], [prediction])
        reference["human_verified"] = True
        result = score_predictions([reference], [prediction])
        self.assertEqual(result["counts"]["omissions"], 1)
        self.assertEqual(result["counts"]["unsupported_fills"], 1)
        self.assertEqual(result["fields"]["visa_number"]["compared"], 0)


if __name__ == "__main__":
    unittest.main()
