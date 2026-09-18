import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services.gigachat_client import parse_json_response


class GigaChatJsonRepairTests(unittest.TestCase):
    def test_repairs_missing_comma_between_properties(self):
        raw = '{"candidates":[{"id":"item_1","value":"110"\n"unit":"А"}]}'
        data = parse_json_response(raw)
        self.assertEqual(data["candidates"][0]["value"], "110")
        self.assertEqual(data["candidates"][0]["unit"], "А")

    def test_repairs_nested_object_and_trailing_commas(self):
        raw = '{"a":{"x":1\n"y":2},"b":[1,2,],}'
        data = parse_json_response(raw)
        self.assertEqual(data, {"a": {"x": 1, "y": 2}, "b": [1, 2]})

    def test_recovers_complete_candidates_from_truncated_json(self):
        raw = '{"candidates":[{"id":"item_1","model":"ТРГ-УЭТМ-110","parameter":"Масса","value":"1000","unit":"кг","source_file":"a.docx"},{"id":"item_2","model":"ТРГ-УЭТМ-110","parameter":"Высота","value":"2123","unit":"мм","source_file":"a.docx"},{"id":"item_3","model":"ТРГ-УЭТМ-110","parameter":"Ширина","value":"819","unit":"мм","source_file":"a.docx"'
        data = parse_json_response(raw, partial_array_key="candidates")
        self.assertEqual(len(data["candidates"]), 2)
        self.assertEqual(data["candidates"][1]["value"], "2123")

    def test_preserves_commas_inside_string_values(self):
        raw = '{"value":"110, 220, 330","unit":"кВ"}'
        data = parse_json_response(raw)
        self.assertEqual(data["value"], "110, 220, 330")


if __name__ == "__main__":
    unittest.main()
