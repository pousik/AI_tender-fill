from pathlib import Path

import services.drawing_reader as drawing_reader
from services.drawing_reader import ReferenceDrawingReader


def test_reference_dimension_figures() -> None:
    root = Path(__file__).resolve().parents[1] / "data_tenders"
    reader = ReferenceDrawingReader(root)
    path = reader._find_reference_file()
    assert path is not None
    figures = reader._extract_dimension_figures(path)
    assert [item["figure_id"] for item in figures] == ["A.1", "A.2", "A.3", "A.4", "A.5", "A.6"]


def test_reference_dimension_vision_covers_all_figures(monkeypatch, tmp_path) -> None:
    root = Path(__file__).resolve().parents[1] / "data_tenders"
    reader = ReferenceDrawingReader(root)
    reader.cache_path = tmp_path / ".drawing_vision_cache.json"

    calls = []

    def fake_vision(client, image_bytes, prompt, schema, *, mime_type="image/png"):
        calls.append(prompt.splitlines()[2])
        return {
            "figure_id": "",
            "caption": "",
            "model_mentions": [],
            "dimensions": [],
            "masses": [],
            "notes": [],
        }

    monkeypatch.setattr(drawing_reader, "ask_vision_json_schema", fake_vision)
    result = reader.read_dimension_drawings(client=object())
    assert result["status"] == "ok"
    assert [x["figure_id"] for x in result["figures"]] == ["A.1", "A.2", "A.3", "A.4", "A.5", "A.6"]
    assert len(calls) == 6
