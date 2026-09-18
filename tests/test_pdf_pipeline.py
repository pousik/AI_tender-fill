from __future__ import annotations

import io
import tempfile
import unittest
from pathlib import Path

import pymupdf
from PIL import Image

from services.tender_engine.pdf_pipeline import PdfTenderReader, PdfTenderWriter
from services.tender_engine.schema import TenderRow


ROOT = Path(__file__).resolve().parents[1]


def font_path() -> str | None:
    candidates = [
        Path(r"C:\Windows\Fonts\arial.ttf"),
        Path(r"C:\Windows\Fonts\calibri.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"),
    ]
    for path in candidates:
        if path.exists():
            return str(path)
    return None


def build_pdf(path: Path, *, raster: bool = False) -> None:
    src = pymupdf.open()
    page = src.new_page(width=595, height=842)
    font = font_path()
    if font:
        page.insert_font(fontname="TestCyr", fontfile=font)
        fn = "TestCyr"
        ff = font
    else:
        fn = "helv"
        ff = None

    x = [80, 120, 365, 430, 550]
    y = [200, 240, 285, 330]
    for xx in x:
        page.draw_line((xx, y[0]), (xx, y[-1]), color=(0, 0, 0), width=0.6)
    for yy in y:
        page.draw_line((x[0], yy), (x[-1], yy), color=(0, 0, 0), width=0.6)

    def put(rect, text, size=8):
        kwargs = {"fontsize": size, "fontname": fn, "align": 0, "overlay": True}
        if ff:
            kwargs["fontfile"] = ff
        page.insert_textbox(rect, text, **kwargs)

    put((150, 205, 340, 235), "Технические требования к оборудованию")
    put((450, 205, 540, 235), "Предлагаемое")
    put((85, 245, 115, 280), "1.1")
    put((122, 245, 350, 280), "Изготовитель")
    put((370, 245, 420, 280), "*")
    put((85, 290, 115, 325), "1.2")
    put((122, 290, 350, 325), "Заводской тип (марка)")
    put((370, 290, 420, 325), "*")
    put((85, 335, 115, 360), "2.")
    put((122, 335, 350, 360), "Раздел")
    if raster:
        pix = page.get_pixmap(dpi=170, alpha=False)
        png = pix.tobytes("png")
        img = Image.open(io.BytesIO(png)).convert("RGB")
        out = pymupdf.open()
        rp = out.new_page(width=595, height=842)
        rp.insert_image(rp.rect, stream=io.BytesIO(png).read())
        src.close()
        out.save(str(path))
        out.close()
        return
    src.save(str(path))
    src.close()


class PdfPipelineTests(unittest.TestCase):
    def test_reader_uses_real_columns_and_row_bands(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "table.pdf"
            build_pdf(path)
            layouts, _ = PdfTenderReader().read(path)
            rows = [x.row for x in layouts]
            target = next(r for r in rows if r.number.rstrip(".") == "1.1")
            self.assertEqual(target.parameter, "Изготовитель")
            self.assertEqual(target.requirement, "*")
            self.assertEqual(target.current_value, "")
            self.assertEqual(target.field_key, "manufacturer")
            self.assertTrue(any(r.number.rstrip(".") == "1.2" for r in rows))

    def test_writer_highlights_ai_value(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "table.pdf"
            build_pdf(path)
            layouts, _ = PdfTenderReader().read(path)
            item = next(x for x in layouts if x.row.number.rstrip(".") == "1.1")
            item.row.proposed_value = "Белый"
            item.row.source = "AI_VISION"
            item.row.mark = "**"
            PdfTenderWriter().write(path, layouts)
            doc = pymupdf.open(path)
            drawings = [d for d in doc[0].get_drawings() if d.get("fill")]
            doc.close()
            self.assertTrue(drawings)

    def test_writer_preserves_cyrillic(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "table.pdf"
            out = Path(td) / "filled.pdf"
            build_pdf(path)
            layouts, _ = PdfTenderReader().read(path)
            row = next(x.row for x in layouts if x.row.number.rstrip(".") == "1.1")
            row.proposed_value = 'ООО "ЭЛЬМАШ (УЭТМ)"'
            row.source = "DATA_TENDERS"
            PdfTenderWriter().write(path, layouts)
            # Writer edits in place; make a fresh extraction from the resulting PDF.
            doc = pymupdf.open(path)
            text = "\n".join(page.get_text("text") for page in doc)
            doc.close()
            self.assertIn("ЭЛЬМАШ", text)
            self.assertIn("УЭТМ", text)

    def test_scanned_page_falls_back_to_ocr_grid(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "scan.pdf"
            build_pdf(path, raster=True)
            layouts, _ = PdfTenderReader().read(path)
            self.assertTrue(layouts)
            rows = [x.row for x in layouts]
            self.assertTrue(any("Изготовитель" in r.parameter for r in rows))


if __name__ == "__main__":
    unittest.main()
