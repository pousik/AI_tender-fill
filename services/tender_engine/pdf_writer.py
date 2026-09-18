from __future__ import annotations

import os
from pathlib import Path

try:
    import pymupdf
except ImportError:  # pragma: no cover
    import fitz as pymupdf

from .pdf_reader import PdfRowLayout
from .text import strip_marks


class PdfTenderWriter:
    """Безопасно записывает значение только внутрь соответствующей ячейки."""

    def __init__(self) -> None:
        self._font_name: str | None = None
        self._font_file: str | None = None

    def write(self, path: str | Path, layouts: list[PdfRowLayout]) -> None:
        doc = pymupdf.open(str(path))
        try:
            for layout in layouts:
                row = layout.row
                if not row.proposed_value:
                    continue
                page = doc[layout.page_index]
                rect = layout.target_rect
                if rect.width < 12 or rect.height < 6:
                    print(f"[PDF][WRITE] skip tiny rect {row.row_id}: {rect}")
                    continue

                if self._needs_background(row):
                    page.draw_rect(rect, color=None, fill=(1, 1, 0), fill_opacity=0.28, overlay=True)
                elif row.current_value:
                    page.draw_rect(rect, color=None, fill=(1, 1, 1), overlay=True)

                text = strip_marks(row.proposed_value) + (row.mark or "")
                self._insert_text_fitted(page, rect, text)
            doc.saveIncr()
        finally:
            doc.close()

    @staticmethod
    def _needs_background(row) -> bool:
        return bool(row.mark) or row.source.startswith("AI")

    def _insert_text_fitted(self, page, rect, text: str) -> None:
        fontname, fontfile = self._font()
        base = min(9.0, max(6.0, rect.height * 0.62))
        for size in (base, 8.0, 7.0, 6.0, 5.0, 4.5):
            if size > base + 0.01:
                continue
            kwargs = {"fontsize": size, "align": 0, "overlay": True}
            if fontfile:
                kwargs["fontname"] = fontname
                kwargs["fontfile"] = fontfile
            else:
                kwargs["fontname"] = "helv"
            try:
                rc = page.insert_textbox(rect, text, **kwargs)
            except TypeError:
                rc = page.insert_textbox(rect, text, fontsize=size, fontname=fontname if fontfile else "helv", align=0, overlay=True)
            if rc >= -0.1:
                return

        kwargs = {"fontsize": 4.2, "align": 0, "overlay": True, "fontname": fontname if fontfile else "helv"}
        if fontfile:
            kwargs["fontfile"] = fontfile
        page.insert_textbox(rect, text, **kwargs)

    def _font(self) -> tuple[str, str | None]:
        if self._font_name:
            return self._font_name, self._font_file
        candidates = []
        env = os.getenv("TENDER_PDF_FONT", "").strip()
        if env:
            candidates.append(env)
        if os.name == "nt":
            windir = os.environ.get("WINDIR", r"C:\Windows")
            candidates.extend([
                str(Path(windir) / "Fonts" / "arial.ttf"),
                str(Path(windir) / "Fonts" / "calibri.ttf"),
                str(Path(windir) / "Fonts" / "times.ttf"),
            ])
        candidates.extend([
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        ])
        for path in candidates:
            if Path(path).exists():
                self._font_name = "TenderCyr"
                self._font_file = path
                return self._font_name, self._font_file
        self._font_name = "helv"
        self._font_file = None
        print("[PDF][FONT] Cyrillic font not found; fallback to Helvetica")
        return self._font_name, self._font_file
