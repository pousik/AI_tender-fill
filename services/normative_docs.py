"""Загрузка нормативного контекста из РЭ/ТУ (1БП.769.001*).

Используется как дополнительный контекст для GigaChat:
текущий документ + таблицы/характеристики из официальных
руководств по эксплуатации и технических условий.
"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_CACHE: str | None = None

# Максимальный размер текста, который безопасно класть в один промпт
# (вместе с текущим документом и БЗ).
MAX_CHARS = 12000


def load_normative_context(max_chars: int = MAX_CHARS) -> str:
    """Возвращает компактный текст характеристик из РЭ/ТУ.

    Источники (в порядке приоритета файла):
    - normative_context.txt  (готовый текст)
    - normative_context.json (собирается в текст)
    """
    global _CACHE
    if _CACHE is not None:
        return _CACHE

    txt_path = ROOT / "normative_context.txt"
    if txt_path.exists():
        text = txt_path.read_text(encoding="utf-8").strip()
        if len(text) > max_chars:
            text = text[:max_chars] + "\n... [сокращено, полный текст в normative_context.txt]"
        _CACHE = text
        return _CACHE

    json_path = ROOT / "normative_context.json"
    if not json_path.exists():
        _CACHE = (
            "Нормативные документы не загружены. "
            "Ожидаются файлы: 1БП.769.001 РЭ, 1БП.769.001-01 РЭ, 1БП 769 001ТУ."
        )
        return _CACHE

    import json

    data = json.loads(json_path.read_text(encoding="utf-8"))
    parts: list[str] = []
    for doc in data:
        parts.append(f"=== {doc.get('source', '')} ===")
        if doc.get("title"):
            parts.append(str(doc["title"]))
        for p in doc.get("key_paragraphs", [])[:8]:
            parts.append(str(p))
        for t in doc.get("tables", [])[:6]:
            parts.append(f"[Таблица {t.get('index', '')}]")
            for row in t.get("rows", [])[:18]:
                line = " | ".join(c for c in row if str(c).strip())
                if line.strip():
                    parts.append(line)
            parts.append("")
        parts.append("")

    text = "\n".join(parts).strip()
    if len(text) > max_chars:
        text = text[:max_chars] + "\n... [сокращено]"
    _CACHE = text
    return _CACHE


def normative_prompt_block() -> str:
    """Готовый блок для вставки в системный промпт GigaChat."""
    ctx = load_normative_context()
    return f"""
=========================================================
НОРМАТИВНЫЕ ДОКУМЕНТЫ (РЭ / ТУ) — ОБЯЗАТЕЛЬНЫЙ КОНТЕКСТ
=========================================================
Первоисточники:
- 1БП.769.001 РЭ от января 2026 г.
- 1БП.769.001-01 РЭ от января 2026 г.
- 1БП 769 001ТУ от января 2026 г.

Порядок заполнения (строго):
1) Если есть подходящий шаблон — использовать его.
2) Если шаблона нет — в первую очередь данные из БД / KnowledgeEntry / профили.
3) Только оставшиеся пустые поля — нейросеть, опираясь на:
   а) текущий тендерный документ;
   б) характеристики и таблицы из указанных РЭ и ТУ (ниже).

Не противоречь значениям из БД. Если БД дала значение — оставь его.
Из РЭ/ТУ бери только подтверждённые технические характеристики
(напряжение, токи, классы точности, климат, габариты и т.п.).

--- ВЫДЕРЖКИ ИЗ РЭ/ТУ ---
{ctx}
--- КОНЕЦ ВЫДЕРЖЕК РЭ/ТУ ---
""".strip()
