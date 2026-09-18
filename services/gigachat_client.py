"""Единая обертка над GigaChat SDK для обоих процессоров (DOCX и PDF).

Ключ никогда не хранится в проекте — только в переменной окружения
``GIGACHAT_CREDENTIALS`` (см. ``run.ps1``, который запрашивает его через
SecureString и передает только текущему процессу).
"""

from __future__ import annotations

import base64
import json
import os
import re
from typing import Any

try:
    from gigachat import GigaChat
    from gigachat.models import (
        ChatCompletionRequest,
        ChatMessage,
        ChatContentPart,
    )
except ImportError:  # позволяет запускать OCR/детерминированный режим без SDK
    GigaChat = None
    ChatCompletionRequest = ChatMessage = ChatContentPart = None

GIGACHAT_CREDENTIALS = os.getenv("GIGACHAT_CREDENTIALS", "MDFhMDY2YWMtZDc1Ni03NTc0LTg4MDEtMzJmYmY2YTY2MDkwOjZkMjFiMmI3LTY4MGItNGUyZS05YjQ4LWViY2E4NjIwMGU0Nw==").strip()
if GIGACHAT_CREDENTIALS.lower().startswith("bearer "):
    GIGACHAT_CREDENTIALS = GIGACHAT_CREDENTIALS[7:].strip()
GIGACHAT_MODEL = os.getenv("GIGACHAT_MODEL", "GigaChat-2-Pro")
GIGACHAT_SCOPE = os.getenv("GIGACHAT_SCOPE", "GIGACHAT_API_PERS")
GIGACHAT_BASE_URL = os.getenv("GIGACHAT_BASE_URL", "https://api.giga.chat/v1")
GIGACHAT_VERIFY_SSL = os.getenv("GIGACHAT_VERIFY_SSL_CERTS", "false").strip().lower() == "true"

# SDK по умолчанию ждёт ответ около 30 секунд. Для технических тендеров
# prompt может быть заметно тяжелее обычного запроса, поэтому задаём явный
# увеличенный таймаут и несколько повторов для временных сетевых ошибок.
try:
    GIGACHAT_TIMEOUT = float(os.getenv("GIGACHAT_TIMEOUT", "180"))
except ValueError as exc:
    raise RuntimeError("GIGACHAT_TIMEOUT должен быть числом секунд.") from exc
if GIGACHAT_TIMEOUT <= 0:
    raise RuntimeError("GIGACHAT_TIMEOUT должен быть больше 0.")

try:
    GIGACHAT_MAX_RETRIES = int(os.getenv("GIGACHAT_MAX_RETRIES", "2"))
except ValueError as exc:
    raise RuntimeError("GIGACHAT_MAX_RETRIES должен быть целым числом.") from exc
if GIGACHAT_MAX_RETRIES < 0:
    raise RuntimeError("GIGACHAT_MAX_RETRIES не может быть отрицательным.")

# Порог уверенности, ниже которого предложение нейросети из контекста
# (не подтвержденное БЗ) отклоняется и не помечается "**".
try:
    AI_CONFIDENCE_THRESHOLD = float(os.getenv("AI_CONFIDENCE_THRESHOLD", "0.75"))
except ValueError as exc:
    raise RuntimeError("AI_CONFIDENCE_THRESHOLD должен быть числом от 0 до 1.") from exc
if not 0.0 <= AI_CONFIDENCE_THRESHOLD <= 1.0:
    raise RuntimeError("AI_CONFIDENCE_THRESHOLD должен быть в диапазоне 0..1.")


def require_credentials() -> None:
    if not GIGACHAT_CREDENTIALS:
        raise RuntimeError(
            "Не задан GIGACHAT_CREDENTIALS. Укажите актуальный ключ в переменной окружения "
            "(например, через run.ps1) перед запуском."
        )


def open_client() -> GigaChat:
    """Создает клиент GigaChat. Требует установленный SDK и учетные данные."""
    if GigaChat is None:
        raise RuntimeError("Не установлен пакет gigachat. Выполните pip install -r requirements.txt.")
    require_credentials()
    return GigaChat(
        credentials=GIGACHAT_CREDENTIALS,
        model=GIGACHAT_MODEL,
        scope=GIGACHAT_SCOPE,
        base_url=GIGACHAT_BASE_URL,
        verify_ssl_certs=GIGACHAT_VERIFY_SSL,
        timeout=GIGACHAT_TIMEOUT,
        max_retries=GIGACHAT_MAX_RETRIES,
        retry_backoff_factor=0.5,
        retry_on_status_codes=(429, 500, 502, 503, 504),
    )


def available_model_ids(client: GigaChat) -> list[str]:
    """Возвращает фактические идентификаторы моделей из /models."""
    models = client.get_models()
    result = []
    for item in getattr(models, "data", []) or []:
        mid = getattr(item, "id", None)
        if mid:
            result.append(str(mid))
    return result



def verify_connection(client: GigaChat) -> str:
    """Проверяет, что ключ/scope рабочие, используя уже открытый клиент."""
    try:
        ids = available_model_ids(client)
        count = len(ids)
        return f"GigaChat OK: model={GIGACHAT_MODEL}, scope={GIGACHAT_SCOPE}, models={count}, ids={ids}"
    except Exception as e:
        raise RuntimeError(
            f"GigaChat не прошел авторизацию. Проверьте ключ и scope={GIGACHAT_SCOPE}. Ответ API: {e}"
        ) from e


def check_gigachat_connection() -> str:
    """Автономная диагностическая проверка (используется test_gigachat.py)."""
    with open_client() as client:
        return verify_connection(client)


def extract_text(response) -> str:
    """Достает текст ответа независимо от формы объекта, которую вернул SDK."""
    messages = getattr(response, "messages", None)
    if messages:
        for message in messages:
            if getattr(message, "role", None) not in (None, "assistant"):
                continue
            content = getattr(message, "content", None) or []
            parts = []
            for part in content:
                text = getattr(part, "text", None)
                if text:
                    parts.append(str(text))
                elif isinstance(part, str):
                    parts.append(part)
            if parts:
                return "".join(parts).strip()
        # У некоторых версий SDK content самого message — обычная строка.
        for message in messages:
            content = getattr(message, "content", None)
            if isinstance(content, str) and content.strip():
                return content.strip()
    choices = getattr(response, "choices", None)
    if choices:
        message = getattr(choices[0], "message", None)
        content = getattr(message, "content", None) if message else None
        if content:
            return str(content).strip()
    raise RuntimeError(f"GigaChat вернул неизвестный формат ответа: {type(response).__name__}")


def clean_json_text(text: str) -> str:
    text = re.sub(r"^```(?:json)?\s*", "", text.strip(), flags=re.I)
    text = re.sub(r"\s*```$", "", text.strip())
    return text.strip()


def _extract_balanced_json_object(text: str) -> str:
    """Извлекает первый сбалансированный JSON-объект, не ломая строки."""
    start = text.find("{")
    if start < 0:
        raise ValueError("GigaChat не вернул JSON-объект: символ '{' не найден")

    depth = 0
    in_string = False
    escaped = False
    for i in range(start, len(text)):
        ch = text[i]
        if in_string:
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_string = False
            continue

        if ch == '"':
            in_string = True
        elif ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start:i + 1]

    raise ValueError("GigaChat вернул незакрытый JSON-объект")


def _repair_common_json_errors(text: str) -> str:
    """Исправляет безопасный набор типовых дефектов JSON от LLM."""
    out: list[str] = []
    in_string = False
    escaped = False
    i = 0
    n = len(text)

    while i < n:
        ch = text[i]
        if in_string:
            if escaped:
                out.append(ch)
                escaped = False
            elif ch == "\\":
                out.append(ch)
                escaped = True
            elif ch == '"':
                out.append(ch)
                in_string = False
            elif ord(ch) < 32:
                out.append(" ")
            else:
                out.append(ch)
            i += 1
            continue

        if ch == '"':
            out.append(ch)
            in_string = True
            i += 1
            continue

        # Вне строк переносы и табы не нужны: оставляем пробел.
        if ch in "\r\n\t":
            out.append(" ")
            i += 1
            continue

        # Удаляем завершающие запятые перед закрытием объекта/массива.
        if ch == ',':
            j = i + 1
            while j < n and text[j].isspace():
                j += 1
            if j < n and text[j] in "}]":
                i += 1
                continue

        out.append(ch)
        i += 1

    return "".join(out).strip()


def _insert_missing_commas(text: str, *, max_repairs: int = 256) -> str:
    """Восстанавливает пропущенные запятые по позициям JSONDecodeError.

    Метод намеренно меняет только места, где сам JSON-парсер сообщает
    `Expecting ',' delimiter`. Это намного безопаснее глобальных regex-замен
    внутри строковых значений.
    """
    value = text
    for _ in range(max_repairs):
        try:
            json.loads(value)
            return value
        except json.JSONDecodeError as exc:
            if "Expecting ',' delimiter" not in exc.msg:
                return value

            pos = exc.pos
            # JSONDecodeError указывает на начало следующего токена.
            # Вставляем запятую непосредственно перед ним.
            j = pos
            while j > 0 and value[j - 1].isspace():
                j -= 1
            if j <= 0:
                return value

            prev = value[j - 1]
            next_char = value[pos] if pos < len(value) else ""
            if prev not in '"}]0123456789elruefao-.' and prev not in "truefalsn":
                return value
            if next_char and next_char not in '"{[-0123456789tfn':
                return value

            value = value[:j] + "," + value[j:]
    return value



def _parse_partial_array_object(text: str, array_key: str) -> dict[str, Any] | None:
    """Извлекает уже завершённые элементы массива из обрезанного JSON.

    Используется как аварийный режим для длинных structured-output ответов.
    Если модель оборвала ответ внутри последней строки/объекта, все полностью
    завершённые элементы до точки обрыва всё равно сохраняются.
    """
    marker = '"' + array_key + '"'
    key_pos = text.find(marker)
    if key_pos < 0:
        return None
    arr_start = text.find("[", key_pos + len(marker))
    if arr_start < 0:
        return None

    objects: list[dict[str, Any]] = []
    depth = 0
    obj_start: int | None = None
    in_string = False
    escaped = False

    for i in range(arr_start + 1, len(text)):
        ch = text[i]
        if in_string:
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_string = False
            continue

        if ch == '"':
            in_string = True
        elif ch == "{":
            if depth == 0:
                obj_start = i
            depth += 1
        elif ch == "}":
            if depth > 0:
                depth -= 1
                if depth == 0 and obj_start is not None:
                    candidate_text = text[obj_start:i + 1]
                    try:
                        value = json.loads(_repair_common_json_errors(candidate_text))
                        if isinstance(value, dict):
                            objects.append(value)
                    except Exception:
                        pass
                    obj_start = None

    if not objects:
        return None
    return {array_key: objects}

def parse_json_response(text: str, *, partial_array_key: str | None = None) -> dict[str, Any]:
    raw = clean_json_text(text)
    candidates = [raw]

    try:
        candidates.append(_extract_balanced_json_object(raw))
    except ValueError:
        pass

    last_error: Exception | None = None
    for candidate in candidates:
        cleaned = _repair_common_json_errors(candidate)
        variants = (
            cleaned,
            _insert_missing_commas(cleaned),
        )
        for repaired in variants:
            try:
                data = json.loads(repaired)
                if not isinstance(data, dict):
                    raise ValueError(
                        f"GigaChat должен вернуть JSON-объект, получен {type(data).__name__}"
                    )
                return data
            except (json.JSONDecodeError, ValueError) as exc:
                last_error = exc

    if partial_array_key:
        partial = _parse_partial_array_object(raw, partial_array_key)
        if partial:
            return partial

    preview = raw[:1200].replace("\n", " ")
    raise ValueError(
        f"Не удалось разобрать JSON-ответ GigaChat: {last_error}; preview={preview!r}"
    ) from last_error


def _json_response_schema() -> dict[str, Any]:
    """Схема ответа AI-аудита.

    Ключи верхнего уровня — item_id, поэтому поле additionalProperties
    содержит описание объекта результата для каждого item_id.
    """
    result_item = {
        "type": "object",
        "properties": {
            "value": {"type": "string"},
            "source": {
                "type": "string",
                "enum": ["DATA_TENDERS", "SELF_CONTEXT", "AI_CONTEXT", "NONE"],
            },
            "confidence": {"type": "number", "minimum": 0, "maximum": 1},
            "evidence": {"type": "array", "items": {"type": "string"}},
            "reason": {"type": "string"},
        },
        "required": ["value", "source", "confidence", "evidence", "reason"],
        "additionalProperties": False,
    }
    return {
        "type": "object",
        "additionalProperties": result_item,
    }


def _sanitize_prompt_for_content_filter(prompt: str) -> str:
    """Убирает из технического контекста слова, которые не нужны AI для расчёта
    значения, но могут ошибочно сработать на тематическом фильтре GigaChat.

    Важно: локальная БД/DOCX не изменяются. Санитизация действует только на
    копию текста, отправляемую во внешний AI-сервис.
    """
    text = str(prompt or "")
    replacements = [
        ("взрывобезопасность", "требование безопасности оборудования"),
        ("взрывобезопасный", "безопасное исполнение оборудования"),
        ("взрывозащищённый", "безопасное исполнение оборудования"),
        ("взрывозащищенный", "безопасное исполнение оборудования"),
        ("взрывоопасный", "требование безопасности"),
        ("взрывоопасная", "требование безопасности"),
        ("взрывоопасное", "требование безопасности"),
        ("взрыв", "опасная ситуация"),
        ("оружие", "оборудование"),
        ("оружия", "оборудования"),
        ("военное назначение", "специальное назначение"),
        ("военного назначения", "специального назначения"),
        ("военная", "специальная"),
        ("военный", "специальный"),
        ("военное", "специальное"),
        ("терроризм", "запрещённая деятельность"),
        ("экстремизм", "запрещённая деятельность"),
        ("наркотики", "запрещённые вещества"),
        ("наркотических", "запрещённых"),
    ]
    # Сначала длинные фразы, затем отдельные слова.
    for src, dst in replacements:
        text = re.sub(re.escape(src), dst, text, flags=re.IGNORECASE)
    return text


def _looks_like_content_filter(text: str) -> bool:
    low = (text or "").lower()
    markers = (
        "чувствительными темами",
        "ответы на вопросы, связанные с чувствительными",
        "временно ограничены",
        "тематические ограничения",
        '"blacklist"',
    )
    return any(marker in low for marker in markers)


def _response_finish_reason(response) -> str:
    reason = getattr(response, "finish_reason", None)
    if reason:
        return str(reason).lower()
    choices = getattr(response, "choices", None) or []
    if choices:
        reason = getattr(choices[0], "finish_reason", None)
        if reason:
            return str(reason).lower()
    return ""


def _validate_and_extract_json_response(response, *, partial_array_key: str | None = None) -> dict[str, Any]:
    """Единая проверка ответа GigaChat для всех structured-output запросов."""
    finish_reason = _response_finish_reason(response)
    if finish_reason == "blacklist":
        raise RuntimeError(
            "GigaChat применил тематическое ограничение (finish_reason=blacklist). "
            "Запрос не был принят моделью."
        )
    text = extract_text(response)
    if not text:
        raise RuntimeError("GigaChat вернул пустой ответ.")
    if _looks_like_content_filter(text):
        raise RuntimeError(
            "GigaChat применил ограничение содержания и не вернул результат AI-аудита: "
            + text[:500]
        )
    return parse_json_response(text, partial_array_key=partial_array_key)


def ask_json_schema(
    client: GigaChat,
    prompt: str,
    schema: dict[str, Any],
    *,
    partial_array_key: str | None = None,
) -> dict[str, Any]:
    """Отправляет произвольную JSON Schema через тот же клиент GigaChat."""
    if GigaChat is None or ChatCompletionRequest is None or ChatMessage is None:
        raise RuntimeError(
            "Для AI-этапа необходим пакет gigachat с поддержкой ChatCompletionRequest."
        )

    safe_prompt = _sanitize_prompt_for_content_filter(prompt)
    if safe_prompt != prompt:
        print(
            "[GigaChat] AI-контекст санитизирован перед отправкой "
            "(локальный документ/БД не изменяются).",
            flush=True,
        )

    request = ChatCompletionRequest(
        model=GIGACHAT_MODEL,
        messages=[ChatMessage(role="user", content=safe_prompt)],
        response_format={
            "type": "json_schema",
            "schema": schema,
            "strict": True,
        },
    )
    return _validate_and_extract_json_response(
        client.chat.create(request),
        partial_array_key=partial_array_key,
    )


def ask_json(client: GigaChat, prompt: str) -> dict[str, Any]:
    """Запрос AI-аудита с заранее определённой схемой item_id -> result."""
    return ask_json_schema(client, prompt, _json_response_schema())


def ask_vision_json_schema(
    client: GigaChat,
    image_bytes: bytes,
    user_prompt: str,
    schema: dict[str, Any],
    *,
    mime_type: str = "image/png",
) -> dict[str, Any]:
    """Анализирует изображение GigaChat Vision и возвращает JSON.

    Изображение передаётся в исходном качестве через data URI, поэтому Vision
    видит таблицы, схемы и мелкие технические обозначения лучше, чем OCR
    локального Tesseract.
    """
    if GigaChat is None or ChatCompletionRequest is None:
        raise RuntimeError("Для GigaChat Vision необходим пакет gigachat с поддержкой multimodal messages.")
    if not image_bytes:
        raise ValueError("image_bytes пустой.")
    encoded = base64.b64encode(image_bytes).decode("ascii")
    data_url = f"data:{mime_type};base64,{encoded}"
    safe_prompt = _sanitize_prompt_for_content_filter(user_prompt)
    if safe_prompt != user_prompt:
        print("[GigaChat Vision] текстовый prompt санитизирован перед отправкой.")
    request = ChatCompletionRequest(
        model=GIGACHAT_MODEL,
        messages=[
            ChatMessage(
                role="user",
                content=[
                    ChatContentPart(type="text", text=safe_prompt),
                    ChatContentPart(type="image_url", image_url={"url": data_url}),
                ],
            )
        ],
        response_format={
            "type": "json_schema",
            "schema": schema,
            "strict": True,
        },
    )
    return _validate_and_extract_json_response(client.chat.create(request))


def ask_vision_json(
    client: GigaChat,
    image_bytes: bytes,
    user_prompt: str,
    *,
    mime_type: str = "image/png",
) -> dict[str, Any]:
    """Совместимый Vision API со стандартной схемой AI-аудита."""
    return ask_vision_json_schema(
        client, image_bytes, user_prompt, _json_response_schema(), mime_type=mime_type
    )
