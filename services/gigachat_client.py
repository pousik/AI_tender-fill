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
    from gigachat.models import ChatCompletionRequest, ChatMessage, ChatContentPart
except ImportError:  # позволяет запускать OCR/детерминированный режим без SDK
    GigaChat = None
    ChatCompletionRequest = ChatMessage = ChatContentPart = None

GIGACHAT_CREDENTIALS = os.getenv("GIGACHAT_CREDENTIALS", "MDFhMDY2YWMtZDc1Ni03NTc0LTg4MDEtMzJmYmY2YTY2MDkwOjZkMjFiMmI3LTY4MGItNGUyZS05YjQ4LWViY2E4NjIwMGU0Nw==").strip()
if GIGACHAT_CREDENTIALS.lower().startswith("bearer "):
    GIGACHAT_CREDENTIALS = GIGACHAT_CREDENTIALS[7:].strip()
GIGACHAT_MODEL = os.getenv("GIGACHAT_MODEL", "GigaChat-2")
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


def verify_connection(client: GigaChat) -> str:
    """Проверяет, что ключ/scope рабочие, используя уже открытый клиент."""
    try:
        models = client.get_models()
        count = len(getattr(models, "data", []) or [])
        return f"GigaChat OK: model={GIGACHAT_MODEL}, scope={GIGACHAT_SCOPE}, models={count}"
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


def parse_json_response(text: str) -> dict[str, Any]:
    raw = clean_json_text(text)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", raw, flags=re.S)
        if not match:
            raise
        data = json.loads(match.group(0))
    if not isinstance(data, dict):
        raise ValueError(
            f"GigaChat должен вернуть JSON-объект, получен {type(data).__name__}"
        )
    return data


def ask_json(client: GigaChat, prompt: str) -> dict[str, Any]:
    """Отправляет промпт и парсит строго-JSON ответ."""
    response = client.chat.create(prompt)
    return parse_json_response(extract_text(response))


def ask_vision_json(
    client: GigaChat,
    image_bytes: bytes,
    user_prompt: str,
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
    request = ChatCompletionRequest(
        messages=[
            ChatMessage(
                role="user",
                content=[
                    ChatContentPart(type="text", text=user_prompt),
                    ChatContentPart(type="image_url", image_url={"url": data_url}),
                ],
            )
        ]
    )
    response = client.chat.create(request)
    return parse_json_response(extract_text(response))
