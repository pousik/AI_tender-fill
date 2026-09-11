"""Быстрая проверка подключения к GigaChat перед основным запуском.

Используется в run.ps1: если ключ/scope неверны, пайплайн не запускается
и не тратит время на OCR/парсинг документа впустую.
"""

from services.gigachat_client import check_gigachat_connection

if __name__ == "__main__":
    print(check_gigachat_connection())
