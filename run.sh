#!/usr/bin/env bash
set -euo pipefail

if [ -z "${GIGACHAT_CREDENTIALS:-}" ]; then
    read -r -s -p "Введите актуальный GigaChat Authorization Key: " GIGACHAT_CREDENTIALS
    echo
    export GIGACHAT_CREDENTIALS
fi

export GIGACHAT_SCOPE="${GIGACHAT_SCOPE:-GIGACHAT_API_PERS}"
export GIGACHAT_MODEL="${GIGACHAT_MODEL:-GigaChat-2}"
export GIGACHAT_BASE_URL="${GIGACHAT_BASE_URL:-https://api.giga.chat/v1}"
export GIGACHAT_VERIFY_SSL_CERTS="${GIGACHAT_VERIFY_SSL_CERTS:-true}"
export GIGACHAT_TIMEOUT="${GIGACHAT_TIMEOUT:-180}"
export GIGACHAT_MAX_RETRIES="${GIGACHAT_MAX_RETRIES:-2}"

python3 test_gigachat.py
python3 main.py "$@"
