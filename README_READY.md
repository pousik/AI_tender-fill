# AI Tender Filler — готовая версия

## Архитектура

`DOCX/PDF -> извлечение -> определение изделия по БД -> константы из БД -> GigaChat-2 Lite -> проверка -> документ`

### Главное правило БД

`transformers.db` — постоянная база знаний. Она не пересоздаётся и не очищается при запуске.

- `product_profiles` — конкретные изделия и паспортные константы;
- `knowledge_entries` — накопленные проверенные факты;
- `field_rules` — синонимы параметров из разных форм ТЗ;
- `tr_type_rules` + справочники — правила допустимых характеристик.

ИИ использует БД как опору, но его ответы автоматически не записываются в БД. Новое знание сначала добавляется через `knowledge_seed.json`, после проверки импортируется в БД.

## Быстрый старт Windows

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

$env:GIGACHAT_CREDENTIALS="ВАШ_КЛЮЧ"
$env:GIGACHAT_SCOPE="GIGACHAT_API_PERS"
$env:GIGACHAT_MODEL="GigaChat-2"
$env:GIGACHAT_BASE_URL="https://api.giga.chat/v1"
$env:GIGACHAT_VERIFY_SSL_CERTS="false"

python test_gigachat.py
python main.py "input.docx" "output.docx"
```

Для B2B/CORP замените `GIGACHAT_SCOPE` на scope, соответствующий вашему ключу.

## Наращивание БД

Добавьте новые проверенные профили/характеристики в `knowledge_seed.json` и выполните:

```powershell
python tools/add_knowledge.py knowledge_seed.json
```

Существующие записи не удаляются.

## Ошибки API

401/ошибка авторизации останавливает процесс до сохранения результата. Ложный `[УСПЕХ]` при ошибке GigaChat невозможен.
