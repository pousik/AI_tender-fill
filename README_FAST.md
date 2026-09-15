# Ускоренный режим

## Главное
Обычный запуск не выполняет новые GigaChat Vision-запросы к изображениям `data_tenders`.
Используется сохранённый `.image_ocr_cache.json`, если он был создан предварительной обработкой.

Для один раз построения Vision-индекса:

```powershell
$env:DATA_TENDERS_VISION_LIVE='1'
python preprocess_data_tenders.py
Remove-Item Env:DATA_TENDERS_VISION_LIVE
```

Для PDF live Vision выключен по умолчанию; включение:

```powershell
$env:PDF_VISION_LIVE='1'
python main.py ...
```

## Ускорения
- точный поиск `data_tenders`: O(1) через хеш-индексы;
- шаблонная БЗ: один запрос всех `TenderParameter` вместо N+1 и индекс точных ключей;
- модель/напряжение вычисляются один раз на документ;
- Vision-результаты кэшируются;
- PDF Vision кэшируется по mtime/size;
- `verify_connection()` не выполняется перед каждым PDF AI-запросом без `TENDER_VERIFY_GIGACHAT=1`;
- PDF OCR default 240 DPI, значение можно вернуть через `OCR_DPI=300` для максимальной точности.
