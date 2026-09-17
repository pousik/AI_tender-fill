# V6

Исправлена ошибка:
TypeError: 'NoneType' object is not iterable

Причина:
`model_profile()` пытался итерировать `self._image_records`, который в некоторых
ветках был None.

Исправления:
- `model_profile()` всегда работает со списком изображений;
- кэш изображений нормализуется в `[]`;
- если модель найдена через fallback, она добавляется в candidates;
- product signature нормализуется;
- document_context безопасно обрабатывает None.

Ожидаемый профиль:
[PROFILE] {'product_type': 'ТРГ-УЭТМ', 'model': 'ТРГ-УЭТМ-110',
'voltage_class': '110', ..., 'candidates': ['ТРГ-УЭТМ-110']}
