# Final fixes

1. DOCX/PDF template lookup is relaxed for an exact source filename or exact product model, so partially corrected specialist templates can be reused.
2. Specialist name is written after Word closes to every DOCX section header as `Заполнил специалист: ФИО` and persisted in `tenders.specialist_name`.
3. `data_tenders` is the sole technical source. Manufacturer is extracted from document paragraphs (`ООО «Эльмаш (УЭТМ)»`) and indexed per product model.
4. Product type for TРГ-УЭТМ is now explicit. Voltage-only documents such as `Технические требования к трансформаторам тока 110 кВ` resolve to `ТРГ-УЭТМ-110`.
5. Instructional requirements such as `Указать` do not block a confirmed `data_tenders` value.
6. False OCR model `ТРГ-УЭТМ-2` is filtered; valid models are 35/110/220/330/500/750.
7. Directory watcher no longer marks a document dirty for unrelated directory events; it compares file size/mtime.
