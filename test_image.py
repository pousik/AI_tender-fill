from io import BytesIO
import docx

from services.gigachat_client import ask_vision_json, open_client

# Импортируем из вашего модуля GigaChat
# from your_gigachat_module import open_client, ask_vision_json

doc_path = "data_tenders/1БП.769.001-01 РЭ от января 2026 г.docx"
doc = docx.Document(doc_path)

with open_client() as client:
    for rel in doc.part.rels.values():
        if not rel.is_external and "image" in rel.target_ref:
            try:
                image_bytes = rel.target_part.blob

                print(f"Обработка изображения: {rel.target_ref} via GigaChat Vision...")

                # Запрос к нейросети
                extracted_data = ask_vision_json(
                    client=client,
                    image_bytes=image_bytes,
                    user_prompt="Извлеки все параметры оборудования и их значения со схемы/изображения."
                )

                print("--- Распознанный JSON с картинки ---")
                print(extracted_data)
                print("=" * 40)

            except Exception as e:
                print(f"Ошибка обработки изображения {rel.target_ref}: {e}")