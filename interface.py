import os
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QApplication, QFileDialog, QLabel, QPushButton, QVBoxLayout, QWidget

from main import Session, process_document
from bootstrap import bootstrap
from services.document_tracker import DocumentTracker
from services.pdf_word_edit import prepare_pdf_for_word_editing


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.file_name: str | None = None
        self.output_file: Path | None = None
        self.processed = False

        self.setWindowTitle("Автозаполнение тендерного документа")
        self.resize(420, 320)

        layout = QVBoxLayout(self)
        layout.addStretch()

        self.FileNameLabel = QLabel("Файл не задан", self)
        self.FileNameLabel.setWordWrap(True)
        self.FileNameLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.button = QPushButton("Выбрать файл")
        self.pending = QPushButton("Обработать")
        self.open_btn = QPushButton("Открыть заполненный файл")

        for btn in (self.button, self.pending, self.open_btn):
            btn.setFixedSize(260, 50)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #007ACC;
                    color: white;
                    font-size: 14px;
                    border-radius: 6px;
                }
                QPushButton:hover { background-color: #005999; }
            """)

        self.button.clicked.connect(self.open_file_dialog)
        self.pending.clicked.connect(self.get_filling_file)
        self.open_btn.clicked.connect(self.open_file)

        layout.addWidget(self.FileNameLabel, alignment=Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(self.open_btn, alignment=Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(self.pending, alignment=Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(self.button, alignment=Qt.AlignmentFlag.AlignHCenter)
        layout.addStretch()

        self.tracker = DocumentTracker(Session, self)
        self.tracker.status.connect(print)
        self.tracker.learned.connect(lambda result: print("[KB] Результат обучения:", result))

    def get_filling_file(self):
        if not self.file_name:
            print("Сначала выберите файл!")
            return

        input_doc = Path(self.file_name)
        suffix = ""
        if(input_doc.suffix.lower() in (".doc", ".docx")):
            suffix = "docx"
        elif (input_doc.suffix.lower() == ".pdf"):
            suffix = "pdf"

        self.output_file = (
                input_doc.parent /
                f"Заполненный_{input_doc.stem}.{suffix}"
        )

        try:
            with Session() as session:
                result = process_document(str(input_doc), str(self.output_file), session)

            if result:
                self.processed = True
                print(result)

                if self.output_file.suffix.lower() == ".pdf":
                    # Вариант 2: инженер редактирует результат в Microsoft Word.
                    # После закрытия Word tracker конвертирует изменённый DOCX
                    # обратно в исходный PDF и записывает исправления в БЗ.
                    editable = self.output_file.with_name(
                        self.output_file.stem + "_для_редактирования.docx"
                    )
                    editable = prepare_pdf_for_word_editing(self.output_file, editable)
                    print(f"[PDF/WORD] Создана редактируемая копия: {editable}")
                    self.tracker.watch(editable, pdf_target=self.output_file)
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(editable)))

                elif self.output_file.suffix.lower() == ".docx":
                    self.tracker.watch(self.output_file)
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_file)))
                else:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_file)))
        except Exception as exc:
            print(f"[PROCESS] Ошибка: {exc}")

    def open_file_dialog(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть файл",
            "",
            "Документы (*.docx *.doc *.docm *.odt *.rtf);;PDF (*.pdf);;Все файлы (*.*)",
        )
        if file_name:
            self.file_name = file_name
            self.FileNameLabel.setText(f"Выбран файл: {Path(file_name).name}")
            print(f"Выбран файл: {file_name}")

    def open_file(self):
        if not self.output_file:
            print("Файл ещё не обработан или не существует!")
            return

        # Для PDF открываем именно Word-компаньон, если он уже создан.
        if self.output_file.suffix.lower() == ".pdf":
            editable = self.output_file.with_name(
                self.output_file.stem + "_для_редактирования.docx"
            )
            if editable.exists():
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(editable)))
                return

        if self.output_file.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_file)))
        else:
            print("Файл ещё не обработан или не существует!")

    def closeEvent(self, event):
        self.tracker.stop()
        event.accept()


if __name__ == "__main__":
    bootstrap()

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
