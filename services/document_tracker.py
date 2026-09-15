from __future__ import annotations

import os
import time
from pathlib import Path

import psutil

from PySide6.QtCore import QObject, QFileSystemWatcher, QTimer, Signal


class DocumentTracker(QObject):
    status = Signal(str)
    learned = Signal(object)
    """
    Отслеживает DOCX, открытый инженером в Word.

    Логика:
        1. Генератор создал DOCX.
        2. Документ открывается в Word.
        3. Word может сохранять файл сколько угодно раз.
        4. Изменения только помечают документ как "грязный".
        5. Когда Word действительно закрыл этот файл,
           выполняется обучение БЗ ровно один раз.
    """

    def __init__(self, session_factory, parent=None):
        super().__init__(parent)

        self.session_factory = session_factory

        self.watcher = QFileSystemWatcher(self)
        self.watcher.fileChanged.connect(self._on_file_changed)
        self.watcher.directoryChanged.connect(self._on_directory_changed)

        self._files: dict[str, dict] = {}

        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._check_documents)

    # ============================================================
    # START
    # ============================================================

    def _log(self, message: str):
        print(message)
        self.status.emit(message)

    def watch(self, path: str | Path, pdf_target: str | Path | None = None):
        """Начать отслеживание документа. Для DOCX-компаньона PDF можно указать pdf_target."""
        path = str(Path(path).resolve())
        pdf_target = str(Path(pdf_target).resolve()) if pdf_target else None

        if not os.path.exists(path):
            self._log(f"[KB] Файл не найден: {path}")
            return

        # Уже отслеживается
        if path in self._files:
            if pdf_target:
                self._files[path]["pdf_target"] = pdf_target
            self._log(f"[KB] Уже отслеживается: {Path(path).name}")
            return

        directory = str(Path(path).parent)

        self._files[path] = {
            "path": path,
            "directory": directory,
            "dirty": False,
            "last_size": os.path.getsize(path),
            "last_mtime": os.path.getmtime(path),
            "last_event": time.monotonic(),
            "processing": False,
            "learned": False,
            "pdf_target": pdf_target,
        }

        if path not in self.watcher.files():
            self.watcher.addPath(path)

        if directory not in self.watcher.directories():
            self.watcher.addPath(directory)

        if not self._timer.isActive():
            self._timer.start()

        self._log(f"[KB] Отслеживается документ: {Path(path).name}")

    # ============================================================
    # FILE CHANGED
    # ============================================================

    def _on_file_changed(self, path: str):
        path = str(Path(path).resolve())

        info = self._files.get(path)
        if not info:
            return

        # Любое изменение только помечает файл.
        # Никакого обучения непосредственно здесь!
        info["dirty"] = True
        info["last_event"] = time.monotonic()

        try:
            info["last_size"] = os.path.getsize(path)
            info["last_mtime"] = os.path.getmtime(path)
        except OSError:
            pass

        self._log(f"[KB] Изменение документа: {Path(path).name}")

        # Word иногда удаляет watcher при замене файла.
        # Поэтому снова добавляем его, если файл существует.
        if os.path.exists(path) and path not in self.watcher.files():
            self.watcher.addPath(path)

    # ============================================================
    # DIRECTORY CHANGED
    # ============================================================

    def _on_directory_changed(self, directory: str):
        directory = str(Path(directory).resolve())

        for path, info in self._files.items():
            if info["directory"] != directory:
                continue

            if os.path.exists(path):
                if path not in self.watcher.files():
                    self.watcher.addPath(path)

                info["dirty"] = True
                info["last_event"] = time.monotonic()

    # ============================================================
    # PERIODIC CHECK
    # ============================================================

    def _check_documents(self):
        if not self._files:
            self._timer.stop()
            return

        for path, info in list(self._files.items()):

            # Уже обработан
            if info["learned"]:
                continue

            # Уже идёт обработка
            if info["processing"]:
                continue

            # Файл должен существовать
            if not os.path.exists(path):
                continue

            # Если изменений ещё не было — ничего не делаем
            if not info["dirty"]:
                continue

            # Должна пройти небольшая пауза после последнего события.
            #
            # Это защищает от ситуации:
            #
            # Word сохраняет:
            #   fileChanged
            #   fileChanged
            #   fileChanged
            #
            # и мы начинаем читать документ посреди сохранения.
            if time.monotonic() - info["last_event"] < 3.0:
                continue

            # Файл должен быть стабилен
            if not self._is_file_stable(info):
                continue

            # Самое главное:
            # Word действительно должен закрыть файл.
            if self._is_file_open_in_word(path):
                continue

            # Только здесь начинаем обучение
            self._learn_document(path, info)

    # ============================================================
    # FILE STABILITY
    # ============================================================

    def _is_file_stable(self, info: dict) -> bool:
        path = info["path"]

        try:
            size = os.path.getsize(path)
            mtime = os.path.getmtime(path)
        except OSError:
            return False

        old_size = info["last_size"]
        old_mtime = info["last_mtime"]

        if size != old_size or mtime != old_mtime:
            info["last_size"] = size
            info["last_mtime"] = mtime
            info["last_event"] = time.monotonic()
            return False

        return True

    # ============================================================
    # WORD OPEN CHECK
    # ============================================================

    def _is_file_open_in_word(self, path: str) -> bool:
        path = str(Path(path).resolve()).lower()

        for process in psutil.process_iter(["name"]):
            try:
                name = (process.info["name"] or "").lower()

                if name not in (
                    "winword.exe",
                    "word.exe",
                ):
                    continue

                for opened in process.open_files() or []:
                    try:
                        opened_path = str(Path(opened.path).resolve()).lower()

                        if opened_path == path:
                            return True

                    except (OSError, ValueError):
                        continue

            except (
                psutil.NoSuchProcess,
                psutil.AccessDenied,
                psutil.ZombieProcess,
            ):
                continue

        return False

    # ============================================================
    # LEARNING
    # ============================================================

    def _learn_document(self, path: str, info: dict):
        info["processing"] = True
        filename = Path(path).name

        self._log(f"[KB] Word закрыт, документ готов к обучению: {filename}")

        try:
            from services.template_knowledge import capture_docx_file
            from services.pdf_word_edit import docx_to_pdf

            session = self.session_factory()
            try:
                # Сначала готовим все результаты. Если PDF-компаньон не может
                # быть пересохранён, транзакция БЗ не фиксируется и документ
                # остаётся отслеживаемым для повторной попытки.
                result = capture_docx_file(session, path)

                pdf_target = info.get("pdf_target")
                if pdf_target:
                    docx_to_pdf(path, pdf_target)
                    self._log(
                        f"[PDF/WORD] Изменения сохранены обратно в PDF: {Path(pdf_target).name}"
                    )

                session.commit()
                self.learned.emit(result)
                info["learned"] = True
                info["dirty"] = False
                self._log(f"[KB] Обучение завершено один раз: {result}")

            except Exception:
                session.rollback()
                raise
            finally:
                session.close()

        except Exception as exc:
            self._log(f"[KB] Ошибка обучения {filename}: {exc}")
            info["processing"] = False
            return

        info["processing"] = False
        self._stop_tracking(path)

    # ============================================================
    # STOP TRACKING
    # ============================================================

    def _stop_tracking(self, path: str):
        path = str(Path(path).resolve())

        info = self._files.pop(path, None)

        if not info:
            return

        try:
            if path in self.watcher.files():
                self.watcher.removePath(path)
        except Exception:
            pass

        if not self._files:
            self._timer.stop()

        self._log(
            f"[KB] Отслеживание завершено: {Path(path).name}"
        )

    # ============================================================
    # STOP ALL
    # ============================================================

    def stop(self):
        self._timer.stop()

        try:
            self.watcher.removePaths(self.watcher.files())
            self.watcher.removePaths(self.watcher.directories())
        except Exception:
            pass

        self._files.clear()

    def close(self):
        self.stop()