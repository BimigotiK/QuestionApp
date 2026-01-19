#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Простой PyQt5‑приложение для работы с Word‑файлами:
• Загрузка .docx → список вопросов (с чек‑боксами)
• Отображение текста и вложенных изображений
• Случайный выбор 30 вопросов
• Сохранение выбранных в новый .docx
"""

import sys, os, random, tempfile
from pathlib import Path
from io import BytesIO

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QScrollArea, QCheckBox, QFrame,
    QFileDialog, QMessageBox, QGroupBox, QSizePolicy
)
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt

try:
    from docx import Document
except ImportError:
    raise RuntimeError("Не найден модуль python-docx. Установите: pip install python-docx")

# ------------------------------------------------------------------
# 1. Парсер вопросов и изображений ---------------------------------
def parse_questions_with_images(doc: Document):
    questions = []
    cur_text_parts, cur_images = [], []

    def add_current():
        if cur_text_parts or cur_images:
            text = "\n".join(cur_text_parts).strip()
            questions.append({"text": text, "images": list(cur_images)})

    for para in doc.paragraphs:
        txt = para.text.strip()

        # Разделители вопросов
        if txt.upper() == "---START---":
            add_current(); cur_text_parts.clear(); cur_images.clear(); continue
        if txt.upper() == "---END---":
            add_current(); cur_text_parts.clear(); cur_images.clear(); continue

        # Текст текущего вопроса (игнорируем только [BILD])
        if txt and txt != "[BILD]":
            cur_text_parts.append(txt)

        # Изображения внутри параграфа
        for run in para.runs:
            drawing = run.element.find(".//{*}drawing")
            if not drawing: continue
            blip = drawing.find(".//{*}blip")
            if not blip: continue
            rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not rId or rId not in doc.part.related_parts: continue
            try:
                part = doc.part.related_parts[rId]
                cur_images.append(part.blob)
            except Exception:
                pass

    add_current()
    return questions


# ------------------------------------------------------------------
# 2. Виджет вопроса -----------------------------------------------
class QuestionWidget(QWidget):
    def __init__(self, data: dict, index: int, parent=None):
        super().__init__(parent)
        self.data = data
        self.index = index
        self.checkbox = QCheckBox()
        self.checkbox.setFixedSize(25, 25)

        main_lay = QHBoxLayout(self)
        main_lay.addWidget(self.checkbox)

        content_wid = QWidget()
        content_lay = QVBoxLayout(content_wid)
        content_lay.setContentsMargins(5, 0, 0, 0)

        # Текст вопроса
        if self.data["text"]:
            for line in filter(None, self.data["text"].splitlines()):
                lbl = QLabel(line.strip())
                lbl.setWordWrap(True)
                lbl.setStyleSheet("font-size:12pt;")
                content_lay.addWidget(lbl)

        # Изображения (разделяем по маркерам [BILD])
        text_parts = self.data['text'].split('[BILD]')
        images = self.data.get('images', [])

        for i, part in enumerate(text_parts):
            if part.strip():
                lbl = QLabel(part.strip())
                lbl.setWordWrap(True)
                lbl.setStyleSheet("font-size:12pt;")
                content_lay.addWidget(lbl)

            if i < len(images):
                try:
                    image = QImage.fromData(images[i])
                    if image.isNull(): continue
                    pixmap = QPixmap.fromImage(image)
                    img_lbl = QLabel()
                    img_lbl.setPixmap(pixmap.scaledToWidth(600, Qt.SmoothTransformation))
                    img_lbl.setAlignment(Qt.AlignCenter)
                    content_lay.addWidget(img_lbl)
                except Exception:
                    pass

        main_lay.addWidget(content_wid, 1)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)
        sep.setStyleSheet("color:#ccc; margin-top:10px;")
        main_lay.addWidget(sep)

    def is_checked(self): return self.checkbox.isChecked()
    def set_checked(self, val): self.checkbox.setChecked(val)


# ------------------------------------------------------------------
# 3. Главное окно -----------------------------------------------
class QuestionApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.questions = []
        self.question_widgets = []

        self._setup_ui()
        self.setWindowTitle("Выбор вопросов из Word")
        self.resize(1200, 900)

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_lay = QHBoxLayout(central)

        # Левая панель
        left_panel = QWidget()
        left_panel.setFixedWidth(300)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignTop)

        btn_load = QPushButton("📂 Загрузить файл")
        btn_load.clicked.connect(self.load_file_dialog)
        left_layout.addWidget(btn_load)

        self.drop_label = QLabel("📄 Перетащите сюда .docx файл")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setFixedHeight(100)
        self.drop_label.setStyleSheet("""
            border: 2px dashed #aaa; border-radius:10px;
            background:#f0f0f0; padding:20px; font-size:14px;
        """)
        self.drop_label.setAcceptDrops(True)
        self.drop_label.dragEnterEvent = self._drag_enter
        self.drop_label.dropEvent = self._drop_file
        left_layout.addWidget(self.drop_label)

        left_layout.addSpacing(20)

        self.counter_lbl = QLabel("Выбрано: 0")
        self.loaded_lbl = QLabel("Загружено: 0")
        for lbl in (self.counter_lbl, self.loaded_lbl):
            lbl.setStyleSheet("font-size:16px;")
            lbl.setAlignment(Qt.AlignCenter)
            left_layout.addWidget(lbl)

        btn_random = QPushButton("🎲 Случайно выбрать 30")
        btn_random.clicked.connect(self.random_select)
        left_layout.addWidget(btn_random)

        btn_save = QPushButton("💾 Сохранить выбранные")
        btn_save.clicked.connect(self.save_selected)
        left_layout.addWidget(btn_save)

        left_layout.addStretch()

        # Правая панель
        right_panel = QWidget()
        right_lay = QVBoxLayout(right_panel)

        title = QLabel("Вопросы")
        title.setStyleSheet("font-size:18pt; font-weight:bold; margin:10px;")
        title.setAlignment(Qt.AlignCenter)
        right_lay.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setAlignment(Qt.AlignTop)
        scroll.setWidget(self.scroll_content)
        right_lay.addWidget(scroll)

        main_lay.addWidget(left_panel)
        main_lay.addWidget(right_panel, 1)

    def _drag_enter(self, event):
        if event.mimeData().hasUrls(): event.acceptProposedAction()

    def _drop_file(self, event):
        for url in event.mimeData().urls():
            path = Path(url.toLocalFile())
            if path.suffix.lower() == ".docx":
                self.load_file(str(path))
                break
        event.acceptProposedAction()

    def load_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Word файл", "", "Word files (*.docx)"
        )
        if file_path: self.load_file(file_path)

    def load_file(self, path: str):
        try:
            doc = Document(path)
            self.questions = parse_questions_with_images(doc)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл:\n{e}")
            return

        for w in self.question_widgets: w.setParent(None)
        self.question_widgets.clear()

        for idx, qdata in enumerate(self.questions):
            qw = QuestionWidget(qdata, idx)
            qw.checkbox.stateChanged.connect(self.update_counter)
            self.question_widgets.append(qw)
            self.scroll_layout.addWidget(qw)

        self.loaded_lbl.setText(f"Загружено: {len(self.questions)}")
        self.update_counter()
        QMessageBox.information(self, "Успех", f"{len(self.questions)} вопросов загружено")

    def update_counter(self):
        count = sum(1 for w in self.question_widgets if w.is_checked())
        self.counter_lbl.setText(f"Выбрано: {count}")

    def random_select(self):
        if len(self.question_widgets) < 30:
            QMessageBox.warning(self, "Ошибка", f"Файл содержит только {len(self.question_widgets)} вопросов")
            return
        for w in self.question_widgets: w.set_checked(False)
        idxs = random.sample(range(len(self.question_widgets)), 30)
        for i in idxs: self.question_widgets[i].set_checked(True)
        self.update_counter()

    def save_selected(self):
        selected_q = [q for w, q in zip(self.question_widgets, self.questions) if w.is_checked()]
        if not selected_q:
            QMessageBox.warning(self, "Ошибка", "Нет выбранных вопросов")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить выбранные вопросы", "", "Word files (*.docx)"
        )
        if not file_path: return

        try:
            out_doc = Document()
            for q in selected_q:
                if q["text"]:
                    for line in q["text"].splitlines():
                        p = out_doc.add_paragraph(line.strip())
                for img_bytes in q.get("images", []):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                        tmp.write(img_bytes)
                        tmp.flush()
                        out_doc.add_picture(tmp.name, width=Inches(5.0))
                        os.unlink(tmp.name)
                out_doc.add_paragraph("---")
            out_doc.save(file_path)
            QMessageBox.information(self, "Готово", f"Файл сохранён: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл:\n{e}")

def main():
    app = QApplication(sys.argv)
    win = QuestionApp()
    win.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
