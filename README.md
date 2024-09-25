# i1TzD4Rk
import sys
import os
import pytesseract
pytesseract.pytesseract.tesseract_cmd=r'C:\Users\DAVHSS-41\Desktop\Eternal Desolate\Tesseract-OCR\tesseract.exe'
import cv2
import docx
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog, QTextEdit, QComboBox
from PyQt5.QtGui import QPixmap, QImage, QIcon
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PIL import Image

class OCRWorker(QThread):
    extracted = pyqtSignal(str)

    def __init__(self, image_path, language, preprocess_options):
        super().__init__()
        self.image_path = image_path
        self.language = language
        self.preprocess_options = preprocess_options

    def run(self):
        if self.image_path:
            try:
                image = Image.open(self.image_path)
                custom_config = r'--oem 3 --psm 6 -l ' + self.language
                if self.preprocess_options == 'Thresholding':
                    image = image.convert('L')
                    image = image.point(lambda x: 0 if x < 200 else 255, '1')
                extracted_text = pytesseract.image_to_string(image, config=custom_config)
                self.extracted.emit(extracted_text)
            except Exception as e:
                print(f"Error during OCR: {e}")
                self.extracted.emit("Error during OCR, please try another image.")

class OCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.languages = pytesseract.get_languages()
        self.initUI()
        self.image_path = None
        self.doc = None
        self.save_format = 'Word'  # Default save format

    def initUI(self):
        self.setWindowTitle('OCR Application')
        self.setGeometry(100, 100, 800, 600)

        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)

        self.result_text_edit = QTextEdit(self)
        self.result_text_edit.setReadOnly(True)

        self.load_button = QPushButton('Load Image', self)
        self.load_button.setStyleSheet("background-color: #4CAF50; color: white;")
        self.load_button.clicked.connect(self.load_image)

        self.extract_button = QPushButton('Extract Text', self)
        self.extract_button.setStyleSheet("background-color: #008CBA; color: white;")
        self.extract_button.clicked.connect(self.extract_text)

        self.word_button = QPushButton('Save as Word', self)
        self.word_button.setStyleSheet("background-color: #f44336; color: white;")
        self.word_button.clicked.connect(self.save_as_word)

        self.text_button = QPushButton('Save as Text', self)
        self.text_button.setStyleSheet("background-color: #f44336; color: white;")
        self.text_button.clicked.connect(self.save_as_text)

        self.excel_button = QPushButton('Save as Excel', self)
        self.excel_button.setStyleSheet("background-color: #f44336; color: white;")
        self.excel_button.clicked.connect(self.save_as_excel)

        self.language_combo_box = QComboBox(self)
        self.language_combo_box.addItems(self.languages)

        self.preprocess_combo_box = QComboBox(self)
        self.preprocess_combo_box.addItems(['None', 'Thresholding'])

        hbox = QHBoxLayout()
        hbox.addWidget(self.load_button)
        hbox.addWidget(self.extract_button)
        hbox.addWidget(self.word_button)
        hbox.addWidget(self.text_button)
        hbox.addWidget(self.excel_button)
        hbox.addWidget(self.language_combo_box)
        hbox.addWidget(self.preprocess_combo_box)

        vbox = QVBoxLayout()
        vbox.addWidget(self.image_label, 1)
        vbox.addWidget(self.result_text_edit, 2)
        vbox.addLayout(hbox)

        central_widget = QWidget(self)
        central_widget.setLayout(vbox)
        self.setCentralWidget(central_widget)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_result)
        self.timer.start(100)

        self.worker = None
        self.video_capture = None
        self.video_timer = None

    def load_image(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, 'Open Image File', '', 'Images (.png .jpg .bmp);;All Files ()', options=options)
        if file_path:
            self.image_path = file_path
            self.stop_live_feed()
            self.update_image()

    def stop_live_feed(self):
        if self.video_timer and self. video_timer.isActive():
            self.video_timer.stop()
        if self.video_capture:
            self.video_capture.release()

    def update_image(self):
        if self.image_path:
            pixmap = QPixmap(self.image_path)
            self.image_label.setPixmap(pixmap.scaledToWidth(600, Qt.SmoothTransformation))
        else:
            self.image_label.setPixmap(QPixmap())

    def extract_text(self):
        if self.image_path:
            self.result_text_edit.clear()
            language = self.languages[self.language_combo_box.currentIndex()]
            preprocess_options = self.preprocess_combo_box.currentText()
            self.worker = OCRWorker(self.image_path, language, preprocess_options)
            self.worker.extracted.connect(self.display_extracted_text)
            self.worker.start()

    def display_extracted_text(self, extracted_text):
        self.result_text_edit.setPlainText(extracted_text)
        self.doc = docx.Document()
        self.doc.add_paragraph(extracted_text)

    def update_result(self):
        if self.image_path or (self.video_capture and self.video_capture.isOpened()):
            self.extract_button.setEnabled(True)
            self.word_button.setEnabled(True)
            self.text_button.setEnabled(True)
            self.excel_button.setEnabled(True)
        else:
            self.extract_button.setEnabled(False)
            self.word_button.setEnabled(False)
            self.text_button.setEnabled(False)
            self.excel_button.setEnabled(False)

    def save_as_word(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save Text', '', 'Word Files (.docx);;All Files ()', options=options)
        if file_path:
            self.save_as_file(file_path, 'Word')

    def save_as_text(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save Text', '', 'Text Files (.txt);;All Files ()', options=options)
        if file_path:
            self.save_as_file(file_path, 'Text')

    def save_as_excel(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save Text', '', 'Excel Files (.xlsx);;All Files ()', options=options)
        if file_path:
            self.save_as_file(file_path, 'Excel')

    def save_as_file(self, file_path, save_format):
        if self.doc:
            if save_format == 'Word':
                self.doc.save(file_path)
            elif save_format == 'Text':
                with open(file_path, 'w') as txt_file:
                    txt_file.write(self.result_text_edit.toPlainText())
            elif save_format == 'Excel':
                # Implement Excel saving logic here
                pass

    def start_live_feed(self):
        self.result_text_edit.clear()
        self.image_path = None
        self.worker = None
        self.stop_live_feed()

        self.video_capture = cv2.VideoCapture(0)
        self.video_timer = QTimer(self)
        self.video_timer.timeout.connect(self.update_live_feed)
        self.video_timer.start(30)

    def update_live_feed(self):
        ret, frame = self.video_capture.read()
        if ret:
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            image = QImage(frame_rgb, frame_rgb.shape[1], frame_rgb.shape[0], QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(image)
            self.image_label.setPixmap(pixmap.scaledToWidth(600, Qt.SmoothTransformation))

    def closeEvent(self, event):
        self.stop_live_feed()
        super().closeEvent(event)

class HomeScreen(QMainWindow):
    def __init__(self, ocr_app):
        super().__init__()
        self.setWindowTitle('Home Screen')
        self.setGeometry(100, 100, 400, 300)
        self.ocr_app = ocr_app

        self.initUI()

    def initUI(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        welcome_label = QLabel('Welcome to the OCR App!')
        welcome_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(welcome_label)

        ocr_button = QPushButton('Open OCR App', self)
        ocr_button.setStyleSheet("background-color: #008CBA; color: white;")
        ocr_button.clicked.connect(self.open_ocr_app)
        layout.addWidget(ocr_button)

        exit_button = QPushButton('Exit', self)
        exit_button.setStyleSheet("background-color: #f44336; color: white;")
        exit_button.clicked.connect(self.close)
        layout.addWidget(exit_button)

        central_widget.setLayout(layout)

    def open_ocr_app(self):
        self.ocr_app.show()
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)

    ocr_app = OCRApp()
    home_screen = HomeScreen(ocr_app)

    app.setWindowIcon(QIcon('icon.png'))  

    home_screen.show()
    sys.exit(app.exec_())
