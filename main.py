import sys
import os
from PIL import Image
import urllib.request as req
import pandas as pd
from PyQt6 import QtCore
from PyQt6.QtGui import QIntValidator

from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QMainWindow, QFileDialog, QLabel, QTextEdit, QLineEdit, \
    QMessageBox, QProgressBar, QErrorMessage


class QtGUI(QMainWindow):

    def __init__(self):
        super().__init__()
        self.success = []
        self.failed = []

        self.setWindowTitle("버그천국")

        self.resize(370, 300)
        title = QLabel("URL TO IMAGE", self)
        title.setStyleSheet("Color : black; font-weight:bold; font-size: 24px")
        title.resize(370, 40)
        title.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        file_label = QLabel("엑셀파일", self)
        file_label.resize(50, 30)
        file_label.move(20, 60)

        self.file_path = QTextEdit("", self)
        self.file_path.setReadOnly(True)
        self.file_path.resize(160, 30)
        self.file_path.move(80, 60)

        button1 = QPushButton("파일선택", self)
        button1.clicked.connect(self.btn_fun_file_load)
        button1.move(250, 60)

        directory_label = QLabel("저장경로", self)
        directory_label.resize(50, 30)
        directory_label.move(20, 110)

        self.directory_path = QTextEdit("", self)
        self.directory_path.setReadOnly(True)
        self.directory_path.resize(160, 30)
        self.directory_path.move(80, 110)

        button2 = QPushButton('폴더선택', self)
        button2.clicked.connect(self.btn_fun_directory_select)
        button2.move(250, 110)

        pixel_label = QLabel("Pixel", self)
        pixel_label.resize(50, 30)
        pixel_label.move(20, 160)

        self.pixel1 = QLineEdit("600", self)
        self.pixel1.resize(70, 30)
        self.pixel1.move(80, 160)
        self.pixel1.setValidator(QIntValidator(self))

        self.pixel2 = QLineEdit("600", self)
        self.pixel2.resize(70, 30)
        self.pixel2.move(170, 160)
        self.pixel2.setValidator(QIntValidator(self))

        self.pbar = QProgressBar(self)
        self.pbar.setValue(0)
        self.pbar.resize(350, 30)
        self.pbar.move(20, 210)

        button3 = QPushButton('변환', self)
        button3.resize(370, 40)
        button3.clicked.connect(self.url_to_image_save)
        button3.move(0, 260)

        self.setWindowTitle("버그천국")

        self.show()

    def btn_fun_file_load(self):
        file_name = QFileDialog.getOpenFileName(self)
        self.file_path.setText(file_name[0])

    def btn_fun_directory_select(self):
        directory_path = QFileDialog.getExistingDirectory(self)
        self.directory_path.setText(directory_path)

    def url_to_image_save(self):
        self.pbar.setValue(10)
        try:
            df = pd.DataFrame(pd.read_excel(self.file_path.toPlainText()))
            urls = list(df['대표이미지'])
            print(urls)
            for i in range(len(urls)):
                try:
                    req.urlretrieve(urls[i], self.directory_path.toPlainText() + "/" + "image" + str(i) + ".jpg")
                    self.success.append(urls[i])
                except:
                    self.failed.append(urls[i])
                    continue

            self.pbar.setValue(50)
            self.change_image_pixel()
            self.pbar.setValue(100)
            self.end_dialog()
            self.pbar.setValue(0)

            print("SUCCESS!!!")
        except Exception as e:
            self.error_dialog(e)
            print("ERROR!!!" + e)

    def change_image_pixel(self):
        image_list = os.listdir(self.directory_path.toPlainText())

        for name in image_list:
            im = Image.open(self.directory_path.toPlainText() + "/" + name)
            im = im.resize((int(self.pixel1.displayText()), int(self.pixel2.displayText())))
            im = im.convert('RGB')
            im.save(self.directory_path.toPlainText() + "/" + name)

    def end_dialog(self):
        msg = QMessageBox()
        msg.setWindowTitle("변환완료")
        msg.resize(300, 300)
        field = ""
        field += f"{len(self.success)}개 이미지 변환완료 \n\n"
        field += f"{len(self.failed)}개 이미지 실패 \n"
        for i in range(len(self.failed)):
            field += f"{self.failed[i]} 실패 \n"
        msg.setText(field)
        msg.exec()

    def error_dialog(self, error):
        msg = QErrorMessage()
        msg.StandardButton(f"오류 발생: {error}")

if __name__ == '__main__':
    app = QApplication(sys.argv)

    ex = QtGUI()

    app.exec()
