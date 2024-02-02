import sys

from PyQt5.QtCore import QDate, pyqtSignal, QObject, pyqtSlot, Qt
from PyQt5.QtWidgets import QApplication, QWidget, QDesktopWidget, QFileDialog, QPushButton, QVBoxLayout, QHBoxLayout, \
    QMessageBox, QLabel, QCalendarWidget, QLineEdit

from excelmanager import ExcelManager


class CalendarWindow(QWidget):
    send_date = pyqtSignal(str)

    def __init__(self):
        super().__init__()

        self.date = None
        self.setWindowTitle("Calendar")
        self.init_ui()

    def init_ui(self):
        self.calendar = QCalendarWidget(self)
        self.calendar.setGridVisible(True)

        vbox = QVBoxLayout()
        vbox.addWidget(self.calendar)
        self.setLayout(vbox)

        self.calendar.clicked.connect(self.select_date)

    def select_date(self):
        self.date = self.calendar.selectedDate().toString('yyyy-MM-dd')
        self.send_date.emit(self.date)


class App(QWidget):
    def __init__(self):
        super().__init__()

        self.base_url = None
        self.text_url = None
        self.save_url = None

        self.date_data = None
        self.calendar_window = None

        self.product_info1 = None
        self.product_info2 = None

        self.excelmanager = ExcelManager()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('엑셀 자동화 프로그램')
        self.setFixedSize(600, 300)

        # 버튼 레이아웃
        self.select_base_button = QPushButton('&base path', self)
        self.select_text_button = QPushButton('&text path', self)
        self.select_save_button = QPushButton('&save path', self)

        self.select_date_button = QPushButton('&날짜 선택', self)
        self.run_button = QPushButton('&실 행', self)

        # 버튼 사이즈
        self.select_base_button.setFixedWidth(100)
        self.select_text_button.setFixedWidth(100)
        self.select_save_button.setFixedWidth(100)

        self.select_date_button.setFixedWidth(100)
        self.run_button.setFixedWidth(100)

        # label 초기화
        self.basepath_label = QLabel("", self)
        self.textpath_label = QLabel("", self)
        self.savepath_label = QLabel("", self)

        self.date_edit = QLineEdit(self)
        self.date_edit.setAlignment(Qt.AlignCenter)
        self.date_edit.setFixedWidth(120)
        self.date_edit.setFixedHeight(30)

        self.product_label1 = QLabel("품 명", self)
        self.product_label2 = QLabel("품 명", self)

        self.product_price_label1 = QLabel("단 가", self)
        self.product_price_label2 = QLabel("단 가", self)

        self.product_edit1 = QLineEdit(self)
        self.product_edit1.setFixedWidth(120)
        self.product_edit1.setFixedHeight(25)

        self.product_edit2 = QLineEdit(self)
        self.product_edit2.setFixedWidth(120)
        self.product_edit2.setFixedHeight(25)

        self.product_price_edit1 = QLineEdit(self)
        self.product_price_edit1.setFixedWidth(120)
        self.product_price_edit1.setFixedHeight(25)

        self.product_price_edit2 = QLineEdit(self)
        self.product_price_edit2.setFixedWidth(120)
        self.product_price_edit2.setFixedHeight(25)

        # label, button layout
        hbox1 = QHBoxLayout()
        hbox1.addStretch(3)
        hbox1.addWidget(self.basepath_label)
        hbox1.addStretch(4)
        hbox1.addWidget(self.select_base_button)

        hbox2 = QHBoxLayout()
        hbox2.addStretch(3)
        hbox2.addWidget(self.textpath_label)
        hbox2.addStretch(4)
        hbox2.addWidget(self.select_text_button)

        hbox3 = QHBoxLayout()
        hbox3.addStretch(3)
        hbox3.addWidget(self.savepath_label)
        hbox3.addStretch(4)
        hbox3.addWidget(self.select_save_button)

        hbox4 = QHBoxLayout()
        hbox4.addStretch(3)
        hbox4.addWidget(self.select_date_button)

        hbox5 = QHBoxLayout()
        hbox5.addStretch(3)
        hbox5.addWidget(self.run_button)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox1)
        vbox.addLayout(hbox2)
        vbox.addLayout(hbox3)
        vbox.addLayout(hbox4)
        vbox.addLayout(hbox5)
        vbox.addStretch(2)
        self.setLayout(vbox)

        # width, height
        self.date_edit.move(470,20)

        self.product_label1.move(22, 163)
        self.product_edit1.move(60, 165)

        self.product_label2.move(22, 203)
        self.product_edit2.move(60, 205)

        self.product_price_label1.move(190, 163)
        self.product_price_edit1.move(225, 163)

        self.product_price_label2.move(190, 205)
        self.product_price_edit2.move(225, 205)

        # 버튼 이벤트 초기화
        self.select_base_button.clicked.connect(lambda: self.open_directory_dialog("base", self.basepath_label))
        self.select_text_button.clicked.connect(self.open_file_dialog)
        self.select_save_button.clicked.connect(lambda: self.open_directory_dialog("save", self.savepath_label))

        self.select_date_button.clicked.connect(self.select_date_event)
        self.run_button.clicked.connect(self.run)

        self.center_window_onscreen()
        self.show()

    def center_window_onscreen(self):
        frame_geometry = self.frameGeometry()
        screen_center_point = QDesktopWidget().availableGeometry().center()
        frame_geometry.moveCenter(screen_center_point)
        self.move(frame_geometry.topLeft())

    def open_directory_dialog(self, m_url: str, label: QLabel):
        folder_path = QFileDialog.getExistingDirectory(self)
        label.setText(folder_path)

        if m_url == "base":
            self.base_url = folder_path
        elif m_url == "save":
            self.save_url = folder_path

    def open_file_dialog(self):
        file_info = QFileDialog.getOpenFileName(self)
        self.text_url = file_info[0]
        self.textpath_label.setText(self.text_url)

    def select_date_event(self):
        self.calendar_window = CalendarWindow()
        self.calendar_window.send_date.connect(self.get_date)
        self.calendar_window.show()

    @pyqtSlot(str)
    def get_date(self, date):
        self.date_data = date
        self.date_edit.setText(self.date_data)
        self.calendar_window.deleteLater()

    def run(self):
        if not self.base_url:
            msg = QMessageBox(self)
            msg.setWindowTitle("Error")
            msg.setText("base파일 경로를 지정 해주세요")
            msg.exec()
            return

        if not self.text_url:
            msg = QMessageBox(self)
            msg.setWindowTitle("Error")
            msg.setText("text파일 경로를 지정 해주세요")
            msg.exec()
            return

        if not self.save_url:
            msg = QMessageBox(self)
            msg.setWindowTitle("Error")
            msg.setText("저장 경로를 지정 해주세요")
            msg.exec()
            return

        product_name1 = self.product_edit1.text()
        product_name2 = self.product_edit2.text()
        product_price1 = self.product_price_edit1.text()
        product_price2 = self.product_price_edit2.text()

        self.excelmanager.init_filepath(self.base_url,self.text_url,self.save_url)
        self.excelmanager.init_date(self.date_data)
        self.excelmanager.init_product(product_name1,product_price1,product_name2,product_price2)
        self.excelmanager.run()


if __name__ == '__main__':
   app = QApplication(sys.argv)
   ex = App()
   sys.exit(app.exec_())
