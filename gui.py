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

        self.excelmanager = ExcelManager()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('엑셀 자동화 프로그램')
        self.resize(500, 250)

        # 버튼 레이아웃
        self.select_base_button = QPushButton('&base path', self)
        self.select_text_button = QPushButton('&text path', self)
        self.select_save_button = QPushButton('&save path', self)

        self.select_product_button = QPushButton('&품명/수량', self)
        self.select_date_button = QPushButton('&날짜', self)
        self.run_button = QPushButton('&실행', self)

        # 버튼 사이즈
        self.select_base_button.setFixedWidth(100)
        self.select_text_button.setFixedWidth(100)
        self.select_save_button.setFixedWidth(100)

        self.select_product_button.setFixedWidth(130)
        self.select_date_button.setFixedWidth(130)
        self.run_button.setFixedWidth(130)

        # label 초기화
        self.basepath_label = QLabel("", self)
        self.textpath_label = QLabel("", self)
        self.savepath_label = QLabel("", self)

        self.date_edit = QLineEdit(self)
        self.date_edit.setAlignment(Qt.AlignCenter)
        self.date_edit.setFixedWidth(120)
        self.date_edit.setFixedHeight(30)


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

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox1)
        vbox.addLayout(hbox2)
        vbox.addLayout(hbox3)
        vbox.addStretch(2)
        self.setLayout(vbox)

        # width, height
        self.select_product_button.move(20, 200)
        self.select_date_button.move(170, 200)
        self.run_button.move(320, 200)
        self.date_edit.move(180,20)

        # 버튼 이벤트 초기화
        self.select_base_button.clicked.connect(lambda: self.open_directory_dialog("base", self.basepath_label))
        self.select_text_button.clicked.connect(lambda: self.open_directory_dialog("text", self.textpath_label))
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
        elif m_url == "text":
            self.text_url = folder_path
        elif m_url == "save":
            self.save_url = folder_path

    # def open_file_dialog(self, label: QLabel):
    #     file_name = QFileDialog.getOpenFileName(self)
    #     url = file_name[0]
    #     self.base_url = url
    #     label.setText(url)

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

        self.excelmanager.init_filepath(self.base_url,self.text_url,self.save_url)
        self.excelmanager.init_date(self.date_data)
        self.excelmanager.run()


if __name__ == '__main__':
   app = QApplication(sys.argv)
   ex = App()
   sys.exit(app.exec_())
