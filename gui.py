import sys

from PyQt5.QtWidgets import QApplication, QWidget, QDesktopWidget, QFileDialog, QPushButton, QVBoxLayout, QHBoxLayout, \
    QMessageBox, QLabel, QCalendarWidget, QLineEdit


class CalendarWindow(QWidget):

    # send_date = pySignal

    def __init__(self):
        super().__init__()

        self.date = None

        self.setWindowTitle("Calendar")
        self.initUI()

    def initUI(self):
        vbox = QVBoxLayout()
        self.calendar = QCalendarWidget(self)
        vbox.addWidget(self.calendar)
        self.setLayout(vbox)






class App(QWidget):

    def __init__(self):
        super().__init__()

        self.base_url = None
        self.text_url = None
        self.save_url = None

        self.date_data = None

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

        self.center_window_onscreen()
        self.show()

    def center_window_onscreen(self):
        frameGeometry = self.frameGeometry()
        screenCenterPoint = QDesktopWidget().availableGeometry().center()
        frameGeometry.moveCenter(screenCenterPoint)
        self.move(frameGeometry.topLeft())

    def open_directory_dialog(self, m_url: str, label: QLabel):

        file_name = QFileDialog.getOpenFileName(self)
        url = file_name[0]
        label.setText(url)

        if m_url == "base":
            self.base_url = url
        elif m_url == "text":
            self.text_url = url
        elif m_url == "save":
            self.save_url = url

    def select_date_event(self):
        self.calendarWindow = CalendarWindow()
        self.calendarWindow.show()




if __name__ == '__main__':
   app = QApplication(sys.argv)
   ex = App()
   sys.exit(app.exec_())