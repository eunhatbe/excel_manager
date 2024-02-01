
from datetime import datetime

# openpyxl 2.6.1
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image


class ExcelManager():
    def __init__(self):
        self.name_list = []
        self.number_list1 = []
        self.number_list2 = []

        self.base_url = None
        self.text_url = None
        self.save_url = None

        self.date = None

    def init_filepath(self, base_url, text_url, save_url):
        self.base_url = base_url
        self.text_url = text_url
        self.save_url = save_url

    def _check_path_exists(self):
        if not self.base_url and self.text_url and self.save_url:
            return False
        return True

    def _read_textfile(self):

        if not self._check_path_exists():
            return

        self.image = Image(f'{self.base_url}/stamp.png')

        try:
            with open(f"{self.text_url}/name.txt", 'r', encoding="utf-8") as f:
                file = f.read()
                self.name_list = file.split("\n")
        except Exception as e:
            print(f"Error : {e}")

        try:
            with open(f"{self.text_url}/number1.txt", 'r') as f:
                file = f.read()
                self.number_list1 = file.split("\n")
        except Exception as e:
            print(f"Error : {e}")

        try:
            with open(f"{self.text_url}/number2.txt", 'r') as f:
                file = f.read()
                self.number_list2 = file.split("\n")
        except Exception as e:
            print(f"Error : {e}")

    def init_date(self, date):
        # 만약 날짜가 지정되지 않았다면 오늘 날짜로 초기화
        if not date:
            self.date = datetime.today().strftime("%Y-%m-%d")
            return

        self.date = date

    def run(self):

        self._read_textfile()

        # 워크북 생성
        wb = openpyxl.load_workbook(f"{self.base_url}/base.xlsx")

        for i in range(len(self.name_list)):
            ws = wb.active

            # 품명
            ws['B5'] = self.name_list[i]

            # 수량
            ws['G14'] = self.number_list1[i]
            ws['G15'] = self.number_list2[i]

            # 날짜
            ws['B3'] = self.date
            ws['B14'] = self.date[5:]
            ws['B15'] = self.date[5:]

            ws['B5'].alignment = Alignment(horizontal='center', vertical='center')
            ws['G14'].alignment = Alignment(horizontal='center', vertical='center')
            ws['G15'].alignment = Alignment(horizontal='center', vertical='center')

            target_cell_j14 = wb['in']['j14']  # 공급가액
            target_cell_j15 = wb['in']['j15']  # 공급가액
            target_cell_i26 = wb['in']['i26']  # 합계
            target_cell_b11 = wb['in']['b11']  # 합계

            target_cell_j14.value = '=G14*H14'
            target_cell_j15.value = '=G15*H15'
            target_cell_i26.value = '=SUM(j14:j15)'
            target_cell_b11.value = '=SUM(j14:j15)'

            wb.save(f'{self.save_url}/{self.name_list[i]}{i}.xlsx')
