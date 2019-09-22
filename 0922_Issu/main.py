# -*- coding: utf-8 -*-
from urllib.parse import quote_plus, urlencode
from urllib.request import urlopen, Request
from PySide2.QtWidgets import QApplication, QMainWindow
from PySide2.QtCore import Qt
from PySide2.QtWidgets import QTableWidgetItem
from ui_Issu import Ui_Main_Issu
import openpyxl
import json
import sys
import time


class Issu(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.ui = Ui_Main_Issu()
        self.ui.setupUi(self)

        self.ui.button_search.setShortcut('Return')
        self.ui.button_insert.setShortcut('Ctrl+Return')

        # declaration class variable
        self.auth_key = '302445398d5248c9b30d8581061a79b9'
        self.dic_location = {'서울특별시교육청': 'B10', '부산광역시교육청': 'C10', '대구광역시교육청': 'D10', '인천광역시교육청': 'E10',
                             '광주광역시교육청': 'F10', '대전광역시교육청': 'G10', '울산광역시교육청': 'H10', '세종특별자치시교육청': 'I10',
                             '경기도교육청': 'J10', '강원도교육청': 'K10', '충청북도교육청': 'M10', '충청남도교육청': 'N10', '전라북도교육청': 'P10',
                             '전라남도교육청': 'Q10', '경상북도교육청': 'R10', '경상남도교육청': 'S10', '제주도교육청': 'T10'}

    def reset(self):
        current_row_cnt = self.ui.table_output.rowCount()

        for i in range(1, current_row_cnt + 1):
            self.ui.table_output.removeRow(current_row_cnt - i)

    def output(self):
        current_row_cnt = self.ui.table_output.rowCount()

        if current_row_cnt is 0:
            return -1

        wb = openpyxl.Workbook()
        sheet = wb.active

        dic_cell = {'번호': ['A', 'A2', 0],
                    '학교명': ['B', 'B2', 1],
                    '관할지역청': ['C', 'C2', 2],
                    '공/사립': ['D', 'D2', 3],
                    '우편번호': ['E', 'E2', 4],
                    '도로명주소': ['F', 'F2', 5]}

        cell_list = list(dic_cell.keys())
        # print(cell_list)

        for i in cell_list:
            sheet[dic_cell[i][1]] = i

        for i in range(0, current_row_cnt):
            for j in cell_list:
                # print(i, j)
                cell = dic_cell[j][0] + str(3 + i)
                sheet[cell] = self.ui.table_output.item(i, dic_cell[j][2]).text()

        dims = {}
        dic_alphabet = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F'}

        for row in sheet.rows:
            for cell in row:
                if cell.value:
                    dims[dic_alphabet[cell.column]] = max((dims.get(cell.column, 0), len(str(cell.value))))

        # print(dims)

        for col, value in dims.items():
            # print(type(value))
            sheet.column_dimensions[col].width = int(value * 4)

        now = time.localtime()

        file_name = "학교주소명단_" + str(now.tm_mday) + str(now.tm_hour) + str(now.tm_min) + '.xlsx'
        wb.save(file_name)

    def delete(self):
        del_item = self.ui.table_output.selectedItems()
        # print(del_item)
        try:
            del_row = self.ui.table_output.row(del_item[0])
        except IndexError:
            return -1

        self.ui.table_output.removeRow(del_row)

        current_row_cnt = self.ui.table_output.rowCount()

        # print("Current : " + str(current_row_cnt))

        for i in range(0, current_row_cnt):
            self.ui.table_output.item(i, 0).setText(str(i + 1))

    def insert(self):
        selected_row = self.ui.table_result.selectedItems()

        if len(selected_row) is 0:
            return -1
        # print(selected_row)

        current_row_cnt = self.ui.table_output.rowCount()
        self.ui.table_output.insertRow(current_row_cnt)

        for i in range(0, 6):
            self.ui.table_output.setItem(current_row_cnt, i, selected_row[i].clone())
            # print(selected_row[i])
            # print(self.ui.table_output.item(current_row_cnt, i))

        for i in range(0, current_row_cnt + 1):
            self.ui.table_output.item(current_row_cnt, 0).setText(str(current_row_cnt + 1))

    def get_juso(self, location_code, sc_name):
        dic_result = {'번호': [], '학교명': [], '관할지역청': [], '공사립': [], '우편번호': [], '도로명주소': []}

        url = 'http://open.neis.go.kr/hub/schoolInfo'
        queryparams = '?' + urlencode(
            {quote_plus('KEY'): self.auth_key, quote_plus('Type'): 'json', quote_plus('pIndex'): '1', quote_plus('pSize'): '100',
             quote_plus('ATPT_OFCDC_SC_CODE'): location_code, quote_plus('SCHUL_NM'): sc_name})

        req = Request(url + queryparams)
        req.get_method = lambda: 'GET'
        response_body = urlopen(req).read()

        root_json = json.loads(response_body)

        # print(root_json)

        try:
            result_code = root_json['RESULT']

            self.ui.table_result.clearContents()
            self.ui.info_result.setText('검색 결과 없음')

            return -1
        except KeyError:
            pass

        result_cnt = root_json['schoolInfo'][0]['head'][0]['list_total_count']

        data_body = root_json['schoolInfo'][1]['row']

        # print(data_body)

        for i in range(0, result_cnt):
            if i > 99:
                result_cnt = 99
                break
            dic_result['번호'].append(i + 1)
            dic_result['학교명'].append(data_body[i]['SCHUL_NM'])
            dic_result['관할지역청'].append(data_body[i]['JU_ORG_NM'])
            dic_result['공사립'].append(data_body[i]['FOND_SC_NM'])
            dic_result['우편번호'].append(data_body[i]['ORG_RDNZC'])
            dic_result['도로명주소'].append(data_body[i]['ORG_RDNMA'])

        # print(dic_result)

        str_result = '검색 결과 : ' + str(result_cnt) + '개'

        self.ui.info_result.setText(str_result)

        row_cnt = self.ui.table_result.rowCount()
        diff_row_cnt = result_cnt - row_cnt

        if diff_row_cnt < 0:
            # print(">>")
            for i in range(1, abs(diff_row_cnt) + 1):
                self.ui.table_result.removeRow(row_cnt - i)
                # print(i)
        else:
            for i in range(0, abs(diff_row_cnt)):
                self.ui.table_result.insertRow(row_cnt + i)

        for i in range(0, self.ui.table_result.rowCount()):
            for j in range(0, self.ui.table_result.columnCount()):
                item = QTableWidgetItem()
                item.setFlags(item.flags() & ~ Qt.ItemIsEditable)
                self.ui.table_result.setItem(i, j, item)

        for i in range(0, result_cnt):
            self.ui.table_result.item(i, 0).setText(str(i + 1))
            self.ui.table_result.item(i, 1).setText(dic_result['학교명'][i])
            self.ui.table_result.item(i, 2).setText(dic_result['관할지역청'][i])
            self.ui.table_result.item(i, 3).setText(dic_result['공사립'][i])
            self.ui.table_result.item(i, 4).setText(dic_result['우편번호'][i])
            self.ui.table_result.item(i, 5).setText(dic_result['도로명주소'][i])

    def search(self):
        location = self.ui.combo_location.currentText()
        location_code = self.dic_location[location]
        # print(self.location_code)

        scname = self.ui.input_scname.text()
        # print(self.scname)

        if len(scname) == 0:
            return -1

        self.get_juso(location_code, scname)


app = QApplication([])
app.setStyle('Fusion')

Issu = Issu()
Issu.show()

sys.exit(app.exec_())