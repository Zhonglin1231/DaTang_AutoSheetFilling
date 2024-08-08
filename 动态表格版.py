import os
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, QLabel, QMessageBox, QScrollArea, QGridLayout, QHBoxLayout, QLineEdit
from PyQt5.QtCore import Qt
import pandas as pd
from openpyxl import load_workbook

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.paths = {}
        self.labels = {}

    def initUI(self):
        self.setGeometry(700, 400, 1000, 500)
        self.setWindowTitle('自动填表生成器')

        # 创建主布局
        main_layout = QHBoxLayout()

        # 左侧文件选择布局
        left_layout = QVBoxLayout()

        # 添加年份输入框和按钮
        year_selection_layout = QHBoxLayout()
        self.year_range_label = QLabel("年份范围:", self)
        year_selection_layout.addWidget(self.year_range_label)
        
        self.start_year_input = QLineEdit(self)
        self.start_year_input.setPlaceholderText("开始年份")
        self.start_year_input.setMaximumWidth(100)  # Set maximum width for the input field
        year_selection_layout.addWidget(self.start_year_input)
        
        self.end_year_input = QLineEdit(self)
        self.end_year_input.setPlaceholderText("结束年份")
        self.end_year_input.setMaximumWidth(100)  # Set maximum width for the input field
        year_selection_layout.addWidget(self.end_year_input)
        
        self.year_range_button = QPushButton("生成文件选择按钮", self)
        self.year_range_button.clicked.connect(self.generate_file_buttons)
        year_selection_layout.addWidget(self.year_range_button)

        left_layout.addLayout(year_selection_layout)

        # 创建文件选择按钮布局
        self.file_selection_layout = QVBoxLayout()
        left_layout.addLayout(self.file_selection_layout)

        self.file_selection_layout.addWidget(QLabel('--------↓--------', self), alignment=Qt.AlignCenter)

        # 创建一个水平滑动框来显示所有路径
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget(self.scroll_area)
        self.scroll_layout =  QGridLayout(self.scroll_content)

        self.scroll_content.setLayout(self.scroll_layout)
        self.scroll_area.setWidget(self.scroll_content)
        left_layout.addWidget(self.scroll_area)

        main_layout.addLayout(left_layout)

        # 右侧布局
        right_layout = QVBoxLayout()

        # 添加固定第一行的布局
        self.fixed_header_layout = QHBoxLayout()
        self.fixed_header_layout.addWidget(QLabel("<b>指标</b>", self))
        self.fixed_header_layout.addWidget(QLabel(f"<b>1年数据</b>", self))
        self.fixed_header_layout.addWidget(QLabel(f"<b>2年数据</b>", self))
        self.fixed_header_layout.addWidget(QLabel(f"<b>3年数据</b>", self))

        right_layout.addLayout(self.fixed_header_layout)

        # 添加滑动框
        self.scroll = QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        self.scroll_content_right = QWidget(self.scroll)
        self.scroll_layout_right = QGridLayout(self.scroll_content_right)
        self.scroll.setWidget(self.scroll_content_right)
        right_layout.addWidget(self.scroll)

        main_layout.addLayout(right_layout)

        self.setLayout(main_layout)
        self.show()

    def generate_file_buttons(self):
        self.paths = {}
        self.labels = {}

        # 清空布局
        self.clear_layout(self.file_selection_layout)
        self.clear_layout(self.fixed_header_layout)
        self.fixed_header_layout.addWidget(QLabel(f"<b>指标</b>", self))
        self.fixed_header_layout.addWidget(QLabel(f"<b>{int(self.end_year_input.text())}年数据</b>", self))
        self.fixed_header_layout.addWidget(QLabel(f"<b>{int(self.end_year_input.text())-1}年数据</b>", self))
        self.fixed_header_layout.addWidget(QLabel(f"<b>{int(self.end_year_input.text())-2}年数据</b>", self))
        self.clear_layout(self.scroll_layout)

        # 获取年份范围
        start_year = int(self.start_year_input.text())
        end_year = int(self.end_year_input.text())

        # 生成文件选择按钮
        for year in range(start_year, end_year + 1):
            path_key = f"{year}年报表路径"
            self.paths[path_key] = ""
            self.add_button(self.file_selection_layout, f'选择{year}年年报', self.select_file, path_key)
            label = QLabel(f'{year}年年报路径: <font color="red">未选择</font>', self)
            self.labels[path_key] = label
            self.scroll_layout.addWidget(label)

        self.paths["back_data_path"] = ""
        self.add_button(self.file_selection_layout, '选择数据底稿', self.select_file, "back_data_path")
        label = QLabel(f'数据底稿路径: <font color="red">未选择</font>', self)
        self.labels["back_data_path"] = label
        self.scroll_layout.addWidget(self.labels["back_data_path"])

        self.paths["target_path"] = ""
        self.add_button(self.file_selection_layout, '评级文件路径', self.select_file, "target_path")
        label = QLabel(f'填入文件路径: <font color="red">未选择</font>', self)
        self.labels["target_path"] = label
        self.scroll_layout.addWidget(self.labels["target_path"])

        self.add_button(self.file_selection_layout, '开始处理', self.startProcess)


    def add_button(self, layout, text, handler, *args):
        button = QPushButton(text, self)
        button.clicked.connect(lambda: handler(*args))
        layout.addWidget(button)


    def select_file(self, path_key):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择文件')
        if file_path:
            self.paths[path_key] = file_path
            self.labels[path_key].setText(f'{path_key}: <font color="green">{file_path}</font>')
            self.labels[path_key].setToolTip(file_path)


    def startProcess(self):
        if any(not path for path in self.paths.values()):
            QMessageBox.warning(self, '警告', '请先选择所有文件')
            return

        if not all(path.endswith(('.xlsx', '.xls')) for path in self.paths.values()):
            QMessageBox.warning(self, '警告', '请选择正确的表格文件文件(.xlsx)/(.xls)')
            return

        report_3, report_2, report_1, back_data = self.read_files()
        if report_3 is None or report_2 is None or report_1 is None or back_data is None:
            return

        data_3, data_2, data_1 = self.extract_data(report_3), self.extract_data(report_2), self.extract_data(report_1)
        if data_3 is None or data_2 is None or data_1 is None:
            QMessageBox.warning(self, '警告', '年报文件中表格数量不正确')
            return
        

        final_data_3, detailed_data_3 = self.calculate_data(data_3, back_data, int(self.start_year_input.text()))
        final_data_2, detailed_data_2 = self.calculate_data(data_2, back_data, int(self.start_year_input.text())+1)
        final_data_1, detailed_data_1 = self.calculate_data(data_1, back_data, int(self.start_year_input.text())+2)

        # 修正数据----------------------------------
            # 资本回报率取均值
        if (final_data_2["资本"]!=0) and (final_data_3["所有者权益合计"]!=0):
            final_data_2["资本回报率"] = final_data_2["EBIT"] / ((final_data_2["资本"] + final_data_3["资本"])/2)

        if (final_data_1["资本"]!=0) and (final_data_2["所有者权益合计"]!=0):
            final_data_1["资本回报率"] = final_data_1["EBIT"] / ((final_data_1["资本"] + final_data_2["资本"])/2)
        # ----------------------------------------

        total_data_3 = {**final_data_3, **detailed_data_3}
        total_data_2 = {**final_data_2, **detailed_data_2}
        total_data_1 = {**final_data_1, **detailed_data_1}

        # 清空滑动框
        self.clear_scroll_area()

        # 显示数据到滑动框
        self.display_data(total_data_1, total_data_2, total_data_3)

        self.write_to_excel(final_data_3, final_data_2, final_data_1)

    def read_files(self):

        for i in self.paths.values():
            if i == "":
                QMessageBox.warning(self, '警告', '请先选择所有文件')
                return None, None, None, None

        # # print(list(self.paths.values())[0])
        try:
            report_3 = pd.read_excel(list(self.paths.values())[0], sheet_name=None)
        except:
            QMessageBox.warning(self, '警告', '第一年年报文件读取失败')
            return None, None, None
        try:
            report_2 = pd.read_excel(list(self.paths.values())[1], sheet_name=None)
        except:
            QMessageBox.warning(self, '警告', '第二年年报文件读取失败')
            return None, None, None
        try:
            report_1 = pd.read_excel(list(self.paths.values())[2], sheet_name=None)
        except:
            QMessageBox.warning(self, '警告', '第三年年报文件读取失败')
            return None, None, None
        try:
            back_data = pd.read_excel(self.paths["back_data_path"])
        except:
            QMessageBox.warning(self, '警告', '数据底稿文件读取失败')
            return None, None, None
        return report_3, report_2, report_1, back_data

    def extract_data(self, report):
        sheets = {name: sheet for name, sheet in report.items() if "资产负债表" in name or "利润表" in name or "现金流量表" in name}
        sheets = dict(list(sheets.items())[:4]) # 只取前四个表 
        if len(sheets) != 4:
            # 未找到四个表
            QMessageBox.warning(self, '警告', '年报文件中表格数量不正确, 应包含资产负债表(资产负债表（续）)、利润表、现金流量表')
            return exit()
        
        for i in sheets.keys():
            # # print(i) 
            if '资产负债表（续）' in i or '资产负债表(续)' in i:
                # # print("找到续表")
                xubiao = True
                break
            else:
                xubiao = False
        if xubiao:
            sheets_list = list(sheets.values())
            bal_sheet = sheets_list[0]
            bal_sheet_con = sheets_list[1]
            profit_sheet = sheets_list[2]
            cash_sheet = sheets_list[3]
            # 汇总到一个sheets中
            sheets = {
                "资产负债表": bal_sheet,
                "资产负债表_con": bal_sheet_con,
                "利润表": profit_sheet,
                "现金流量表": cash_sheet
            }
            # # print(sheets.keys()) 
            return sheets
        else:
            # # print("未找到续表")
            sheets_list = list(sheets.values())
            bal_sheet = sheets_list[0]
            profit_sheet = sheets_list[1]
            cash_sheet = sheets_list[2]
            sheets = {
                "资产负债表": bal_sheet,
                "资产负债表_con": bal_sheet,
                "利润表": profit_sheet,
                "现金流量表": cash_sheet
            }
            # # print(sheets.keys())
            return sheets

    def calculate_data(self, sheets, back_data, year):
        data = {}
        data_set = self.extract_values(sheets, back_data, year)
        # # print(year, "---------------")
        # # print(data_set)

        data["EBITDA"] = EBITDA(
            data_set["营业利润"],
            data_set["财务费用"],
            data_set["折旧费"],
            data_set["公允价值变动"],
            data_set["投资收益"],
            data_set["取得投资收益收到的现金"],
            data_set["政府补助"],
            data_set["经营租赁费用调整"],
            data_set["资本化开发成本"],
            data_set["勘探费用"]
        )

        data["EBIT"] = EBIT(
            data_set["营业利润"],
            data_set["财务费用"],
            data_set["利息收入"],
            data_set["公允价值变动"],
            data_set["投资收益"],
            data_set["对联营企业和合营企业的投资收益"],
            data_set["政府补助"],
            data_set["经营租赁的利息调整"]
        )

        data["自由运营现金流(FOCF)"] = FOCF(
            data_set["经营活动产生的现金流量净额"],
            data_set["购建固定资产无形资产和其他长期资产支付的现金"],
            data_set["取得投资收益收到的现金"],
            data_set["收到其他与投资活动有关的现金附注利息部分"],
            data_set["分配股利利润或偿付利息支付的现金"],
            data_set["对所有者或股东的分配"],
            data_set["经营租赁折旧调整"],
            data_set["资本化开发成本"]
        )

        data["经营活动产生的现金(FFO)"] = FFO(
            data["EBITDA"],
            data_set["利息费用"],
            data_set["利息收入"],
            data_set["所得税费用"],
            data_set["经营租赁费用调整"],
            data_set["经营租赁折旧调整"],
            data_set["资本化利息"]
        )

        data["总负债"] = Total_debt(
            data_set["短期借款"],
            data_set["应付利息"],
            data_set["一年内到期的长期借款"],
            data_set["一年内到期的应付债券"],
            data_set["其它流动负债短期应付债券"],
            data_set["一年内应付融资租赁款"],
            data_set["长期借款"],
            data_set["应付债券"],
            data_set["长期应付融资租赁款"],
            data_set["重大合同及履行状况担保情况"],
            data_set["货币资金"],
            data_set["以公允价值计量且其变动计入当期损益的金融资产"],
            data_set["其他货币资金"],
            data_set["卖出回购金融资产款"],
            data_set["特定行业或公司现金盈余不做调整扣除的部分加回"],
            data_set["经营租赁调整"],
            data_set["永续债"]
        )

        data["资本"] = Capital(
            data_set["所有者权益合计"],
            data_set["短期借款"],
            data_set["应付利息"],
            data_set["一年内到期的长期借款"],
            data_set["一年内到期的应付债券"],
            data_set["其它流动负债短期应付债券"],
            data_set["一年内应付融资租赁款"],
            data_set["长期借款"],
            data_set["应付债券"],
            data_set["长期应付融资租赁款"],
            data_set["递延所得税负债"],
            data_set["重大合同及履行状况担保情况"],
            data_set["货币资金"],
            data_set["以公允价值计量且其变动计入当期损益的金融资产"],
            data_set["其他货币资金"],
            data_set["卖出回购金融资产款"],
            data_set["特定行业或公司现金盈余不做调整扣除的部分加回"],
            data_set["经营租赁调整"],
            data_set["永续债"]
        )

        data["EBITDA利润率"] = EBITDA_profit_rate(
            data["EBITDA"],
            data_set["营业收入"]
        )

        data["资本回报率"] = Capital_RR(
            data["EBIT"],
            data["资本"]
        )

        data["经营活动产生的资金/债务"] = Operating_cash_to_debt(
            data["经营活动产生的现金(FFO)"],
            data["总负债"]
        )

        data["债务/息税摊折前利润"] = debt_to_PBITA(
            data["总负债"],
            data["EBITDA"]
        )

        data["自由运营现金流/债务"] = FOCF_to_debt(
            data["自由运营现金流(FOCF)"],
            data["总负债"]
        )

        data["息税摊折前利润 / 利息支出"] = EBITDA_to_interest_expense(
            data["EBITDA"],
            data_set["利息费用"],
        )

        data["营业收入"] = data_set["营业收入"]

        data["总资产"] = data_set["总资产"]

        data["所有者权益合计"] = data_set["所有者权益合计"]

        return data, data_set

    def display_data(self, data_1, data_2, data_3):
        row = 0
        col = 0

        years_data = [("1年数据", data_1), ("2年数据", data_2), ("3年数据", data_3)]
        
        

        for year_label, data in years_data:
            for key, value in data.items():
                if year_label == "1年数据": #为了避免重复，只显示第一例地数据标签
                    self.scroll_layout_right.addWidget(QLabel(f'{key}:'), row, col)
                self.scroll_layout_right.addWidget(QLabel(str(round(value, 2))), row, col + 1)
                row += 1
            col += 1
            row = 0

    def extract_values(self, sheets, back_data, year):
        # 从底稿中找到对应的列数，列名就是年份
        year_col = 0
        for col in back_data.columns:
            if col == year:
                year_col
                break
            year_col += 1
        if year_col >= len(back_data.columns):
            QMessageBox.warning(self, '警告', f"数据底稿未找到{year}年的数据")

        if "资产负债表_con" not in sheets:
            sheets["资产负债表_con"] = sheets["资产负债表"] # 如果没有续表，就用资产负债表表代替, 这样可以减少代码修改

        data_set = {
            "营业利润": find_cell(self, sheets["利润表"], "营业利润"),
            "财务费用": find_cell(self, sheets["利润表"], "财务费用"),
            "折旧费": back_data.iloc[row_to_num(2), year_col],
            "公允价值变动": find_cell(self, sheets["利润表"], "公允价值变动"),
            "投资收益": find_cell(self, sheets["利润表"], "投资收益"),
            "取得投资收益收到的现金": find_cell(self, sheets["现金流量表"], "取得投资收益收到的现金"),
            "政府补助": find_cell(self, sheets["利润表"], "政府补助"),
            "经营租赁费用调整": back_data.iloc[row_to_num(3), year_col],
            "资本化开发成本": back_data.iloc[row_to_num(4), year_col],
            "勘探费用": back_data.iloc[row_to_num(5), year_col],
            "利息收入": find_cell(self, sheets["利润表"], "利息收入"),
            "对联营企业和合营企业的投资收益": find_cell(self, sheets["利润表"], "对联营企业和合营企业的投资收益"),
            "经营租赁的利息调整": back_data.iloc[row_to_num(6), year_col],
            "经营活动产生的现金流量净额": find_cell(self, sheets["现金流量表"], "经营活动产生的现金流量净额"),
            "购建固定资产无形资产和其他长期资产支付的现金": find_cell(self, sheets["现金流量表"], "购建固定资产"),
            "收到其他与投资活动有关的现金附注利息部分": back_data.iloc[row_to_num(7), year_col],
            "分配股利利润或偿付利息支付的现金": find_cell(self, sheets["现金流量表"], "分配股利"),
            "对所有者或股东的分配": back_data.iloc[row_to_num(8), year_col],
            "经营租赁折旧调整": back_data.iloc[row_to_num(9), year_col],
            "所得税费用": find_cell(self, sheets["利润表"], "所得税费用"),
            "资本化利息": back_data.iloc[row_to_num(10), year_col],
            "短期借款": find_cell(self, sheets["资产负债表_con"], "短期借款"),
            "应付利息": back_data.iloc[row_to_num(11), year_col],
            "一年内到期的长期借款": back_data.iloc[row_to_num(12), year_col],
            "一年内到期的应付债券": back_data.iloc[row_to_num(13), year_col],
            "其它流动负债短期应付债券": back_data.iloc[row_to_num(14), year_col],
            "一年内应付融资租赁款": back_data.iloc[row_to_num(15), year_col],
            "长期借款": find_cell(self, sheets["资产负债表_con"], "长期借款"),
            "应付债券": find_cell(self, sheets["资产负债表_con"], "应付债券"),
            "长期应付融资租赁款": back_data.iloc[row_to_num(16), year_col],
            "重大合同及履行状况担保情况": back_data.iloc[row_to_num(17), year_col],
            "货币资金": find_cell(self, sheets["资产负债表"], "货币资金"),
            "以公允价值计量且其变动计入当期损益的金融资产": find_cell(self, sheets["资产负债表"], "以公允价值计量且其变动计入当期损益的金融资产"),
            "其他货币资金": back_data.iloc[row_to_num(21), year_col],
            "卖出回购金融资产款": back_data.iloc[row_to_num(18), year_col],
            "特定行业或公司现金盈余不做调整扣除的部分加回": back_data.iloc[row_to_num(19), year_col],
            "经营租赁调整": back_data.iloc[row_to_num(20), year_col],
            "永续债": find_cell(self, sheets["资产负债表_con"], "永续债"),
            "所有者权益合计": find_cell(self, sheets["资产负债表_con"], "所有者权益（或股东权益）合计"),
            "递延所得税负债": find_cell(self, sheets["资产负债表_con"], "递延所得税负债"),
            "营业收入": find_cell(self, sheets["利润表"], "营业收入"),
            "总资产": find_cell(self, sheets["资产负债表"], "资  产  总  计"),
            "利息费用": find_cell(self, sheets["利润表"], "利息费用")
        }

           # 将所有空着的数据填充为0
        for i in data_set:
            if str(data_set[i]) == "nan":
                data_set[i] = 0
            elif type(data_set[i]) == str:
                QMessageBox.warning(self, '警告', f'数据{data_set[i]}不是数字, 请检查数据是否对应正确')

        return data_set

    def write_to_excel(self, data_3, data_2, data_1):
        try:
            report = load_workbook(self.paths["target_path"])
            sheet = report["Inputs"]
            # # print("Writing to Excel...")
            # 将数据写入Excel
                # 逐年分析            
            sheet["D46"], sheet["E46"], sheet["F46"] = data_1["EBITDA利润率"], data_2["EBITDA利润率"], data_3["EBITDA利润率"]
            sheet["D47"], sheet["E47"], sheet["F47"] = data_1["资本回报率"], data_2["资本回报率"], data_3["资本回报率"]
            sheet["D48"], sheet["E48"], sheet["F48"] = data_1["营业收入"], data_2["营业收入"], data_3["营业收入"]
            sheet["D49"], sheet["E49"], sheet["F49"] = data_1["总资产"], data_2["总资产"], data_3["总资产"]
            sheet["D54"], sheet["E54"], sheet["F54"] = data_1["经营活动产生的资金/债务"], data_2["经营活动产生的资金/债务"], data_3["经营活动产生的资金/债务"]
            sheet["D55"], sheet["E55"], sheet["F55"] = data_1["债务/息税摊折前利润"], data_2["债务/息税摊折前利润"], data_3["债务/息税摊折前利润"]
            sheet["D56"], sheet["E56"], sheet["F56"] = data_1["自由运营现金流/债务"], data_2["自由运营现金流/债务"], data_3["自由运营现金流/债务"]
            sheet["D57"], sheet["E57"], sheet["F57"] = data_1["息税摊折前利润 / 利息支出"], data_2["息税摊折前利润 / 利息支出"], data_3["息税摊折前利润 / 利息支出"]
            sheet["D58"], sheet["E58"], sheet["F58"] = data_1["经营活动产生的现金(FFO)"], data_2["经营活动产生的现金(FFO)"], data_3["经营活动产生的现金(FFO)"]
            sheet["D59"], sheet["E59"], sheet["F59"] = data_1["总负债"], data_2["总负债"], data_3["总负债"]
            sheet["D60"], sheet["E60"], sheet["F60"] = data_1["EBITDA"], data_2["EBITDA"], data_3["EBITDA"]
            sheet["D63"], sheet["E63"], sheet["F63"] = data_1["营业收入"], data_2["营业收入"], data_3["营业收入"]
            sheet["D64"], sheet["E64"], sheet["F64"] = data_1["总资产"], data_2["总资产"], data_3["总资产"]

            # print("赋值完成")

            # 检查特殊情况并设置为 "NM"
            for i in sheet["B46:G64"]:
                for j in i:
                    if str(j.value) == "inf" or str(j.value) == "-inf" or str(j.value) == "nan":
                        j.value = "NM"

            if data_3["经营活动产生的现金(FFO)"] < 0 and data_3["总负债"] < 0:
                sheet["F54"] = "NM"
            if data_3["总负债"] < 0 and data_3["EBITDA"] < 0:
                sheet["F55"] = "NM"
            if data_3["自由运营现金流(FOCF)"] < 0 and data_3["总负债"] < 0:
                sheet["F56"] = "NM"

            if data_2["经营活动产生的现金(FFO)"] < 0 and data_2["总负债"] < 0:
                sheet["E54"] = "NM"
            if data_2["总负债"] < 0 and data_2["EBITDA"] < 0:
                sheet["E55"] = "NM"
            if data_2["自由运营现金流(FOCF)"] < 0 and data_2["总负债"] < 0:
                sheet["E56"] = "NM"

            if data_1["经营活动产生的现金(FFO)"] < 0 and data_1["总负债"] < 0:
                sheet["D54"], sheet["B54"] = "NM", "NM"
            if data_1["总负债"] < 0 and data_1["EBITDA"] < 0:
                sheet["D55"], sheet["B55"] = "NM", "NM"
            if data_1["自由运营现金流(FOCF)"] < 0 and data_1["总负债"] < 0:
                sheet["D56"], sheet["B56"] = "NM", "NM"
            # print("特殊情况处理完成")

            report.save(self.paths["target_path"])

            QMessageBox.information(self, "提示", "数据填充完成", QMessageBox.Yes)
            os.startfile(self.paths["target_path"])
        except Exception as e:
            QMessageBox.warning(self, '警告', f'文件写入失败: {str(e)}, 请确保填入文件没被打开')


    def clear_scroll_area(self):
        while self.scroll_layout_right.count():
            child = self.scroll_layout_right.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def clear_layout(self, layout):
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
            elif child.layout():
                self.clear_layout(child.layout())


# 计算公式
def EBITDA(营业利润, 财务费用, 折旧费, 公允价值变动, 投资收益, 取得投资收益收到的现金, 政府补助, 经营租赁费用调整, 资本化开发成本, 勘探费用):
    return 营业利润 + 财务费用 + 折旧费 - 公允价值变动 - 投资收益 + 取得投资收益收到的现金 + 政府补助 + 经营租赁费用调整 - 资本化开发成本 + 勘探费用

def EBIT(营业利润, 财务费用, 利息收入, 公允价值变动, 投资收益, 对联营企业和合营企业的投资收益, 政府补助, 经营租赁的利息调整):
    return 营业利润 + 财务费用 + 利息收入 - 公允价值变动 - 投资收益 + 对联营企业和合营企业的投资收益 + 政府补助 + 经营租赁的利息调整

def FOCF(经营活动产生的现金流量净额, 购建固定资产无形资产和其他长期资产支付的现金, 取得投资收益收到的现金, 收到其他与投资活动有关的现金附注利息部分, 分配股利利润或偿付利息支付的现金, 对所有者或股东的分配, 经营租赁折旧调整, 资本化开发成本):
    return 经营活动产生的现金流量净额 - 购建固定资产无形资产和其他长期资产支付的现金 + 取得投资收益收到的现金 + 收到其他与投资活动有关的现金附注利息部分 - 分配股利利润或偿付利息支付的现金 - 对所有者或股东的分配 + 经营租赁折旧调整 - 资本化开发成本

def FFO(EBITDA, 利息费用, 利息收入, 所得税费用, 经营租赁费用调整, 经营租赁折旧调整, 资本化利息):
    return EBITDA - 利息费用 + 利息收入 - 所得税费用 - 经营租赁费用调整 + 经营租赁折旧调整 - 资本化利息

def Total_debt(短期借款, 应付利息, 一年内到期的长期借款, 一年内到期的应付债券, 其它流动负债短期应付债券, 一年内应付融资租赁款, 长期借款, 应付债券, 长期应付融资租赁款, 重大合同及履行状况担保情况, 货币资金, 以公允价值计量且其变动计入当期损益的金融资产, 其他货币资金, 卖出回购金融资产款, 特定行业或公司现金盈余不做调整扣除的部分加回, 经营租赁调整, 永续债):
    return 短期借款 + 应付利息 + 一年内到期的长期借款 + 一年内到期的应付债券 + 其它流动负债短期应付债券 + 一年内应付融资租赁款 + 长期借款 + 应付债券 + 长期应付融资租赁款 + 重大合同及履行状况担保情况 - (货币资金 + 以公允价值计量且其变动计入当期损益的金融资产 - 其他货币资金) * 0.75 + 卖出回购金融资产款 + 特定行业或公司现金盈余不做调整扣除的部分加回 + 经营租赁调整 + 永续债

def Capital(所有者权益合计, 短期借款, 应付利息, 一年内到期的长期借款, 一年内到期的应付债券, 其它流动负债短期应付债券, 一年内应付融资租赁款, 长期借款, 应付债券, 长期应付融资租赁款, 递延所得税负债, 重大合同及履行状况担保情况, 货币资金, 以公允价值计量且其变动计入当期损益的金融资产, 其他货币资金, 卖出回购金融资产款, 特定行业或公司现金盈余不做调整扣除的部分加回, 经营租赁调整, 永续债):
    return 所有者权益合计 + 短期借款 + 应付利息 + 一年内到期的长期借款 + 一年内到期的应付债券 + 其它流动负债短期应付债券 + 一年内应付融资租赁款 + 长期借款 + 应付债券 + 长期应付融资租赁款 + 递延所得税负债 + 重大合同及履行状况担保情况 - (货币资金 + 以公允价值计量且其变动计入当期损益的金融资产 - 其他货币资金) * 0.75 + 卖出回购金融资产款 + 特定行业或公司现金盈余不做调整扣除的部分加回 + 经营租赁调整 + 永续债

def EBITDA_profit_rate(EBITDA, 营业收入):
    return EBITDA / 营业收入

def Capital_RR(EBIT, 资本):
    return EBIT / 资本

def Operating_cash_to_debt(FFO, 总负债):
    return FFO / 总负债

def debt_to_PBITA(总负债, EBITDA):
    return 总负债 / EBITDA

def FOCF_to_debt(FOCF, 总负债):
    return FOCF / 总负债

def EBITDA_to_interest_expense(EBITDA, 利息费用):
    return EBITDA / 利息费用

def col_to_num(col_str):
    num = 0
    for i, c in enumerate(reversed(col_str)):
        num += (ord(c) - ord('A') + 1) * (26 ** i)
    return num - 1

def row_to_num(num):
    return int(num) - 2


def find_cell(self, sheet, keyword):
    # 定位行 --------------------------------
    for row in range(sheet.shape[0]):
        for col in range(sheet.shape[1]):
            # # print(keyword, sheet.iloc[row, col])
            if keyword in str(sheet.iloc[row, col]):
                # 特殊情况：利息收入有重复的名字，这里属于是利息费用中的一部分
                if keyword == "利息收入" and  "利息费用" not in sheet.iloc[row-1, col]:
                    continue
                final_row = row
                semi_col = col
                # print("final_row:", final_row, "semi_col:", semi_col)
                break
        if keyword in str(sheet.iloc[row, col]):
            break

    # 检查是否找到
    if "final_row" not in locals():
        QMessageBox.warning(self, '警告', f"未找到{keyword}数据")
        return None

    # 定位列 是在行名之后的列 列名为 "本期金额" / "期末余额" / "年末余额" --------------------------
    current_list = ["本期金额", "期末余额", "年末余额", "本年金额"]


    for row in range(sheet.shape[0]):
        for col in range(semi_col, sheet.shape[1]):
            # # print("row:", row, "col:", col, "value:", sheet.iloc[row, col])
            if str(sheet.iat[row, col])[:4] in current_list:
                final_col = col
                return sheet.iloc[final_row, final_col]
            
    # 如果有表头
    col = 0
    if semi_col == 0:
        for col_name in sheet.columns:
            # print("col_name:", col_name)
            if str(col_name)[:4] in current_list:
                final_col = col
                return sheet.iloc[final_row, final_col]
            col += 1
    else:
        count_appear_time = 0
        for col_name in sheet.columns:
            # print("col_name:", col_name)
            if str(col_name)[:4] in current_list:
                count_appear_time += 1
                if count_appear_time > 1:
                    final_col = col
                    return sheet.iloc[final_row, final_col]
            col += 1

    return None, None


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    sys.exit(app.exec_())
