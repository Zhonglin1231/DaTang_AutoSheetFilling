import os
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, QLabel, QMessageBox, QScrollArea, QGridLayout, QHBoxLayout
from PyQt5.QtCore import Qt
import pandas as pd
from openpyxl import load_workbook

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.paths = {
            "report_2021_path": "",
            "report_2022_path": "",
            "report_2023_path": "",
            "back_data_path": "",
            "target_path": ""
        }

    def initUI(self):
        self.setGeometry(700, 400, 1000, 500)
        self.setWindowTitle('自动填表生成器')

        # 创建主布局
        main_layout = QHBoxLayout()

        # 左侧文件选择布局
        left_layout = QVBoxLayout()

        # 创建文件选择按钮布局
        file_selection_layout = QVBoxLayout()
        self.add_button(file_selection_layout, '选择2021年年报', self.select_file, "report_2021_path")
        self.add_button(file_selection_layout, '选择2022年年报', self.select_file, "report_2022_path")
        self.add_button(file_selection_layout, '选择2023年年报', self.select_file, "report_2023_path")
        self.add_button(file_selection_layout, '选择数据底稿', self.select_file, "back_data_path")
        self.add_button(file_selection_layout, '评级文件路径', self.select_file, "target_path")

        file_selection_layout.addWidget(QLabel('--------↓--------', self), alignment=Qt.AlignCenter)
        self.add_button(file_selection_layout, '开始处理', self.startProcess)

        # 创建一个水平滑动框来显示所有路径
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget(self.scroll_area)
        self.scroll_layout = QVBoxLayout(self.scroll_content)

        self.labels = {
            "report_2021_path": QLabel('2021年年报路径: <font color="red">未选择</font>', self),
            "report_2022_path": QLabel('2022年年报路径: <font color="red">未选择</font>', self),
            "report_2023_path": QLabel('2023年年报路径: <font color="red">未选择</font>', self),
            "back_data_path": QLabel('数据底稿路径: <font color="red">未选择</font>', self),
            "target_path": QLabel('评级文件路径: <font color="red">未选择</font>', self)
        }

        for label in self.labels.values():
            self.scroll_layout.addWidget(label)

        self.scroll_content.setLayout(self.scroll_layout)
        self.scroll_area.setWidget(self.scroll_content)
        file_selection_layout.addWidget(self.scroll_area)

        left_layout.addLayout(file_selection_layout)

        main_layout.addLayout(left_layout)

        # 右侧布局
        right_layout = QVBoxLayout()

        # 添加固定第一行的布局
        fixed_header_layout = QHBoxLayout()
        fixed_header_layout.addWidget(QLabel("<b>指标</b>", self))
        fixed_header_layout.addWidget(QLabel("<b>2023年数据</b>", self))
        fixed_header_layout.addWidget(QLabel("<b>2022年数据</b>", self))
        fixed_header_layout.addWidget(QLabel("<b>2021年数据</b>", self))

        right_layout.addLayout(fixed_header_layout)

        # 添加滑动框
        self.scroll = QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        self.scroll_content = QWidget(self.scroll)
        self.scroll_layout = QGridLayout(self.scroll_content)
        self.scroll.setWidget(self.scroll_content)
        right_layout.addWidget(self.scroll)

        main_layout.addLayout(right_layout)

        self.setLayout(main_layout)
        self.show()

    def add_button(self, layout, text, handler, *args):
        button = QPushButton(text, self)
        button.clicked.connect(lambda: handler(*args))
        layout.addWidget(button)

    def select_file(self, path_key):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择文件')
        if file_path:
            self.paths[path_key] = file_path
            self.labels[path_key].setText(f'{path_key}: <font color="green">{file_path}</font>')
            self.labels[path_key].setToolTip(file_path)  # 设置完整路径为工具提示


    def startProcess(self):
        if any(not path for path in self.paths.values()):
            QMessageBox.warning(self, '警告', '请先选择所有文件')
            return

        if not all(path.endswith(('.xlsx', '.xls')) for path in self.paths.values()):
            QMessageBox.warning(self, '警告', '请选择正确的表格文件文件(.xlsx)/(.xls)')
            return

        report_2021, report_2022, report_2023, back_data = self.read_files()
        if report_2021 is None or report_2022 is None or report_2023 is None or back_data is None:
            return

        data_2021, data_2022, data_2023 = self.extract_data(report_2021), self.extract_data(report_2022), self.extract_data(report_2023)
        if data_2021 is None or data_2022 is None or data_2023 is None:
            QMessageBox.warning(self, '警告', '年报文件中表格数量不正确')
            return
        

        final_data_2021, detailed_data_2021 = self.calculate_data(data_2021, back_data, 2021)
        final_data_2022, detailed_data_2022 = self.calculate_data(data_2022, back_data, 2022)
        final_data_2023, detailed_data_2023 = self.calculate_data(data_2023, back_data, 2023)

        # 修正数据----------------------------------
            # 资本回报率取均值
        if (final_data_2022["资本"]!=0) and (final_data_2021["所有者权益合计"]!=0):
            final_data_2022["资本回报率"] = final_data_2022["EBIT"] / ((final_data_2022["资本"] + final_data_2021["资本"])/2)

        if (final_data_2023["资本"]!=0) and (final_data_2022["所有者权益合计"]!=0):
            final_data_2023["资本回报率"] = final_data_2023["EBIT"] / ((final_data_2023["资本"] + final_data_2022["资本"])/2)
        # ----------------------------------------

        total_data_2021 = {**final_data_2021, **detailed_data_2021}
        total_data_2022 = {**final_data_2022, **detailed_data_2022}
        total_data_2023 = {**final_data_2023, **detailed_data_2023}

        # 显示数据到滑动框
        self.display_data(total_data_2023, total_data_2022, total_data_2021)

        self.write_to_excel(final_data_2021, final_data_2022, final_data_2023)

    def read_files(self):
        try:
            report_2021 = pd.read_excel(self.paths["report_2021_path"], sheet_name=None)
        except:
            QMessageBox.warning(self, '警告', '2021年年报文件读取失败')
            return None, None, None
        try:
            report_2022 = pd.read_excel(self.paths["report_2022_path"], sheet_name=None)
        except:
            QMessageBox.warning(self, '警告', '2022年年报文件读取失败')
            return None, None, None
        try:
            report_2023 = pd.read_excel(self.paths["report_2023_path"], sheet_name=None)
        except:
            QMessageBox.warning(self, '警告', '2023年年报文件读取失败')
            return None, None, None
        try:
            back_data = pd.read_excel(self.paths["back_data_path"])
        except:
            QMessageBox.warning(self, '警告', '数据底稿文件读取失败')
            return None, None, None
        return report_2021, report_2022, report_2023, back_data

    def extract_data(self, report):
        sheets = {name: sheet for name, sheet in report.items() if "资产负债表" in name or "利润表" in name or "现金流量表" in name}
        sheets = dict(list(sheets.items())[:4]) # 只取前四个表 
        if len(sheets) != 4:
            return None
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
        return sheets

    def calculate_data(self, sheets, back_data, year):
        data = {}
        if year == 2021:
            data_set = self.extract_values_2021(sheets, back_data)
        elif year == 2022:
            data_set = self.extract_values_2022(sheets, back_data)
        elif year == 2023:
            data_set = self.extract_values_2023(sheets, back_data)

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
            data_set["财务费用"],
            data_set["资本化利息"],
            data_set["经营租赁的利息调整"]
        )

        data["营业收入"] = data_set["营业收入"]

        data["总资产"] = data_set["总资产"]

        data["所有者权益合计"] = data_set["所有者权益合计"]

        return data, data_set

    def display_data(self, data_2023, data_2022, data_2021):
        row = 0
        col = 0

        years_data = [("2023年数据", data_2023), ("2022年数据", data_2022), ("2021年数据", data_2021)]
        
        

        for year_label, data in years_data:
            for key, value in data.items():
                if year_label == "2023年数据": #为了避免重复，只显示第一例地数据标签
                    self.scroll_layout.addWidget(QLabel(f'{key}:'), row, col)
                self.scroll_layout.addWidget(QLabel(str(round(value, 2))), row, col + 1)
                row += 1
            col += 1
            row = 0

    def extract_values_2021(self, sheets, back_data):
        return {
            "营业利润": sheets["利润表"].iloc[row_to_num(7), col_to_num("G")],
            "财务费用": sheets["利润表"].iloc[row_to_num(33), col_to_num("C")],
            "折旧费": back_data.iloc[row_to_num(2), col_to_num("D")],
            "公允价值变动": sheets["利润表"].iloc[row_to_num(44), col_to_num("C")],
            "投资收益": sheets["利润表"].iloc[row_to_num(39), col_to_num("C")],
            "取得投资收益收到的现金": sheets["现金流量表"].iloc[row_to_num(6), col_to_num("H")],
            "政府补助": sheets["利润表"].iloc[row_to_num(9), col_to_num("G")],
            "经营租赁费用调整": back_data.iloc[row_to_num(3), col_to_num("D")],
            "资本化开发成本": back_data.iloc[row_to_num(4), col_to_num("D")],
            "勘探费用": back_data.iloc[row_to_num(5), col_to_num("D")],
            "利息收入": sheets["利润表"].iloc[row_to_num(35), col_to_num("C")],
            "对联营企业和合营企业的投资收益": sheets["利润表"].iloc[row_to_num(40), col_to_num("C")],
            "经营租赁的利息调整": back_data.iloc[row_to_num(6), col_to_num("D")],
            "经营活动产生的现金流量净额": sheets["现金流量表"].iloc[row_to_num(32), col_to_num("D")],
            "购建固定资产无形资产和其他长期资产支付的现金": sheets["现金流量表"].iloc[row_to_num(11), col_to_num("H")],
            "收到其他与投资活动有关的现金附注利息部分": back_data.iloc[row_to_num(7), col_to_num("D")],
            "分配股利利润或偿付利息支付的现金": sheets["现金流量表"].iloc[row_to_num(25), col_to_num("H")],
            "对所有者或股东的分配": back_data.iloc[row_to_num(8), col_to_num("D")],
            "经营租赁折旧调整": back_data.iloc[row_to_num(9), col_to_num("D")],
            "所得税费用": sheets["利润表"].iloc[row_to_num(12), col_to_num("G")],
            "资本化利息": back_data.iloc[row_to_num(10), col_to_num("D")],
            "短期借款": sheets["资产负债表_con"].iloc[row_to_num(6), col_to_num("D")],
            "应付利息": back_data.iloc[row_to_num(11), col_to_num("D")],
            "一年内到期的长期借款": back_data.iloc[row_to_num(12), col_to_num("D")],
            "一年内到期的应付债券": back_data.iloc[row_to_num(13), col_to_num("D")],
            "其它流动负债短期应付债券": back_data.iloc[row_to_num(14), col_to_num("D")],
            "一年内应付融资租赁款": back_data.iloc[row_to_num(15), col_to_num("D")],
            "长期借款": sheets["资产负债表_con"].iloc[row_to_num(37), col_to_num("D")],
            "应付债券": sheets["资产负债表_con"].iloc[row_to_num(38), col_to_num("D")],
            "长期应付融资租赁款": back_data.iloc[row_to_num(16), col_to_num("D")],
            "重大合同及履行状况担保情况": back_data.iloc[row_to_num(17), col_to_num("D")],
            "货币资金": sheets["资产负债表"].iloc[row_to_num(6), col_to_num("D")],
            "以公允价值计量且其变动计入当期损益的金融资产": sheets["资产负债表"].iloc[row_to_num(10), col_to_num("D")],
            "其他货币资金": back_data.iloc[row_to_num(21), col_to_num("D")],
            "卖出回购金融资产款": back_data.iloc[row_to_num(18), col_to_num("D")],
            "特定行业或公司现金盈余不做调整扣除的部分加回": back_data.iloc[row_to_num(19), col_to_num("D")],
            "经营租赁调整": back_data.iloc[row_to_num(20), col_to_num("D")],
            "永续债": sheets["资产负债表_con"].iloc[row_to_num(40), col_to_num("D")],
            "所有者权益合计": sheets["资产负债表_con"].iloc[row_to_num(80), col_to_num("D")],
            "递延所得税负债": sheets["资产负债表_con"].iloc[row_to_num(48), col_to_num("D")],
            "营业收入": sheets["利润表"].iloc[row_to_num(5), col_to_num("C")],
            "总资产": sheets["资产负债表"].iloc[row_to_num(61), col_to_num("D")],
            "利息费用": sheets["利润表"].iloc[row_to_num(34), col_to_num("C")]
        }


    def extract_values_2022(self, sheets, back_data):
        return {
            "营业利润": sheets["利润表"].iloc[row_to_num(7), col_to_num("G")],
            "财务费用": sheets["利润表"].iloc[row_to_num(33), col_to_num("C")],
            "折旧费": back_data.iloc[row_to_num(2), col_to_num("C")],
            "公允价值变动": sheets["利润表"].iloc[row_to_num(44), col_to_num("C")],
            "投资收益": sheets["利润表"].iloc[row_to_num(39), col_to_num("C")],
            "取得投资收益收到的现金": sheets["现金流量表"].iloc[row_to_num(6), col_to_num("H")],
            "政府补助": sheets["利润表"].iloc[row_to_num(9), col_to_num("G")],
            "经营租赁费用调整": back_data.iloc[row_to_num(3), col_to_num("C")],
            "资本化开发成本": back_data.iloc[row_to_num(4), col_to_num("C")],
            "勘探费用": back_data.iloc[row_to_num(5), col_to_num("C")],
            "利息收入": sheets["利润表"].iloc[row_to_num(35), col_to_num("C")],
            "对联营企业和合营企业的投资收益": sheets["利润表"].iloc[row_to_num(40), col_to_num("C")],
            "经营租赁的利息调整": back_data.iloc[row_to_num(6), col_to_num("C")],
            "经营活动产生的现金流量净额": sheets["现金流量表"].iloc[row_to_num(32), col_to_num("D")],
            "购建固定资产无形资产和其他长期资产支付的现金": sheets["现金流量表"].iloc[row_to_num(11), col_to_num("H")],
            "收到其他与投资活动有关的现金附注利息部分": back_data.iloc[row_to_num(7), col_to_num("C")],
            "分配股利利润或偿付利息支付的现金": sheets["现金流量表"].iloc[row_to_num(25), col_to_num("H")],
            "对所有者或股东的分配": back_data.iloc[row_to_num(8), col_to_num("C")],
            "经营租赁折旧调整": back_data.iloc[row_to_num(9), col_to_num("C")],
            "所得税费用": sheets["利润表"].iloc[row_to_num(12), col_to_num("G")],
            "资本化利息": back_data.iloc[row_to_num(10), col_to_num("C")],
            "短期借款": sheets["资产负债表_con"].iloc[row_to_num(6), col_to_num("D")],
            "应付利息": back_data.iloc[row_to_num(11), col_to_num("C")],
            "一年内到期的长期借款": back_data.iloc[row_to_num(12), col_to_num("C")],
            "一年内到期的应付债券": back_data.iloc[row_to_num(13), col_to_num("C")],
            "其它流动负债短期应付债券": back_data.iloc[row_to_num(14), col_to_num("C")],
            "一年内应付融资租赁款": back_data.iloc[row_to_num(15), col_to_num("C")],
            "长期借款": sheets["资产负债表_con"].iloc[row_to_num(37), col_to_num("D")],
            "应付债券": sheets["资产负债表_con"].iloc[row_to_num(38), col_to_num("D")],
            "长期应付融资租赁款": back_data.iloc[row_to_num(16), col_to_num("C")],
            "重大合同及履行状况担保情况": back_data.iloc[row_to_num(17), col_to_num("C")],
            "货币资金": sheets["资产负债表"].iloc[row_to_num(6), col_to_num("D")],
            "以公允价值计量且其变动计入当期损益的金融资产": sheets["资产负债表"].iloc[row_to_num(10), col_to_num("D")],
            "其他货币资金": back_data.iloc[row_to_num(21), col_to_num("C")],
            "卖出回购金融资产款": back_data.iloc[row_to_num(18), col_to_num("C")],
            "特定行业或公司现金盈余不做调整扣除的部分加回": back_data.iloc[row_to_num(19), col_to_num("C")],
            "经营租赁调整": back_data.iloc[row_to_num(20), col_to_num("C")],
            "永续债": sheets["资产负债表_con"].iloc[row_to_num(40), col_to_num("D")],
            "所有者权益合计": sheets["资产负债表_con"].iloc[row_to_num(80), col_to_num("D")],
            "递延所得税负债": sheets["资产负债表_con"].iloc[row_to_num(48), col_to_num("D")],
            "营业收入": sheets["利润表"].iloc[row_to_num(5), col_to_num("C")],
            "总资产": sheets["资产负债表"].iloc[row_to_num(62), col_to_num("D")],
            "利息费用": sheets["利润表"].iloc[row_to_num(34), col_to_num("C")]
        }

    def extract_values_2023(self, sheets, back_data):
        return {
            # 类似的处理逻辑，用适当的列和行号
            "营业利润": sheets["利润表"].iloc[row_to_num(7), col_to_num("I")],
            "财务费用": sheets["利润表"].iloc[row_to_num(33), col_to_num("D")],
            "折旧费": back_data.iloc[row_to_num(2), col_to_num("B")],
            "公允价值变动": sheets["利润表"].iloc[row_to_num(44), col_to_num("D")],
            "投资收益": sheets["利润表"].iloc[row_to_num(39), col_to_num("D")],
            "取得投资收益收到的现金": sheets["现金流量表"].iloc[row_to_num(6), col_to_num("H")],
            "政府补助": sheets["利润表"].iloc[row_to_num(9), col_to_num("I")],
            "经营租赁费用调整": back_data.iloc[row_to_num(3), col_to_num("B")],
            "资本化开发成本": back_data.iloc[row_to_num(4), col_to_num("B")],
            "勘探费用": back_data.iloc[row_to_num(5), col_to_num("B")],
            "利息收入": sheets["利润表"].iloc[row_to_num(35), col_to_num("D")],
            "对联营企业和合营企业的投资收益": sheets["利润表"].iloc[row_to_num(40), col_to_num("D")],
            "经营租赁的利息调整": back_data.iloc[row_to_num(6), col_to_num("B")],
            "经营活动产生的现金流量净额": sheets["现金流量表"].iloc[row_to_num(32), col_to_num("D")],
            "购建固定资产无形资产和其他长期资产支付的现金": sheets["现金流量表"].iloc[row_to_num(11), col_to_num("H")],
            "收到其他与投资活动有关的现金附注利息部分": back_data.iloc[row_to_num(7), col_to_num("B")],
            "分配股利利润或偿付利息支付的现金": sheets["现金流量表"].iloc[row_to_num(25), col_to_num("H")],
            "对所有者或股东的分配": back_data.iloc[row_to_num(8), col_to_num("B")],
            "经营租赁折旧调整": back_data.iloc[row_to_num(9), col_to_num("B")],
            "所得税费用": sheets["利润表"].iloc[row_to_num(12), col_to_num("I")],
            "资本化利息": back_data.iloc[row_to_num(10), col_to_num("B")],
            "短期借款": sheets["资产负债表_con"].iloc[row_to_num(6), col_to_num("D")],
            "应付利息": back_data.iloc[row_to_num(11), col_to_num("B")],
            "一年内到期的长期借款": back_data.iloc[row_to_num(12), col_to_num("B")],
            "一年内到期的应付债券": back_data.iloc[row_to_num(13), col_to_num("B")],
            "其它流动负债短期应付债券": back_data.iloc[row_to_num(14), col_to_num("B")],
            "一年内应付融资租赁款": back_data.iloc[row_to_num(15), col_to_num("B")],
            "长期借款": sheets["资产负债表_con"].iloc[row_to_num(37), col_to_num("D")],
            "应付债券": sheets["资产负债表_con"].iloc[row_to_num(38), col_to_num("D")],
            "长期应付融资租赁款": back_data.iloc[row_to_num(16), col_to_num("B")],
            "重大合同及履行状况担保情况": back_data.iloc[row_to_num(17), col_to_num("B")],
            "货币资金": sheets["资产负债表"].iloc[row_to_num(6), col_to_num("D")],
            "以公允价值计量且其变动计入当期损益的金融资产": sheets["资产负债表"].iloc[row_to_num(10), col_to_num("D")],
            "其他货币资金": back_data.iloc[row_to_num(21), col_to_num("B")],
            "卖出回购金融资产款": back_data.iloc[row_to_num(18), col_to_num("C")],
            "特定行业或公司现金盈余不做调整扣除的部分加回": back_data.iloc[row_to_num(19), col_to_num("B")],
            "经营租赁调整": back_data.iloc[row_to_num(20), col_to_num("B")],
            "永续债": sheets["资产负债表_con"].iloc[row_to_num(40), col_to_num("D")],
            "所有者权益合计": sheets["资产负债表_con"].iloc[row_to_num(80), col_to_num("D")],
            "递延所得税负债": sheets["资产负债表_con"].iloc[row_to_num(48), col_to_num("D")],
            "营业收入": sheets["利润表"].iloc[row_to_num(5), col_to_num("D")],
            "总资产": sheets["资产负债表"].iloc[row_to_num(62), col_to_num("D")],
            "利息费用": sheets["利润表"].iloc[row_to_num(34), col_to_num("D")]
        }


    def write_to_excel(self, data_2021, data_2022, data_2023):
        try:
            report = load_workbook(self.paths["target_path"])
            sheet = report["Inputs"]
            print("Writing to Excel...")
            # 将数据写入Excel
                # 逐年分析
            sheet["D46"], sheet["E46"], sheet["F46"] = data_2023["EBITDA利润率"], data_2022["EBITDA利润率"], data_2021["EBITDA利润率"]
            sheet["D47"], sheet["E47"], sheet["F47"] = data_2023["资本回报率"], data_2022["资本回报率"], data_2021["资本回报率"]
            sheet["D48"], sheet["E48"], sheet["F48"] = data_2023["营业收入"], data_2022["营业收入"], data_2021["营业收入"]
            sheet["D49"], sheet["E49"], sheet["F49"] = data_2023["总资产"], data_2022["总资产"], data_2021["总资产"]
            sheet["D54"], sheet["E54"], sheet["F54"] = data_2023["经营活动产生的资金/债务"], data_2022["经营活动产生的资金/债务"], data_2021["经营活动产生的资金/债务"]
            sheet["D55"], sheet["E55"], sheet["F55"] = data_2023["债务/息税摊折前利润"], data_2022["债务/息税摊折前利润"], data_2021["债务/息税摊折前利润"]
            sheet["D56"], sheet["E56"], sheet["F56"] = data_2023["自由运营现金流/债务"], data_2022["自由运营现金流/债务"], data_2021["自由运营现金流/债务"]
            sheet["D57"], sheet["E57"], sheet["F57"] = data_2023["息税摊折前利润 / 利息支出"], data_2022["息税摊折前利润 / 利息支出"], data_2021["息税摊折前利润 / 利息支出"]
            sheet["D58"], sheet["E58"], sheet["F58"] = data_2023["经营活动产生的现金(FFO)"], data_2022["经营活动产生的现金(FFO)"], data_2021["经营活动产生的现金(FFO)"]
            sheet["D59"], sheet["E59"], sheet["F59"] = data_2023["总负债"], data_2022["总负债"], data_2021["总负债"]
            sheet["D60"], sheet["E60"], sheet["F60"] = data_2023["EBITDA"], data_2022["EBITDA"], data_2021["EBITDA"]
            sheet["D63"], sheet["E63"], sheet["F63"] = data_2023["营业收入"], data_2022["营业收入"], data_2021["营业收入"]
            sheet["D64"], sheet["E64"], sheet["F64"] = data_2023["总资产"], data_2022["总资产"], data_2021["总资产"]

                # 三年平均
            for i in range(46, 65):
                if i not in [50, 51, 52, 53, 61, 62]:
                    sheet[f"B{i}"] = f"=AVERAGE(D{i}:F{i})"
            print("赋值完成")

            # 检查特殊情况并设置为 "NM"
            for i in sheet["B46:G64"]:
                for j in i:
                    # print(j.value)
                    if str(j.value) == "inf" or str(j.value) == "-inf" or str(j.value) == "nan":
                        j.value = "NM"

            if data_2021["经营活动产生的现金(FFO)"] < 0 and data_2021["总负债"] < 0:
                sheet["F54"] = "NM"
            if data_2021["总负债"] < 0 and data_2021["EBITDA"] < 0:
                sheet["F55"] = "NM"
            if data_2021["自由运营现金流(FOCF)"] < 0 and data_2021["总负债"] < 0:
                sheet["F56"] = "NM"

            if data_2022["经营活动产生的现金(FFO)"] < 0 and data_2022["总负债"] < 0:
                sheet["E54"] = "NM"
            if data_2022["总负债"] < 0 and data_2022["EBITDA"] < 0:
                sheet["E55"] = "NM"
            if data_2022["自由运营现金流(FOCF)"] < 0 and data_2022["总负债"] < 0:
                sheet["E56"] = "NM"

            if data_2023["经营活动产生的现金(FFO)"] < 0 and data_2023["总负债"] < 0:
                sheet["D54"], sheet["B54"] = "NM", "NM"
            if data_2023["总负债"] < 0 and data_2023["EBITDA"] < 0:
                sheet["D55"], sheet["B55"] = "NM", "NM"
            if data_2023["自由运营现金流(FOCF)"] < 0 and data_2023["总负债"] < 0:
                sheet["D56"], sheet["B56"] = "NM", "NM"
            print("特殊情况处理完成")

            report.save(self.paths["target_path"])

            QMessageBox.information(self, "提示", "数据填充完成", QMessageBox.Yes)
            os.startfile(self.paths["target_path"])
        except Exception as e:
            QMessageBox.warning(self, '警告', f'文件写入失败: {str(e)}, 请确保填入文件没被打开')

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

def EBITDA_to_interest_expense(EBITDA, 财务费用, 资本化利息, 经营租赁的利息调整):
    return EBITDA / (财务费用 + 资本化利息 + 经营租赁的利息调整)

def col_to_num(col_str):
    num = 0
    for i, c in enumerate(reversed(col_str)):
        num += (ord(c) - ord('A') + 1) * (26 ** i)
    return num - 1

def row_to_num(num):
    return int(num) - 2

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    sys.exit(app.exec_())
