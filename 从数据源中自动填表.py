import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, QLabel

import numpy as np
import pandas as pd
from openpyxl import load_workbook

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.report_2022_path = ""
        self.report_2023_path = ""
        self.back_data_path = ""
        self.target_path = ""

    def initUI(self):
        self.setGeometry(700, 400, 600, 420)
        self.setWindowTitle('数据源选择')

        layout = QVBoxLayout()

        # 按钮设置
        self.button1 = QPushButton('选择2022年年报', self)
        self.button1.clicked.connect(self.selectReport2022)
        layout.addWidget(self.button1)

        self.button2 = QPushButton('选择2023年年报', self)
        self.button2.clicked.connect(self.selectReport2023)
        layout.addWidget(self.button2)

        self.button3 = QPushButton('选择数据底稿', self)
        self.button3.clicked.connect(self.selectBackData)
        layout.addWidget(self.button3)


        self.button4 = QPushButton('生成位置', self)
        self.button4.clicked.connect(self.selectTargetPath)
        layout.addWidget(self.button4)  

        self.button5 = QPushButton('开始处理', self)
        self.button5.clicked.connect(self.startProcess) # 开始处理
        layout.addWidget(self.button5)

        # 标签设置
        self.label_2022 = QLabel('2022年年报路径: 未选择', self)
        layout.addWidget(self.label_2022)

        self.label_2023 = QLabel('2023年年报路径: 未选择', self)
        layout.addWidget(self.label_2023)

        self.label_back_data = QLabel('备份数据路径: 未选择', self)
        layout.addWidget(self.label_back_data)

        self.label_target_path = QLabel('生成位置: 未选择', self)
        layout.addWidget(self.label_target_path)

        self.setLayout(layout)
        self.show()

    def selectReport2022(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择2022年年报文件')
        if file_path:
            self.report_2022_path = file_path
            self.label_2022.setText(f'2022年年报路径: {file_path}')

    def selectReport2023(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择2023年年报文件')
        if file_path:
            self.report_2023_path = file_path
            self.label_2023.setText(f'2023年年报路径: {file_path}')

    def selectBackData(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择数据底稿文件')
        if file_path:
            self.back_data_path = file_path
            self.label_back_data.setText(f'数据底稿路径: {file_path}')

    def selectTargetPath(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择生成位置')
        if file_path:
            self.target_path = file_path
            self.label_target_path.setText(f'生成位置: {file_path}')

    def startProcess(self):
        # 设置路径
        report_path = self.target_path
        report_2022_path = self.report_2022_path
        report_2023_path = self.report_2023_path
        back_data_path = self.back_data_path

        # 读取文件以及需要的数据
        report_2022 = pd.read_excel(report_2022_path, sheet_name=None)
        report_2023 = pd.read_excel(report_2023_path, sheet_name=None)
        back_data = pd.read_excel(back_data_path)

        # 数据预处理
            # 2022年年报数据
        bal_sheet_2022 = report_2022["NB003-资产负债表"]
        bal_sheet_2022_con = report_2022["NB004-资产负债表（续）"]
        profit_sheet_2022 = report_2022["NB005-利润表"]
        cash_sheet_2022 = report_2022["NB006-现金流量表"]


            # 2023年年报数据
        bal_sheet_2023 = report_2023["Z01 资产负债表"]
        profit_sheet_2023 = report_2023["Z02 利润表"]
        cash_sheet_2023 = report_2023["Z03 现金流量表"]


            # 数据底稿
        back_data = back_data

        # 数据分配
            # 建立字典
        raw_data_2022 = {"EBITDA": 0,
                        "EBIT": 0,
                        "自由运营现金流(FOCF)": 0,
                        "经营活动产生的现金(FFO)": 0,
                        "总负债": 0,
                        "资本": 0,
                        "EBITDA利润率": 0,
                        "资本回报率": 0,
                        "经营活动产生的资金/债务": 0,
                        "债务/息税摊折前利润": 0,
                        "自由运营现金流/债务": 0,
                        "息税摊折前利润 / 利息支出": 0
                        }

        raw_data_2023 = {"EBITDA": 0,
                        "EBIT": 0,
                        "自由运营现金流(FOCF)": 0,
                        "经营活动产生的现金(FFO)": 0,
                        "总负债": 0,
                        "资本": 0,
                        "EBITDA利润率": 0,
                        "资本回报率": 0,
                        "经营活动产生的资金/债务": 0,
                        "债务/息税摊折前利润": 0,
                        "自由运营现金流/债务": 0,
                        "息税摊折前利润 / 利息支出": 0
                        }

        # 计算
            # 2022
        营业利润_2022 = profit_sheet_2022.iloc[(row_to_num(7)), col_to_num("G")]
        财务费用_2022 = profit_sheet_2022.iloc[(row_to_num(33)), col_to_num("C")]
        折旧费_2022 = back_data.iloc[(row_to_num(2)), col_to_num("C")]
        公允价值变动_2022 = profit_sheet_2022.iloc[(row_to_num(44)), col_to_num("C")]
        投资收益_2022 = profit_sheet_2022.iloc[(row_to_num(39)), col_to_num("C")]
        取得投资收益收到的现金_2022 = cash_sheet_2022.iloc[(row_to_num(6)), col_to_num("H")]
        政府补助_2022 = profit_sheet_2022.iloc[(row_to_num(9)), col_to_num("G")]
        经营租赁费用调整_2022 = back_data.iloc[(row_to_num(3)), col_to_num("C")]
        资本化开发成本_2022 = back_data.iloc[(row_to_num(4)), col_to_num("C")]
        勘探费用_2022 = back_data.iloc[(row_to_num(5)), col_to_num("C")]
        利息收入_2022 = profit_sheet_2022.iloc[(row_to_num(35)), col_to_num("C")]
        对联营企业和合营企业的投资收益_2022 = profit_sheet_2022.iloc[(row_to_num(40)), col_to_num("C")]
        经营租赁的利息调整_2022 = back_data.iloc[(row_to_num(6)), col_to_num("C")]
        经营活动产生的现金流量净额_2022 = cash_sheet_2022.iloc[(row_to_num(32)), col_to_num("D")]
        购建固定资产无形资产和其他长期资产支付的现金_2022 = cash_sheet_2022.iloc[(row_to_num(11)), col_to_num("H")]
        收到其他与投资活动有关的现金附注利息部分_2022 = back_data.iloc[(row_to_num(7)), col_to_num("C")]
        分配股利利润或偿付利息支付的现金_2022 = cash_sheet_2022.iloc[(row_to_num(25)), col_to_num("H")]
        对所有者或股东的分配_2022 = back_data.iloc[(row_to_num(8)), col_to_num("C")]
        经营租赁折旧调整_2022 = back_data.iloc[(row_to_num(9)), col_to_num("C")]
        所得税费用_2022 = profit_sheet_2022.iloc[(row_to_num(12)), col_to_num("G")]
        资本化利息_2022 = back_data.iloc[(row_to_num(10)), col_to_num("C")]
        短期借款_2022 = bal_sheet_2022_con.iloc[(row_to_num(6)), col_to_num("D")]
        应付利息_2022 = back_data.iloc[(row_to_num(11)), col_to_num("C")]
        一年内到期的长期借款_2022 = back_data.iloc[(row_to_num(12)), col_to_num("C")]
        一年内到期的应付债券_2022 = back_data.iloc[(row_to_num(13)), col_to_num("C")]
        其它流动负债短期应付债券_2022 = back_data.iloc[(row_to_num(14)), col_to_num("C")]
        一年内应付融资租赁款_2022 = back_data.iloc[(row_to_num(15)), col_to_num("C")]
        长期借款_2022 = bal_sheet_2022_con.iloc[(row_to_num(37)), col_to_num("D")]
        应付债券_2022 = bal_sheet_2022_con.iloc[(row_to_num(38)), col_to_num("D")]
        长期应付融资租赁款_2022 = back_data.iloc[(row_to_num(16)), col_to_num("C")]
        重大合同及履行状况担保情况_2022 = back_data.iloc[(row_to_num(17)), col_to_num("C")]
        货币资金_2022 = bal_sheet_2022.iloc[(row_to_num(6)), col_to_num("D")]
        以公允价值计量且其变动计入当期损益的金融资产_2022 = bal_sheet_2022.iloc[(row_to_num(10)), col_to_num("D")]
        其他货币资金_2022 = back_data.iloc[(row_to_num(21)), col_to_num("C")]
        卖出回购金融资产款_2022 = back_data.iloc[(row_to_num(18)), col_to_num("C")]
        特定行业或公司现金盈余不做调整扣除的部分加回_2022 = back_data.iloc[(row_to_num(19)), col_to_num("C")]
        经营租赁调整_2022 = back_data.iloc[(row_to_num(20)), col_to_num("C")]
        永续债_2022 = bal_sheet_2022_con.iloc[(row_to_num(40)), col_to_num("D")]
        所有者权益合计_2022 = bal_sheet_2022_con.iloc[(row_to_num(80)), col_to_num("D")]
        递延所得税负债_2022 = bal_sheet_2022_con.iloc[(row_to_num(48)), col_to_num("D")]
        营业收入_2022 = profit_sheet_2022.iloc[(row_to_num(5)), col_to_num("C")]
        总资产_2022 = bal_sheet_2022.iloc[(row_to_num(62)), col_to_num("D")]

            # 2023
        营业利润_2023 = profit_sheet_2023.iloc[row_to_num(8), col_to_num("G")]
        财务费用_2023 = profit_sheet_2023.iloc[row_to_num(40), col_to_num("C")]
        折旧费_2023 = back_data.iloc[row_to_num(2), col_to_num("B")]
        公允价值变动_2023 = profit_sheet_2023.iloc[row_to_num(51), col_to_num("C")]
        投资收益_2023 = profit_sheet_2023.iloc[row_to_num(46), col_to_num("C")]
        取得投资收益收到的现金_2023 = cash_sheet_2023.iloc[row_to_num(10), col_to_num("G")]
        政府补助_2023 = profit_sheet_2023.iloc[row_to_num(10), col_to_num("G")]
        经营租赁费用调整_2023 = back_data.iloc[row_to_num(3), col_to_num("B")]
        资本化开发成本_2023 = back_data.iloc[row_to_num(4), col_to_num("B")]
        勘探费用_2023 = back_data.iloc[row_to_num(5), col_to_num("B")]
        利息收入_2023 = profit_sheet_2023.iloc[row_to_num(13), col_to_num("C")]
        对联营企业和合营企业的投资收益_2023 = profit_sheet_2023.iloc[row_to_num(47), col_to_num("C")]
        经营租赁的利息调整_2023 = back_data.iloc[row_to_num(6), col_to_num("B")]
        经营活动产生的现金流量净额_2023 = cash_sheet_2023.iloc[row_to_num(40), col_to_num("C")]
        购建固定资产无形资产和其他长期资产支付的现金_2023 = cash_sheet_2023.iloc[row_to_num(15), col_to_num("G")]
        收到其他与投资活动有关的现金附注利息部分_2023 = back_data.iloc[row_to_num(7), col_to_num("B")]
        分配股利利润或偿付利息支付的现金_2023 = cash_sheet_2023.iloc[row_to_num(29), col_to_num("G")]
        对所有者或股东的分配_2023 = back_data.iloc[row_to_num(8), col_to_num("B")]
        经营租赁折旧调整_2023 = back_data.iloc[row_to_num(9), col_to_num("B")]
        所得税费用_2023 = profit_sheet_2023.iloc[row_to_num(13), col_to_num("G")]
        资本化利息_2023 = back_data.iloc[row_to_num(10), col_to_num("B")]
        短期借款_2023 = bal_sheet_2023.iloc[row_to_num(9), col_to_num("G")]
        应付利息_2023 = back_data.iloc[row_to_num(11), col_to_num("B")]
        一年内到期的长期借款_2023 = back_data.iloc[row_to_num(12), col_to_num("B")]
        一年内到期的应付债券_2023 = back_data.iloc[row_to_num(13), col_to_num("B")]
        其它流动负债短期应付债券_2023 = back_data.iloc[row_to_num(14), col_to_num("B")]
        一年内应付融资租赁款_2023 = back_data.iloc[row_to_num(15), col_to_num("B")]
        长期借款_2023 = bal_sheet_2023.iloc[row_to_num(41), col_to_num("G")]
        应付债券_2023 = bal_sheet_2023.iloc[row_to_num(42), col_to_num("G")]
        长期应付融资租赁款_2023 = back_data.iloc[row_to_num(16), col_to_num("B")]
        重大合同及履行状况担保情况_2023 = back_data.iloc[row_to_num(17), col_to_num("B")]
        货币资金_2023 = bal_sheet_2023.iloc[row_to_num(9), col_to_num("C")]
        以公允价值计量且其变动计入当期损益的金融资产_2023 = bal_sheet_2023.iloc[row_to_num(13), col_to_num("C")]
        其他货币资金_2023 = back_data.iloc[row_to_num(21), col_to_num("B")]
        卖出回购金融资产款_2023 = back_data.iloc[row_to_num(18), col_to_num("B")]
        特定行业或公司现金盈余不做调整扣除的部分加回_2023 = back_data.iloc[row_to_num(19), col_to_num("B")]
        经营租赁调整_2023 = back_data.iloc[row_to_num(20), col_to_num("B")]
        永续债_2023 = bal_sheet_2023.iloc[row_to_num(44), col_to_num("G")]
        所有者权益合计_2023 = bal_sheet_2023.iloc[row_to_num(86), col_to_num("G")]
        递延所得税负债_2023 = bal_sheet_2023.iloc[row_to_num(54), col_to_num("G")]
        营业收入_2023 = profit_sheet_2023.iloc[row_to_num(6), col_to_num("C")]
        总资产_2023 = bal_sheet_2023.iloc[row_to_num(87), col_to_num("C")]

            # 数据打磨
        data_set = {
            "营业利润_2022": 营业利润_2022, "财务费用_2022": 财务费用_2022, "折旧费_2022": 折旧费_2022, "公允价值变动_2022": 公允价值变动_2022, "投资收益_2022": 投资收益_2022, "取得投资收益收到的现金_2022": 取得投资收益收到的现金_2022, "政府补助_2022": 政府补助_2022, "经营租赁费用调整_2022": 经营租赁费用调整_2022, "资本化开发成本_2022": 资本化开发成本_2022, "勘探费用_2022": 勘探费用_2022,
            "利息收入_2022": 利息收入_2022, "对联营企业和合营企业的投资收益_2022": 对联营企业和合营企业的投资收益_2022, "经营租赁的利息调整_2022": 经营租赁的利息调整_2022, "经营活动产生的现金流量净额_2022": 经营活动产生的现金流量净额_2022, "购建固定资产无形资产和其他长期资产支付的现金_2022": 购建固定资产无形资产和其他长期资产支付的现金_2022, "收到其他与投资活动有关的现金附注利息部分_2022": 收到其他与投资活动有关的现金附注利息部分_2022, "分配股利利润或偿付利息支付的现金_2022": 分配股利利润或偿付利息支付的现金_2022, "对所有者或股东的分配_2022": 对所有者或股东的分配_2022, "经营租赁折旧调整_2022": 经营租赁折旧调整_2022, "所得税费用_2022": 所得税费用_2022, "资本化利息_2022": 资本化利息_2022,
            "短期借款_2022": 短期借款_2022, "应付利息_2022": 应付利息_2022, "一年内到期的长期借款_2022": 一年内到期的长期借款_2022, "一年内到期的应付债券_2022": 一年内到期的应付债券_2022, "其它流动负债短期应付债券_2022": 其它流动负债短期应付债券_2022, "一年内应付融资租赁款_2022": 一年内应付融资租赁款_2022, "长期借款_2022": 长期借款_2022, "应付债券_2022": 应付债券_2022, "长期应付融资租赁款_2022": 长期应付融资租赁款_2022, "重大合同及履行状况担保情况_2022": 重大合同及履行状况担保情况_2022, "货币资金_2022": 货币资金_2022, "以公允价值计量且其变动计入当期损益的金融资产_2022": 以公允价值计量且其变动计入当期损益的金融资产_2022, "其他货币资金_2022": 其他货币资金_2022, "卖出回购金融资产款_2022": 卖出回购金融资产款_2022, 
            "特定行业或公司现金盈余不做调整扣除的部分加回_2022": 特定行业或公司现金盈余不做调整扣除的部分加回_2022, "经营租赁调整_2022": 经营租赁调整_2022, "永续债_2022": 永续债_2022, "所有者权益合计_2022": 所有者权益合计_2022, "递延所得税负债_2022": 递延所得税负债_2022, "营业收入_2022": 营业收入_2022,
            "营业利润_2023": 营业利润_2023, "财务费用_2023": 财务费用_2023, "折旧费_2023": 折旧费_2023, "公允价值变动_2023": 公允价值变动_2023, "投资收益_2023": 投资收益_2023, "取得投资收益收到的现金_2023": 取得投资收益收到的现金_2023, "政府补助_2023": 政府补助_2023, "经营租赁费用调整_2023": 经营租赁费用调整_2023, "资本化开发成本_2023": 资本化开发成本_2023, "勘探费用_2023": 勘探费用_2023,
            "利息收入_2023": 利息收入_2023, "对联营企业和合营企业的投资收益_2023": 对联营企业和合营企业的投资收益_2023, "经营租赁的利息调整_2023": 经营租赁的利息调整_2023, "经营活动产生的现金流量净额_2023": 经营活动产生的现金流量净额_2023, "购建固定资产无形资产和其他长期资产支付的现金_2023": 购建固定资产无形资产和其他长期资产支付的现金_2023, "收到其他与投资活动有关的现金附注利息部分_2023": 收到其他与投资活动有关的现金附注利息部分_2023, "分配股利利润或偿付利息支付的现金_2023": 分配股利利润或偿付利息支付的现金_2023, "对所有者或股东的分配_2023": 对所有者或股东的分配_2023, "经营租赁折旧调整_2023": 经营租赁折旧调整_2023, "所得税费用_2023": 所得税费用_2023, "资本化利息_2023": 资本化利息_2023,
            "短期借款_2023": 短期借款_2023, "应付利息_2023": 应付利息_2023, "一年内到期的长期借款_2023": 一年内到期的长期借款_2023, "一年内到期的应付债券_2023": 一年内到期的应付债券_2023, "其它流动负债短期应付债券_2023": 其它流动负债短期应付债券_2023, "一年内应付融资租赁款_2023": 一年内应付融资租赁款_2023, "长期借款_2023": 长期借款_2023, "应付债券_2023": 应付债券_2023, "长期应付融资租赁款_2023": 长期应付融资租赁款_2023, "重大合同及履行状况担保情况_2023": 重大合同及履行状况担保情况_2023, "货币资金_2023": 货币资金_2023, "以公允价值计量且其变动计入当期损益的金融资产_2023": 以公允价值计量且其变动计入当期损益的金融资产_2023, "其他货币资金_2023": 其他货币资金_2023, "卖出回购金融资产款_2023": 卖出回购金融资产款_2023,
            "特定行业或公司现金盈余不做调整扣除的部分加回_2023": 特定行业或公司现金盈余不做调整扣除的部分加回_2023, "经营租赁调整_2023": 经营租赁调整_2023, "永续债_2023": 永续债_2023, "所有者权益合计_2023": 所有者权益合计_2023, "递延所得税负债_2023": 递延所得税负债_2023, "营业收入_2023": 营业收入_2023
        }

                # 将所有空着的数据填充为0
        for i in data_set:
            if np.isnan(data_set[i]):
                data_set[i] = 0

        raw_data_2022["EBITDA"] = EBITDA(
            data_set["营业利润_2022"],
            data_set["财务费用_2022"],
            data_set["折旧费_2022"],
            data_set["公允价值变动_2022"],
            data_set["投资收益_2022"],
            data_set["取得投资收益收到的现金_2022"],
            data_set["政府补助_2022"],
            data_set["经营租赁费用调整_2022"],
            data_set["资本化开发成本_2022"],
            data_set["勘探费用_2022"]
        )

        raw_data_2022["EBIT"] = EBIT(
            data_set["营业利润_2022"],
            data_set["财务费用_2022"],
            data_set["利息收入_2022"],
            data_set["公允价值变动_2022"],
            data_set["投资收益_2022"],
            data_set["对联营企业和合营企业的投资收益_2022"],
            data_set["政府补助_2022"],
            data_set["经营租赁的利息调整_2022"]
        )

        raw_data_2022["自由运营现金流(FOCF)"] = FOCF(
            data_set["经营活动产生的现金流量净额_2022"],
            data_set["购建固定资产无形资产和其他长期资产支付的现金_2022"],
            data_set["取得投资收益收到的现金_2022"],
            data_set["收到其他与投资活动有关的现金附注利息部分_2022"],
            data_set["分配股利利润或偿付利息支付的现金_2022"],
            data_set["对所有者或股东的分配_2022"],
            data_set["经营租赁折旧调整_2022"],
            data_set["资本化开发成本_2022"]
        )

        raw_data_2022["经营活动产生的现金(FFO)"] = FFO(
            raw_data_2022["EBITDA"],
            data_set["财务费用_2022"],
            data_set["利息收入_2022"],
            data_set["所得税费用_2022"],
            data_set["经营租赁费用调整_2022"],
            data_set["经营租赁折旧调整_2022"],
            data_set["资本化利息_2022"]
        )

        raw_data_2022["总负债"] = Total_debt(
            data_set["短期借款_2022"],
            data_set["应付利息_2022"],
            data_set["一年内到期的长期借款_2022"],
            data_set["一年内到期的应付债券_2022"],
            data_set["其它流动负债短期应付债券_2022"],
            data_set["一年内应付融资租赁款_2022"],
            data_set["长期借款_2022"],
            data_set["应付债券_2022"],
            data_set["长期应付融资租赁款_2022"],
            data_set["重大合同及履行状况担保情况_2022"],
            data_set["货币资金_2022"],
            data_set["以公允价值计量且其变动计入当期损益的金融资产_2022"],
            data_set["其他货币资金_2022"],
            data_set["卖出回购金融资产款_2022"],
            data_set["特定行业或公司现金盈余不做调整扣除的部分加回_2022"],
            data_set["经营租赁调整_2022"],
            data_set["永续债_2022"]
        )

        raw_data_2022["资本"] = Capital(
            data_set["所有者权益合计_2022"],
            data_set["短期借款_2022"],
            data_set["应付利息_2022"],
            data_set["一年内到期的长期借款_2022"],
            data_set["一年内到期的应付债券_2022"],
            data_set["其它流动负债短期应付债券_2022"],
            data_set["一年内应付融资租赁款_2022"],
            data_set["长期借款_2022"],
            data_set["应付债券_2022"],
            data_set["长期应付融资租赁款_2022"],
            data_set["递延所得税负债_2022"],
            data_set["重大合同及履行状况担保情况_2022"],
            data_set["货币资金_2022"],
            data_set["以公允价值计量且其变动计入当期损益的金融资产_2022"],
            data_set["其他货币资金_2022"],
            data_set["卖出回购金融资产款_2022"],
            data_set["特定行业或公司现金盈余不做调整扣除的部分加回_2022"],
            data_set["经营租赁调整_2022"],
            data_set["永续债_2022"]
        )

        raw_data_2022["EBITDA利润率"] = EBITDA_profit_rate(
            raw_data_2022["EBITDA"],
            data_set["营业收入_2022"]
        )

        raw_data_2022["资本回报率"] = Capital_RR(
            raw_data_2022["EBIT"],
            raw_data_2022["资本"]
        )

        raw_data_2022["经营活动产生的资金/债务"] = Operating_cash_to_debt(
            raw_data_2022["经营活动产生的现金(FFO)"],
            raw_data_2022["总负债"]
        )

        raw_data_2022["债务/息税摊折前利润"] = debt_to_PBITA(
            raw_data_2022["总负债"],
            raw_data_2022["EBITDA"]
        )

        raw_data_2022["自由运营现金流/债务"] = FOCF_to_debt(
            raw_data_2022["自由运营现金流(FOCF)"],
            raw_data_2022["总负债"]
        )

        raw_data_2022["息税摊折前利润 / 利息支出"] = EBITDA_to_interest_expense(
            raw_data_2022["EBITDA"],
            data_set["财务费用_2022"],
            data_set["资本化利息_2022"],
            data_set["经营租赁的利息调整_2022"]
        )

        raw_data_2023["EBITDA"] = EBITDA(
            data_set["营业利润_2023"],
            data_set["财务费用_2023"],
            data_set["折旧费_2023"],
            data_set["公允价值变动_2023"],
            data_set["投资收益_2023"],
            data_set["取得投资收益收到的现金_2023"],
            data_set["政府补助_2023"],
            data_set["经营租赁费用调整_2023"],
            data_set["资本化开发成本_2023"],
            data_set["勘探费用_2023"]
        )

        raw_data_2023["EBIT"] = EBIT(
            data_set["营业利润_2023"],
            data_set["财务费用_2023"],
            data_set["利息收入_2023"],
            data_set["公允价值变动_2023"],
            data_set["投资收益_2023"],
            data_set["对联营企业和合营企业的投资收益_2023"],
            data_set["政府补助_2023"],
            data_set["经营租赁的利息调整_2023"]
        )

        raw_data_2023["自由运营现金流(FOCF)"] = FOCF(
            data_set["经营活动产生的现金流量净额_2023"],
            data_set["购建固定资产无形资产和其他长期资产支付的现金_2023"],
            data_set["取得投资收益收到的现金_2023"],
            data_set["收到其他与投资活动有关的现金附注利息部分_2023"],
            data_set["分配股利利润或偿付利息支付的现金_2023"],
            data_set["对所有者或股东的分配_2023"],
            data_set["经营租赁折旧调整_2023"],
            data_set["资本化开发成本_2023"]
        )

        raw_data_2023["经营活动产生的现金(FFO)"] = FFO(
            raw_data_2023["EBITDA"],
            data_set["财务费用_2023"],
            data_set["利息收入_2023"],
            data_set["所得税费用_2023"],
            data_set["经营租赁费用调整_2023"],
            data_set["经营租赁折旧调整_2023"],
            data_set["资本化利息_2023"]
        )

        raw_data_2023["总负债"] = Total_debt(
            data_set["短期借款_2023"],
            data_set["应付利息_2023"],
            data_set["一年内到期的长期借款_2023"],
            data_set["一年内到期的应付债券_2023"],
            data_set["其它流动负债短期应付债券_2023"],
            data_set["一年内应付融资租赁款_2023"],
            data_set["长期借款_2023"],
            data_set["应付债券_2023"],
            data_set["长期应付融资租赁款_2023"],
            data_set["重大合同及履行状况担保情况_2023"],
            data_set["货币资金_2023"],
            data_set["以公允价值计量且其变动计入当期损益的金融资产_2023"],
            data_set["其他货币资金_2023"],
            data_set["卖出回购金融资产款_2023"],
            data_set["特定行业或公司现金盈余不做调整扣除的部分加回_2023"],
            data_set["经营租赁调整_2023"],
            data_set["永续债_2023"]
        )

        raw_data_2023["资本"] = Capital(
            data_set["所有者权益合计_2023"],
            data_set["短期借款_2023"],
            data_set["应付利息_2023"],
            data_set["一年内到期的长期借款_2023"],
            data_set["一年内到期的应付债券_2023"],
            data_set["其它流动负债短期应付债券_2023"],
            data_set["一年内应付融资租赁款_2023"],
            data_set["长期借款_2023"],
            data_set["应付债券_2023"],
            data_set["长期应付融资租赁款_2023"],
            data_set["递延所得税负债_2023"],
            data_set["重大合同及履行状况担保情况_2023"],
            data_set["货币资金_2023"],
            data_set["以公允价值计量且其变动计入当期损益的金融资产_2023"],
            data_set["其他货币资金_2023"],
            data_set["卖出回购金融资产款_2023"],
            data_set["特定行业或公司现金盈余不做调整扣除的部分加回_2023"],
            data_set["经营租赁调整_2023"],
            data_set["永续债_2023"]
        )

        raw_data_2023["EBITDA利润率"] = EBITDA_profit_rate(
            raw_data_2023["EBITDA"],
            data_set["营业收入_2023"]
        )

        raw_data_2023["资本回报率"] = Capital_RR(
            raw_data_2023["EBIT"],
            raw_data_2023["资本"]
        )

        raw_data_2023["经营活动产生的资金/债务"] = Operating_cash_to_debt(
            raw_data_2023["经营活动产生的现金(FFO)"],
            raw_data_2023["总负债"]
        )

        raw_data_2023["债务/息税摊折前利润"] = debt_to_PBITA(
            raw_data_2023["总负债"],
            raw_data_2023["EBITDA"]
        )

        raw_data_2023["自由运营现金流/债务"] = FOCF_to_debt(
            raw_data_2023["自由运营现金流(FOCF)"],
            raw_data_2023["总负债"]
        )

        raw_data_2023["息税摊折前利润 / 利息支出"] = EBITDA_to_interest_expense(
            raw_data_2023["EBITDA"],
            data_set["财务费用_2023"],
            data_set["资本化利息_2023"],
            data_set["经营租赁的利息调整_2023"]
        )

        # 输出数据 用于查看
        print("--------2022年----------")
        for i in raw_data_2022:
            print(f"{i}: {raw_data_2022[i]}")

        print("\n--------2023年----------")

        for i in raw_data_2023:
            print(f"{i}: {raw_data_2022[i]}")

        # 写入数据
        report = load_workbook(report_path)
        sheet = report["Inputs"]
        sheet["E46"] = raw_data_2023["EBITDA利润率"]
        sheet["F46"] = raw_data_2022["EBITDA利润率"]
        sheet["E47"] = raw_data_2023["资本回报率"]
        sheet["F47"] = raw_data_2022["资本回报率"]
        sheet["E48"] = 营业收入_2023
        sheet["F48"] = 营业收入_2022
        sheet["E49"] = 总资产_2023
        sheet["F49"] = 总资产_2022
        sheet["E54"] = raw_data_2023["经营活动产生的资金/债务"]
        sheet["F54"] = raw_data_2022["经营活动产生的资金/债务"]
        sheet["E55"] = raw_data_2023["债务/息税摊折前利润"]
        sheet["F55"] = raw_data_2022["债务/息税摊折前利润"]
        sheet["E56"] = raw_data_2023["自由运营现金流/债务"]
        sheet["F56"] = raw_data_2022["自由运营现金流/债务"]
        sheet["E57"] = raw_data_2023["息税摊折前利润 / 利息支出"]
        sheet["F57"] = raw_data_2022["息税摊折前利润 / 利息支出"]
        sheet["E58"] = raw_data_2023["经营活动产生的现金(FFO)"]
        sheet["F58"] = raw_data_2022["经营活动产生的现金(FFO)"]
        sheet["E59"] = raw_data_2023["总负债"]
        sheet["F59"] = raw_data_2022["总负债"]
        sheet["E60"] = raw_data_2023["EBITDA"]
        sheet["F60"] = raw_data_2022["EBITDA"]
        sheet["E63"] = 营业收入_2023
        sheet["F63"] = 营业收入_2022
        sheet["E64"] = 总资产_2023
        sheet["F64"] = 总资产_2022

        sheet["B46"] = raw_data_2023["EBITDA利润率"]
        sheet["B47"] = raw_data_2023["资本回报率"]
        sheet["B48"] = 营业收入_2023
        sheet["B49"] = 总资产_2023
        sheet["B54"] = raw_data_2023["经营活动产生的资金/债务"]
        sheet["B55"] = raw_data_2023["债务/息税摊折前利润"]
        sheet["B56"] = raw_data_2023["自由运营现金流/债务"]
        sheet["B57"] = raw_data_2023["息税摊折前利润 / 利息支出"]
        sheet["B58"] = raw_data_2023["经营活动产生的现金(FFO)"]
        sheet["B59"] = raw_data_2023["总负债"]
        sheet["B60"] = raw_data_2023["EBITDA"]
        sheet["B63"] = 营业收入_2023
        sheet["B64"] = 总资产_2023

        # 排查数据
            # 如果发生以下任意一种情况，请将财务比率输入为 "NM"：
                # a) 如果任何财务比率的分母为零；
                # b) 如果任意财务比率的分子和分母均为负数。
        for i in sheet["B46:F64"]:
            for j in i:
                print(j.value)
                if j.value == "inf" or j.value == "-inf" or j.value == "nan":
                    j.value = "NM"


        report.save(report_path)
        


# excel里列的字母转数字
def col_to_num(col_str):
    num = 0
    for i, c in enumerate(reversed(col_str)):
        num += (ord(c) - ord('A') + 1) * (26 ** i)
    return num - 1

# excel里行的数字转换
def row_to_num(num):
    return int(num) - 2

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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    sys.exit(app.exec_())
