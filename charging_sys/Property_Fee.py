import numpy as np
import pandas as pd
import time
from datetime import datetime
from datetime import timedelta
from PyQt5.Qt import *


class PropertyFee():
    # processing_finished_signal = pyqtSignal()

    def __init__(self):
        self.xlsx = None
        self.df_resource = None
        self.df_clear_pre = None
        self.df_clear_ed = None

    def read_data(self, addr):
        # 刷新界面
        QApplication.processEvents()

        # 读取数据
        self.xlsx = pd.ExcelFile(addr)
        self.df_resource = pd.read_excel(self.xlsx, '项目应收清单')

        # 刷新界面
        QApplication.processEvents()

    def data_processing(self):
        # 选取物业收费项目
        df2 = self.df_resource[((self.df_resource["收费项目"] == "0102_小高层物业服务费") | (self.df_resource["收费项目"] == "0103_多层物业服务费") |
            (self.df_resource["收费项目"] == "0112_商务公寓物业服务费") | (self.df_resource["收费项目"] == "0108_联排别墅物业服务费") |
            (self.df_resource["收费项目"] == "0104_商业物业服务费") | (self.df_resource["收费项目"] == "0114_叠排别墅物业服务费"))]

        # 索引（应收月份）重命名
        df2 = df2.rename(columns={"应收月份": "应收年月"})

        # 拆分欠费年月
        df2["应收年份"] = df2["应收年月"].astype('str').str[:4]
        df2["应收月份"] = df2["应收年月"].astype('str').str[4:6]

        # 数据重新索引
        df3 = df2.reindex(columns=["房产名称", "客户名称", "收费项目", "应收年月", "应收年份", "应收月份", "未缴金额", "挂起金额"])

        # NA处理
        df3 = df3.fillna(0)

        # 将未缴金额为0的剔除
        df3 = df3[df3["未缴金额"] != 0]

        # 按年对物业欠费聚合
        self.df_clear_pre = df3.groupby(["房产名称", "客户名称", "应收年份"], as_index=False)[["未缴金额", "挂起金额"]].sum()

        # 聚合运算后，用重新索引 添加4个时间索引列
        self.df_clear_pre = self.df_clear_pre.reindex(columns=["房产名称", "客户名称", "应收年份", "未缴金额", "挂起金额", "起始月份", "终止月份", "起始欠费日期", "终止欠费日期", "欠费区间"])
        self.df_clear_pre = self.df_clear_pre.fillna(0)

        self.df_clear_pre["起始月份"] = self.df_clear_pre["起始月份"].astype("int64")
        self.df_clear_pre["终止月份"] = self.df_clear_pre["终止月份"].astype("int64")

        # 刷新界面
        QApplication.processEvents()

        # 按年计算欠费的起始、终止月份
        for i in range(len(self.df_clear_pre)):
            agg_value = (df3[((df3["房产名称"] == self.df_clear_pre.iloc[i]["房产名称"]) & (df3["客户名称"] == self.df_clear_pre.iloc[i]["客户名称"]) & \
                              (df3["应收年份"] == self.df_clear_pre.iloc[i]["应收年份"]))]).loc[:, "应收月份"].min()
            self.df_clear_pre.iloc[i, 5] = int(agg_value)

            # 刷新界面
            QApplication.processEvents()

        for i in range(len(self.df_clear_pre)):
            agg_value = (df3[((df3["房产名称"] == self.df_clear_pre.iloc[i]["房产名称"]) & (df3["客户名称"] == self.df_clear_pre.iloc[i]["客户名称"]) & \
                              (df3["应收年份"] == self.df_clear_pre.iloc[i]["应收年份"]))]).loc[:, "应收月份"].max()
            self.df_clear_pre.iloc[i, 6] = int(agg_value)

            # 刷新界面
            QApplication.processEvents()

        # 按年计算欠费的起始年月日
        for i in range(len(self.df_clear_pre)):
            section = f"{self.df_clear_pre.iloc[i, 2]}-{self.df_clear_pre.iloc[i, 5]:02d}-01"
            self.df_clear_pre.iloc[i, 7] = section

        # 按年计算欠费的终止年月日
        m_days = {1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31}

        for i in range(len(self.df_clear_pre)):
            section = f"{self.df_clear_pre.iloc[i, 2]}-{self.df_clear_pre.iloc[i, 6]:02d}-{m_days[self.df_clear_pre.iloc[i, 6]]:d}"
            self.df_clear_pre.iloc[i, 8] = section

        # 按年计算欠费区间
        for i in range(len(self.df_clear_pre)):
            self.df_clear_pre.iloc[i, 9] = '~'.join([self.df_clear_pre.iloc[i, 7], self.df_clear_pre.iloc[i, 8]])

        # 去除挂起的金额,房号为0的行，未缴金额为0的行
        self.df_clear_ed = self.df_clear_pre[self.df_clear_pre["挂起金额"] == 0]
        self.df_clear_ed = self.df_clear_ed[self.df_clear_ed["房产名称"] != 0]
        self.df_clear_ed = self.df_clear_ed[self.df_clear_ed["未缴金额"] != 0]

        self.df_clear_ed = self.df_clear_ed.reindex(
            columns=["序号", "房产名称", "客户名称", "应收年份", "未缴金额", "挂起金额", "起始月份", "终止月份", "起始欠费日期", "终止欠费日期", "欠费区间"])

        self.df_clear_ed.loc[:, "序号"] = [x for x in range(len(self.df_clear_ed))]
        self.df_clear_ed = self.df_clear_ed.set_index(['序号'])

        # 去除断档金额
        n = len(self.df_clear_ed) - 1
        for i in range(n, -1, -1):
            if i == n:
                if datetime.strptime(self.df_clear_ed.iloc[i, 8], '%Y-%m-%d') != datetime(2020, 12, 31):
                    self.df_clear_ed.drop(i, inplace=True)
            else:
                if self.df_clear_ed.iloc[i, 0] == self.df_clear_ed.iloc[i + 1, 0]:
                    if datetime.strptime(self.df_clear_ed.iloc[i + 1, 7], '%Y-%m-%d') - datetime.strptime(self.df_clear_ed.iloc[i, 8],'%Y-%m-%d') != timedelta(1):
                        self.df_clear_ed.drop(i, inplace=True)
                else:
                    if datetime.strptime(self.df_clear_ed.iloc[i, 8], '%Y-%m-%d') != datetime(2020, 12, 31):
                        self.df_clear_ed.drop(i, inplace=True)

            # 刷新界面
            QApplication.processEvents()

        # 序号整理
        self.df_clear_ed.loc[:, "序号"] = [x for x in range(len(self.df_clear_ed))]
        self.df_clear_ed = self.df_clear_ed.set_index(['序号'])

        # self.processing_finished_signal.emit()

    def write_data(self, addr):
        # 将数据写入excel
        writer = pd.ExcelWriter(addr)
        self.df_resource.to_excel(writer, sheet_name='项目应收清单')

        # 刷新界面
        QApplication.processEvents()

        self.df_clear_pre.to_excel(writer, sheet_name='断档去除前')

        # 刷新界面
        QApplication.processEvents()

        self.df_clear_ed.to_excel(writer, sheet_name='去除断档后')

        writer.save()


