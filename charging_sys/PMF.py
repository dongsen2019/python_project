import pandas as pd
from PyQt5.Qt import *


class PublicMaintenanceFund():
    def __init__(self):
        self.xlsx = None
        self.df_resource = None
        self.df_processed = None

    def read_data(self, addr):
        # 刷新界面
        QApplication.processEvents()

        # 读取数据
        self.xlsx = pd.ExcelFile(addr)
        self.df_resource = pd.read_excel(self.xlsx, "项目应收清单")

        # 刷新界面
        QApplication.processEvents()

    def data_processing(self):
        # 选取需要统计的收费项目
        self.df_resource = self.df_resource[self.df_resource["收费项目"] == "0201_日常收取的维修资金"]

        # 列索引名的修改和添加
        self.df_processed = self.df_resource.rename(columns={"应收月份": "应收年月"})
        self.df_processed = self.df_processed.reindex(columns=["房产名称", "客户名称", "应收年月", "应收年份", "应收月份", "未缴金额", "挂起金额"])
        self.df_processed.fillna(0, inplace=True)

        # 列数据的类型转换
        self.df_processed["应收年份"] = self.df_processed["应收年份"].astype("str")
        self.df_processed["应收月份"] = self.df_processed["应收月份"].astype("str")
        self.df_processed["应收年月"] = self.df_processed["应收年月"].astype("str")

        # 提取年份和月份
        self.df_processed["应收年份"] = self.df_processed["应收年月"].str[0:4]
        self.df_processed["应收月份"] = self.df_processed["应收年月"].str[4:6]

        # 刷新界面
        QApplication.processEvents()

        # 聚合运算
        self.df_processed = self.df_processed.groupby(["房产名称", "客户名称", "应收年份"], as_index=False).sum()

        # 剔除挂起和空白房号的行
        self.df_processed = self.df_processed[self.df_processed["挂起金额"] == 0]
        self.df_processed = self.df_processed[self.df_processed["房产名称"] != 0]
        self.df_processed = self.df_processed[self.df_processed["未缴金额"] != 0]

        # 序号重编
        self.df_processed = self.df_processed.reindex(columns=["序号", "房产名称", "客户名称", "应收年份", "未缴金额", "挂起金额"])
        self.df_processed["序号"] = list(range(len(self.df_processed)))

        self.df_processed = self.df_processed.set_index(["序号"])

        # 刷新界面
        QApplication.processEvents()

    def write_data(self, addr):
        # 将数据写入excel
        writer = pd.ExcelWriter(addr)
        self.df_resource.to_excel(writer, sheet_name="项目应收清单")

        # 刷新界面
        QApplication.processEvents()

        self.df_processed.to_excel(writer, sheet_name="日常维修基金")

        # 刷新界面
        QApplication.processEvents()

        writer.save()


