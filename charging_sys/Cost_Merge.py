import pandas as pd
from PyQt5.Qt import *


class CostMerge():

    def __init__(self):
        self.xlsx_ppf = None
        self.xlsx_pmf = None
        self.xlsx_sfbz = None
        # self.xlsx_merged = None
        self.df_ppf = None
        self.df_pmf = None
        self.df_merged = None
        self.df_sfbz = None
        self.df_year_stat = None
        self.df_total_stat = None

    # 读取物业费和日常维修基金数据
    def read_data(self, addr_ppf, addr_pmf):
        # 刷新界面
        QApplication.processEvents()

        self.xlsx_ppf = pd.ExcelFile(addr_ppf)
        self.df_ppf = pd.read_excel(self.xlsx_ppf, sheet_name="去除断档后")

        # 刷新界面
        QApplication.processEvents()

        self.xlsx_pmf = pd.ExcelFile(addr_pmf)
        self.df_pmf = pd.read_excel(self.xlsx_pmf, sheet_name="日常维修基金")

        # 刷新界面
        QApplication.processEvents()

    def data_processing(self, sfbz_addr):
        # 物业费的列重新索引以及列名重命名
        self.df_ppf = self.df_ppf.reindex(columns=["房产名称", "客户名称", "应收年份", "未缴金额", "起始欠费日期", "终止欠费日期", "欠费区间"])
        self.df_ppf = self.df_ppf.rename(columns={"未缴金额": "物业费"})

        # 维修基金列重命名
        self.df_pmf = self.df_pmf.reindex(columns=["房产名称", "客户名称", "应收年份", "未缴金额"])
        self.df_pmf = self.df_pmf.rename(columns={"未缴金额": "日常维修基金"})

        # 根据key["房产名称","客户名称","应收年份"]将物业费和维修基金进行左连接
        self.df_merged = pd.merge(self.df_ppf, self.df_pmf, on=["房产名称", "客户名称", "应收年份"], how="left")

        # 物业费、维修基金数据合并后重命名
        self.df_merged = self.df_merged.reindex(columns=["房产名称", "客户名称", "应收年份", "物业费", "日常维修基金", "起始欠费日期", "终止欠费日期", "欠费区间"])

        # 读取收费标准
        self.xlsx_sfbz = pd.ExcelFile(sfbz_addr)
        self.df_sfbz = pd.read_excel(self.xlsx_sfbz, sheet_name="收费标准选用")

        # 刷新界面
        QApplication.processEvents()

        # 收费标准的NA值填充
        self.df_sfbz.fillna(0, inplace=True)

        # 去除收费标准中异常的信息录入值
        self.df_sfbz = self.df_sfbz[self.df_sfbz["收费标准名称"] != 0]

        # 对收费标准列进行重新索引以及列的重命名
        self.df_sfbz = self.df_sfbz.reindex(columns=["房产名称", "产权归属", "物业面积", "收费标准名称", "公共维修收费标准"])
        self.df_sfbz = self.df_sfbz.rename(columns={"产权归属": "客户名称", "收费标准名称": "物业费收费标准"})

        # 抛弃原断号索引，使用新的连续索引
        self.df_sfbz = self.df_sfbz.reindex(columns=["序号", "房产名称", "客户名称", "物业面积", "物业费收费标准", "公共维修收费标准"])
        self.df_sfbz["序号"] = range(len(self.df_sfbz))
        self.df_sfbz = self.df_sfbz.set_index("序号")

        # 包含"物业"二字的bool信息
        self.df_sfbz = self.df_sfbz.reindex(columns=["房产名称", "客户名称", "物业面积", "物业费收费标准", "公共维修收费标准", "是否包含日常维修字符串"])
        self.df_sfbz["是否包含日常维修字符串"] = self.df_sfbz["物业费收费标准"].str.contains("物业")

        # 将收费标准重新整理位置
        for i in range(1, len(self.df_sfbz), 2):
            if self.df_sfbz.iloc[i, 5] == True:
                self.df_sfbz.iloc[i, 4] = self.df_sfbz.iloc[i - 1, 3]
            else:
                self.df_sfbz.iloc[i, 4] = self.df_sfbz.iloc[i, 3]
                self.df_sfbz.iloc[i, 3] = self.df_sfbz.iloc[i - 1, 3]

            # 刷新界面
            QApplication.processEvents()

        # 将收费标准带NA值的行清除
        self.df_sfbz = self.df_sfbz.dropna()

        # 将物业费、日常维修费清单与收费标准连接
        df1 = pd.merge(self.df_merged, self.df_sfbz, left_on=["房产名称", "客户名称"], right_on=["房产名称", "客户名称"], how="left")

        # 对带有收费标准的物业费清单进行重新索引
        self.df_year_stat = df1.reindex(columns=["房产名称", "客户名称", "应收年份", "物业费", "日常维修基金", "起始欠费日期", "终止欠费日期", "欠费区间", "物业面积", "公共维修费单价", \
                                    "物业费收费标准", "公共维修收费标准", "公共维修费是否异常"])

        # 正则表达式
        import re
        pattern = r"[0-9]{1}.[0-9]{1,2}|[0-9]{1}"
        regex = re.compile(pattern)

        # 将匹配到的正则表达式提取出来
        self.df_year_stat["公共维修费单价"] = self.df_year_stat["公共维修收费标准"].str.findall(pattern)

        # 将没有匹配到的收费标准NA数据填充为0
        self.df_year_stat = self.df_year_stat.fillna(0)

        # 创建没有匹配到的收费标准NA数据填充为0的bool列表
        s1 = (self.df_year_stat.iloc[:, 9] == 0)

        # 将匹配到的正则表达式的值从列表中提取出来
        for i in range(len(self.df_year_stat["公共维修费单价"])):
            if s1[i] == False:
                self.df_year_stat.iloc[i, 9] = self.df_year_stat.iloc[i, 9][0]

        #  进行将起始年月日和终止年月日的字符串转化为datetime类型
        from datetime import datetime
        from datetime import timedelta

        for i in range(len(self.df_year_stat)):
            self.df_year_stat.iloc[i, 5] = datetime.strptime(self.df_year_stat.iloc[i, 5], "%Y-%m-%d")
            self.df_year_stat.iloc[i, 6] = datetime.strptime(self.df_year_stat.iloc[i, 6], "%Y-%m-%d")

        # 刷新界面
        QApplication.processEvents()

        # 检验公共维修基金的账期是否与物业费的账期相同
        for i in range(len(self.df_year_stat)):
            self.df_year_stat.iloc[i, 12] = ((self.df_year_stat.iloc[i, 8] * float(self.df_year_stat.iloc[i, 9])) / 365) * ((self.df_year_stat.iloc[i, 6] - self.df_year_stat.iloc[i, 5] + timedelta(1)) / timedelta(1))

        # 给予检测后的标识
        for i in range(len(self.df_year_stat)):
            self.df_year_stat.iloc[i, 12] = abs((self.df_year_stat.iloc[i, 12] - self.df_year_stat.iloc[i, 4])) < 2

        # 欠费按户进行聚合总欠费
        self.df_total_stat = self.df_year_stat.groupby(["房产名称", "客户名称"]).sum()[["物业费", "日常维修基金"]]

        # 对总欠费列表添加列和去除索引
        self.df_total_stat = self.df_total_stat.reindex(columns=["物业费", "日常维修基金", "起始时间", "终止时间", "欠费区间"])
        self.df_total_stat = self.df_total_stat.reset_index()

        # 计算欠费区间的起始和终止时间
        for i in range(len(self.df_total_stat)):
            self.df_total_stat.iloc[i, 4] = ((self.df_year_stat[(self.df_year_stat.iloc[:, 0] == self.df_total_stat.iloc[i, 0]) & (self.df_year_stat.iloc[:, 1] == self.df_total_stat.iloc[i, 1])]).iloc[:, 5]).min()
            self.df_total_stat.iloc[i, 5] = ((self.df_year_stat[(self.df_year_stat.iloc[:, 0] == self.df_total_stat.iloc[i, 0]) & (self.df_year_stat.iloc[:, 1] == self.df_total_stat.iloc[i, 1])]).iloc[:, 6]).max()

        # 刷新界面
        QApplication.processEvents()

        # 转换日期为字符串
        for i in range(len(self.df_total_stat)):
            self.df_total_stat.iloc[i, 4] = (self.df_total_stat.iloc[i, 4]).strftime("%Y-%m-%d")
            self.df_total_stat.iloc[i, 5] = (self.df_total_stat.iloc[i, 5]).strftime("%Y-%m-%d")

        # 连接日期字符串
        for i in range(len(self.df_total_stat)):
            self.df_total_stat.iloc[i, 6] = self.df_total_stat.iloc[i, 4] + '~' + self.df_total_stat.iloc[i, 5]

        # 刷新界面
        QApplication.processEvents()

    # 写入文件并保存
    def write_data(self, addr):
        writer = pd.ExcelWriter(addr)
        self.df_year_stat.to_excel(writer, sheet_name="按年统计欠费清单")

        # 刷新界面
        QApplication.processEvents()

        self.df_total_stat.to_excel(writer, sheet_name="按户统计总欠费清单")

        writer.save()


