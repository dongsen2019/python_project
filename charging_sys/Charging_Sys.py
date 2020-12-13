from PyQt5.Qt import *
from source.charging_sys import Ui_sys_interface
from PyQt5 import QtCore
from CustomDialog import DialogIpt, DialogCal, DialogAddr, DialogAbnormity
import time
import pandas as pd


# 定义多线程类
class Th(QThread):
    value_signal = pyqtSignal(int)

    def __init__(self):
        QThread.__init__(self)

    def __del__(self):
        self.wait()

    def run(self):
        for i in range(50):
            time.sleep(1)
            self.value_signal.emit(i)   # 每隔1秒输出一个信号，传递出一个int值
            # QApplication.processEvents()


class SysInterface(QWidget, Ui_sys_interface):
    # 自定义信号
    import_data_signal = pyqtSignal(str)
    btn_ppt_cal_clicked_signal = pyqtSignal()
    btn_pmf_cal_clicked_signal = pyqtSignal()
    btn_fee_merge_clicked_signal = pyqtSignal()
    directions_for_use_signal = pyqtSignal()
    btn_browse_signal = pyqtSignal(str)

    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.setupUi(self)

        # 添加列表模型以便后续查找列表内的信息
        self.data_model = QtCore.QStringListModel()

        # 添加进度条
        self.pbar = QProgressBar(self)
        self.pbar.resize(720, 40)
        self.pbar.setRange(0,100)
        self.pbar.move(100, 250)
        self.pbar.hide()

    # 数据导入触发的槽函数
    def btn_import_clicked(self):
        addr = QFileDialog.getOpenFileName(self, "选择一个excel文件", "./", "All(*.*);;Excel 工作簿(*.xlsx)",
                                           "All(*.*)")

        # 发送选择的文件地址传递给 main
        if len(addr[0]) != 0:
            if addr[0].split('/')[-1][0:6] == '项目应收清单' or addr[0].split('/')[-1][0:4] == '收费标准':
                # 将用户选择的地址发送给程序主入口
                self.import_data_signal.emit(addr[0])

                # 将 self.data_list 中的数据取出到 mylist 并添加新加入的文件名
                mylist = self.data_model.stringList()
                mylist.append(addr[0])

                self.data_model.setStringList(mylist)

                # 将新的列表返回给data_list
                self.data_list.setModel(self.data_model)

                """
                for i in self.data_list.children():
                    print(i, isinstance(i, QStringListModel))
                    if isinstance(i, QStringListModel) == True:
                        print(i.stringList())
                """

            # 如果导入的是其他文件，则弹出提示对话框
            else:
                d = DialogIpt()
                d.exec()

    # 点击数据列表显示文件数据
    def data_list_clicked(self, qModelIndex):
        # 根据数据列表中选中的条目读取数据到数据展示文本中
        addr = self.data_model.stringList()[qModelIndex.row()]
        xlsx = pd.ExcelFile(addr)

        # 提取地址中的数据文件名称
        filename = addr.split('/')[-1]
        print(filename)

        # 将对应名称的数据文件中的内容读取到数据文本中
        try:
            if filename[0] == '项':
                df_resource = pd.read_excel(xlsx, '项目应收清单')
                self.textEdit_data.setText(str(df_resource))
            else:
                df_resource = pd.read_excel(xlsx, '收费标准选用')
                self.textEdit_data.setText(str(df_resource))
        except:
            print("传入项目应收清单和收费标准，请不要传入其他文件")

    # 定义点击物业费统计按钮触发的槽函数
    def btn_ppt_cal_clicked(self):
        # 创建文件导入齐备的标签
        xmysqd_mark = 0
        sfbz_mark = 0
        # 遍历导入的文件，查看是否导入文件齐全
        for i in self.data_model.stringList():
            if i.split('/')[-1][0] == '项':
                xmysqd_mark += 1
            elif i.split('/')[-1][0] == '收':
                sfbz_mark += 1
        if xmysqd_mark >= 1 and sfbz_mark >= 1:
            if len(self.lineEdit.text()) != 0 or self.radio_btn_state() == 1:  # 激活保存在原文件夹按钮或者地址文本框不为空，则执行算法，反之弹出提示框
                # 加入进度条
                self.pbar.setValue(0)
                self.pbar.show()
                self.th = Th()
                self.th.value_signal.connect(lambda i: self.pbar.setValue(i*1.9))
                self.th.start()

                self.data_list.hide()
                self.textEdit_data.hide()

                # 将触发物业费统计按钮的信号传递给 main
                self.btn_ppt_cal_clicked_signal.emit()
            else:
                d = DialogAddr()
                d.exec()
        else:
            dlg = DialogCal()
            dlg.exec()

    # 定义点击维修基金统计按钮触发的槽函数
    def btn_pmf_cal_clicked(self):
        # 创建文件导入齐备的标签
        xmysqd_mark = 0
        sfbz_mark = 0
        # 遍历导入的文件，查看是否导入文件齐全
        for i in self.data_model.stringList():
            if i.split('/')[-1][0] == '项':
                xmysqd_mark += 1
            elif i.split('/')[-1][0] == '收':
                sfbz_mark += 1
        if xmysqd_mark >= 1 and sfbz_mark >= 1:    # 文件齐全各标记 >= 1
            if len(self.lineEdit.text()) != 0 or self.radio_btn_state() == 1:   # 激活保存在原文件夹按钮或者地址文本框不为空，则执行算法，反之弹出提示框
                # 加入进度条
                self.pbar.setValue(0)
                self.pbar.show()
                self.th = Th()
                self.th.value_signal.connect(lambda i: self.pbar.setValue(i * 3.9))
                self.th.start()

                self.data_list.hide()
                self.textEdit_data.hide()

                # 将触发公共维修基金按钮的信号传递给 main
                self.btn_pmf_cal_clicked_signal.emit()
            else:
                d = DialogAddr()
                d.exec()
        else:
            dlg = DialogCal()
            dlg.exec()

    def btn_fee_merge_clicked(self):
        # 将触发 费用合并按钮 的信号传递给 main
        self.btn_fee_merge_clicked_signal.emit()

    def directions_for_use(self):
        self.directions_for_use_signal.emit()

    def abnormity_analysis(self):
        d = DialogAbnormity()
        d.exec()

    def btn_exit_clicked(self):
        # 触发 退出按钮 关闭界面
        self.close()

    # 点击浏览触发的槽函数
    def btn_browse(self):
        # 切换radioButton按钮
        self.radioButton_save_custom.setChecked(True)
        fd = QFileDialog(self, "选择一个文件", "../", "Excel 工作簿(*.xlsx)")
        fd.setAcceptMode(QFileDialog.AcceptSave)
        fd.setDefaultSuffix("xlsx")
        fd.open()

        # 当保存的地址被选择，发送信号到main，并让文本框显示地址
        def addr_sure(addr):
            self.btn_browse_signal.emit(addr)
            self.lineEdit.setText(addr)

        fd.fileSelected.connect(addr_sure)

    # 返回哪个radioButton按钮被选中
    def radio_btn_state(self):
        if self.radioButton_save_orig.isChecked():
            return 1
        else:
            return 2

    # 定义费用合并的进度条
    def fee_merge_pbar(self):
        # 加入进度条
        self.pbar.setValue(0)
        self.pbar.show()
        self.th = Th()
        self.th.value_signal.connect(lambda i: self.pbar.setValue(i * 3.9))
        self.th.start()

        self.data_list.hide()
        self.textEdit_data.hide()

    # 当计算模块运行完毕，将进度条拉满，并隐藏进度条，同时再次显示数据文本框和数据文件列表框
    def pbar_val_max(self):
        self.pbar.setValue(self.pbar.maximum())

        # 刷新界面
        QApplication.processEvents()

        time.sleep(3)
        self.pbar.hide()  # 隐藏进度条

        # 再次显示数据文本框和数据文件列表框
        self.data_list.show()
        self.textEdit_data.show()


if __name__ == '__main__':
    import sys

    app = QApplication(sys.argv)

    window = SysInterface()
    window.show()

    sys.exit(app.exec_())
