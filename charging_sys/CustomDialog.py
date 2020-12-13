from PyQt5.Qt import *


class DialogIpt(QDialog):
    def __init__(self):
        QDialog.__init__(self)
        self.setFixedSize(200, 150)
        self.setWindowTitle("信息提示")

        label = QLabel(self)
        label.setText("请不要导入不相关的文件！")
        label.move(30, 45)

        btn = QPushButton(self)
        btn.setText("确定")
        btn.move(63, 100)

        btn.clicked.connect(lambda: self.accept())


class DialogCal(QDialog):
    def __init__(self):
        QDialog.__init__(self)
        self.setFixedSize(200, 150)
        self.setWindowTitle("信息提示")

        label = QLabel(self)
        label.setText("缺少项目应收清单或收费标\n准！")
        label.move(25, 30)

        btn = QPushButton(self)
        btn.setText("确定")
        btn.move(63, 100)

        btn.clicked.connect(lambda: self.accept())


class DialogMerge(QDialog):
    def __init__(self):
        QDialog.__init__(self)
        self.setFixedSize(200, 150)
        self.setWindowTitle("信息提示")

        label = QLabel(self)
        label.setText("生成物业费和维修基金后，\n请不要关闭系统，直接点击\n费用合并生成最终文件！\n\n请关闭界面重新操作一次，选\n择默认保存地址，新文件会覆\n盖老的文件！")
        label.move(30, 10)

        btn = QPushButton(self)
        btn.setText("确定")
        btn.move(63, 120)

        btn.clicked.connect(lambda: self.accept())


class DialogAddr(QDialog):
    def __init__(self):
        QDialog.__init__(self)
        self.setFixedSize(200, 150)
        self.setWindowTitle("信息提示")

        label = QLabel(self)
        label.setText("请先选择要保存的地址！")
        label.move(25, 30)

        btn = QPushButton(self)
        btn.setText("确定")
        btn.move(63, 100)

        btn.clicked.connect(lambda: self.accept())

class DialogAbnormity(QDialog):
    def __init__(self):
        QDialog.__init__(self)
        self.setFixedSize(200, 150)
        self.setWindowTitle("信息提示")

        label = QLabel(self)
        label.setText("异常分析功能已经纳入\n费用合并！")
        label.move(25, 30)

        btn = QPushButton(self)
        btn.setText("确定")
        btn.move(63, 100)

        btn.clicked.connect(lambda: self.accept())
