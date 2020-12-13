from PyQt5.Qt import *
from source.direction_for_user import Ui_Form


class DirectionDialog(QWidget, Ui_Form):

    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.setupUi(self)


if __name__ == '__main__':
    import sys

    app = QApplication(sys.argv)

    window = DirectionDialog()
    window.show()

    sys.exit(app.exec_())