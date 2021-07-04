import sys
from kiwoom.kiwoom import *
from PyQt5.QtWidgets import *
class Ui_class():
    def __init__(self):
        print("실행할 UI_class 클래스")

        self.app = QApplication(sys.argv)

        self.kiwoom = Kiwoom( )
        self.app.exec_()




