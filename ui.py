import sys
from PyQt5.QtWidgets import QApplication, QWidget

if __name__ == '__main__':

    #App object.
    app = QApplication(sys.argv)

    #base class of all user interface objects in PyQt5
    #We provide the default constructor for QWidget
    #The default constructor has no parent
    #A widget with no parent is called a window.
    w = QWidget()

    #Window size
    w.resize(500, 500)

    #Window position on screen
    w.move(500, 500)

    #Window title
    w.setWindowTitle('CATITA')

    #Window on screen
    w.show()

    #Close window
    sys.exit(app.exec_())
