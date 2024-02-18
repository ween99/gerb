import sys
import os
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QFileDialog, QGridLayout, QMenuBar
from PyQt5.QtGui import QPixmap
from PyQt5 import QtGui, QtCore
from PyQt5.QtGui import QCursor

#dynamically changing widgets
widgets = {
    "logo": [],
    "Process": [],
}

#GUI application
app = QApplication(sys.argv)

#window and settings
w = QWidget()
w.setWindowTitle("Process")
w.setFixedWidth(550)
w.move(470, 270)
w.setStyleSheet("background: #121519;")
w.setWindowIcon(QtGui.QIcon('earth.png'))

#grid layout
grid = QGridLayout()


def clear_widgets():
    ''' hide all existing widgets'''
    for widget in widgets:
        if widgets[widget] != []:
            widgets[widget][-1].hide()
        for i in range(0, len(widgets[widget])):
            widgets[widget].pop()

def show_frame1():
    '''display frame 1'''
    clear_widgets()
    frame1()



def frame1():
    #logo widget
    image = QPixmap("earth.png")
    logo = QLabel()
    logo.setPixmap(image)
    logo.setAlignment(QtCore.Qt.AlignCenter)
    logo.setStyleSheet("margin-top: 0px;")
    widgets["logo"].append(logo)

    #button widget
    Process = QPushButton("Process")
    Process.setCursor(QCursor(QtCore.Qt.OpenHandCursor))
    Process.setStyleSheet(
        '''
        *{
            border: 7px solid '#000000';
            border-radius: 25px;
            font-size: 15px;
            color: 'white';
            padding: 10px 0;
            margin: 0px 20px;
        }
        *:hover{
            background: '#00aaf0';
        }
        '''
    )
    #button callback
    Process.clicked.connect(show_frame1)
    widgets["Process"].append(Process)

    #place global widgets on the grid
    grid.addWidget(widgets["logo"][-1], 0, 0, 1, 2)
    grid.addWidget(widgets["Process"][-1], 3, 0, 1, 0) #page 1


def click1():
    os.system('cmd /c "mstsc /v:10.10.126.135"')

frame1()

w.setLayout(grid)

w.show()
sys.exit(app.exec()) #terminate the app
