import sys
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QWidget, QFileDialog, QGridLayout, QTextEdit
from PyQt5.QtGui import QPixmap, QIcon, QFont, QCursor
from PyQt5 import QtGui, QtCore

# dynamically changing widgets
widgets = {
    "logo": [],
    "Process": [],
    "text_edit": [],
}

def clear_widgets():
    ''' hide all existing widgets'''
    for widget in widgets:
        if widgets[widget] != []:
            widgets[widget][-1].hide()
        for i in range(0, len(widgets[widget])):
            widgets[widget].pop()

def show_frame1(grid):
    '''display frame 1'''
    clear_widgets()
    frame1(grid)

def frame1(grid):
    # logo widget
    image = QPixmap("earth.png")
    logo = QLabel()
    logo.setPixmap(image)
    logo.setAlignment(QtCore.Qt.AlignCenter)
    logo.setStyleSheet("margin-top: 0px;")
    widgets["logo"].append(logo)

    # button widget
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
    # button callback
    Process.clicked.connect(lambda: click1(grid))
    widgets["Process"].append(Process)

    # text_edit widget
    text_edit = QTextEdit()
    text_edit.setStyleSheet(
        '''
        *{
            background-color: #f0f0f0;  
            font-size: 15px;
            color: 'black';
        }
        QScrollBar:vertical {
            background: #d3d3d3; /* Set the background color of the scroll bar */
            width: 20px; /* Set the width of the scroll bar */
        }
        QScrollBar::handle:vertical {
            background: #808080; /* Set the color of the handle (thumb) */
            min-height: 20px; /* Set the minimum height of the handle */
        }
        QScrollBar::add-line:vertical {
            background: #d3d3d3; /* Set the color of the add line */
            height: 20px; /* Set the height of the add line */
        }
        QScrollBar::sub-line:vertical {
            background: #d3d3d3; /* Set the color of the sub line */
            height: 20px; /* Set the height of the sub line */
        }
        '''
    )
    widgets["text_edit"].append(text_edit)

    # place global widgets on the grid
    grid.addWidget(widgets["text_edit"][-1], 1, 0, 2, 0)
    grid.addWidget(widgets["logo"][-1], 0, 0, 1, 2)
    grid.addWidget(widgets["Process"][-1], 3, 0, 1, 0)  # page 1

def click1(grid):
    file_path, _ = QFileDialog.getOpenFileName(None, "Open Text File", "", "Text files (*.txt)")
    if file_path:
        with open(file_path, 'r') as file:
            text = file.read()
            for line in file:
                linestrip = line.strip().split(" ")

            widgets["text_edit"][-1].setPlainText(text)

# GUI application
app = QApplication(sys.argv)

# window and settings
w = QWidget()
w.setWindowTitle("Process")
w.setFixedWidth(550)
w.move(470, 270)
w.setStyleSheet("background: #121519;")
w.setWindowIcon(QtGui.QIcon('earth.png'))

# grid layout
grid = QGridLayout()

frame1(grid)

w.setLayout(grid)
w.show()
sys.exit(app.exec_())