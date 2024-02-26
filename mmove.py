import sys
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QWidget, QGridLayout
from PyQt5.QtGui import QCursor, QIcon, QPixmap, QMovie
from PyQt5.QtCore import QThread, pyqtSignal, QMutex, QMutexLocker, Qt
import pyautogui
import time


class MouseMoverThread(QThread):
    update_signal = pyqtSignal(str)
    _lock = QMutex()

    def __init__(self, size=100, delay=5):
        super().__init__()
        self.size = size
        self.delay = delay
        self.active = False

    def run(self):
        self.active = True
        check_interval = 0.1  # Shorter interval for more frequent checks
        steps_per_movement = int(self.delay / check_interval)

        while self.active:
            for _ in range(2):  # Repeat twice for a square
                for movement in [(self.size, 0), (0, self.size), (-self.size, 0), (0, -self.size)]:
                    if not self.active:
                        return  # Exit the loop immediately if stopped
                    pyautogui.moveRel(*movement, duration=1)  # Execute movement

                    # Split the delay into shorter intervals with active checks
                    for _ in range(steps_per_movement):
                        time.sleep(check_interval)
                        if not self.active:
                            return  # Exit the loop immediately if stopped
        self.update_signal.emit("Mouse movement stopped.")

    def stop(self):
        locker = QMutexLocker(self._lock)
        self.active = False
        locker.unlock()


class AnimatedGifLabel(QLabel):
    def __init__(self, gif_path, static_image_path, parent=None):
        super().__init__(parent)
        # Save the paths
        self.gif_path = gif_path
        self.static_image_path = static_image_path
        # Setup the static image
        self.static_image = QPixmap(static_image_path)
        self.setPixmap(self.static_image)  # Start with the static image

    def start_animation(self):
        self.movie = QMovie(self.gif_path)  # Create the movie
        self.setMovie(self.movie)
        self.movie.start()

    def stop_animation(self):
        self.movie.stop()
        self.setPixmap(self.static_image)  # Display the static image


# GUI application setup
app = QApplication(sys.argv)
w = QWidget()
w.setWindowTitle("Work, work, work")
w.setFixedWidth(550)
w.move(470, 270)
w.setStyleSheet("background: #121519;")
w.setWindowIcon(QIcon(r'C:\Users\Puter\Downloads\ezgif-7-0af21d0106-gif-im\frame_0_delay-0.11s.gif'))  # Static image as the window icon

# Layout and widgets
grid = QGridLayout()

# Use the AnimatedGifLabel for the animated GIF
logo = AnimatedGifLabel(
    r'C:\Users\Puter\Downloads\ezgif-7-0af21d0106-gif-im\2310385203c488124b7eafd79947cc48.gif',  # Path to the animated GIF
    r'C:\Users\Puter\Downloads\ezgif-7-0af21d0106-gif-im\frame_0_delay-0.11s.gif'  # Path to the static image
)
logo.setAlignment(Qt.AlignCenter)
logo.setStyleSheet("margin-top: 20px;")

process_button = QPushButton("Start")
process_button.setCursor(QCursor(Qt.OpenHandCursor))
process_button.setStyleSheet('''
    *{
        border: 4px solid '#333';
        border-radius: 15px;
        font-size: 16px;
        color: 'white';
        padding: 15px 25px;
        margin: 10px 50px;
    }
    *:hover{
        background: '#0057b7';
    }
''')

# Mouse movement thread instance
mouse_mover_thread = MouseMoverThread()

def toggle_mouse_movement():
    if not mouse_mover_thread.isRunning():
        process_button.setText("Stop")
        mouse_mover_thread.start()  # Start the thread
        logo.start_animation()  # Start the GIF animation
    else:
        process_button.setText("Start")
        mouse_mover_thread.stop()  # Signal the thread to stop
        logo.stop_animation()  # Stop the GIF animation and show static image

process_button.clicked.connect(toggle_mouse_movement)

# Adding widgets to the layout
grid.addWidget(logo, 0, 0, 1, 2)
grid.addWidget(process_button, 1, 0, 1, 2)

w.setLayout(grid)
w.show()

sys.exit(app.exec_())