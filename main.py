import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox, QHBoxLayout, \
    QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QFont, QPixmap, QColor

from excel_to_list import excel_to_list
from text_to_speech import text_to_speech


class TimerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.hours = 0
        self.minutes = 0
        self.academic_rules_file = None
        self.university_logo_file = None
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        academic_rules_layout = QHBoxLayout()
        academic_rules_label = QLabel("Academic Rules (Excel):", self)
        academic_rules_label.setFont(QFont('Arial', 14))
        academic_rules_button = QPushButton("Upload", self)
        academic_rules_button.clicked.connect(self.upload_academic_rules)

        academic_rules_layout.addWidget(academic_rules_label)
        academic_rules_layout.addWidget(academic_rules_button)

        university_logo_layout = QHBoxLayout()
        university_logo_label = QLabel("University Logo (PNG or JPEG):", self)
        university_logo_label.setFont(QFont('Arial', 14))
        university_logo_button = QPushButton("Upload", self)
        university_logo_button.clicked.connect(self.upload_university_logo)
        university_logo_layout.addWidget(university_logo_label)
        university_logo_layout.addWidget(university_logo_button)
        hours_minutes_layout = QHBoxLayout()
        hours_label = QLabel("Hours:", self)
        hours_label.setFont(QFont('Arial', 14))
        minutes_label = QLabel("Minutes:", self)
        minutes_label.setFont(QFont('Arial', 14))
        self.hours_combo = QComboBox(self)
        self.hours_combo.setFont(QFont('Arial', 14))
        for i in range(24):
            self.hours_combo.addItem(str(i))
        self.minutes_combo = QComboBox(self)
        self.minutes_combo.setFont(QFont('Arial', 14))
        for i in range(60):
            self.minutes_combo.addItem(str(i))
        hours_minutes_layout.addWidget(hours_label)
        hours_minutes_layout.addWidget(self.hours_combo)
        hours_minutes_layout.addWidget(minutes_label)
        hours_minutes_layout.addWidget(self.minutes_combo)
        layout.addLayout(academic_rules_layout)
        layout.addLayout(university_logo_layout)
        layout.addLayout(hours_minutes_layout)
        start_button = QPushButton("Start Timer", self)
        start_button.clicked.connect(self.start_timer)
        start_button.setFont(QFont('Arial', 16))
        layout.addWidget(start_button)
        self.setLayout(layout)
        self.setGeometry(100, 100, 600, 250)

    def upload_academic_rules(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "Upload Academic Rules (Excel)", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            self.academic_rules_file = file_name
            print(f"Uploaded Academic Rules: {self.academic_rules_file}")

    def upload_university_logo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "Upload University Logo (PNG or JPEG)", "", "Image Files (*.png *.jpeg *.jpg);;All Files (*)", options=options)
        if file_name:
            self.university_logo_file = file_name
            pixmap = QPixmap(self.university_logo_file)
            pixmap = pixmap.scaledToWidth(100)
            university_logo_label = QLabel(self)
            university_logo_label.setPixmap(pixmap)

    def start_timer(self):
        if not self.academic_rules_file or not self.university_logo_file:
            message = "Please upload both academic rules and university logo before starting the timer!!!"
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setText(message)
            msg_box.setWindowTitle('Error')
            msg_box.exec_()
            return
        self.hours = int(self.hours_combo.currentText())
        self.minutes = int(self.minutes_combo.currentText())
        self.timer_window = TimerWindow(self.hours, self.minutes, self.academic_rules_file, self.university_logo_file)
        self.timer_window.showFullScreen()
        self.close()

class TimerWindow(QMainWindow):
    def __init__(self, hours, minutes, academic_rules_file, university_logo_file):
        super().__init__()
        self.hours = hours
        self.minutes = minutes
        self.seconds = hours * 3600 + minutes * 60
        self.timer_label = QLabel(self)
        self.timer_label.setAlignment(Qt.AlignCenter)
        self.timer_label.setFont(QFont("SansSerif", 50))  # Set the font family and size
        self.timer_label.setStyleSheet("color: white;")  # Set font color to white
        self.update_timer_label()
        exit_button = QPushButton("Exit", self)
        exit_button.clicked.connect(self.close_app)
        exit_button.setFont(QFont('Arial', 24))
        exit_button.setMaximumWidth(200)

        # academic_rules_file, university_logo_file
        #set the background color of the entire window to black
        palette = self.palette()
        palette.setColor(self.backgroundRole(), QColor(0, 0, 0))
        self.setPalette(palette)
        layout = QVBoxLayout()
        #showlogo logo
        image_label = QLabel(self)
        image = QPixmap(university_logo_file)
        image_label.setFixedSize(400, 300)
        image = image.scaled(image_label.size(), aspectRatioMode=1, transformMode=0)
        image_label.setPixmap(image)
        layout.addWidget(image_label,alignment=Qt.AlignCenter)
        layout.addWidget(self.timer_label, alignment=Qt.AlignCenter)
        layout.addWidget(exit_button, alignment=Qt.AlignCenter)
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)



        #before starting the timer
        welcome_message = "Greetings all, wish you all luck at your exam, before starting the exam I need to remind you with our regulations!"
        text_to_speech(welcome_message)
        rules = excel_to_list(academic_rules_file)
        if rules:
            for rule in rules:
                text_to_speech(rule)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_timer)
        self.timer.start(1000)

    def update_timer(self):
        self.seconds -= 1
        if self.seconds <= 0:
            self.timer.stop()
            print("Timer finished!")
        else:
            self.update_timer_label()

    def update_timer_label(self):
        hours = self.seconds // 3600
        minutes = (self.seconds % 3600) // 60
        seconds = self.seconds % 60
        self.timer_label.setText(f"{hours:02d}:{minutes:02d}:{seconds:02d}")

    def close_app(self):
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TimerApp()
    window.show()
    sys.exit(app.exec_())
