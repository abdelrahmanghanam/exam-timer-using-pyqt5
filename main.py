import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox, QHBoxLayout, \
    QMainWindow, QFileDialog, QMessageBox, QCheckBox, QLineEdit
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QFont, QPixmap, QColor

from excel_to_list import excel_to_list
from text_to_speech import text_to_speech
import threading
import traceback



class TimerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.hours = 0
        self.minutes = 0
        self.academic_rules_file = None
        self.university_logo_file = None
        self.course_name = ""
        self.course_instructor = ""
        self.time_reminder_enabled = False
        self.leave_allowed_after_half_time = False
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Add course name and instructor input fields
        course_name_label = QLabel("Course Name:", self)
        course_name_label.setFont(QFont('Arial', 12))
        self.course_name_edit = QLineEdit(self)
        self.course_name_edit.setFont(QFont('Arial', 12))

        course_instructor_label = QLabel("Course Instructor:", self)
        course_instructor_label.setFont(QFont('Arial', 12))
        self.course_instructor_edit = QLineEdit(self)
        self.course_instructor_edit.setFont(QFont('Arial', 12))

        layout.addWidget(course_name_label)
        layout.addWidget(self.course_name_edit)
        layout.addWidget(course_instructor_label)
        layout.addWidget(self.course_instructor_edit)

        academic_rules_layout = QHBoxLayout()
        academic_rules_label = QLabel("Academic Rules (Excel):", self)
        academic_rules_label.setFont(QFont('Arial', 12))
        self.academic_rules_button = QPushButton("Upload Rules", self)
        self.academic_rules_button.clicked.connect(self.upload_academic_rules)
        self.academic_rules_button.setMaximumWidth(150)
        academic_rules_layout.addWidget(academic_rules_label)
        academic_rules_layout.addWidget(self.academic_rules_button)

        university_logo_layout = QHBoxLayout()
        university_logo_label = QLabel("University / School Logo (PNG or JPEG):", self)
        university_logo_label.setFont(QFont('Arial', 12))
        self.university_logo_button = QPushButton("Upload Logo", self)
        self.university_logo_button.setMaximumWidth(150)
        self.university_logo_button.clicked.connect(self.upload_university_logo)
        university_logo_layout.addWidget(university_logo_label)
        university_logo_layout.addWidget(self.university_logo_button)

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

        # Add checkboxes for time reminder and leaving allowed after half time
        reminder_checkbox = QCheckBox("Enable Time Reminder", self)
        reminder_checkbox.stateChanged.connect(self.toggle_time_reminder)

        leave_checkbox = QCheckBox("Allow Leaving After Half Time", self)
        leave_checkbox.stateChanged.connect(self.toggle_leave_after_half_time)

        layout.addLayout(academic_rules_layout)
        layout.addLayout(university_logo_layout)
        layout.addLayout(hours_minutes_layout)
        layout.addWidget(reminder_checkbox)
        layout.addWidget(leave_checkbox)

        start_button = QPushButton("Start Timer", self)
        start_button.clicked.connect(self.start_timer)
        start_button.setFont(QFont('Arial', 16))
        layout.addWidget(start_button)

        self.setLayout(layout)
        self.setGeometry(100, 100, 600, 350)

    def upload_academic_rules(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "Upload Academic Rules (Excel)", "",
                                                   "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            self.academic_rules_file = file_name
            print(f"Uploaded Academic Rules: {self.academic_rules_file}")
            self.academic_rules_button.setStyleSheet('background-color: #4BB543; color: white')
            self.academic_rules_button.setText("File Uploaded!")

    def upload_university_logo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "Upload University Logo (PNG or JPEG)", "",
                                                   "Image Files (*.png *.jpeg *.jpg);;All Files (*)", options=options)
        if file_name:
            self.university_logo_file = file_name
            self.university_logo_button.setStyleSheet('background-color: #4BB543; color: white')
            self.university_logo_button.setText("Logo Uploaded!")

    def toggle_time_reminder(self, state):
        self.time_reminder_enabled = state == Qt.Checked

    def toggle_leave_after_half_time(self, state):
        self.leave_allowed_after_half_time = state == Qt.Checked

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
        self.course_name = self.course_name_edit.text()
        self.course_instructor = self.course_instructor_edit.text()

        self.timer_window = TimerWindow(self.hours, self.minutes, self.academic_rules_file, self.university_logo_file,
                                        self.course_name, self.course_instructor, self.time_reminder_enabled,
                                        self.leave_allowed_after_half_time)
        self.timer_window.showFullScreen()
        self.close()


class TimerWindow(QMainWindow):
    def __init__(self, hours, minutes, academic_rules_file, university_logo_file,course_name,course_instructor, time_reminder_enabled,half_time_leave=False):

        super().__init__()
        try:
            self.half_time_passed = False
            self.three_forth_passed = False

            self.course_label = QLabel(self)
            self.academic_rules_file = academic_rules_file

            self.course_label.setText(f"Course: {course_name}")
            self.course_label.setAlignment(Qt.AlignCenter)
            self.course_label.setFont(QFont("SansSerif", 14))
            self.course_label.setStyleSheet("color: white")

            self.instructor_label = QLabel(self)
            self.instructor_label.setText(f"Instructor: {course_instructor}")
            self.instructor_label.setAlignment(Qt.AlignCenter)
            self.instructor_label.setFont(QFont("SansSerif", 14))
            self.instructor_label.setStyleSheet("color: white")

            self.exam_time = QLabel(self)
            self.exam_time.setText(f"Exam Duration: {hours:02d}:{minutes:02d}:00")
            self.exam_time.setAlignment(Qt.AlignCenter)
            self.exam_time.setFont(QFont("SansSerif", 14))
            self.exam_time.setStyleSheet("color: white")




            self.time_reminder_enabled = time_reminder_enabled
            self.half_time_leave = half_time_leave
            self.hours = hours
            self.minutes = minutes
            self.seconds = hours * 3600 + minutes * 60
            self.start_seconds = self.seconds
            self.timer_label = QLabel(self)
            self.timer_label.setAlignment(Qt.AlignCenter)
            self.timer_label.setFont(QFont("SansSerif", 75))  # Set the font family and size
            self.timer_label.setStyleSheet("color: white;")  # Set font color to white
            self.update_timer_label()


            self.instructions_button = QPushButton("Read Instructions", self)
            self.instructions_button.clicked.connect(self.start_reading_instructions)
            self.instructions_button.setFont(QFont('Arial', 24))
            self.instructions_button.setMaximumWidth(350)

            exit_button = QPushButton("Exit", self)
            exit_button.clicked.connect(self.close_app)
            exit_button.setFont(QFont('Arial', 24))
            exit_button.setMaximumWidth(200)
            self.start_timer_button = QPushButton("Start", self)
            self.start_timer_button.clicked.connect(self.start_stop_timer)
            self.start_timer_button.setFont(QFont('Arial', 24))
            self.start_timer_button.setMaximumWidth(250)



            # academic_rules_file, university_logo_file
            #set the background color of the entire window to black
            palette = self.palette()
            palette.setColor(self.backgroundRole(), QColor(0, 0, 0))
            self.setPalette(palette)

            #make hbox for upper layout
            image_label = QLabel(self)
            image = QPixmap(university_logo_file)
            image_label.setFixedSize(400, 300)
            image = image.scaled(image_label.size(), aspectRatioMode=1, transformMode=0)
            image_label.setPixmap(image)

            upper_left_layout = QVBoxLayout()
            upper_left_layout.addSpacing(5)
            upper_left_layout.addWidget(self.course_label,alignment=Qt.AlignLeft)
            upper_left_layout.addWidget(self.instructor_label,alignment=Qt.AlignLeft)
            upper_left_layout.addWidget(self.exam_time,alignment=Qt.AlignLeft)
            upper_left_layout.addWidget(self.instructions_button,alignment=Qt.AlignLeft)

            upper_layout = QHBoxLayout()
            upper_layout.addLayout(upper_left_layout)
            upper_layout.addWidget(image_label, alignment=Qt.AlignRight)


            layout = QVBoxLayout()
            central_layout = QVBoxLayout()
            central_layout.addWidget(self.timer_label,0, alignment=Qt.AlignCenter)
            central_layout.setContentsMargins(0, 10, 0, 0)
            central_layout.setSpacing(10)
            central_layout.addWidget(self.start_timer_button,0, alignment=Qt.AlignCenter)

            layout.addLayout(upper_layout)
            layout.addLayout(central_layout)
            layout.addWidget(exit_button, alignment=Qt.AlignCenter)
            widget = QWidget()
            widget.setLayout(layout)
            self.setCentralWidget(widget)

            #before starting the timer

            self.timer = QTimer(self)
            self.timer.timeout.connect(self.update_timer)
            self.started = False
        except Exception as e:
            exception_info = traceback.format_exc()
            print(exception_info)

    def start_reading_instructions(self):
        self.instructions_button.setEnabled(False)
        # self.instructions_button.destroy()
        thread = threading.Thread(target=self.voice_function)
        thread.start()
        thread.join(2)

    def voice_function(self):
        welcome_message = "Greetings all, wish you all luck at your exam, before starting the exam I need to remind you with our regulations!"
        text_to_speech(welcome_message)
        rules = excel_to_list(self.academic_rules_file)
        if rules:
            for rule in rules:
                text_to_speech(rule)



    def start_stop_timer(self):
        if self.started:
            self.timer.stop()
            self.start_timer_button.setText("Start")
            self.started = False
        else:
            self.instructions_button.setEnabled(False)
            self.timer.start(1000)
            self.start_timer_button.setText("Pause")
            self.started = True

    def update_timer(self):
        self.seconds -= 1
        if self.seconds <= 0:
            self.timer.stop()
            self.timer_label.setText(f"{0:02d}:{0:02d}:{0:02d}")
            self.start_timer_button.setEnabled(False)
            self.start_timer_button.setText("Time is Up!")
            palette = self.palette()
            palette.setColor(self.backgroundRole(), QColor(0, 255, 0))
            self.setPalette(palette)
            message = "Time is up, everyone. Please stop writing and submit your papers. Make sure all your answers are complete. Thank you for your cooperation. You are now free to leave. Have a great day!"
            t = threading.Thread(target=text_to_speech, args=(message,))
            t.start()
            t.join(1)
        else:
            self.update_timer_label()

    def update_timer_label(self):
        hours = self.seconds // 3600
        minutes = (self.seconds % 3600) // 60
        seconds = self.seconds % 60
        self.timer_label.setText(f"{hours:02d}:{minutes:02d}:{seconds:02d}")
        if self.half_time_passed == False and self.seconds / self.start_seconds <= 0.5:
            self.half_time_passed = True
            self.timer_label.setStyleSheet("color : #00FF00")
            if self.half_time_leave:
                if self.time_reminder_enabled:
                    message= f"you have reached the halfway point of the exam, you have less than {hours} hours and {minutes} minutes and {seconds} seconds left . If you wish to do so, you are now allowed to submit your answers and leave."
                    t = threading.Thread(target=text_to_speech, args=(message,))
                    t.start()
                    t.join(1)
            else:
                if self.time_reminder_enabled:
                    message = f"You have reached the halfway point of the exam, you have less than {hours} hours and {minutes} minutes and {seconds} seconds left ."
                    t = threading.Thread(target=text_to_speech, args=(message,))
                    t.start()
                    t.join(1)
        if self.three_forth_passed == False and self.seconds/self.start_seconds <0.25:
            self.three_forth_passed = True
            self.timer_label.setStyleSheet("color : #FF0000")
            if self.time_reminder_enabled:
                message = f"We are three-fourths of the way through the exam, you have less than {hours} hours and {minutes} minutes and {seconds} seconds left."
                t = threading.Thread(target=text_to_speech, args=(message,))
                t.start()
                t.join(1)


    def close_app(self):
        sys.exit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TimerApp()
    window.show()
    sys.exit(app.exec_())
