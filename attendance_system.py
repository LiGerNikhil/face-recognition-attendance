from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.image import Image
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.dropdown import DropDown
from kivy.clock import Clock
from kivy.graphics.texture import Texture
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.core.window import Window
from kivy.utils import get_color_from_hex
from kivy.properties import StringProperty, BooleanProperty, ObjectProperty
from datetime import datetime, timedelta
import cv2
import cvzone
from cvzone.FaceDetectionModule import FaceDetector
import os
import csv
import numpy as np
import pyttsx3
import logging
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import schedule
import threading
import face_recognition
from PIL import Image as PILImage
import io

# Disable comtypes debug logging
logging.getLogger('comtypes').setLevel(logging.WARNING)

Builder.load_string('''
#:import get_color_from_hex kivy.utils.get_color_from_hex
#:import dp kivy.metrics.dp

<NotificationPopup>:
    message: ''
    bg_color: '#1E1E2E'
    size_hint: (0.8, 0.2)
    title: ''
    separator_height: 0
    BoxLayout:
        orientation: 'vertical'
        padding: dp(15)
        spacing: dp(10)
        canvas.before:
            Color:
                rgba: get_color_from_hex('#2E2E3E')
            RoundedRectangle:
                size: self.size
                pos: self.pos
                radius: [dp(20),]
        Label:
            text: root.message
            font_size: dp(18)
            halign: 'center'
            valign: 'middle'
            color: get_color_from_hex('#F8F8FF')
            bold: True
        Button:
            size_hint_y: None
            height: dp(40)
            text: 'OK'
            background_normal: ''
            background_color: get_color_from_hex('#00FFC6')
            color: get_color_from_hex('#1E1E2E')
            bold: True
            on_press: root.dismiss()

<CustomButton@Button>:
    button_color: '#00FFC6'
    background_normal: ''
    background_color: (0, 0, 0, 0)
    color: get_color_from_hex('#FFFFFF')
    font_size: dp(16)
    bold: True
    canvas.before:
        Color:
            rgba: get_color_from_hex(self.button_color) if not self.disabled else get_color_from_hex('#555555')
        RoundedRectangle:
            size: self.size
            pos: self.pos
            radius: [dp(12),]

<CustomLabel@Label>:
    color: get_color_from_hex('#C3FF00')
    halign: 'left'
    valign: 'middle'
    font_size: dp(17)
    size_hint_y: None
    height: dp(30)

<CustomTextInput@TextInput>:
    background_normal: ''
    background_active: ''
    background_color: get_color_from_hex('#2E2E3E')
    foreground_color: get_color_from_hex('#00FFC6')
    padding: [dp(10), dp(10), dp(10), dp(10)]
    multiline: False
    size_hint_y: None
    height: dp(50)
    font_size: dp(16)
    canvas.before:
        Color:
            rgba: get_color_from_hex('#00FFC6')
        Line:
            width: 1.5
            rounded_rectangle: (self.x, self.y, self.width, self.height, dp(8))

<EmailSettingsContent>:
    cols: 1
    padding: dp(20)
    spacing: dp(15)
    size_hint_y: None
    height: dp(400)
    canvas.before:
        Color:
            rgba: get_color_from_hex('#1E1E2E')
        Rectangle:
            size: self.size
            pos: self.pos

    CustomLabel:
        text: 'Email Settings'
        font_size: dp(22)
        bold: True
        halign: 'center'
        size_hint_y: None
        height: dp(30)
        color: get_color_from_hex('#00FFC6')

    CustomTextInput:
        id: sender_email
        hint_text: "Sender Email"
        text: root.sender_email

    CustomTextInput:
        id: sender_password
        hint_text: "Sender Password"
        password: True
        text: root.sender_password

    CustomTextInput:
        id: receiver_email
        hint_text: "Receiver Email"
        text: root.receiver_email

    Spinner:
        id: frequency
        text: root.frequency
        values: ['daily', 'weekly', 'monthly']
        size_hint_y: None
        height: dp(50)
        background_color: get_color_from_hex('#2E2E3E')
        color: get_color_from_hex('#00FFC6')

    BoxLayout:
        size_hint_y: None
        height: dp(50)
        spacing: dp(10)
        CustomButton:
            text: 'Cancel'
            button_color: '#FF4F4F'
            on_press: root.cancel()
        CustomButton:
            text: 'Save'
            button_color: '#00C896'
            on_press: root.save()

<AdminDropdownButton>:
    size_hint_y: None
    height: dp(40)
    text: 'Admin Menu'
    button_color: '#8A2BE2'
    on_release: root.show_dropdown()

<AttendanceApp>:
    orientation: 'vertical'
    padding: dp(10)
    spacing: dp(10)
    canvas.before:
        Color:
            rgba: get_color_from_hex('#101018')
        Rectangle:
            size: self.size
            pos: self.pos

    BoxLayout:
        size_hint: (1, 0.12)
        padding: [dp(10), 0]
        spacing: dp(10)
        canvas.before:
            Color:
                rgba: get_color_from_hex('#29293D')
            Rectangle:
                size: self.size
                pos: self.pos
        CustomLabel:
            id: date_label
            text: 'Date: '
            color: get_color_from_hex('#00FFC6')
        CustomLabel:
            id: time_label
            text: 'Time: '
            color: get_color_from_hex('#00FFC6')
            halign: 'right'

    BoxLayout:
        size_hint: (1, 0.5)
        padding: dp(5)
        canvas.before:
            Color:
                rgba: get_color_from_hex('#1E1E2E')
            RoundedRectangle:
                size: self.size
                pos: self.pos
                radius: [dp(10),]
            Color:
                rgba: get_color_from_hex('#00FFC6')
            Line:
                width: 1.2
                rounded_rectangle: (self.x, self.y, self.width, self.height, dp(10))
        Image:
            id: camera_feed

    BoxLayout:
        size_hint: (1, 0.25)
        orientation: 'vertical'
        padding: dp(10)
        spacing: dp(5)
        canvas.before:
            Color:
                rgba: get_color_from_hex('#1E1E2E')
            RoundedRectangle:
                size: self.size
                pos: self.pos
                radius: [dp(10),]
            Color:
                rgba: get_color_from_hex('#00FFC6')
            Line:
                width: 1
                rounded_rectangle: (self.x, self.y, self.width, self.height, dp(10))
        CustomLabel:
            id: emp_id_label
            text: 'Employee ID: '
        CustomLabel:
            id: emp_name_label
            text: 'Name: '
        CustomLabel:
            id: status_label
            text: 'Status: '

    AdminDropdownButton:
        id: admin_dropdown_btn
        size_hint: (1, 0.08)
        app: root

<RegisterPopupContent>:
    cols: 1
    padding: dp(20)
    spacing: dp(15)
    size_hint_y: None
    height: dp(350)
    canvas.before:
        Color:
            rgba: get_color_from_hex('#1E1E2E')
        Rectangle:
            size: self.size
            pos: self.pos

    CustomLabel:
        text: 'Register Employee'
        font_size: dp(20)
        bold: True
        halign: 'center'
        size_hint_y: None
        height: dp(30)
        color: get_color_from_hex('#00FFC6')

    CustomTextInput:
        id: admin_password
        hint_text: "Admin Password"
        password: True

    CustomTextInput:
        id: popup_emp_id
        hint_text: "Employee ID"

    CustomTextInput:
        id: popup_emp_name
        hint_text: "Employee Name"

    BoxLayout:
        size_hint_y: None
        height: dp(50)
        spacing: dp(10)
        CustomButton:
            text: 'Cancel'
            button_color: '#FF4F4F'
            on_press: root.popup_instance.dismiss()
        CustomButton:
            text: 'Register'
            button_color: '#00FFC6'
            on_press: root.register_employee()

<DeletePopupContent>:
    cols: 1
    padding: dp(20)
    spacing: dp(15)
    size_hint_y: None
    height: dp(300)
    canvas.before:
        Color:
            rgba: get_color_from_hex('#1E1E2E')
        Rectangle:
            size: self.size
            pos: self.pos

    CustomLabel:
        text: 'Delete Employee'
        font_size: dp(20)
        bold: True
        halign: 'center'
        size_hint_y: None
        height: dp(30)
        color: get_color_from_hex('#FF4F4F')

    CustomTextInput:
        id: admin_password
        hint_text: "Admin Password"
        password: True

    CustomTextInput:
        id: emp_id_to_delete
        hint_text: "Employee ID"

    BoxLayout:
        size_hint_y: None
        height: dp(50)
        spacing: dp(10)
        CustomButton:
            text: 'Cancel'
            button_color: '#888888'
            on_press: root.popup_instance.dismiss()
        CustomButton:
            text: 'Delete'
            button_color: '#FF4F4F'
            on_press: root.delete_employee()

''')

class NotificationPopup(Popup):
    message = StringProperty('')
    bg_color = StringProperty('#4A6FA5')  # Default blue color
    
    def __init__(self, message='', bg_color='#4A6FA5', **kwargs):
        super().__init__(**kwargs)
        self.message = message
        self.bg_color = bg_color

class EmailSettingsContent(GridLayout):
    sender_email = StringProperty('')
    sender_password = StringProperty('')
    receiver_email = StringProperty('')
    frequency = StringProperty('weekly')

    def __init__(self, main_app, popup_instance, **kwargs):
        super().__init__(**kwargs)
        self.main_app = main_app
        self.popup_instance = popup_instance
        self.sender_email = self.main_app.email_config['sender_email']
        self.sender_password = self.main_app.email_config['sender_password']
        self.receiver_email = self.main_app.email_config['receiver_email']
        self.frequency = self.main_app.email_config['report_frequency']

    def save(self):
        self.main_app.update_email_settings(
            self.ids.sender_email.text,
            self.ids.sender_password.text,
            self.ids.receiver_email.text,
            self.ids.frequency.text
        )
        self.popup_instance.dismiss()

    def cancel(self):
        self.popup_instance.dismiss()

class CustomButton(Button):
    button_color = StringProperty('#4A6FA5')

class AdminDropdownButton(CustomButton):
    app = ObjectProperty(None)
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.dropdown = DropDown()
        self.create_dropdown_items()
        
    def create_dropdown_items(self):
        # Register button
        btn_register = CustomButton(
            text='Register', 
            size_hint_y=None, 
            height=dp(40),
            button_color='#27AE60'
        )
        btn_register.bind(on_release=lambda btn: self.app.show_register_popup())
        self.dropdown.add_widget(btn_register)
        
        # Delete button
        btn_delete = CustomButton(
            text='Delete', 
            size_hint_y=None, 
            height=dp(40),
            button_color='#E74C3C'
        )
        btn_delete.bind(on_release=lambda btn: self.app.show_delete_popup())
        self.dropdown.add_widget(btn_delete)
        
        # Reset button
        btn_reset = CustomButton(
            text='Reset', 
            size_hint_y=None, 
            height=dp(40),
            button_color='#95A5A6'
        )
        btn_reset.bind(on_release=lambda btn: self.app.reset_interface())
        self.dropdown.add_widget(btn_reset)
        
        # Email Settings button
        btn_email = CustomButton(
            text='Email Settings', 
            size_hint_y=None, 
            height=dp(40),
            button_color='#3498DB'
        )
        btn_email.bind(on_release=lambda btn: self.app.show_email_settings())
        self.dropdown.add_widget(btn_email)
        
        # Send Report button
        btn_report = CustomButton(
            text='Send Report Now', 
            size_hint_y=None, 
            height=dp(40),
            button_color='#9B59B6'
        )
        btn_report.bind(on_release=lambda btn: self.app.send_email_with_report())
        self.dropdown.add_widget(btn_report)
    
    def show_dropdown(self):
        self.dropdown.open(self)

class RegisterPopupContent(GridLayout):
    def __init__(self, main_app, popup_instance, **kwargs):
        super().__init__(**kwargs)
        self.main_app = main_app
        self.popup_instance = popup_instance
        self.admin_password = "kwikdelete"  # Hardcoded admin password

    def register_employee(self):
        password = self.ids.admin_password.text.strip()
        emp_id = self.ids.popup_emp_id.text.strip()
        emp_name = self.ids.popup_emp_name.text.strip()

        if password != self.admin_password:
            self.main_app.show_notification("Incorrect admin password", '#E74C3C')
            self.main_app.speak("Incorrect admin password.")
            return

        if not emp_id or not emp_name:
            self.main_app.show_notification("Please enter both Employee ID and Name", '#E74C3C')
            self.main_app.speak("Please enter both Employee ID and Name.")
            return

        # Check if employee ID already exists
        if emp_id in self.main_app.known_employees:
            self.main_app.show_notification(f"Employee ID {emp_id} already exists", '#E74C3C')
            self.main_app.speak(f"Employee ID {emp_id} already exists.")
            return

        # Check if employee name already exists with a different ID
        for existing_id, existing_name in self.main_app.known_employees.items():
            if existing_name.lower() == emp_name.lower() and existing_id != emp_id:
                self.main_app.show_notification(f"Name {emp_name} already registered with ID {existing_id}", '#E74C3C')
                self.main_app.speak(f"Name {emp_name} already registered with ID {existing_id}")
                return

        if self.main_app.registration_face is None:
            self.main_app.show_notification("No face captured for registration", '#E74C3C')
            self.main_app.speak("No face captured for registration.")
            return

        # Check if the face is already registered
        try:
            # Convert the face to RGB (required by face_recognition)
            rgb_face = cv2.cvtColor(self.main_app.registration_face, cv2.COLOR_BGR2RGB)
            
            # Get face encodings
            face_encodings = face_recognition.face_encodings(rgb_face)
            
            if not face_encodings:
                self.main_app.show_notification("No face detected in the captured image", '#E74C3C')
                self.main_app.speak("No face detected in the captured image.")
                return
                
            new_face_encoding = face_encodings[0]
            
            # Compare with all registered faces
            for existing_id, existing_encoding in self.main_app.known_face_encodings.items():
                matches = face_recognition.compare_faces([existing_encoding], new_face_encoding)
                if matches[0]:
                    self.main_app.show_notification(f"This face is already registered as {self.main_app.known_employees[existing_id]} (ID: {existing_id})", '#E74C3C')
                    self.main_app.speak(f"This face is already registered as {self.main_app.known_employees[existing_id]} with ID {existing_id}")
                    return
        except Exception as e:
            print(f"Error in face recognition: {e}")
            self.main_app.show_notification("Error in face recognition", '#E74C3C')
            self.main_app.speak("Error in face recognition.")
            return

        # Ensure directory exists
        os.makedirs('employee_data', exist_ok=True)
        
        # Convert the face to grayscale for better recognition
        gray_face = cv2.cvtColor(self.main_app.registration_face, cv2.COLOR_BGR2GRAY)
        
        # Save the face image
        filepath = f"employee_data/{emp_id}.jpg"
        cv2.imwrite(filepath, gray_face)

        # Save to CSV
        csv_file = "employee_data/employee_names.csv"
        file_exists = os.path.exists(csv_file)
        
        with open(csv_file, 'a', newline='') as file:
            writer = csv.writer(file)
            if not file_exists:
                writer.writerow(["Employee ID", "Name"])
            writer.writerow([emp_id, emp_name])

        # Add to known encodings
        self.main_app.known_face_encodings[emp_id] = new_face_encoding
        self.main_app.known_employees[emp_id] = emp_name

        self.main_app.show_notification(f"Employee {emp_name} registered successfully", '#27AE60')
        self.main_app.speak(f"Employee {emp_name} registered successfully.")
        self.main_app.train_recognizer()
        self.main_app.reset_interface()
        self.popup_instance.dismiss()

class DeletePopupContent(GridLayout):
    def __init__(self, main_app, popup_instance, **kwargs):
        super().__init__(**kwargs)
        self.main_app = main_app
        self.popup_instance = popup_instance
        self.admin_password = "kwikdelete"  # Hardcoded admin password

    def delete_employee(self):
        password = self.ids.admin_password.text.strip()
        emp_id = self.ids.emp_id_to_delete.text.strip()

        if password != self.admin_password:
            self.main_app.show_notification("Incorrect admin password", '#E74C3C')
            self.main_app.speak("Incorrect admin password.")
            return

        if not emp_id:
            self.main_app.show_notification("Please enter an Employee ID", '#E74C3C')
            self.main_app.speak("Please enter an Employee ID.")
            return

        # Check if employee exists
        if emp_id not in self.main_app.known_employees:
            self.main_app.show_notification(f"Employee ID {emp_id} not found", '#E74C3C')
            self.main_app.speak(f"Employee ID {emp_id} not found.")
            return

        # Delete the employee image file
        img_path = f"employee_data/{emp_id}.jpg"
        if os.path.exists(img_path):
            try:
                os.remove(img_path)
            except Exception as e:
                self.main_app.show_notification(f"Error deleting employee image", '#E74C3C')
                self.main_app.speak(f"Error deleting employee image: {str(e)}")
                return

        # Remove from CSV
        csv_file = "employee_data/employee_names.csv"
        temp_file = "employee_data/temp.csv"
        
        try:
            with open(csv_file, 'r') as infile, open(temp_file, 'w', newline='') as outfile:
                reader = csv.reader(infile)
                writer = csv.writer(outfile)
                header = next(reader)
                writer.writerow(header)
                for row in reader:
                    if len(row) >= 2 and row[0] != emp_id:
                        writer.writerow(row)
            
            # Replace old file with new file
            os.remove(csv_file)
            os.rename(temp_file, csv_file)
            
            # Remove from known encodings
            if emp_id in self.main_app.known_face_encodings:
                del self.main_app.known_face_encodings[emp_id]
            
            self.main_app.show_notification(f"Employee {emp_id} deleted successfully", '#27AE60')
            self.main_app.speak(f"Employee {emp_id} deleted successfully.")
            self.main_app.known_employees = self.main_app.load_employees()
            self.main_app.train_recognizer()
            self.main_app.reset_interface()
            self.popup_instance.dismiss()
        except Exception as e:
            self.main_app.show_notification(f"Error deleting employee", '#E74C3C')
            self.main_app.speak(f"Error deleting employee: {str(e)}")
            if os.path.exists(temp_file):
                os.remove(temp_file)

class AttendanceApp(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.engine = pyttsx3.init()
        self.detector = FaceDetector(minDetectionCon=0.7)
        self.face_recognizer = cv2.face.LBPHFaceRecognizer_create()
        self.known_employees = self.load_employees()
        self.known_face_encodings = self.load_face_encodings()
        self.punch_status = {}  # Track punch status and timestamps
        self.current_face = None
        self.registration_face = None  # Stores the face image for registration
        self.current_emp_id = None
        self.last_punch_time = {}  # Track last punch time for each employee
        self.face_detected_time = None
        self.cooldown_period = 60  # 1 minutes cooldown in seconds
        self.punch_delay = 2  # 2 seconds delay before auto punch
        self.duplicate_check_window = 5  # 5 seconds window to prevent duplicate punches
        self.face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
        self.trained = False
        self.liveness_threshold = 0.3  # Threshold for liveness detection
        self.last_frame_time = 0
        self.frame_count = 0

        # Email configuration
        self.email_config = {
            'sender_email': 'saakshinikhil@gmail.com',
            'sender_password': 'axnl kklq fwdw tlwh',  # Use app password for Gmail
            'receiver_email': '2040124@sliet.ac.in',
            'smtp_server': 'smtp.gmail.com',
            'smtp_port': 587,
            'report_frequency': 'daily'  # Can be 'daily', 'weekly', or 'monthly'
        }

        # Ensure necessary folders exist
        os.makedirs('employee_data', exist_ok=True)
        os.makedirs('attendance_records', exist_ok=True)

        self.capture = cv2.VideoCapture(0)
        if not self.capture.isOpened():
            self.speak("Error: Could not open camera.")
            return

        # Train the face recognizer with existing employees
        self.train_recognizer()

        # Start the email scheduler
        self.start_email_scheduler()

        Clock.schedule_interval(self.update, 1.0 / 30.0)
        Clock.schedule_interval(self.update_time, 1)

    def load_face_encodings(self):
        encodings = {}
        for emp_id in self.known_employees:
            path = f"employee_data/{emp_id}.jpg"
            if os.path.exists(path):
                try:
                    image = face_recognition.load_image_file(path)
                    face_encodings = face_recognition.face_encodings(image)
                    if face_encodings:
                        encodings[emp_id] = face_encodings[0]
                except Exception as e:
                    print(f"Error loading face encoding for {emp_id}: {e}")
        return encodings

    def start_email_scheduler(self):
        """Start the email scheduler in a background thread"""
        def run_scheduler():
            freq = self.email_config['report_frequency']
            
            if freq == 'daily':
                schedule.every().day.at("06:00").do(self.send_email_with_report)  # 6 PM daily
            elif freq == 'weekly':
                schedule.every().monday.at("18:00").do(self.send_email_with_report)  # Every Monday at 6 PM
            else:  # monthly
                schedule.every().month.at("18:00").do(self.send_email_with_report)  # 1st of month at 6 PM
            
            while True:
                schedule.run_pending()
                time.sleep(60)  # Check every minute

        # Start the scheduler in a daemon thread
        scheduler_thread = threading.Thread(target=run_scheduler)
        scheduler_thread.daemon = True
        scheduler_thread.start()

    def generate_attendance_report(self):
        """Generate CSV report of attendance data"""
        report_filename = "attendance_records/report.csv"
        try:
            with open("attendance_records/attendance_log.csv", 'r') as infile, \
                 open(report_filename, 'w', newline='') as outfile:
                reader = csv.reader(infile)
                writer = csv.writer(outfile)
                writer.writerow(["Employee ID", "Name", "Punch Type", "Date", "Time"])
                next(reader)  # Skip header
                for row in reader:
                    if len(row) >= 5:  # Ensure we have enough columns
                        writer.writerow([row[0], row[1], row[2], row[3], row[4]])
            return report_filename
        except Exception as e:
            print(f"Error generating report: {e}")
            return None

    def send_email_with_report(self):
        """Send email with attendance report attachment"""
        report_file = self.generate_attendance_report()
        if not report_file:
            return False

        msg = MIMEMultipart()
        msg['From'] = self.email_config['sender_email']
        msg['To'] = self.email_config['receiver_email']
        
        # Set email subject based on frequency
        freq = self.email_config['report_frequency']
        if freq == 'daily':
            msg['Subject'] = f"Daily Attendance Report - {datetime.now().strftime('%Y-%m-%d')}"
        elif freq == 'weekly':
            msg['Subject'] = f"Weekly Attendance Report - {datetime.now().strftime('%Y-%m-%d')}"
        else:  # monthly
            msg['Subject'] = f"Monthly Attendance Report - {datetime.now().strftime('%Y-%m')}"

        # Email body
        body = f"Please find attached the {freq} attendance report."
        msg.attach(MIMEText(body, 'plain'))

        # Attach the report file
        attachment = open(report_file, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {report_file.split('/')[-1]}")
        msg.attach(part)

        # Send email
        try:
            server = smtplib.SMTP(self.email_config['smtp_server'], self.email_config['smtp_port'])
            server.starttls()
            server.login(self.email_config['sender_email'], self.email_config['sender_password'])
            server.send_message(msg)
            server.quit()
            self.show_notification(f"{freq.capitalize()} report sent successfully", '#27AE60')
            return True
        except Exception as e:
            print(f"Error sending email: {e}")
            self.show_notification(f"Failed to send email: {e}", '#E74C3C')
            return False

    def update_email_settings(self, sender_email, sender_password, receiver_email, frequency):
        """Update email configuration settings"""
        self.email_config = {
            'sender_email': sender_email,
            'sender_password': sender_password,
            'receiver_email': receiver_email,
            'smtp_server': 'smtp.gmail.com',  # Default to Gmail
            'smtp_port': 587,
            'report_frequency': frequency
        }
        self.show_notification("Email settings updated", '#27AE60')

    def show_email_settings(self):
        popup = Popup(title="Email Settings", size_hint=(0.85, 0.7))
        content = EmailSettingsContent(main_app=self, popup_instance=popup)
        popup.content = content
        popup.open()

    def train_recognizer(self):
        faces = []
        labels = []
        label_ids = {}
        current_id = 0

        for emp_id in self.known_employees:
            path = f"employee_data/{emp_id}.jpg"
            if os.path.exists(path):
                img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
                if img is not None:
                    img = cv2.resize(img, (200, 200))
                    faces.append(img)
                    labels.append(current_id)
                    label_ids[current_id] = emp_id
                    current_id += 1

        if faces:
            self.face_recognizer.train(faces, np.array(labels))
            self.label_ids = label_ids
            self.trained = True
        else:
            self.trained = False

    def load_employees(self):
        employees = {}
        csv_file = "employee_data/employee_names.csv"
        if os.path.exists(csv_file):
            with open(csv_file, 'r') as file:
                reader = csv.reader(file)
                next(reader, None)  # Skip header
                for row in reader:
                    if len(row) >= 2:
                        emp_id, name = row
                        employees[emp_id] = name
        return employees

    def show_notification(self, message, bg_color='#4A6FA5'):
        popup = NotificationPopup(message=message, bg_color=bg_color)
        popup.open()

    def speak(self, text):
        try:
            self.engine.say(text)
            self.engine.runAndWait()
        except Exception as e:
            print(f"Error in speech synthesis: {e}")

    def update_time(self, dt):
        now = datetime.now()
        self.ids.date_label.text = f"Date: {now.strftime('%Y-%m-%d')}"
        self.ids.time_label.text = f"Time: {now.strftime('%H:%M:%S')}"

    def check_liveness(self, frame):
        """Simple liveness check by looking for eye blinking or facial movements"""
        # Convert to grayscale
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        
        # Detect faces
        faces = self.face_cascade.detectMultiScale(gray, 1.3, 5)
        
        if len(faces) == 0:
            return False
        
        # For now, just return True if a face is detected
        # In a real application, you would implement more sophisticated checks
        return True

    def update(self, dt):
        try:
            ret, frame = self.capture.read()
            if not ret:
                return

            current_time = time.time()
            self.frame_count += 1

            # Flip frame horizontally for mirror effect
            frame = cv2.flip(frame, 1)
            
            # Convert to grayscale for face detection
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            
            # Detect faces using Haar Cascade (more reliable)
            faces = self.face_cascade.detectMultiScale(
                gray,
                scaleFactor=1.1,
                minNeighbors=5,
                minSize=(100, 100),
                flags=cv2.CASCADE_SCALE_IMAGE
            )
            
            if len(faces) > 0:
                (x, y, w, h) = faces[0]
                
                # Draw rectangle around the face
                cv2.rectangle(frame, (x, y), (x+w, y+h), (255, 0, 0), 2)
                
                # Get the face region (store both color and grayscale versions)
                self.current_face = gray[y:y+h, x:x+w]  # Grayscale for recognition
                self.registration_face = frame[y:y+h, x:x+w]  # Color for registration
                
                # Check liveness (prevent photo spoofing)
                is_live = self.check_liveness(frame)
                
                if not is_live:
                    self.ids.emp_id_label.text = "Employee ID: "
                    self.ids.emp_name_label.text = "Name: "
                    self.ids.status_label.text = "Status: Possible spoofing attempt"
                    self.face_detected_time = None
                    buf = cv2.flip(frame, 0).tobytes()
                    texture = Texture.create(size=(frame.shape[1], frame.shape[0]), colorfmt='bgr')
                    texture.blit_buffer(buf, colorfmt='bgr', bufferfmt='ubyte')
                    self.ids.camera_feed.texture = texture
                    return
                
                # Recognize the face if we have trained data
                if self.trained:
                    face_img = cv2.resize(self.current_face, (200, 200))
                    label, confidence = self.face_recognizer.predict(face_img)
                    
                    # Confidence threshold (lower is better)
                    if confidence < 80 and label in self.label_ids:
                        self.current_emp_id = self.label_ids[label]
                    else:
                        self.current_emp_id = None
                else:
                    self.current_emp_id = None

                if self.current_emp_id:
                    # Additional verification with face encodings
                    rgb_face = cv2.cvtColor(self.registration_face, cv2.COLOR_BGR2RGB)
                    current_encodings = face_recognition.face_encodings(rgb_face)
                    
                    if current_encodings:
                        current_encoding = current_encodings[0]
                        matches = face_recognition.compare_faces(
                            [self.known_face_encodings[self.current_emp_id]], 
                            current_encoding
                        )
                        
                        if not matches[0]:
                            # Face doesn't match the stored encoding
                            self.current_emp_id = None
                            self.ids.emp_id_label.text = "Employee ID: Verification failed"
                            self.ids.emp_name_label.text = "Name: "
                            self.ids.status_label.text = "Status: Face mismatch"
                            self.face_detected_time = None
                            buf = cv2.flip(frame, 0).tobytes()
                            texture = Texture.create(size=(frame.shape[1], frame.shape[0]), colorfmt='bgr')
                            texture.blit_buffer(buf, colorfmt='bgr', bufferfmt='ubyte')
                            self.ids.camera_feed.texture = texture
                            return

                    self.ids.emp_id_label.text = f"Employee ID: {self.current_emp_id}"
                    self.ids.emp_name_label.text = f"Name: {self.known_employees.get(self.current_emp_id, 'Unknown')}"
                    
                    # Check if face is newly detected
                    if self.face_detected_time is None:
                        self.face_detected_time = current_time
                    
                    # Check if face has been detected long enough for auto punch
                    if current_time - self.face_detected_time >= self.punch_delay:
                        self.handle_auto_punch()
                else:
                    self.ids.emp_id_label.text = "Employee ID: New"
                    self.ids.emp_name_label.text = "Name: Unknown"
                    self.face_detected_time = None
            else:
                self.face_detected_time = None
                self.current_face = None
                self.registration_face = None
                self.ids.emp_id_label.text = "Employee ID: "
                self.ids.emp_name_label.text = "Name: "
                self.ids.status_label.text = "Status: "

            buf = cv2.flip(frame, 0).tobytes()
            texture = Texture.create(size=(frame.shape[1], frame.shape[0]), colorfmt='bgr')
            texture.blit_buffer(buf, colorfmt='bgr', bufferfmt='ubyte')
            self.ids.camera_feed.texture = texture
        except Exception as e:
            print(f"Error in update: {e}")

    def handle_auto_punch(self):
        if not self.current_emp_id:
            return
            
        current_time = time.time()
        last_punch = self.last_punch_time.get(self.current_emp_id, 0)
        
        # Check if employee is currently punched in
        if self.punch_status.get(self.current_emp_id) == "IN":
            # Check if cooldown period has passed (2 minutes)
            if current_time - last_punch >= self.cooldown_period:
                self.mark_attendance(self.current_emp_id, "OUT")
                self.punch_status[self.current_emp_id] = "OUT"
                self.last_punch_time[self.current_emp_id] = current_time
                self.ids.status_label.text = "Status: OUT"
                self.show_notification(f"{self.known_employees.get(self.current_emp_id, 'Employee')} punched OUT", '#3498DB')
                self.reset_interface()
        else:
            # Check if cooldown period has passed (2 minutes)
            if current_time - last_punch >= self.cooldown_period:
                self.mark_attendance(self.current_emp_id, "IN")
                self.punch_status[self.current_emp_id] = "IN"
                self.last_punch_time[self.current_emp_id] = current_time
                self.ids.status_label.text = "Status: IN"
                self.show_notification(f"{self.known_employees.get(self.current_emp_id, 'Employee')} punched IN", '#3498DB')
                self.reset_interface()

    def mark_attendance(self, emp_id, punch_type):
        try:
            now = datetime.now()
            date_str = now.strftime("%Y-%m-%d")
            time_str = now.strftime("%H:%M:%S")
            filename = "attendance_records/attendance_log.csv"

            # Check for duplicate entry within the time window
            if os.path.exists(filename):
                with open(filename, 'r') as file:
                    reader = csv.reader(file)
                    last_row = None
                    for row in reader:
                        if row and row[0] == emp_id and row[2] == punch_type:
                            last_row = row
                    
                    if last_row:
                        last_time = datetime.strptime(last_row[4], "%H:%M:%S")
                        current_time = datetime.strptime(time_str, "%H:%M:%S")
                        time_diff = (current_time - last_time).total_seconds()
                        if time_diff < self.duplicate_check_window:
                            self.speak(f"Duplicate {punch_type} detected for {self.known_employees.get(emp_id, 'Employee')}")
                            return

            file_exists = os.path.exists(filename)
            
            with open(filename, 'a', newline='') as file:
                writer = csv.writer(file)
                if not file_exists:
                    writer.writerow(["Employee ID", "Name", "Punch Type", "Date", "Time", "Timestamp"])
                
                writer.writerow([
                    emp_id,
                    self.known_employees.get(emp_id, "Unknown"),
                    punch_type,
                    date_str,
                    time_str,
                    now.timestamp()
                ])

            self.speak(f"{self.known_employees.get(emp_id, 'Employee')} {punch_type.lower()} at {time_str}")
        except Exception as e:
            print(f"Error marking attendance: {e}")
            self.speak("Error recording attendance")

    def show_register_popup(self):
        if self.registration_face is None:
            self.show_notification("No face detected for registration", '#E74C3C')
            self.speak("No face detected for registration")
            return
            
        popup = Popup(title="", size_hint=(0.85, 0.6), separator_height=0)
        content = RegisterPopupContent(main_app=self, popup_instance=popup)
        popup.content = content
        popup.open()

    def show_delete_popup(self):
        popup = Popup(title="", size_hint=(0.85, 0.5), separator_height=0)
        content = DeletePopupContent(main_app=self, popup_instance=popup)
        popup.content = content
        popup.open()

    def reset_interface(self):
        self.current_emp_id = None
        self.current_face = None
        self.registration_face = None
        self.face_detected_time = None
        self.ids.emp_id_label.text = "Employee ID: "
        self.ids.emp_name_label.text = "Name: "
        self.ids.status_label.text = "Status: "

    def on_stop(self):
        if hasattr(self, 'capture') and self.capture.isOpened():
            self.capture.release()

class MobileAttendanceApp(App):
    def build(self):
        # Set window size for mobile (adjust as needed)
        Window.size = (360, 640)  # Typical mobile portrait size
        self.title = 'Face Attendance System'
        return AttendanceApp()

    def on_stop(self):
        self.root.on_stop()

if __name__ == '__main__':
    MobileAttendanceApp().run()