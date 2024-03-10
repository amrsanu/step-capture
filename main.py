"""
App to save screenshots and format a document.
"""
import sys
import os
import datetime

import pyautogui
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QWidget
from PyQt5.QtCore import QObject, Qt
from pynput import mouse
from docx import Document
from docx.shared import Inches

from app import Ui_Form
from captureWindow import select_window

FOLDER_NAME = f'step-Capture-{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}'
FOLDER_NAME = os.path.join(os.path.dirname(
    os.path.abspath(__file__)), FOLDER_NAME)


def ensure_directory(directory):
    """
    Ensure the directory exists, creating it if necessary.

    Args:
        directory (str): The directory path.
    """
    if not os.path.exists(directory):
        os.makedirs(directory)


class MainWindow(QWidget, QObject):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.setup_ui_events()
        self.setMouseTracking(True)

        self.capture_area = None
        self.take_screenshot = False
        ensure_directory(FOLDER_NAME)
        # Set the window flag to keep the window on top
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)

    def setup_ui_events(self):
        self.ui.selectAreaPushButton.clicked.connect(self.select_area)
        self.ui.startCapturePushButton.clicked.connect(self.start_capture)
        self.ui.createDocumentPushButton.clicked.connect(self.create_document)

    def select_area(self):
        print("Select Area button clicked")
        self.capture_area = select_window()
        print(f"Area: {self.capture_area}")

    def start_capture(self):
        print("Start Capture button clicked")
        try:
            self.mouseListener = mouse.Listener(on_click=self.on_click)

            self.alt_pressed = False
            self.mouseListener.start()
            print("Take capture from window")
            self.take_screenshot = True
        except Exception as ex:
            print(f"Start Capture Error: {ex}")

    def on_click(self, x, y, button, pressed):
        if self.take_screenshot and pressed and self.is_within_capture_area(x, y):
            print(
                f"Save capture state - {self.capture_area} - Ctrl+Clicks pressed")
            self.screenshot()

    def screenshot(self):
        # Get the screen geometry
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = os.path.join(FOLDER_NAME,  f"screenshot_{timestamp}.png")
        (x1, y1), (x2, y2) = self.capture_area
        width = x2 - x1
        height = y2 - y1
        screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
        screenshot.save(filename)

        print("Screenshot saved as screenshot.png")

    def is_within_capture_area(self, x, y):
        (x1, y1), (x2, y2) = self.capture_area
        return x1 <= x <= x2 and y1 <= y <= y2

    def create_document(self):
        print("Create Document button clicked")
        self.take_screenshot = False
        self.mouseListener.stop()

        doc = Document()
        doc.add_heading('Document with Images', level=1)
        doc.add_paragraph("This document contains images.")
        image_files = []  # List of image file paths
        for image in os.listdir(FOLDER_NAME):
            if image.endswith('png'):
                image_files.append(os.path.join(FOLDER_NAME, image))
        for image_file in image_files:
            # Adjust width as needed
            doc.add_heading(os.path.basename(image_file), level=2)
            doc.add_picture(image_file, width=Inches(5))
        # Save the document
        doc.save(f'{FOLDER_NAME}.docx')


if __name__ == "__main__":
    try:
        app = QtWidgets.QApplication(sys.argv)
        mainWindow = MainWindow()
        mainWindow.show()
        app.installEventFilter(mainWindow)
        sys.exit(app.exec_())
    except Exception as ex:
        print(f"MAIN Error: {ex}")
