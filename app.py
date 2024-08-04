import sys
import re
import zipfile
import os
import traceback
import ctypes
from io import StringIO
from PyQt5 import QtWidgets, QtCore, QtGui

from docx import Document  # Ensure this import is at the top

# Store history and code data
history = []
code_store = {}

def preprocess_code(code):
    # Add the import statement at the beginning
    import_statement = "from docx import Document\n"
    
    # Remove '/mnt/data/' from file paths
    code = re.sub(r'/mnt/data/', '', code)
    
    return import_statement + code

class PythonToDOCXApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.is_dark_mode = True  # Set to dark mode by default
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Python to DOCX")
        self.setGeometry(100, 100, 1000, 700)  # Adjusted window size to be larger

        # Layout
        main_layout = QtWidgets.QVBoxLayout(self)

        # Code Input (Multiline text area for code)
        self.code_input = QtWidgets.QPlainTextEdit(self)
        self.code_input.setPlaceholderText("Enter your Python code here...")
        self.code_input.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        main_layout.addWidget(self.code_input, stretch=2)  # Stretch factor added

        # Output Text
        self.output_text = QtWidgets.QPlainTextEdit(self)
        self.output_text.setReadOnly(True)
        self.output_text.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)  # Adjusted to Minimum
        main_layout.addWidget(self.output_text, stretch=1)  # Stretch factor added

        # Buttons
        button_layout = QtWidgets.QHBoxLayout()

        self.clear_button = QtWidgets.QPushButton("Clear All", self)
        self.clear_button.clicked.connect(self.clear_all)
        button_layout.addWidget(self.clear_button)

        self.paste_button = QtWidgets.QPushButton("Paste", self)
        self.paste_button.clicked.connect(self.paste_text)
        button_layout.addWidget(self.paste_button)

        self.execute_button = QtWidgets.QPushButton("Enter", self)
        self.execute_button.clicked.connect(self.execute_code)
        button_layout.addWidget(self.execute_button)

        self.history_button = QtWidgets.QPushButton("History", self)
        self.history_button.clicked.connect(self.show_history)
        button_layout.addWidget(self.history_button)

        # Dark/Light Mode Switch
        self.mode_button = QtWidgets.QPushButton("Switch to Light Mode", self)
        self.mode_button.clicked.connect(self.toggle_mode)
        button_layout.addWidget(self.mode_button)

        main_layout.addLayout(button_layout)

        # Footer
        footer = QtWidgets.QLabel("Made by Sayan", self)
        footer.setAlignment(QtCore.Qt.AlignCenter)
        footer.setStyleSheet("color: blue; font-style: italic;")
        footer.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        main_layout.addWidget(footer)

        self.setLayout(main_layout)
        self.apply_theme()  # Apply the initial theme

    def apply_theme(self):
        if self.is_dark_mode:
            self.setStyleSheet("""
                QWidget {
                    background-color: #2e2e2e;
                    color: #f0f0f0;
                }
                QPlainTextEdit {
                    background-color: #1e1e1e;
                    color: #dcdcdc;
                }
                QPushButton {
                    background-color: #3c3c3c;
                    color: #f0f0f0;
                }
                QPushButton:hover {
                    background-color: #4a4a4a;
                }
                QLabel {
                    color: #f0f0f0;
                }
            """)
            self.mode_button.setText("Switch to Light Mode")
        else:
            self.setStyleSheet("""
                QWidget {
                    background-color: #ffffff;
                    color: #000000;
                }
                QPlainTextEdit {
                    background-color: #f5f5f5;
                    color: #000000;
                }
                QPushButton {
                    background-color: #e0e0e0;
                    color: #000000;
                }
                QPushButton:hover {
                    background-color: #c0c0c0;
                }
                QLabel {
                    color: #000000;
                }
            """)
            self.mode_button.setText("Switch to Dark Mode")

    def toggle_mode(self):
        self.is_dark_mode = not self.is_dark_mode
        self.apply_theme()

    def execute_code(self):
        code = self.code_input.toPlainText()
        
        if not code.strip():
            self.output_text.appendPlainText("No code entered.\n")
            return
        
        # Preprocess the code
        processed_code = preprocess_code(code)
        
        # Clear previous output
        self.output_text.clear()
        
        # Redirect stdout and stderr to capture print statements and errors
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        new_stdout = StringIO()
        new_stderr = StringIO()
        sys.stdout = new_stdout
        sys.stderr = new_stderr

        # File tracking
        docx_files = []
        zip_files = []

        # Add to history with serial number
        serial_number = len(history) + 1
        timestamp = QtCore.QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        file_type = ""

        try:
            # Execute the processed code
            exec(processed_code)
            
            # Example: If code creates a DOCX file named 'example.docx'
            # Add the names of generated DOCX files to the list
            docx_files.append('questions.docx')  # Adjust based on actual output
            
            # Determine which type of file to provide
            if zip_files:
                # Provide ZIP file
                zip_file_name = zip_files[0]  # Assuming there's only one ZIP file
                file_type = "ZIP"
                self.output_text.appendPlainText(f"ZIP file created: {zip_file_name}\n")
            elif docx_files:
                # Provide DOCX file
                docx_file_name = docx_files[0]  # Assuming there's only one DOCX file
                file_type = "DOCX"
                self.output_text.appendPlainText(f"DOCX file created: {docx_file_name}\n")
            else:
                self.output_text.appendPlainText("No DOCX or ZIP files created.\n")

        except Exception as e:
            # Get the exception traceback
            traceback_str = traceback.format_exc()
            self.output_text.appendPlainText(f"An error occurred:\n{traceback_str}")
        finally:
            # Restore original stdout and stderr
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            
            # Get output from StringIO
            output = new_stdout.getvalue()
            error = new_stderr.getvalue()
            
            if output:
                self.output_text.appendPlainText(output)
            if error:
                self.output_text.appendPlainText(f"Error: {error}")

            # Store code and history
            code_store[serial_number] = code
            history.append((serial_number, timestamp, file_type))

    def paste_text(self):
        # Paste text from clipboard into the code_input text widget
        clipboard = QtWidgets.QApplication.clipboard()
        self.code_input.insertPlainText(clipboard.text())

    def clear_all(self):
        self.code_input.clear()

    def show_history(self):
        self.history_window = HistoryWindow()
        self.history_window.show()

class HistoryWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("History")
        self.setGeometry(150, 150, 800, 600)  # Size of the history window

        # Layout
        layout = QtWidgets.QVBoxLayout(self)

        # History Table
        self.history_table = QtWidgets.QTableWidget()
        self.history_table.setColumnCount(5)
        self.history_table.setHorizontalHeaderLabels(["Serial No", "Timestamp", "File Type", "Open", "Copy Text"])
        self.history_table.setRowCount(len(history))
        self.update_history_table()
        
        layout.addWidget(self.history_table)

        self.setLayout(layout)

    def update_history_table(self):
        self.history_table.setRowCount(len(history))
        for row, (serial, timestamp, file_type) in enumerate(history):
            self.history_table.setItem(row, 0, QtWidgets.QTableWidgetItem(str(serial)))
            self.history_table.setItem(row, 1, QtWidgets.QTableWidgetItem(timestamp))
            self.history_table.setItem(row, 2, QtWidgets.QTableWidgetItem(file_type))
            
            # Open Button
            open_button = QtWidgets.QPushButton("Open")
            open_button.clicked.connect(lambda _, s=serial: self.open_code(s))
            self.history_table.setCellWidget(row, 3, open_button)
            
            # Copy Text Button
            copy_button = QtWidgets.QPushButton("Copy Text")
            copy_button.clicked.connect(lambda _, s=serial: self.copy_text(s))
            self.history_table.setCellWidget(row, 4, copy_button)

    def open_code(self, serial):
        code = code_store.get(serial, "")
        if code:
            code_window = QtWidgets.QWidget()
            code_window.setWindowTitle(f"Code for Entry {serial}")
            code_window.setGeometry(200, 200, 800, 600)
            
            # Layout
            layout = QtWidgets.QVBoxLayout(code_window)
            
            # Code Display
            code_text_edit = QtWidgets.QPlainTextEdit()
            code_text_edit.setPlainText(code)
            code_text_edit.setReadOnly(True)
            layout.addWidget(code_text_edit)
            
            code_window.setLayout(layout)
            code_window.show()

    def copy_text(self, serial):
        code = code_store.get(serial, "")
        if code:
            clipboard = QtWidgets.QApplication.clipboard()
            clipboard.setText(code)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = PythonToDOCXApp()
    window.show()
    sys.exit(app.exec_())
