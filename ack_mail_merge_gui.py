from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton, QVBoxLayout, QWidget, QLabel, QMessageBox, QTextEdit
from PySide6.QtCore import QIODevice, QThread, Signal, Qt
from PySide6.QtGui import QTextCursor
import sys
import os
import glob
from mail_merge import MailMerge, find_latest_complete_csv, Logger
from ack_letter import AckLetterProcessor
from labels import LabelProcessor

class EmittingStream(QIODevice):
    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit
        self.open(QIODevice.WriteOnly)

    def writeData(self, data):
        text = str(data, 'utf-8', errors='replace')
        cursor = self.text_edit.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(text)
        self.text_edit.setTextCursor(cursor)
        self.text_edit.ensureCursorVisible()
        return len(data)

    def write(self, text):
        self.writeData(text.encode('utf-8'))

    def flush(self):
        pass

class Worker(QThread):
    progress = Signal(str, bool)
    finished = Signal(str)

    def __init__(self, output_dir, template_path):
        super().__init__()
        self.output_dir = output_dir
        self.template_path = template_path

    def run(self):
        try:
            latest_csv = find_latest_complete_csv(self.output_dir)
            self.progress.emit(f"", False)
            logger = Logger()
            logger.log_signal.connect(self.log)
            mail_merge = MailMerge(self.output_dir, logger)
            mail_merge.merge(latest_csv, self.template_path)
            self.finished.emit("")
        except Exception as e:
            self.finished.emit(f"An error occurred during mail merge: {e}")

    def log(self, message, update_only=False):
        self.progress.emit(message, update_only)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.input_dir = os.getcwd()
        self.output_dir = os.getcwd()
        self.template_path = ''
        self.processor = None
        self.initUI()
        self.check_required_files()

        # Update the Logger instantiation in the GUI
        self.logger = Logger()
        self.logger.log_signal.connect(self.log_message)

    def initUI(self):
        self.setWindowTitle('Acknowledgment Letter and Mail Merge Tool')
        self.resize(800, 600)  # Set the initial size of the window

        # Layout and widgets
        layout = QVBoxLayout()

        self.select_folder_button = QPushButton('Click to change default Input/Output Folder')
        self.select_folder_button.clicked.connect(self.select_folder)

        self.run_labels_button = QPushButton('Click to begin process to correct Addressee and Salutation in _export.csv')
        self.run_labels_button.clicked.connect(self.run_labels)
        self.run_labels_button.setEnabled(True)

        self.run_ack_button = QPushButton('Click to begin process to Format mailing data')
        self.run_ack_button.clicked.connect(self.run_ack_letter)
        self.run_ack_button.setEnabled(False)

        self.select_template_button = QPushButton('Select DOCX template for Mail Merge')
        self.select_template_button.clicked.connect(self.select_template)
        self.select_template_button.setEnabled(False)

        self.run_mail_merge_button = QPushButton('Click to begin process to merge mailing data into Word template')
        self.run_mail_merge_button.clicked.connect(self.run_mail_merge)
        self.run_mail_merge_button.setEnabled(False)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)

        # Add widgets to layout
        layout.addWidget(self.select_folder_button)
        layout.addWidget(self.run_labels_button)
        layout.addWidget(self.run_ack_button)
        layout.addWidget(self.select_template_button)
        layout.addWidget(self.run_mail_merge_button)
        layout.addWidget(self.log_output)

        # Set the central widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Redirect stdout to the QTextEdit widget
        sys.stdout = EmittingStream(self.log_output)
        # sys.stderr = EmittingStream(self.log_output)  # Keep stderr commented for now

        # Initialize logger
        self.logger = Logger()
        self.logger.log_signal.connect(self.log_message)

        # Print the initial folder info to the terminal window
        last_folder = os.path.basename(self.input_dir)
        self.logger.log(f"Default input/output folder: {last_folder}")

    def check_required_files(self):
        required_files = {
            'DOCX template': '*.docx',
            '_mail CSV': '*_mail.[Cc][Ss][Vv]',
            '_export CSV': '*_export.[Cc][Ss][Vv]'
        }
        missing_files = []
        for file_type, pattern in required_files.items():
            found_files = glob.glob(os.path.join(self.input_dir, pattern))
            if not found_files:
                missing_files.append(file_type)

        if missing_files:
            print("Missing required files:")
            for file_type in missing_files:
                print(f" - {file_type}")

        return missing_files

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folder:
            self.input_dir = folder
            self.output_dir = folder
            last_folder = os.path.basename(folder)
            print(f"\nSelected Folder: {last_folder}")
            self.run_labels_button.setEnabled(True)
            self.check_required_files()

    def run_labels(self):
        missing_files = self.check_required_files()
        if missing_files:
            print("Cannot continue. Missing required files:")
            for file_type in missing_files:
                print(f" - {file_type}")
            return
        
        warning_message = (
            "          BEFORE RUNNING THIS PROCESS:"
            "\nYou must manually review Genders, Titles and notes\n"
            "\n Have you manually reviewed the '_export.csv' file?"
        )
        response = QMessageBox.warning(
            self,
            "Manual Review Confirmation",
            warning_message,
            QMessageBox.Yes | QMessageBox.No
        )
        if response == QMessageBox.Yes:
            label_processor = LabelProcessor(self.input_dir)
            if label_processor.process_files():
                self.run_ack_button.setEnabled(True)
            else:
                print("Despite your 'review' of the data, errors in Genders and titles were found. These are only potentially some errors. Review _export.csv again")
        else:
            QMessageBox.information(self, "Process Stopped")

    def run_ack_letter(self):
        confirm_message = (
            "This will use the cleaned data from _export while it formats the mailing data\n"
            "\n                               Do you want to continue?"
        )
        response = QMessageBox.question(
            self,
            "Confirmation",
            confirm_message,
            QMessageBox.Yes | QMessageBox.No
        )
        if response == QMessageBox.No:
            self.logger.log("Process stopped by the user.")
            return

        logger = Logger()
        logger.log_signal.connect(self.log_message)
        self.processor = AckLetterProcessor(self.input_dir, logger)
        fidelis_files = self.processor.find_csv_files('Fidelis.[Cc][Ss][Vv]')
        if not fidelis_files:
            response = QMessageBox.question(
                self, 
                "Fidelis.csv Not Found",
                "OCA TY letters use Fidelis.csv, but is not found.\n                 Do you want to continue?",
                QMessageBox.Yes | QMessageBox.No
            )
            if response == QMessageBox.No:
                QMessageBox.information(self, "Process Stopped", "Please add the Fidelis.csv file and run the script again.")
                return
            else:
                self.processor.set_continue_without_fidelis(True)
        result = self.processor.process_files()
        if result:
            print(result)
        self.select_template_button.setEnabled(True)

    def select_template(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        template, _ = QFileDialog.getOpenFileName(self, "Select DOCX Template", self.input_dir, "DOCX Files (*.docx)", options=options)
        if template:
            self.template_path = template
            self.run_mail_merge_button.setEnabled(True)
            self.logger.log(f"Selected template: {template}")
        else:
            self.logger.log("No template selected.")

    def run_mail_merge(self):
        confirm_message = (
            "This will use do a word merge using the docx template in this folder\n"
            "\nDo you want to continue?"
        )
        response = QMessageBox.question(
            self,
            "Confirmation",
            confirm_message,
            QMessageBox.Yes | QMessageBox.No
        )
        if response == QMessageBox.No:
            self.logger.log("Process stopped by the user.")
            return

        # Call the worker for the mail merge process
        self.worker = Worker(self.output_dir, self.template_path)
        self.worker.progress.connect(self.log_message)
        self.worker.finished.connect(self.on_merge_finished)
        self.worker.start()

    def on_merge_finished(self, message):
        self.logger.log(message)

    def log_message(self, message, update_only=False):
        cursor = self.log_output.textCursor()
        if update_only:
            cursor.movePosition(QTextCursor.End)
            cursor.select(QTextCursor.LineUnderCursor)
            cursor.removeSelectedText()
            cursor.insertText(message)
        else:
            cursor.movePosition(QTextCursor.End)
            if not self.log_output.toPlainText().endswith('\n'):
                cursor.insertBlock()
            cursor.insertText(message)
        self.log_output.setTextCursor(cursor)
        self.log_output.ensureCursorVisible()
        QApplication.processEvents()  # Ensure the GUI updates in real-time


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec())
