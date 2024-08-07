from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton, QVBoxLayout, QWidget, QLabel, QMessageBox, QTextEdit
from PySide6.QtCore import QIODevice
import sys
import os
import glob
from ack_letter import AckLetterProcessor, Logger
from mail_merge import MailMerge, find_latest_complete_csv, find_docx_template
from labels import LabelProcessor

class EmittingStream(QIODevice):
    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit
        self.open(QIODevice.WriteOnly)

    def writeData(self, data):
        text = str(data, 'utf-8', errors='replace')
        cursor = self.text_edit.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        cursor.insertText(text)
        self.text_edit.setTextCursor(cursor)
        self.text_edit.ensureCursorVisible()
        return len(data)

    def write(self, text):
        self.writeData(text.encode('utf-8'))

    def flush(self):
        pass

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.input_dir = os.getcwd()
        self.output_dir = os.getcwd()
        self.template_path = ''
        self.processor = None
        self.initUI()
        self.check_required_files()

    def initUI(self):
        self.setWindowTitle('Acknowledgment Letter and Mail Merge Tool')
        self.resize(800, 600)  # Set the initial size of the window

        # Layout and widgets
        layout = QVBoxLayout()

        self.select_folder_button = QPushButton('Change Input Output Folder')
        self.select_folder_button.clicked.connect(self.select_folder)

        self.run_labels_button = QPushButton('Run LabelProcessor')
        self.run_labels_button.clicked.connect(self.run_labels)
        self.run_labels_button.setEnabled(True)

        self.run_ack_button = QPushButton('Run AckLetterProcessor')
        self.run_ack_button.clicked.connect(self.run_ack_letter)
        self.run_ack_button.setEnabled(False)

        self.select_template_button = QPushButton('Select DOCX for Mail Merge')
        self.select_template_button.clicked.connect(self.select_template)
        self.select_template_button.setEnabled(False)

        self.run_mail_merge_button = QPushButton('Run MailMerge')
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

        # Print the initial folder info to the terminal window
        last_folder = os.path.basename(self.input_dir)
        print(f"Default input/output folder: {last_folder}")

    def check_required_files(self):
        required_files = {
            'DOCX template': '*.docx',
            '_mail CSV': '*_mail.[Cc][Ss][Vv]',
            '_export CSV': '*_export.[Cc][Ss][Vv]'
        }

        missing_files = []
        for file_type, pattern in required_files.items():
            if not glob.glob(os.path.join(self.input_dir, pattern)):
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
            print(f"Selected Folder: {last_folder}")
            self.run_labels_button.setEnabled(True)
            self.check_required_files()

    def run_labels(self):
        missing_files = self.check_required_files()
        if missing_files:
            print("Cannot run LabelProcessor. Missing required files:")
            for file_type in missing_files:
                print(f" - {file_type}")
            return
        
        warning_message = (
            "Warning - This will ensure Add/Sal are correct. "
            "This ONLY works if Gender, Titles and notes have been manually reviewed for HoH and Spouse. "
            "The file MUST have '_export' at the end of the file name. "
            "Have you manually reviewed?"
        )
        response = QMessageBox.warning(
            self,
            "Manual Review Confirmation",
            warning_message,
            QMessageBox.Yes | QMessageBox.No
        )
        if response == QMessageBox.Yes:
            label_processor = LabelProcessor(self.input_dir)
            label_processor.process_files()
        else:
            QMessageBox.information(self, "Process Stopped", "Please manually review the data and try again.")

    def run_ack_letter(self):
        logger = Logger(print)
        self.processor = AckLetterProcessor(self.input_dir, logger)
        fidelis_files = self.processor.find_csv_files('Fidelis.[Cc][Ss][Vv]')
        if not fidelis_files:
            response = QMessageBox.question(
                self, 
                "OCA TY letters use Fidelis.csv, but is not found.\nDo you want to continue without this file and without the column 'Fidelis Society' being added?",
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

    def run_mail_merge(self):
        try:
            latest_csv = find_latest_complete_csv(self.output_dir)
            mail_merge = MailMerge(self.output_dir)
            mail_merge.merge(latest_csv, self.template_path)
            print("Mail merge completed successfully.")
        except Exception as e:
            print(f"An error occurred during mail merge: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec())
