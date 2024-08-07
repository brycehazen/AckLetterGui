import pandas as pd
from docx import Document
import os
import glob
from docxcompose.composer import Composer
from docx import Document as Document_compose

class MailMerge:
    def __init__(self, output_dir):
        self.output_dir = output_dir

    def read_csv_file(self, file_path):
        """Attempt to read a CSV file with multiple encodings."""
        try:
            # Try reading with UTF-8 first
            return pd.read_csv(file_path, encoding='utf-8'), 'utf-8'
        except UnicodeDecodeError:
            # If UTF-8 fails, try ISO-8859-1
            return pd.read_csv(file_path, encoding='ISO-8859-1'), 'ISO-8859-1'

    def merge(self, data_path, template_path):
        # Read the data
        df, encoding = self.read_csv_file(data_path)

        # Process each row in the data
        for index, row in df.iterrows():
            # Load the template document
            doc = Document(template_path)

            # Replace placeholders with actual data
            for paragraph in doc.paragraphs:
                self._replace_placeholders(paragraph, row)

            # Save the personalized document
            output_path = os.path.join(self.output_dir, f"merged_letter_{index}.docx")
            doc.save(output_path)

        # Combine individual documents into a single document
        self.combine_documents()

        # Clean up individual merged files
        self.cleanup_individual_files()

    def _replace_placeholders(self, paragraph, data):
        for key, value in data.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(f"«{key}»", str(value))

    def combine_documents(self):
        files = sorted(glob.glob(os.path.join(self.output_dir, 'merged_letter_*.docx')))
        merged_document = Document_compose(files[0])
        composer = Composer(merged_document)

        for file in files[1:]:
            doc = Document_compose(file)
            composer.append(doc)

        output_path = os.path.join(self.output_dir, "combined_ack_letters.docx")
        composer.save(output_path)

    def cleanup_individual_files(self):
        files = glob.glob(os.path.join(self.output_dir, 'merged_letter_*.docx'))
        for file in files:
            os.remove(file)

def find_latest_complete_csv(output_dir):
    list_of_files = glob.glob(os.path.join(output_dir, '*_complete.csv'))  # Adjust the pattern if necessary
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file

def find_docx_template(input_dir):
    list_of_files = glob.glob(os.path.join(input_dir, '*.docx'))
    if len(list_of_files) == 1:
        return list_of_files[0]
    else:
        raise FileNotFoundError("Exactly one .docx template must be present in the directory.")

# Example usage:
if __name__ == "__main__":
    input_dir = os.getcwd()  # Use the current directory for input
    output_dir = os.getcwd()  # Use the current directory for output

    try:
        template_path = find_docx_template(input_dir)
        data_path = find_latest_complete_csv(output_dir)

        mail_merge = MailMerge(output_dir)
        mail_merge.merge(data_path, template_path)

        print("Mail merge completed successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")
