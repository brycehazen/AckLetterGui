import pandas as pd
import glob
import os
from tabulate import tabulate

class Logger:
    def __init__(self, output_function):
        self.output_function = output_function

    def log(self, message):
        self.output_function(message)

class AckLetterProcessor:
    def __init__(self, input_dir, logger):
        self.input_dir = input_dir
        self.logger = logger
        self.continue_without_fidelis = False

    def set_continue_without_fidelis(self, decision):
        self.continue_without_fidelis = decision

    def find_csv_files(self, pattern):
        """Find CSV files in the specified directory matching a given pattern."""
        return glob.glob(os.path.join(self.input_dir, pattern))

    def read_csv_file(self, file_path):
        """Attempt to read a CSV file with multiple encodings."""
        try:
            # Try reading with UTF-8 first
            return pd.read_csv(file_path, encoding='utf-8'), 'utf-8'
        except UnicodeDecodeError:
            # If UTF-8 fails, try ISO-8859-1
            return pd.read_csv(file_path, encoding='ISO-8859-1'), 'ISO-8859-1'

    def check_for_missing_records(self, df_mail, df_clean):
        """Check for missing Constituent IDs in mail and clean files and print them, including Addressee details, excluding visitors."""
        # Filter out records containing 'Visitors -' or 'Visitor -' in the Addressee fields
        df_mail_filtered = df_mail.loc[~df_mail['Addressee'].str.contains('Visitors -|Visitor -', regex=True, na=False)]
        df_clean_filtered = df_clean.loc[~df_clean['CnAdrSal_Addressee'].str.contains('Visitors -|Visitor -', regex=True, na=False)]

        # Create sets of Constituent IDs
        mail_ids = set(df_mail_filtered['Constituent ID'])
        clean_ids = set(df_clean_filtered['CnBio_ID'])

        # Identify missing records
        missing_in_mail = clean_ids - mail_ids
        missing_in_clean = mail_ids - clean_ids

        # Prepare data for table format
        missing_mail_data = [(cid, df_clean_filtered[df_clean_filtered['CnBio_ID'] == cid]['CnAdrSal_Addressee'].values[0])
                             for cid in missing_in_mail]
        missing_clean_data = [(cid, df_mail_filtered[df_mail_filtered['Constituent ID'] == cid]['Addressee'].values[0])
                              for cid in missing_in_clean]

        # Print details of missing records in table format
        if missing_in_mail:
            self.logger.log("\nRecords found in '_clean.csv' but not in '_mail.csv':")
            self.logger.log(tabulate(missing_mail_data, headers=["ID", "Addressee"], tablefmt="grid"))

        if missing_in_clean:
            self.logger.log("\nRecords found in '_mail.csv' but not in '_clean.csv':")
            self.logger.log(tabulate(missing_clean_data, headers=["ID", "Addressee"], tablefmt="grid"))

    def process_files(self):
        # Find files
        mail_files = self.find_csv_files('*_mail.[Cc][Ss][Vv]')
        clean_files = self.find_csv_files('*_clean.[Cc][Ss][Vv]')
        fidelis_files = self.find_csv_files('Fidelis.[Cc][Ss][Vv]')

        # Read files and capture encoding used
        if mail_files:
            df_mail, mail_encoding = self.read_csv_file(mail_files[0])
        else:
            df_mail, mail_encoding = None, None

        if clean_files:
            df_clean, clean_encoding = self.read_csv_file(clean_files[0])
        else:
            df_clean, clean_encoding = None, None

        # Check if required files are loaded
        if df_mail is None or df_clean is None:
            self.logger.log("Required files '_mail.csv' and/or '_clean.csv' not found. Please ensure both files are present and try again.")
            return

        # Check for Fidelis.csv and handle accordingly
        fidelis_used = False
        if not fidelis_files:
            if not self.continue_without_fidelis:
                self.logger.log("Exiting, please add the Fidelis.csv file and run the script again")
                return
            else:
                df_fidelis = None
        else:
            df_fidelis, fidelis_encoding = self.read_csv_file(fidelis_files[0])
            fidelis_used = True

        # Check for missing records
        self.check_for_missing_records(df_mail, df_clean)

        # Process files according to defined functions
        processed_data = self.process_data(df_mail, df_clean, df_fidelis, fidelis_used)

        # Use the encoding that was successful for saving
        output_path = os.path.join(self.input_dir, f"{pd.Timestamp.now().strftime('%Y-%m-%d')} OCA Ack_complete.csv")
        processed_data.to_csv(output_path, index=False, encoding=mail_encoding)

        return f"Processed data saved to {output_path}"

    def process_data(self, df_mail, df_clean, df_fidelis, fidelis_used):
        df_mail['Amount'] = df_mail['Amount'].astype(str).str.replace('[^\d.]', '', regex=True)
        df_mail['Amount'] = pd.to_numeric(df_mail['Amount'], errors='coerce')
        df_mail['Gift date'] = pd.to_datetime(df_mail['Gift date'], errors='coerce')
        df_mail = df_mail.loc[~df_mail['Addressee'].str.contains('Visitors -|Visitor -', regex=True, na=False)].copy()

        # Ensure that merge does not create duplicates by keeping the clean data unique per 'Constituent ID'
        df_clean_unique = df_clean.drop_duplicates(subset=['CnBio_ID'])
        df_mail = df_mail.merge(df_clean_unique[['CnBio_ID', 'CnAdrSal_Addressee', 'CnAdrSal_Salutation']],
                                left_on='Constituent ID', right_on='CnBio_ID', how='left')
        df_mail['Addressee'] = df_mail['CnAdrSal_Addressee'].combine_first(df_mail['Addressee'])
        df_mail['Salutation'] = df_mail['CnAdrSal_Salutation'].combine_first(df_mail['Salutation'])
        df_mail.drop(columns=['CnBio_ID', 'CnAdrSal_Addressee', 'CnAdrSal_Salutation'], inplace=True)

        if fidelis_used:
            fidelis_ids = df_fidelis['Constituent ID'].unique()
            df_mail['Fidelis Society'] = df_mail['Constituent ID'].apply(lambda x: 'Fidelis Society' if x in fidelis_ids else '')
        else:
            df_mail['Fidelis Society'] = ''

        grouped = df_mail.groupby(['Constituent ID', 'Fund description_1'])
        final_rows = []

        for name, group in grouped:
            selected_data = group[group['Gift type'] == 'Pledge'] if 'Pledge' in group['Gift type'].values else group
            amount_sum = selected_data['Amount'].sum()
            latest_date = selected_data['Gift date'].max()
            gift_type = 'Pledge' if 'Pledge' in selected_data['Gift type'].values else 'Cash'

            final_row = {
                'Constituent ID': name[0],
                'Addressee': group['Addressee'].iloc[0],
                'Salutation': group['Salutation'].iloc[0],
                'Address_Line_1': group['Address line 1'].iloc[0],
                'Address line 2': group['Address line 2'].iloc[0],
                'Address line 3': group['Address line 3'].iloc[0],
                'City': group['City'].iloc[0],
                'State': group['State'].iloc[0],
                'ZIP_Code': group['ZIP Code'].iloc[0],
                'Gift type': gift_type,
                'Gift subtype': group['Gift subtype'].iloc[0],
                'Amount': amount_sum,
                'Fund description_1': name[1],
                'Gift date': latest_date,
                'Pay Method': group['Pay Method'].iloc[0],
                'Installment Frequency': group['Installment Frequency'].iloc[0],
            }

            if fidelis_used:
                final_row['Fidelis Society'] = group['Fidelis Society'].iloc[0]

            final_rows.append(final_row)

        final_data = pd.DataFrame(final_rows)
        column_order = ['Constituent ID', 'Addressee', 'Salutation', 'Address_Line_1', 'Address line 2', 'Address line 3',
                        'City', 'State', 'ZIP_Code', 'Gift type', 'Gift subtype', 'Amount', 'Fund description_1', 'Gift date', 'Pay Method', 'Installment Frequency']

        if fidelis_used:
            column_order.append('Fidelis Society')
        # reorder columns
        final_data = final_data[column_order]
        # change gift type wording
        final_data['Gift type'] = final_data['Gift type'].replace({'Cash': 'gift', 'Pledge': 'pledge'})

        return final_data

def main():
    input_dir = os.getcwd()  # Use the current working directory
    logger = Logger(print)  # Use print function for logging
    processor = AckLetterProcessor(input_dir, logger)
    result = processor.process_files()
    print(result)

if __name__ == "__main__":
    main()
