import glob
import pandas as pd
import os
from tabulate import tabulate as tb
strictly_male_titles = ['Rev. Mr.', 'Deacon', 'Father', 'Brother', 'Monsignor', 'Reverend Monsignor', 'Mr.', 'Sr.']
strictly_female_titles = ['Mrs.', 'Miss', 'Sister', 'Ms.']
# List of all RE titles
AllREtitles = ['Dr.', 'The Honorable', 'Col.', 'Cmsgt. Ret.', 'Rev. Mr.', 'Deacon', 'Judge', 
                'Lt. Col.', 'Col. Ret.', 'Major', 'Capt.', 'Maj. Gen.', 'Family of', 'Senator', 'Reverend', 
                'Lt.', 'Cmdr.', 'Msgt.', 'Sister', 'Drs.', 'Master', 'Sgt. Maj.', 'SMSgt.', 'Prof.', 'Lt. Col. Ret.', 'Rev. Dr.', 
                'Father', 'Brother', 'Bishop', 'Gen.', 'Admiral', 'Very Reverend', 'MMC', 'Monsignor', '1st Lt.', 'Reverend Monsignor', 
                'Maj.', 'Most Reverend', 'Bishop Emeritus','Mrs.', 'Mr.', 'Ms.', 'Sra.', 'Señor', 'Miss','Sr.', 'Family of']

# List of  special  titles
specialTitle = ['Dr.', 'The Honorable', 'Col.', 'Cmsgt. Ret.', 'Rev. Mr.', 'Deacon', 'Judge', 
                'Lt. Col.', 'Col. Ret.', 'Major', 'Capt.', 'Maj. Gen.', 'Family of', 'Senator', 'Reverend', 
                'Lt.', 'Cmdr.', 'Msgt.', 'Sister', 'Drs.', 'Master', 'Sgt. Maj.', 'SMSgt.', 'Prof.', 'Lt. Col. Ret.', 'Rev. Dr.', 
                'Father', 'Brother', 'Bishop', 'Gen.', 'Admiral', 'Very Reverend', 'MMC', 'Monsignor', '1st Lt.', 'Reverend Monsignor', 
                'Maj.', 'Most Reverend', 'Bishop Emeritus','Family of']

# List of common titles
commonTitles = ['Mrs.', 'Mr.', 'Ms.', 'Miss','Sr.','Sra.', 'Señor'] 

class LabelProcessor:
    def __init__(self, input_dir):
        self.input_dir = input_dir
    
    def process_files(self):
        files = glob.glob(os.path.join(self.input_dir, '*_export.CSV')) + glob.glob(os.path.join(self.input_dir, '*_export.csv'))
        for file in files:
            try:
                df = pd.read_csv(file, encoding='utf-8', low_memory=False, dtype=str)
                file_encoding = 'utf-8'
            except UnicodeDecodeError:
                df = pd.read_csv(file, encoding='ISO-8859-1', low_memory=False, dtype=str)
                file_encoding = 'ISO-8859-1'

            if not self.check_titles_and_genders(df):
                print("ERROR:")
                return False
        
        def remove_data_based_on_condition1(row):
            # Check if 'CnBio_First_Name' is equal to 'CnSpSpBio_First_Name' and remove data if True
            if row['CnBio_First_Name'] == row['CnSpSpBio_First_Name']:
                row['CnSpSpBio_Gender'] = None
                row['CnSpSpBio_Title_1'] = None
                row['CnSpSpBio_First_Name'] = None
                row['CnSpSpBio_Last_Name'] = None
                row['CnBio_Marital_status'] = 'WidSinDiv_0'

            return row
        df = df.apply(remove_data_based_on_condition1, axis=1)
    
        def remove_data_based_on_condition2(row):
            # Check if 'CnSpSpBio_Inactive' or 'CnSpSpBio_Deceased' is 'Yes' and remove data if True
            if row['CnSpSpBio_Inactive'] == 'Yes' or row['CnSpSpBio_Deceased'] == 'Yes' or row['CnBio_Marital_status'] == 'Widowed' or row['CnBio_Marital_status'] == 'Divorced' or row['CnBio_Marital_status'] == 'Separated'or row['CnBio_Marital_status'] == 'Annuled':
                row['CnSpSpBio_Gender'] = None
                row['CnSpSpBio_Title_1'] = None
                row['CnSpSpBio_First_Name'] = None
                row['CnSpSpBio_Last_Name'] = None
                row['CnBio_Marital_status'] = 'WidSinDiv_0'

            return row
        df = df.apply(remove_data_based_on_condition2, axis=1)

        def swap_rows_based_on_gender(row):
            if row['CnBio_Gender'] == 'Female' and row['CnSpSpBio_Gender'] == 'Male':
                temp_gender = row['CnBio_Gender']
                temp_first_name = row['CnBio_First_Name']
                temp_last_name = row['CnBio_Last_Name']
                temp_title = row['CnBio_Title_1']

                row['CnBio_Gender'] = row['CnSpSpBio_Gender']
                row['CnBio_First_Name'] = row['CnSpSpBio_First_Name']
                row['CnBio_Last_Name'] = row['CnSpSpBio_Last_Name']
                row['CnBio_Title_1'] = row['CnSpSpBio_Title_1']

                row['CnSpSpBio_Gender'] = temp_gender
                row['CnSpSpBio_First_Name'] = temp_first_name
                row['CnSpSpBio_Last_Name'] = temp_last_name
                row['CnSpSpBio_Title_1'] = temp_title
            return row
        df = df.apply(swap_rows_based_on_gender, axis=1)
        
        def update_titles_if_married(row):
        # This function update Ms and Miss to mrs if the last names are the same and marital status is married 
            if (row['CnBio_Last_Name'] == row['CnSpSpBio_Last_Name']) and (row['CnBio_Marital_status'] == 'Married' or row['CnSpSpBio_Marital_status'] == 'Married') and (row['CnBio_Title_1'] != 'Mr.') and (row['CnSpSpBio_Title_1'] == 'Miss' 
            or row['CnSpSpBio_Title_1'] == 'Ms.' or row['CnBio_Title_1'] == 'Miss' or row['CnBio_Title_1'] == 'Ms.'):
                row['CnSpSpBio_Title_1'] = 'Mrs.'
                row['CnBio_Title_1'] = 'Mrs.'
            return row
        df = df.apply(update_titles_if_married, axis=1)

        def update_titles_if_blank_mr(row):
        # This function update blanks titles to mr if gender is male or sptitle is mrs, ms, or miss
            if (pd.isnull(row['CnBio_Title_1']) and pd.notnull(row['CnBio_Last_Name']) and (row['CnBio_Gender'] == 'Male' or row['CnSpSpBio_Title_1'] == 'Mrs.' or row['CnSpSpBio_Title_1'] == 'Ms.' or row['CnSpSpBio_Title_1'] == 'Miss')):
                row['CnBio_Title_1'] = 'Mr.'
            return row
        df = df.apply(update_titles_if_blank_mr, axis=1)

        def update_titles_if_blank_ms(row):
        # This function updates blank titles to ms if gender is female or sptitle Mr.
            if (pd.isnull(row['CnBio_Title_1']) and pd.notnull(row['CnBio_Last_Name']) and (row['CnBio_Gender'] == 'Female' or row['CnSpSpBio_Title_1'] == 'Mr.')):
                row['CnBio_Title_1'] = 'Ms.'
            return row
        df = df.apply(update_titles_if_blank_ms, axis=1)

        def update_sptitles_if_blank_mr(row):
        # This function updates blank sptitles to mr if gender is male or title is mrs, ms, or miss
            if (pd.isnull(row['CnSpSpBio_Title_1']) and pd.notnull(row['CnSpSpBio_Last_Name']) and (row['CnSpSpBio_Gender'] == 'Male' or row['CnBio_Title_1'] == 'Mrs.' or row['CnBio_Title_1'] == 'Ms.' or row['CnBio_Title_1'] == 'Miss')):
                row['CnSpSpBio_Title_1'] = 'Mr.'
            return row
        df = df.apply(update_sptitles_if_blank_mr, axis=1)

        def update_sptitles_if_blank_ms(row):
        # This function updates blanks sptitles to ms if gender is female or title is mr
            if (pd.isnull(row['CnSpSpBio_Title_1']) and pd.notnull(row['CnSpSpBio_Last_Name']) and (row['CnSpSpBio_Gender'] == 'Female' or row['CnBio_Title_1'] == 'Mr.')):
                row['CnSpSpBio_Title_1'] = 'Ms.'
            return row
        df = df.apply(update_sptitles_if_blank_ms, axis=1)
        
        def update_marital_status_if_blank_married(row):
        # If marital status is blank and Last names are  equal, fill in with married. They might be brother and sister, but Add/sal will be mostly the same. 
            if (((pd.isnull(row['CnBio_Marital_status']) or (row['CnBio_Marital_status'] == 'Single')) and (row['CnSpSpBio_Last_Name'] == row['CnBio_Last_Name'])) or ((row['CnSpSpBio_Last_Name'] != row['CnBio_Last_Name']) and pd.notnull(row['CnSpSpBio_Last_Name']) )):
                row['CnBio_Marital_status'] = 'Married'
            return row
        df = df.apply(update_marital_status_if_blank_married, axis=1)

        def update_marital_status_Widowed(row):
            # updates marital status to Widowed if Deceased or Inactive = yes. it will also change instnaces where person is married to themselves (wrong status but avoids bad add/sal) 
            # Check if spouse-related fields are all blank
            spouse_info_blank = all(pd.isnull(row[field] ) or row[field].strip() == '' for field in ['CnSpSpBio_Title_1', 'CnSpSpBio_First_Name', 'CnSpSpBio_Last_Name'])
            # Update marital status to 'Widowed' based on specified conditions
            if ((row['CnSpSpBio_Deceased'] == 'Yes') or 
                (row['CnSpSpBio_Inactive'] == 'Yes') or 
                (row['CnBio_Marital_status'] == 'Divorced') or 
                (row['CnBio_Marital_status'] == 'Separated') or
                ((row['CnBio_Marital_status'] in ['Single', 'Married', 'Unknown', None] or pd.isnull(row['CnBio_Marital_status'])) and spouse_info_blank)):
                row['CnBio_Marital_status'] = 'WidSinDiv_0'
            return row
        df = df.apply(update_marital_status_Widowed, axis=1)

        def Different_Last_Name_1(row):
            # Check if last names are different, marital status is 'Married', 
            # and either first name or last name of the spouse is not null/blank
            if (row['CnBio_Last_Name'] != row['CnSpSpBio_Last_Name']) and \
            (row['CnBio_Marital_status'] == 'Married') and \
            ((pd.notnull(row['CnSpSpBio_First_Name']) and row['CnSpSpBio_First_Name'].strip()) or \
            (pd.notnull(row['CnSpSpBio_Last_Name']) and row['CnSpSpBio_Last_Name'].strip())):
                row['CnBio_Marital_status'] = 'DifferentLastName_1'
            return row
        df = df.apply(Different_Last_Name_1, axis=1)

        def Same_Last_Name_Same_Title_NonSpecial_2(row):
            # If last names are different but titles are the same and neither are special
            global specialTitle # uses list of titles, global is used so that it can be accessed inside the functions
            if ((row['CnBio_Last_Name'] == row['CnSpSpBio_Last_Name']) and (row['CnBio_Title_1'] == row['CnSpSpBio_Title_1']) and (row['CnBio_Marital_status'] == 'Married') and (row['CnBio_Title_1'] not in specialTitle) ): # 
                row['CnBio_Marital_status'] = 'SameLastNameSameTitleNonSpecial_2'
            return row
        df = df.apply(Same_Last_Name_Same_Title_NonSpecial_2, axis=1)
        
        def Same_Last_Name_Same_Title_Special_3(row):
            # If Last names are the same and the title is the same 
            global specialTitle
            if ((row['CnBio_Last_Name'] == row['CnSpSpBio_Last_Name']) and (row['CnBio_Marital_status'] == 'Married') and (row['CnBio_Title_1'] == row['CnSpSpBio_Title_1']) and (row['CnBio_Title_1'] in specialTitle) ):
                row['CnBio_Marital_status'] = 'SameLastNameSameTitleSpecial_3'
            return row
        df = df.apply(Same_Last_Name_Same_Title_Special_3, axis=1)
        
        def Same_Last_Name_Both_Specical_Title_4(row):
            # If Last names are the same and both have a special title
            global specialTitle
            if ((row['CnBio_Last_Name'] == row['CnSpSpBio_Last_Name']) and (row['CnBio_Marital_status'] == 'Married') and (row['CnBio_Title_1'] in specialTitle and row['CnSpSpBio_Title_1'] in specialTitle) ):
                row['CnBio_Marital_status'] = 'SameLastNameBothSpecicalTitle_4'
            return row
        df = df.apply(Same_Last_Name_Both_Specical_Title_4, axis=1)
        
        def Same_Last_Name_Main_Specical_Title_5(row):
            # If Last names are the same only main has special title
            global specialTitle
            if ((row['CnBio_Last_Name'] == row['CnSpSpBio_Last_Name']) and (row['CnBio_Marital_status'] == 'Married') and (row['CnBio_Title_1'] in specialTitle) ):
                row['CnBio_Marital_status'] = 'SameLastNameMainSpecicalTitle_5'
            return row
        df = df.apply(Same_Last_Name_Main_Specical_Title_5, axis=1) 
        
        def Same_Last_Name_Sp_Specical_Title_6(row): 
            # If Last names are the same only spouse has special title
            global specialTitle
            if ((row['CnBio_Last_Name'] == row['CnSpSpBio_Last_Name']) and (row['CnBio_Marital_status'] == 'Married') and (row['CnSpSpBio_Title_1'] in specialTitle) ):
                row['CnBio_Marital_status'] = 'SameLastNameSpSpecicalTitle_6'
            return row
        df = df.apply(Same_Last_Name_Sp_Specical_Title_6, axis=1)

        def Standard_Add_Sal_7(row): 
            # Standard Add/sal for married couple
            global commonTitles
            if ((row['CnBio_Last_Name'] == row['CnSpSpBio_Last_Name']) and (row['CnBio_Marital_status'] != 'Widowed') and 
                (row['CnBio_Marital_status'] == 'Married') and ((row['CnBio_Title_1'] in commonTitles) or (row['CnSpSpBio_Title_1'] in commonTitles)  ) ):
                row['CnBio_Marital_status'] = 'StandardAddSal_7'
            return row
        df = df.apply(Standard_Add_Sal_7, axis=1) 
        
        def Standard_Add_Sal_MaleSp_8(row):
            global commonTitles
            if (row['CnBio_Last_Name'] == row['CnSpSpBio_Last_Name']
                and row['CnBio_Marital_status'] != 'Widowed'
                and row['CnBio_Marital_status'] == 'Married'
                and (row['CnBio_Title_1'] in commonTitles or row['CnSpSpBio_Title_1'] in commonTitles)
                and row['CnSpSpBio_Gender'] == 'Male'
            ):
                row['CnBio_Marital_status'] = 'StandardAddSal_MaleSp_8'
            return row
        df = df.apply(Standard_Add_Sal_MaleSp_8, axis=1)
    
        def blank_names_Unchanged_AddSal(row):
            # Name info is blank, cannot concatenate a addsal
            if (pd.isnull(row['CnBio_Last_Name']) and pd.isnull(row['CnBio_First_Name'])):
                row['CnBio_Marital_status'] = 'Unchanged'
            return row
        df = df.apply(blank_names_Unchanged_AddSal, axis=1)

        # fills back in a blank space otherwise it would fill cell with 'nan'
        df['CnBio_First_Name'] = df['CnBio_First_Name'].loc[:].fillna('')
        df['CnBio_Last_Name'] = df['CnBio_Last_Name'].loc[:].fillna('')
        df['CnSpSpBio_First_Name'] = df['CnSpSpBio_First_Name'].loc[:].fillna('')
        df['CnSpSpBio_Last_Name'] = df['CnSpSpBio_Last_Name'].loc[:].fillna('')
        df['CnBio_Title_1'] = df['CnBio_Title_1'].loc[:].fillna('')
        df['CnSpSpBio_Title_1'] = df['CnSpSpBio_Title_1'].loc[:].fillna('')

        def concate_add_sal(row):

            if (row['CnBio_Marital_status'] == 'Unchanged' ):
            # Unchanged
            # Not enough data to concatenate a add/sal  
                addressee = str(row['CnAdrSal_Addressee'])
                salutation = str(row['CnAdrSal_Salutation'])
            
            elif (row['CnBio_Marital_status'] == 'WidSinDiv_0'):
            # WidSinDiv_0
            # Mr. Bryce Howard 
            # Mr. Howard
                # Check if Last Name is not blank
                if pd.notnull(row['CnBio_Last_Name']) and row['CnBio_Last_Name'].strip():
                    addressee = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_First_Name']) + ' ' + str(row['CnBio_Last_Name'])
                    salutation = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_Last_Name'])
                # If Last Name is blank, use First Name
                elif pd.notnull(row['CnBio_First_Name']) and row['CnBio_First_Name'].strip():
                    addressee = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_First_Name'])
                    salutation = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_First_Name'])
        
            elif (row['CnBio_Marital_status'] == 'DifferentLastName_1'):
            # Different_Last_Name_1
            # Mr. Bryce Howard and Mrs. Jennifer Ha 
            # Mr. Howard and Mrs. Ha
                addressee = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_First_Name']) + ' ' + str(row['CnBio_Last_Name']) +' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_First_Name']) + ' ' +  str( row['CnSpSpBio_Last_Name'])
                salutation = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_Last_Name']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_Last_Name'])

            elif (row['CnBio_Marital_status'] == 'SameLastNameSameTitleNonSpecial_2'):
            # Same_Last_Name_Same_Title_NonSpecial_2
            # Mr. Bryce Howard and Mr. Branden Howard
            # Mr Howard and Mr. Howard
                addressee = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_First_Name']) + ' ' + str(row['CnBio_Last_Name']) +' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_First_Name']) + ' ' +  str( row['CnSpSpBio_Last_Name'])
                salutation = str(row['CnBio_Title_1']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnBio_Last_Name'])

            elif (row['CnBio_Marital_status'] == 'SameLastNameSameTitleSpecial_3'):
            # Gives: Same_Last_Name_Same_Title_Special_3
            # Dr. Bryce Howard and Dr. Jen Howard
            # Dr. Howard and Dr. Howard
                addressee = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_First_Name']) + ' ' + str(row['CnBio_Last_Name']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_First_Name']) + ' ' +  str( row['CnSpSpBio_Last_Name'])
                salutation = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_Last_Name']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_Last_Name'])

            elif (row['CnBio_Marital_status'] == 'SameLastNameBothSpecicalTitle_4'):
            # Same_Last_Name_Both_Specical_Title_4 
            # Senator Bryce Howard and Dr. Jen Howard
            # Senator Howard and Dr. Howard
                addressee = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_First_Name']) + ' ' + str(row['CnBio_Last_Name']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_First_Name']) + ' ' +  str( row['CnSpSpBio_Last_Name'])
                salutation = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_Last_Name']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_Last_Name'])

            elif (row['CnBio_Marital_status'] == 'SameLastNameMainSpecicalTitle_5'):
            # Same_Last_Name_Main_Specical_Title_5 
            # Dr. Bryce Howard and Mrs. Howard
            # Dr. Howard and Mrs. Howard
                addressee = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_First_Name']) + ' ' + str(row['CnBio_Last_Name']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' +  str( row['CnSpSpBio_Last_Name'])
                salutation = str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_Last_Name']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_Last_Name'])       
            
            elif (row['CnBio_Marital_status'] == 'SameLastNameSpSpecicalTitle_6'):
            # Same_Last_Name_Sp_Specical_Title_6 
            # Dr. Jennifer Howard and Mr. Bryce Howard
            # Dr. Howard and Mr. Howard
                addressee = str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_First_Name']) + ' ' + str(row['CnSpSpBio_Last_Name']) + ' and ' + str(row['CnBio_Title_1']) + ' ' +  str( row['CnBio_Last_Name'])
                salutation = str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnSpSpBio_Last_Name']) + ' and ' + str(row['CnBio_Title_1']) + ' ' + str(row['CnBio_Last_Name'])      
            
            elif (row['CnBio_Marital_status'] == 'StandardAddSal_7'):
            # Standard_Add_Sal_7
            # Mr. and Mrs. Bryce Howard
            # Mr. and Mrs. Howard
                addressee = str(row['CnBio_Title_1']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnBio_First_Name']) + ' ' + str(row['CnBio_Last_Name'])
                salutation = str(row['CnBio_Title_1']) + ' and ' + str(row['CnSpSpBio_Title_1']) + ' ' + str(row['CnBio_Last_Name'])

            elif (row['CnBio_Marital_status'] == 'Standard_Add_Sal_MaleSp_8'):
                addressee = str(row['CnSpSpBio_Title_1']) + ' and ' + str(row['CnBio_Title_1']) + ' ' + str(row['CnSpSpBio_First_Name']) + ' ' + str(row['CnSpSpBio_Last_Name'])
                salutation = str(row['CnSpSpBio_Title_1']) + ' and ' + str(row['CnBio_Title_1']) + ' ' + str(row['CnSpSpBio_Last_Name'])

            else:
            # This will make the add/sal blank, which will help find edge cases. Any add/sal that did not fit the criteria above, will come out blank.
                addressee = ''
                salutation = ''
            return pd.Series({'CnAdrSal_Addressee': addressee, 'CnAdrSal_Salutation': salutation})
        df[['CnAdrSal_Addressee', 'CnAdrSal_Salutation']] = df.apply(concate_add_sal, axis=1)

        def add_bishop_fields(df):
            bishop_addressee_data = []
            bishop_salutation_data = []
            
            for index, row in df.iterrows():
                if pd.isna(row['CnBio_First_Name']) and pd.isna(row['CnBio_Last_Name']):
                    bishop_addressee_data.append('')
                    bishop_salutation_data.append('')
                    continue
                
                first_name = row['CnBio_First_Name'] if not pd.isna(row['CnBio_First_Name']) else ''
                last_name = row['CnBio_Last_Name'] if not pd.isna(row['CnBio_Last_Name']) else ''
                spouse_first_name = row['CnSpSpBio_First_Name'] if not pd.isna(row['CnSpSpBio_First_Name']) else ''

                if spouse_first_name:
                    addressee = f"{first_name} and {spouse_first_name} {last_name}"
                    salutation = f"{first_name} and {spouse_first_name}"
                else:
                    addressee = f"{first_name} {last_name}"
                    salutation = first_name
                
                bishop_addressee_data.append(addressee)
                bishop_salutation_data.append(salutation)
            
            new_columns = pd.DataFrame({
                'Bishop_Addressee': bishop_addressee_data,
                'Bishop_Salutation': bishop_salutation_data
            })
            
            df = pd.concat([df, new_columns], axis=1)

            columns = df.columns.tolist()
            columns.insert(12, columns.pop(columns.index('Bishop_Addressee')))
            columns.insert(13, columns.pop(columns.index('Bishop_Salutation')))
            df = df[columns]
            
            return df
        df = add_bishop_fields(df)


        base, ext = os.path.splitext(file)
        new_file = base + '_clean' + ext
        df.to_csv(f'{new_file}', index=False, encoding=file_encoding)
        print("\nAddress and Salutation processing completed. \n_export_clean.csv created")
        return True

    def check_titles_and_genders(self, df):
        # Collect rows that fail the checks

        failed_rows = []

        for _, row in df.iterrows():
            if row['CnBio_Gender'] == 'Male' and row['CnBio_Title_1'] in strictly_female_titles:
                failed_rows.append([row['CnBio_ID'], row['CnBio_Gender'], row['CnBio_Title_1'], row['CnSpSpBio_Gender'], row['CnSpSpBio_Title_1']])
            if row['CnBio_Gender'] == 'Female' and row['CnBio_Title_1'] in strictly_male_titles:
                failed_rows.append([row['CnBio_ID'], row['CnBio_Gender'], row['CnBio_Title_1'], row['CnSpSpBio_Gender'], row['CnSpSpBio_Title_1']])
            if row['CnSpSpBio_Gender'] == 'Male' and row['CnSpSpBio_Title_1'] in strictly_female_titles:
                failed_rows.append([row['CnBio_ID'], row['CnBio_Gender'], row['CnBio_Title_1'], row['CnSpSpBio_Gender'], row['CnSpSpBio_Title_1']])
            if row['CnSpSpBio_Gender'] == 'Female' and row['CnSpSpBio_Title_1'] in strictly_male_titles:
                failed_rows.append([row['CnBio_ID'], row['CnBio_Gender'], row['CnBio_Title_1'], row['CnSpSpBio_Gender'], row['CnSpSpBio_Title_1']])
            if row['CnBio_First_Name'] and row['CnBio_Gender'] == 'Unknown' and row['CnBio_Title_1'] in strictly_male_titles + strictly_female_titles:
                failed_rows.append([row['CnBio_ID'], row['CnBio_Gender'], row['CnBio_Title_1'], row['CnSpSpBio_Gender'], row['CnSpSpBio_Title_1']])
            if row['CnSpSpBio_First_Name'] and row['CnSpSpBio_Gender'] == 'Unknown' and row['CnSpSpBio_Title_1'] in strictly_male_titles + strictly_female_titles:
                failed_rows.append([row['CnBio_ID'], row['CnBio_Gender'], row['CnBio_Title_1'], row['CnSpSpBio_Gender'], row['CnSpSpBio_Title_1']])

        if failed_rows:
            print("\nThese some of errors found '_export'. There are potentially more:")
            print(tb(failed_rows, headers=["CnBio_ID", "Gender", "Title", "SpSpGender", "SpSpTitle"], tablefmt="grid"))
            return False
        
        return True
    
