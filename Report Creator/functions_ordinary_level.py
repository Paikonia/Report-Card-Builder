import numpy as np
import pandas as pd
import os


def calculate_average(r1, r2, r3):
    values = [r1, r2, r3]
    non_null_values = [value for value in values if pd.notnull(value) and pd.to_numeric(value, errors='coerce') == value]
    
    if len(non_null_values) == 3:
        return np.nanmean(non_null_values) 
    elif len(non_null_values) == 2:
        return np.nanmean(non_null_values)
    elif len(non_null_values) == 1:
        return non_null_values[0]
    else:
        return np.nan

def calc_grading(total):
    if(np.isnan(total)):
        return ''
    if total >= 90:
        return 'A*'
    elif total >= 80:
        return 'A'
    elif total >= 70:
        return 'B'
    elif total >= 60:
        return 'C'
    elif total >= 50:
        return 'D'
    elif total >= 40:
        return 'E'
    elif total >= 30:
        return 'F'
    else:
        return 'G'
    
def abbreviate_column_name(column_name):
    column_name = column_name.replace('Formative Score', 'fs')
    column_name = column_name.replace('EOT Score', 'es')
    column_name = column_name.replace('Total Score', 'to')
    column_name = column_name.replace('Grade', 'grade')
    column_name = column_name.replace('AVERAGE', 'avg')
    column_name = column_name.replace('Average', 'avg')
    column_name = column_name.replace('A01', 'a01')
    column_name = column_name.replace('A02', 'a02')
    column_name = column_name.replace('A03', 'a03')
    column_name = column_name.replace('Mid Term', 'mt')
    return column_name.replace(' ', '_')


def grading_ordinary_level(folder_path: str):
    if folder_path is None:
        raise FileNotFoundError('You have not given a valid folder path')
    if not os.path.isdir(folder_path):
        raise FileExistsError(f'The Folder you are trying to open {folder_path} does not exist')
    
    calculated_averages = {}
    sheet_name_list = ['Senior One', 'Senior Two', 'Senior Three', 'Senior Four']
    expected_columns = ['Student ID', 'Name', 'A01', 'A02', 'A03', 'EOT', 'Comment']
    
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            try:
                subject_name = os.path.splitext(filename)[0]
                parts = subject_name.split(' ')
                if len(parts) == 3:
                    subject_name = parts[2]
                elif len(parts) > 3:
                    subject_name = f"{parts[2]} {parts[3]}"
                
                expected_sheets = set(sheet_name_list)
                actual_sheets = set(pd.ExcelFile(os.path.join(folder_path, filename)).sheet_names)
                               
                missing_sheets = expected_sheets - actual_sheets
                if missing_sheets:
                    raise ValueError(f"Missing sheets {missing_sheets} in file {filename}")
                
                dfs_end_of_term = {
                    sheet_name: pd.read_excel(os.path.join(folder_path, filename), sheet_name=sheet_name)
                    for sheet_name in sheet_name_list
                }
                columns = ['student_id', 'Name', 'A01', 'A02', 'A03', 'Formative Score', 'EOT Score', 'Total Score', 'Grade', 'Comment']
                new_dfs = {
                    'Senior One': pd.DataFrame(columns=columns), 
                    'Senior Two': pd.DataFrame(columns=columns), 
                    'Senior Three': pd.DataFrame(columns=columns), 
                    'Senior Four': pd.DataFrame(columns=columns), 
                }
                
                for sheet_name in sheet_name_list:
                    df = dfs_end_of_term[sheet_name]
                    if list(df.columns) != expected_columns:
                        raise ValueError(f"Sheet {sheet_name} in file {filename} does not have the expected columns.\nExpected Columns {expected_columns} Passed column {df.columns}")
                    
                    new_df = pd.DataFrame(columns=['Name', 'A01', 'A02', 'A03', 'Average Score', 'Formative Score', 'EOT Score', 'Total Score', 'Grade'])
                    for i, row in df.iterrows():
                        name = row['Name']
                        student_id = row['Student ID']
                        
                        fs = calculate_average(row['A01'], row['A02'], row['A03'])
                        value = row['EOT']
                        es = round(value * 0.8) if not pd.isnull(value) and not isinstance(value, str) else np.nan
                        if np.isnan(fs) and np.isnan(es):
                            to = np.nan
                        else:
                            fs = round(fs) if not np.isnan(fs) else 0
                            es = round(es) if not np.isnan(es) else 0
                            to = round((fs *0.2) + es)
                        grade = calc_grading(to)
                        average_grade = calc_grading(fs)
                        new_df.at[i, 'Student ID'] = student_id
                        new_df.at[i, 'Name'] = name
                        new_df.at[i, 'A01'] = row['A01']
                        new_df.at[i, 'A02'] = row['A02']
                        new_df.at[i, 'A03'] = row['A03']
                        new_df.at[i, 'Average Score'] = fs if np.isnan(fs) else round(fs)
                        new_df.at[i, 'Average Grade'] = average_grade
                        new_df.at[i, 'Formative Score'] = fs if np.isnan(fs) else (round(fs) * 0.2)     # fs for Formative Score 
                        new_df.at[i, 'EOT Score'] = es                                                  # es for EOT Score
                        new_df.at[i, 'Total Score'] = to                                                # to for Total Score
                        new_df.at[i, 'Grade'] = grade                                                   # grade for Grade
                    
                    new_dfs[sheet_name] = new_df
                
                calculated_averages[subject_name] = new_dfs
            
            except Exception as e:
                raise e
    
    return calculated_averages



