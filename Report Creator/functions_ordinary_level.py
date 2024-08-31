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

def convert_to_percentage(marks, total):
    if np.isnan(marks):
        return np.nan
    if np.isnan(total):
        raise ValueError('Marks entered without a total score to work it out of.')
    
    return (marks/total)*100



def grading_ordinary_level(folder_path: str):
    if folder_path is None:
        raise FileNotFoundError('You have not given a valid folder path')
    if not os.path.isdir(folder_path):
        raise FileExistsError(f'The Folder you are trying to open {folder_path} does not exist')
    
    calculated_averages = {}
    sheet_name_list = ['Senior One', 'Senior Two', 'Senior Three', 'Senior Four']
    expected_columns = set(['Student ID', 'Name', 'A01', 'A02', 'A03', 'EOT', 'Comment'])
    
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            try:
                subject_name = os.path.splitext(filename)[0]
                subject_name = subject_name.split(' - ')[1]
                
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
                    
                    
                    df['Name'] = df['Name'].str.strip().str.upper()
                    
                    actual_columns = set(df.columns)
                    
                    if actual_columns != expected_columns:
                        raise ValueError(f"Sheet {sheet_name} in file {filename} does not have the expected columns.\nExpected Columns {expected_columns} Passed column {df.columns}")
                    
                    df['A01'] = pd.to_numeric(df['A01'], errors='coerce')
                    df['A02'] = pd.to_numeric(df['A02'], errors='coerce')
                    df['A03'] = pd.to_numeric(df['A03'], errors='coerce')
                    df['EOT'] = pd.to_numeric(df['EOT'], errors='coerce')
                    
                    total_marks_row = df[df['Name'] == 'TOTAL MARKS']
                    
                    if total_marks_row.empty:
                        raise ValueError(f"Subject {subject_name} - Sheet {sheet_name} does not contain a row with 'Total Marks'")
                    
                    total_marks_index = total_marks_row.index[0]
                    total_marks = df.loc[total_marks_index]
                    df = df.drop(total_marks_index)
                    
                    total_a01 = total_marks['A01']
                    total_a02 = total_marks['A02']
                    total_a03 = total_marks['A03']
                    total_eot = total_marks['EOT'] if not pd.isnull(total_marks['EOT']) else 100
                    
                    try:
                        df['A01'] = df['A01'].apply(lambda x: convert_to_percentage(x, total_a01))
                        df['A02'] = df['A02'].apply(lambda x: convert_to_percentage(x, total_a02))
                        df['A03'] = df['A03'].apply(lambda x: convert_to_percentage(x, total_a03))
                    except Exception as e:
                        raise Exception(f'An error "{str(e)}" occurred when processing:\nSubject {subject_name}, class {sheet_name}')
                    
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
                            to = round((fs * 0.2) + es)
                        grade = calc_grading(to)
                        average_grade = calc_grading(fs)
                        
                        new_df.at[i, 'Student ID'] = student_id
                        new_df.at[i, 'Name'] = name
                        new_df.at[i, 'A01'] = row['A01']
                        new_df.at[i, 'A02'] = row['A02']
                        new_df.at[i, 'A03'] = row['A03']
                        new_df.at[i, 'Average Score'] = fs if np.isnan(fs) else round(fs)
                        new_df.at[i, 'Average Grade'] = average_grade
                        new_df.at[i, 'Formative Score'] = fs if np.isnan(fs) else (round(fs) * 0.2)
                        new_df.at[i, 'EOT Score'] = es if not pd.isnull(row['EOT']) else np.nan
                        new_df.at[i, 'Total Score'] = to
                        new_df.at[i, 'Grade'] = grade
                    
                    new_dfs[sheet_name] = new_df
                
                calculated_averages[subject_name] = new_dfs
            
            except Exception as e:
                raise e
    
    return calculated_averages

