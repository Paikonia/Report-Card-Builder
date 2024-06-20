import os
import pandas as pd
import numpy as np
from mailmerge import MailMerge

from functions import calculate_average, calc_grading


def calculate_end_of_term(folder_path: str):
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
                # print(actual_sheets)
                missing_sheets = expected_sheets - actual_sheets
                if missing_sheets:
                    raise ValueError(f"Missing sheets {missing_sheets} in file {filename}")
                
                dfs_end_of_term = {
                    sheet_name: pd.read_excel(os.path.join(folder_path, filename), sheet_name=sheet_name)
                    for sheet_name in sheet_name_list
                }
                
                new_dfs = {
                    'Senior One': pd.DataFrame(columns=['student_id', 'name', 'fs', 'es', 'to', 'grade', 'comment']), 
                    'Senior Two': pd.DataFrame(columns=['student_id', 'name', 'fs', 'es', 'to', 'grade', 'comment']), 
                    'Senior Three': pd.DataFrame(columns=['student_id', 'name', 'fs', 'es', 'to', 'grade', 'comment']), 
                    'Senior Four': pd.DataFrame(columns=['student_id', 'name', 'fs', 'es', 'to', 'grade', 'comment']), 
                }
                
                for sheet_name in sheet_name_list:
                    df = dfs_end_of_term[sheet_name]
                    if list(df.columns) != expected_columns:
                        raise ValueError(f"Sheet {sheet_name} in file {filename} does not have the expected columns {expected_columns}")
                    
                    new_df = pd.DataFrame(columns=['name', 'fs', 'es', 'to', 'grade'])
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
                            to = round(fs + es)
                        grade = calc_grading(to)
                        new_df.at[i, 'Student ID'] = student_id
                        new_df.at[i, 'name'] = name
                        new_df.at[i, 'fs'] = fs if np.isnan(fs) else round(fs)
                        new_df.at[i, 'es'] = es
                        new_df.at[i, 'to'] = to
                        new_df.at[i, 'grade'] = grade
                    
                    new_dfs[sheet_name] = new_df
                
                calculated_averages[subject_name] = new_dfs
            
            except Exception as e:
                raise e
    
    return calculated_averages


def combine_to_report_merger_format(sheet_name, data_frame_list, processed_data):
    for subject_name in processed_data:
        df = processed_data[subject_name][sheet_name]
        cols = df.columns
        new_columns = {col: f'{subject_name[:3]}_{col}' for col in df.columns[1:-1]}
            
        for index, row in df.iterrows():
            if 'Student ID' not in data_frame_list[sheet_name].columns:
                data_frame_list[sheet_name]['Student ID'] = []
            if 'Name' not in data_frame_list[sheet_name].columns:
                data_frame_list[sheet_name]['Name'] = []

            student_name = str(row[cols[0]]).upper()
            if student_name not in data_frame_list[sheet_name]['Name'].astype(str).str.upper().tolist():
                new_row = {'Name': student_name}
                for col_name in new_columns.values():
                    new_row[col_name] = None
                
                data_frame_list[sheet_name] = pd.concat([data_frame_list[sheet_name], pd.DataFrame([new_row])], ignore_index=True)

            for col_name, col_value in new_columns.items():
                name_index = data_frame_list[sheet_name][data_frame_list[sheet_name]['Name'].astype(str).str.upper() == student_name].index
                data_frame_list[sheet_name].loc[name_index, col_value] = row[col_name]


def run_mail_merge(data_frame, template_path, output_path, class_name):
    n_path = os.path.join(output_path, class_name)
    os.makedirs(n_path, exist_ok=True)
    non_numeric_columns = data_frame.select_dtypes(exclude=['int', 'float']).columns
    data_frame[non_numeric_columns] = data_frame[non_numeric_columns].fillna('')
    
    numeric_columns = data_frame.select_dtypes(include=['int', 'float']).columns
    data_frame[numeric_columns] = data_frame[numeric_columns].fillna('').astype(str)
    
    records = data_frame.astype(str).to_dict('records')
    
    for record in records:
        record = {k: str(v) for k, v in record.items()}
        
        if 'Name' not in record:
            raise KeyError("'Name' field is missing in the record.")
        
        document = MailMerge(template_path)
        document.merge(**record)
        
        output_file = os.path.join(n_path, f'Report Card {record["Name"]}.docx')
        
        document.write(output_file)
        document.close()

def make_reports(unprocessed_marks, templete_path, output_path):
    subject_sheets = calculate_end_of_term(unprocessed_marks)
    dfs_end_of_term = {
        'Senior One': pd.DataFrame(), 
        'Senior Two': pd.DataFrame(), 
        'Senior Three': pd.DataFrame(), 
        'Senior Four': pd.DataFrame(), 
    }

    for sheet_name in dfs_end_of_term.keys():
        combine_to_report_merger_format(sheet_name, dfs_end_of_term, subject_sheets)

    for c in dfs_end_of_term.keys():
        file_name = f'Report Card Template - {c}.docx'
        path = os.path.join(templete_path, file_name)
        if os.path.exists(path=path):
            run_mail_merge(data_frame=dfs_end_of_term[c], template_path=path, output_path=output_path, class_name=c)
        else:
            raise FileExistsError(f'Template for {c} does not exist.')




if __name__ == '__main__':
    make_reports('./Marksheet Term One', './G', './')
    
    
    