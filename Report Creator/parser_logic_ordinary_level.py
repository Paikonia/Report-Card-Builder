import os
import pandas as pd
import numpy as np

from functions_ordinary_level import calculate_average, calc_grading, abbreviate_column_name


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
               
                missing_sheets = expected_sheets - actual_sheets
                if missing_sheets:
                    raise ValueError(f"Missing sheets {missing_sheets} in file {filename}")
                
                dfs_end_of_term = {
                    sheet_name: pd.read_excel(os.path.join(folder_path, filename), sheet_name=sheet_name)
                    for sheet_name in sheet_name_list
                }
                columns = ['student_id', 'Name', 'Formative Score', 'EOT Score', 'Total Score', 'Grade', 'Comment']
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
                    
                    new_df = pd.DataFrame(columns=['Name', 'Formative Score', 'EOT Score', 'Total Score', 'Grade'])
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
                        new_df.at[i, 'Name'] = name
                        new_df.at[i, 'Formative Score'] = fs if np.isnan(fs) else round(fs)     # fs for Formative Score 
                        new_df.at[i, 'EOT Score'] = es                                          # es for EOT Score
                        new_df.at[i, 'Total Score'] = to                                        # to for Total Score
                        new_df.at[i, 'Grade'] = grade                                           #grade for Grade
                    
                    new_dfs[sheet_name] = new_df
                
                calculated_averages[subject_name] = new_dfs
            
            except Exception as e:
                raise e
    
    return calculated_averages


def combine_subjects_to_report_format(processed_subject_data, report_type='STUDENT_REPORT'):
    data_frame_list = {
        'Senior One': pd.DataFrame(),
        'Senior Two': pd.DataFrame(),
        'Senior Three': pd.DataFrame(),
        'Senior Four': pd.DataFrame()
    }
    
    report_types = ['STUDENT_REPORT', 'MARKS_SUMMARY']
    
    if report_type not in report_types:
        raise ValueError(f'The report type {report_types} is not a valid report type')
    
    
    for subject_name,  subject_dfs in processed_subject_data.items():
        
        for sheet_name, df in subject_dfs.items():
            cols = df.columns
            new_columns = {col: f'{subject_name[:3]} {col}' for col in df.columns[1:-1]}
                
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
    
    for class_name, df in data_frame_list.items():
        data_frame_list[class_name]['Total Points'] = df.apply(
            lambda row: data_frame_list(row['Subject 1 Grade'], row['Subject 2 Grade'], row['Subject 3 Grade'], row['Subsidiary Grade'], row['GP Grade']), axis=1
        )
    
    if report_type == 'STUDENT_REPORT':
        for class_name, df in data_frame_list.items():
            columns = df.columns
            cols = {}
            for i in columns:
                cols.update({i: abbreviate_column_name(i)})
            
            df.rename(columns=cols, inplace=True)
    return data_frame_list


def parse_ordinary_level_marksheets(folder_path:str, report_type= 'STUDENT_REPORT'):
    try:
        parsed_subjects = calculate_end_of_term(folder_path=folder_path)
        combined_subjects = combine_subjects_to_report_format(processed_subject_data=parsed_subjects, report_type=report_type)
        return combined_subjects
    except Exception as e:
        raise e


if __name__ == '__main__':
    parsed_subjects = calculate_end_of_term('./test_subjects/Marks Sheet O level')
    combined_subjects = combine_subjects_to_report_format(processed_subject_data=parsed_subjects)
    
    with pd.ExcelWriter('combined_data_frames_o_level.xlsx') as writer:
        for class_name, df in combined_subjects.items():
            df.to_excel(writer, sheet_name=class_name, index=False)
    
    