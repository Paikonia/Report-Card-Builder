import os
import pandas as pd
import numpy as np

def convert_to_percentage(marks, total):
    if np.isnan(marks):
        return np.nan
    if np.isnan(total):
        raise ValueError('Marks entered without a total score to work it out of.')
    
    return (marks / total) * 100 if total and marks else np.nan

def make_subject_report_o_level(folder_path: str):
    if folder_path is None:
        raise FileNotFoundError('You have not given a valid folder path')
    if not os.path.isdir(folder_path):
        raise FileExistsError(f'The Folder you are trying to open {folder_path} does not exist')
    
    subject_reports = {}
    sheet_name_list = ['Senior One', 'Senior Two', 'Senior Three', 'Senior Four']
    expected_columns = ['Student ID', 'Name', 'A01', 'A02', 'A03', 'EOT', 'Comment']
    columns = ['Metric', 'Senior One', 'Senior Two', 'Senior Three', 'Senior Four']
    
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            try:
                new_df = pd.DataFrame(columns=columns)
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
                
                metrics_dict = {
                    'Metric': ['Students', 'A01 Entered', 'A01 Average', 'A01 Best', 'A01 Worst',
                               'A02 Entered', 'A02 Average', 'A02 Best', 'A02 Worst',
                               'A03 Entered', 'A03 Average', 'A03 Best', 'A03 Worst',
                               'EOT Entered', 'EOT Average', 'EOT Best', 'EOT Worst']
                }

                for class_name, df in dfs_end_of_term.items():
                    df['A01'] = pd.to_numeric(df['A01'], errors='coerce')
                    df['A02'] = pd.to_numeric(df['A02'], errors='coerce')
                    df['A03'] = pd.to_numeric(df['A03'], errors='coerce')
                    df['EOT'] = pd.to_numeric(df['EOT'], errors='coerce')
                    if all(col in df.columns for col in expected_columns):
                        total_marks_row = df[df['Name'] == 'Total Marks']
                        # if total_marks_row.empty:
                        #     raise Exception(f'The marks sheet for {subject_name} - {class_name} is missing the "Total Marks" row. It is a requirement for correct calculations')
                        total_marks_index = total_marks_row.index[0]
                        total_marks = df.loc[total_marks_index]
                        df = df.drop(total_marks_index)

                        total_a01 = total_marks['A01']
                        total_a02 = total_marks['A02']
                        total_a03 = total_marks['A03']
                        total_eot = total_marks['EOT']

                        df['A01'] = df['A01'].apply(lambda x: convert_to_percentage(x, total_a01))
                        df['A02'] = df['A02'].apply(lambda x: convert_to_percentage(x, total_a02))
                        df['A03'] = df['A03'].apply(lambda x: convert_to_percentage(x, total_a03))
                        df['EOT'] = df['EOT'].apply(lambda x: convert_to_percentage(x, total_eot))
                    
                        metrics_dict[class_name] = [
                            len(df),
                            df['A01'].count(), df['A01'].mean(), df['A01'].max(), df['A01'].min(),
                            df['A02'].count(), df['A02'].mean(), df['A02'].max(), df['A02'].min(),
                            df['A03'].count(), df['A03'].mean(), df['A03'].max(), df['A03'].min(),
                            df['EOT'].count(), df['EOT'].mean(), df['EOT'].max(), df['EOT'].min()
                        ]
                        
                new_df = pd.DataFrame(metrics_dict)
                subject_reports[subject_name] = new_df
                
            except Exception as e:
                raise e

    return subject_reports

if __name__ == '__main__':
    data = make_subject_report_o_level('test_subjects/Marks Sheet Term II 2024/O level Marks Sheet Term II 2024')
    for i in data:
        print(i)
        print(data[i])
