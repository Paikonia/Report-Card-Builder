

import os
import pandas as pd

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
                
                metrics_dict = {
                    'Metric': ['Students', 'A01 Entered', 'A01 Average', 'A01 Best', 'A01 Worst',
                               'A02 Entered', 'A02 Average', 'A02 Best', 'A02 Worst',
                               'A03 Entered', 'A03 Average', 'A03 Best', 'A03 Worst',
                               'EOT Entered', 'EOT Average', 'EOT Best', 'EOT Worst']
                }

                for class_name, df in dfs_end_of_term.items():
                    if all(col in df.columns for col in expected_columns):
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
    data = make_subject_report_o_level('test_subjects/Marks Sheet O level')
    for i in data:
        print(i)
        print(data[i])