import os
import pandas as pd

def make_subject_report_a_level(folder_path: str):
    
    if folder_path is None:
        raise FileNotFoundError('You have not given a valid folder path')
    if not os.path.isdir(folder_path):
        raise FileExistsError(f'The Folder you are trying to open {folder_path} does not exist')
    
    subject_reports = {}
    sheet_name_list = ['Senior Five', 'Senior Six']
    
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            try:
                new_df = pd.DataFrame()
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
                
                metrics_dict = {'Metric': ['Class', 'Students']}

                for class_name, df in dfs_end_of_term.items():
                    current_columns = set(df.columns)
                    current_columns.discard('Name')
                    current_columns.discard('Student ID')
                    current_columns.discard('Comments')

                    for i in current_columns:
                        if 'paper' not in i.lower():
                            current_columns.discard(i)
                        else:
                            df[i] = pd.to_numeric(df[i], errors='coerce')
                    
                    class_metrics = [class_name, len(df)]
                    
                    for col in current_columns:
                        class_metrics.extend([
                            df[col].count(), df[col].mean(), df[col].max(), df[col].min()
                        ])
                    metrics_dict[class_name] = class_metrics
                
                for col in current_columns:
                    metrics_dict['Metric'].extend([f'{col} Entered', f'{col} Average', f'{col} Best', f'{col} Worst'])
                
                new_df = pd.DataFrame(metrics_dict)
                new_df.columns = ['Class'] + list(new_df.iloc[0, 1:])
                new_df = new_df.drop(index=0)
                subject_reports[subject_name] = new_df
                
            except Exception as e:
                raise e

    return subject_reports



if __name__ == '__main__':
    data = make_subject_report_a_level('test_subjects/Marks Sheet Term II 2024/A level Marks Sheet Term II 2024')
    for i  in data:
        print(i)
        print(data[i])