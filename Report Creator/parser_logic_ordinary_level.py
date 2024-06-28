import pandas as pd

from functions_ordinary_level import calc_grading, abbreviate_column_name, grading_ordinary_level

def combine_mid_term(processed_subject_data):
    data_frame_list = {
        'Senior One': pd.DataFrame(),
        'Senior Two': pd.DataFrame(),
        'Senior Three': pd.DataFrame(),
        'Senior Four': pd.DataFrame()
    }
    
    for subject_name,  subject_dfs in processed_subject_data.items():
        for sheet_name, df in subject_dfs.items():
            cols = list(df.columns)
            cols.remove('EOT Score')
            cols.remove('Formative Score')
            cols.remove('Total Score')
            cols.remove('Grade')
            cols.remove('Name')
            cols.remove('Student ID')
            new_columns = {col: f'{subject_name[:3]} {col}' for col in cols}
            
            for index, row in df.iterrows():
                if 'Student ID' not in data_frame_list[sheet_name].columns:
                    data_frame_list[sheet_name]['Student ID'] = []
                if 'Name' not in data_frame_list[sheet_name].columns:
                    data_frame_list[sheet_name]['Name'] = []

                student_name = str(row['Name']).upper()
                if student_name not in data_frame_list[sheet_name]['Name'].astype(str).str.upper().tolist():
                    new_row = {'Name': student_name}
                    for col_name in new_columns.values():
                        new_row[col_name] = None
                    
                    data_frame_list[sheet_name] = pd.concat([data_frame_list[sheet_name], pd.DataFrame([new_row])], ignore_index=True)

                for col_name, col_value in new_columns.items():
                    name_index = data_frame_list[sheet_name][data_frame_list[sheet_name]['Name'].astype(str).str.upper() == student_name].index
                    data_frame_list[sheet_name].loc[name_index, col_value] = row[col_name]
    
    for class_name, df in data_frame_list.items():
        A01 = [col for col in df.columns if 'A01' in col]
        A02 = [col for col in df.columns if 'A02' in col]
        A03 = [col for col in df.columns if 'A03' in col]
        average = [col for col in df.columns if 'Average Score' in col]
        df['A01 Average'] = df[A01].mean(axis=1)
        df['A02 Average'] = df[A02].mean(axis=1)
        df['A03 Average'] = df[A03].mean(axis=1)
        df['Mid Term Average'] = df[average].mean(axis=1)
        df['Mid Term Grade'] = df['Mid Term Average'].apply(calc_grading)
        df['Mid Term Position'] = df['Mid Term Average'].rank(method='min', ascending=False)
        df['Mid Term Position'] = df['Mid Term Position'].fillna(-1)
        df['Mid Term Position'] = df['Mid Term Position'].astype(int)    
    
    return data_frame_list


def combine_end_of_term(processed_subject_data):
    data_frame_list = {
        'Senior One': pd.DataFrame(),
        'Senior Two': pd.DataFrame(),
        'Senior Three': pd.DataFrame(),
        'Senior Four': pd.DataFrame()
    }
    for subject_name,  subject_dfs in processed_subject_data.items():
        print(subject_name)
        for sheet_name, df in subject_dfs.items():
            cols = list(df.columns)
            
            cols.remove('Average Score')
            cols.remove('A01')
            cols.remove('A02')
            cols.remove('A03')
            cols.remove('Name')
            cols.remove('Average Grade')
            cols.remove('Student ID')
            new_columns = {col: f'{subject_name[:3]} {col}' for col in cols}
            
            for index, row in df.iterrows():
                if 'Student ID' not in data_frame_list[sheet_name].columns:
                    data_frame_list[sheet_name]['Student ID'] = []
                if 'Name' not in data_frame_list[sheet_name].columns:
                    data_frame_list[sheet_name]['Name'] = []

                student_name = str(row['Name']).upper()
                if student_name not in data_frame_list[sheet_name]['Name'].astype(str).str.upper().tolist():
                    new_row = {'Name': student_name}
                    for col_name in new_columns.values():
                        new_row[col_name] = None
                    
                    data_frame_list[sheet_name] = pd.concat([data_frame_list[sheet_name], pd.DataFrame([new_row])], ignore_index=True)

                for col_name, col_value in new_columns.items():
                    name_index = data_frame_list[sheet_name][data_frame_list[sheet_name]['Name'].astype(str).str.upper() == student_name].index
                    data_frame_list[sheet_name].loc[name_index, col_value] = row[col_name]
    
    
    for class_name, df in data_frame_list.items():
        formative_score_cols = [col for col in df.columns if 'Formative Score' in col]
        eot_score_cols = [col for col in df.columns if 'EOT Score' in col]
        total_score_cols = [col for col in df.columns if 'Total Score' in col]
        df['Average Formative Score'] = df[formative_score_cols].mean(axis=1)
        df['Average EOT Score'] = df[eot_score_cols].mean(axis=1)
        df['Average Total Score'] = df[total_score_cols].mean(axis=1)
        df['Average Grade'] = df['Average Total Score'].apply(calc_grading)
        df['Position'] = df['Average Total Score'].rank(method='min', ascending=False)
        df['Position'] = df['Position'].fillna(-1)
        df['Position'] = df['Position'].astype(int)    
    return data_frame_list


def combine_subjects_to_report_format(processed_subject_data, report_type='MARKS_SUMMARY_REPORT'):
    data_frame_list = {
        'Senior One': pd.DataFrame(),
        'Senior Two': pd.DataFrame(),
        'Senior Three': pd.DataFrame(),
        'Senior Four': pd.DataFrame()
    }
    
    
    report_types = ['END_OF_TERM_REPORT', 'MARKS_SUMMARY_REPORT', 'MID_TERM_REPORT']
    
    if report_type not in report_types:
        raise ValueError(f'The report type {report_type} is not a valid report type')
    
    if report_type== 'END_OF_TERM_REPORT':
        combined_data_frames_end_of_term = combine_end_of_term(processed_subject_data=processed_subject_data)
        for class_name, df in combined_data_frames_end_of_term.items():
            columns = df.columns
            cols = {}
            for i in columns:
                cols.update({i: abbreviate_column_name(i)})
            df.rename(columns=cols, inplace=True)
        return (combined_data_frames_end_of_term, None)
    
    if report_type == 'MID_TERM_REPORT':
        combined_data_frames_mid_term =combine_mid_term(processed_subject_data=processed_subject_data)
        for class_name, df in combined_data_frames_mid_term.items():
            columns = df.columns
            cols = {}
            for i in columns:
                cols.update({i: abbreviate_column_name(i)})
            df.rename(columns=cols, inplace=True)
        return (None, combined_data_frames_mid_term)
        
        
    if report_type == 'MARKS_SUMMARY_REPORT':    
        midterm_combine = combine_mid_term(processed_subject_data)
        end_of_term_combine = combine_end_of_term(processed_subject_data)
        return (end_of_term_combine, midterm_combine)
        
        
def parse_ordinary_level_marksheets(folder_path:str, report_type= 'MARKS_SUMMARY_REPORT'):
    try:
        parsed_subjects = grading_ordinary_level(folder_path=folder_path)
        
        combined_subjects = combine_subjects_to_report_format(processed_subject_data=parsed_subjects, report_type=report_type)
        
        return combined_subjects
    except Exception as e:
        raise e


if __name__ == '__main__':
    (eot, mid) = parse_ordinary_level_marksheets('./test_subjects/Marks Sheet Term II 2024/O level Marks Sheet Term II 2024', 'MARKS_SUMMARY_REPORT')
    for i in eot:
        print(eot[i])
    for i in mid:
        print(mid[i])