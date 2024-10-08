import pandas as pd
import numpy as np
import os
from functions_advanced_level import paper_grading, subject_grading_two_papers, subject_grading_three_papers, subject_grading_four_papers, calculate_ICT_total, calc_total_points, abbreviate_column_name

def apply_grading(df, column_set, grade_list, apd=''):
    
    for i in column_set:
        g = f"{i} Grade"
        df[g] = df[i].apply(paper_grading)
        grade_list.append(g)
    if len(grade_list) == 4:
        df[f'{apd}Subject Grade'] = df.apply(lambda row: subject_grading_four_papers(row[grade_list[0]], row[grade_list[1]], row[grade_list[2]], row[grade_list[3]]), axis=1)
    if len(grade_list) == 3:
        df[f'{apd}Subject Grade'] = df.apply(lambda row: subject_grading_three_papers(row[grade_list[0]], row[grade_list[1]], row[grade_list[2]]), axis=1)
    if len(grade_list) == 2:
        df[f'{apd}Subject Grade'] = df.apply(lambda row: subject_grading_two_papers(row[grade_list[0]], row[grade_list[1]]), axis=1)
    if len(grade_list)== 1:
        df[f'{apd}Subject Grade'] = df[grade_list[0]]
    return len(df), df

def apply_grading_ict(df, column_set, apd=''):
    column_set = list(column_set)
    if len(column_set) ==1: 
        df[f'{apd}Total Marks'] = df.apply(lambda row: calculate_ICT_total(row[column_set[0]], np.nan, np.nan), axis=1)
    if len(column_set) == 2:
        df[f'{apd}Total Marks'] = df.apply(lambda row: calculate_ICT_total(row[column_set[0]], row[column_set[1]], np.nan), axis=1)
    if len(column_set) == 3:
        df[f'{apd}Total Marks'] = df.apply(lambda row: calculate_ICT_total(row[column_set[0]], row[column_set[1]], row[column_set[2]]), axis=1)
    
    df[f'{apd}Subject Grade'] = df['Total Marks'].apply(paper_grading)

def calculating_subject_grade(file_path:str, sheet_name:str, subject = '', mid=False)-> tuple[int, pd.DataFrame]:
    try:
        if file_path == None:
            raise FileExistsError('Path must be a valid string')
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f'The file whose path you have entered {file_path} does not exist')
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df['Name'] = df['Name'].str.strip().str.upper()
        s = set(df.columns)
        column_set = set()
        
        column_len = len(s)
        if( not column_len > 5 and not column_len < 11 ):
            raise ValueError(f'The columns are not of the right number.\nFile: {file_path}')
        for column_name in s:
            if 'Unnamed: ' in column_name:
                raise ValueError(f'The file subject {subject} in class {sheet_name} in the file {file_path} has columns that are not named.')
            if 'paper' in column_name.lower():
                column_set.add(column_name)
        
        for column in column_set:
            df[column] = pd.to_numeric(df[column], errors='coerce')
        
        
        mid_term_columns = { col for col in column_set if 'Mid' in col}
        end_of_term_columns = { col for col in column_set if 'Mid' not in col}
        grade_list = []
        grade_list_mid_term =[]
        
        if subject != 'ICT':
            apply_grading(df, end_of_term_columns, grade_list)
            apply_grading(df, mid_term_columns, grade_list_mid_term, 'Mid ')
        else:
            apply_grading_ict(df, end_of_term_columns)
            apply_grading_ict(df, end_of_term_columns, apd='Mid ')
       
        return len(df), df
    except KeyError as e:
        raise KeyError(f'A mandatory co')
    except ValueError as e:
        raise ValueError(f'An error "{e}" was raise')
    

def grading_advanced_level(subject_dir_path):
    if subject_dir_path ==None:
        raise ValueError('You have not passed a directory where the marks are stored')
    
    if not os.path.isdir(subject_dir_path):
        raise ValueError(f'The path you have passed {subject_dir_path} is not a directory')
    parsed_subjects = {}
    for file_name in os.listdir(subject_dir_path):
        if file_name.endswith('.xlsx'):
            subject_name = os.path.splitext(file_name)[0]
            subject_name = subject_name.split(' - ')[1]
            
            # if("A'" not in parts):
            #     raise NameError(f"The name of the file {file_name} is incorrect, it must include A' level to differentiate it from O level class")
            
            (s5_n, subject_data_frame_senior_five) = calculating_subject_grade(os.path.join(subject_dir_path, file_name), sheet_name='Senior Five', subject=subject_name)
            # (s6_n, subject_data_frame_senior_six) = calculating_subject_grade(os.path.join(subject_dir_path, file_name), sheet_name='Senior Six', subject=subject_name)
            
            sub_dict = {
                'Senior Five': subject_data_frame_senior_five,
                # 'Senior Six': subject_data_frame_senior_six
            }
            
            parsed_subjects[subject_name] = sub_dict
        
    return parsed_subjects    


def  merge_subjects(parsed_subjects, combined_data_frames, apd=''):
    subject_count = 0
    for subject, classes in parsed_subjects.items():
        subject_count=subject_count+1
        for class_name, class_df in classes.items():
              
            old_cols = set(class_df.columns)
            cols = []
            for i in old_cols:
                if f'Paper' in i and 'Grade' not in i:
                    if apd.strip().lower() == '' and 'Mid' not in i:
                        cols.append(i)
                    elif apd.strip().lower() == 'mid' and 'Mid' in i:
                        cols.append(i)
            cols.sort()
            class_df = class_df.dropna(subset=[f'{apd}Subject Grade'])
            try:
                for i, row in class_df.iterrows():
                    
                    if pd.notna(row[f'{apd}Subject Grade']):
                        student_index = combined_data_frames[class_name][combined_data_frames[class_name]['Name'] == row['Name']].index
                       
                        if subject.strip().upper() == 'GP' or subject.strip().upper() == 'GENERAL PAPER':
                            if student_index.empty:
                                new_row = {
                                    'Student ID': row['Student ID'],
                                    'Name': row['Name'],
                                }
                                new_row.update({
                                    'GP Marks': row[f'{apd}Paper 1'],
                                    'GP Grade': row[f'{apd}Paper 1 Grade'],
                                    'GP Comment': row['Comments']
                                })
                                combined_data_frames[class_name] = pd.concat([combined_data_frames[class_name], pd.DataFrame([new_row])], ignore_index=True)      
                            else:
                                combined_data_frames[class_name].loc[student_index, f'GP Marks'] = row[f'Paper 1']
                                combined_data_frames[class_name].loc[student_index, f'GP Comment'] = row[f'Comments']
                                combined_data_frames[class_name].loc[student_index, f'GP Grade'] = row[f'Paper 1 Grade']
                            continue
                                    
                                
                        if subject.strip().upper() == 'SUBMATH':
                            if student_index.empty:
                                new_row = {
                                    'Student ID': row['Student ID'],
                                    'Name': row['Name'],
                                }
                                new_row.update({
                                    'Subsidiary Marks': row[f'{apd}Paper 1'],
                                    'Subsidiary Grade': row[f'{apd}Paper 1 Grade'],
                                    'Subsidiary Subject': 'SUBMATH',
                                    'Subsidiary Comment': row['Comments']
                                })
                                combined_data_frames[class_name] = pd.concat([combined_data_frames[class_name], pd.DataFrame([new_row])], ignore_index=True)      
                            else:
                                combined_data_frames[class_name].loc[student_index, 'Subsidiary Marks'] = row['Paper 1']
                                combined_data_frames[class_name].loc[student_index, 'Subsidiary Comment'] = row['Comments']
                                combined_data_frames[class_name].loc[student_index, 'Subsidiary Grade'] = row['Paper 1 Grade']
                                combined_data_frames[class_name].loc[student_index, 'Subsidiary Subject'] = 'SUBMATH'
                            continue
                        
                        if subject.strip().upper() == 'ICT':
                            
                            if student_index.empty:
                                new_row = {
                                    'Student ID': row['Student ID'],
                                    'Name': row['Name'],
                                }
                                new_row.update({
                                    'Subsidiary Marks': row[f'{apd}Total Marks'],
                                    'Subsidiary Grade': row[f'Subject Grade'],
                                    'Subsidiary Subject': 'ICT',
                                    'Subsidiary Comment': row['Comments']
                                })
                                combined_data_frames[class_name] = pd.concat([combined_data_frames[class_name], pd.DataFrame([new_row])], ignore_index=True)      
                            else:
                                combined_data_frames[class_name].loc[student_index, 'Subsidiary Marks'] = row[f'{apd}Total Marks']
                                combined_data_frames[class_name].loc[student_index, 'Subsidiary Comment'] = row['Comments']
                                combined_data_frames[class_name].loc[student_index, 'Subsidiary Grade'] = row[f'Subject Grade']
                                combined_data_frames[class_name].loc[student_index, 'Subsidiary Subject'] = 'ICT'
                            continue
                            
                        if student_index.empty:
                            new_row = {
                                'Student ID': row['Student ID'],
                                'Name': row['Name'],
                            }
                            
                            new_row.update({
                                f'Subject 1': subject,
                                f'Subject 1 Comment': row['Comments'],
                                f'Subject 1 First Paper': cols[0],
                                f'Subject 1 First Paper Marks': row[f'{cols[0]}'],
                                f'Subject 1 First Paper Grade': row[f'{cols[0]} Grade'],
                            })
                            
                            if len(cols) > 1:
                                
                                new_row.update({
                                    f'Subject 1 Second Paper': cols[1],
                                    f'Subject 1 Second Paper Marks': row[f'{cols[1]}'],
                                    f'Subject 1 Second Paper Grade': row[f'{cols[1]} Grade']
                                })
                            if  len(cols) > 2:
                                new_row.update({
                                    f'Subject 1 Third Paper': cols[2],
                                    f'Subject 1 Third Paper Marks': row[f'{cols[2]}'],
                                    f'Subject 1 Third Paper Grade': row[f'{cols[2]} Grade']
                                })
                            
                            if  len(cols) > 3:
                                new_row.update({
                                    f'Subject 1 Fourth Paper': cols[3],
                                    f'Subject 1 Fourth Paper Marks': row[f'{cols[3]}'],
                                    f'Subject 1 Fourth Paper Grade': row[f'{cols[3]} Grade']
                                })
                            
                            new_row.update({f'Subject {subject_num} Grade': row[f'Subject Grade']})
                            
                            s_pd = pd.DataFrame([new_row])
                            
                            combined_data_frames[class_name] = pd.concat([combined_data_frames[class_name], s_pd], ignore_index=True)
                        else:
                            
                            for subject_num in range(1, 4):
                                if pd.isna(combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Grade']).all():
                                    combined_data_frames[class_name].loc[student_index, f'Subject {subject_num}'] = subject
                                    combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Comment'] = row['Comments']
                                    
                                    combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} First Paper'] = cols[0]
                                    combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} First Paper Marks'] = row[f'{cols[0]}']
                                    combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} First Paper Grade'] = row[f'{cols[0]} Grade']
                                    if len(cols) >1:
                                        combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Second Paper'] = cols[1]
                                        combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Second Paper Marks'] = row[f'{cols[1]}']
                                        combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Second Paper Grade'] = row[f'{cols[1]} Grade']
                                    if len(cols) > 2:
                                        combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Third Paper'] = cols[2]
                                        combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Third Paper Marks'] = row[f'{cols[2]}']
                                        combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Third Paper Grade'] = row[f'{cols[2]} Grade']
                                    if len(cols) > 3:
                                        combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Fourth Paper'] = cols[3]
                                        combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Fourth Paper Marks'] = row[f'{cols[3]}']
                                        combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Fourth Paper Grade'] = row[f'{cols[0]} Grade']
                                        
                                    combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Grade'] = row['Subject Grade']
                                    break
                
            except Exception as e:
                pass                 
        
    for class_name, df in combined_data_frames.items():
        combined_data_frames[class_name]['Total Points'] = df.apply(
            lambda row: calc_total_points(row['Subject 1 Grade'], row['Subject 2 Grade'], row['Subject 3 Grade'], row['Subsidiary Grade'], row['GP Grade']), axis=1
        )



def combine_subjects_to_report_format(parsed_subjects, output_folder='', report_type='MARKS_SUMMARY_REPORT'):
    """
    Combine subjects data for students in senior five and senior six classes.

    Args:
        parsed_subjects (dict): A dictionary of dictionary. A subject is the key to another
        dictionary that contains two classes: Senior Five and Senior Six.
        Each class contains the dataframes with these columns: ['Student ID', 'Name', 'Comment', ..., 'Subject Grade'].
        The ... represents two or three papers depending on the subject. The column name is the paper name and can be paper 1 through 6, 
        and each paper has a grade, e.g., 'Paper 1 Grade' and 'Paper 2 Grade'. Each subject also has an overall grade.
        
    Returns:
        dict: A dictionary containing the combined DataFrames for senior five and senior six.
    """
    
    report_types = ['END_OF_TERM_REPORT', 'MARKS_SUMMARY_REPORT', 'MID_TERM_REPORT']
    
    if report_type not in report_types:
        raise ValueError(f'The report type {report_type} is not a valid report type')
    
    columns = [
        'Student ID', 
        'Name',
        'Subject 1', 'Subject 1 Comment', 'Subject 1 First Paper', 'Subject 1 First Paper Marks', 'Subject 1 First Paper Grade',
        'Subject 1 Second Paper', 'Subject 1 Second Paper Marks', 'Subject 1 Second Paper Grade', 'Subject 1 Third Paper',
        'Subject 1 Third Paper Marks', 'Subject 1 Third Paper Grade', 'Subject 1 Grade',
        'Subject 2', 'Subject 2 Comment', 'Subject 2 First Paper', 'Subject 2 First Paper Marks', 'Subject 2 First Paper Grade',
        'Subject 2 Second Paper', 'Subject 2 Second Paper Marks', 'Subject 2 Second Paper Grade', 'Subject 2 Third Paper',
        'Subject 2 Third Paper Marks', 'Subject 2 Third Paper Grade', 'Subject 2 Grade',
        'Subject 3', 'Subject 3 Comment', 'Subject 3 First Paper', 'Subject 3 First Paper Marks', 'Subject 3 First Paper Grade',
        'Subject 3 Second Paper', 'Subject 3 Second Paper Marks', 'Subject 3 Second Paper Grade', 'Subject 3 Third Paper',
        'Subject 3 Third Paper Marks', 'Subject 3 Third Paper Grade', 'Subject 3 Grade',
        'GP Marks', 'GP Grade', 'GP Comment', 'Subsidiary Subject', 'Subsidiary Comment', 'Subsidiary Marks', 'Subsidiary Grade', 'Total Points'
    ]

    combined_data_frames_end_of_term = {
        'Senior Five': pd.DataFrame(columns=columns),
        'Senior Six': pd.DataFrame(columns=columns)
    }
    
    combined_data_frames_mid_term = {
        'Senior Five': pd.DataFrame(columns=columns),
        'Senior Six': pd.DataFrame(columns=columns)
    }
    
    caught_errors = {
        
    }
    
    if report_type =='END_OF_TERM_REPORT' or report_type == 'MARKS_SUMMARY_REPORT':
        try:
            
            merge_subjects(parsed_subjects=parsed_subjects, combined_data_frames=combined_data_frames_end_of_term)
        except Exception as e:
            
            with pd.ExcelWriter(os.path.join(output_folder, 'Errored_end_of_term_output.xlsx')) as writer:
                for class_name, dataframe in combined_data_frames_end_of_term.items():
                    dataframe.to_excel(writer, sheet_name=class_name, index=False)
            
            combined_data_frames_end_of_term = None
            caught_errors['end of term process errors'] = str(e)
            
    if report_type =='MID_TERM_REPORT' or report_type == 'MARKS_SUMMARY_REPORT':
        try: 
            merge_subjects(parsed_subjects=parsed_subjects, combined_data_frames=combined_data_frames_mid_term, apd='Mid ')
        except Exception as e:
            
            with pd.ExcelWriter(os.path.join(output_folder, 'Errored_mid_term_output.xlsx')) as writer:
                for class_name, dataframe in combined_data_frames_mid_term.items():
                    dataframe.to_excel(writer, sheet_name=class_name, index=False)
            
            combined_data_frames_mid_term = None
            caught_errors['mid term process errors'] = str(e)
        
    if report_type == 'END_OF_TERM_REPORT':
        if combined_data_frames_end_of_term == None:
            raise Exception(caught_errors)
        
        for class_name, df in combined_data_frames_end_of_term.items():
            columns = df.columns
            cols = {}
            for i in columns:
                cols.update({i: abbreviate_column_name(i)})
            df.rename(columns=cols, inplace=True)
        return (combined_data_frames_end_of_term, None)
    if report_type == 'MID_TERM_REPORT':
        if combined_data_frames_mid_term == None:
            raise Exception(caught_errors)
        
        for class_name, df in combined_data_frames_mid_term.items():
            columns = df.columns
            cols = {}
            for i in columns:
                cols.update({i: abbreviate_column_name(i)})
            df.rename(columns=cols, inplace=True)
        return (None, combined_data_frames_mid_term)
    
    
    if combined_data_frames_end_of_term == None:
        raise Exception(caught_errors)
        
    return (combined_data_frames_end_of_term, combined_data_frames_mid_term)

def parse_advanced_level_marksheet(folder_path:str, report_type='END_OF_TERM_REPORT'):
    try:
        parsed_subjects = grading_advanced_level(folder_path)
        combined_subjects = combine_subjects_to_report_format(parsed_subjects=parsed_subjects, report_type=report_type)
        return combined_subjects
    except Exception as e:
        raise e


if __name__ == '__main__':
    parsed_subjects = grading_advanced_level('./Marks Sheet Term II 2024/A level Marks Sheet Term II 2024')
    (end_of_term, mid_term) = combine_subjects_to_report_format(parsed_subjects=parsed_subjects,report_type='END_OF_TERM_REPORT')
    # report_type='MID_TERM_REPORT'
    # with pd.ExcelWriter('combined_data_frames_a_level.xlsx') as writer:
    #     for class_name, df in end_of_term.items():
    #         df.to_excel(writer, sheet_name=class_name, index=False)
    
    with pd.ExcelWriter('combined_data_frames_a_level_end_of_term.xlsx') as writer:
        for class_name, df in end_of_term.items():
            df.to_excel(writer, sheet_name=class_name, index=False)
    