import pandas as pd
import numpy as np
import os
from functions_advanced_level import paper_grading, subject_grading_two_papers, subject_grading_three_papers, calculate_ICT_total, calc_total_points, abbreviate_column_name

def subject_average_calculator(file_path:str, sheet_name:str, subject = '')-> tuple[int, pd.DataFrame]:
    try:
        if file_path == None:
            raise FileExistsError('Path must be a valid string')
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f'The file whose path you have entered {file_path} does not exist')
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        column_set = set(df.columns)
        # column_len = len(column_set)
        # if(not column_len == 5 and not column_len == 6):
        #     raise ValueError('The columns are not of the right number.')
        column_set.discard('Comments')
        column_set.discard('Name')
        column_set.discard('Student ID')
        grade_list = []
        if subject != 'ICT':
            for i in column_set:
                g = f"{i} Grade"
                df[g] = df[i].apply(paper_grading)
                grade_list.append(g)
            if len(grade_list) == 3:
                df['Subject Grade'] = df.apply(lambda row: subject_grading_three_papers(row[grade_list[0]], row[grade_list[1]], row[grade_list[2]]), axis=1)
            if len(grade_list) == 2:
                df['Subject Grade'] = df.apply(lambda row: subject_grading_two_papers(row[grade_list[0]], row[grade_list[1]]), axis=1)
            if len(grade_list)== 1:
                df['Subject Grade'] = df[grade_list[0]]
        else:
            column_set = list(column_set)
            
            if len(column_set) ==1: 
                df['Total Marks'] = df.apply(lambda row: calculate_ICT_total(row[column_set[0]], np.nan, np.nan), axis=1)
            if len(column_set) == 2:
                df['Total Marks'] = df.apply(lambda row: calculate_ICT_total(row[column_set[0]], row[column_set[1]], np.nan), axis=1)
            if len(column_set) == 3:
                df['Total Marks'] = df.apply(lambda row: calculate_ICT_total(row[column_set[0]], row[column_set[1]], row[column_set[2]]), axis=1)
            
            df['Subject Grade'] = df['Total Marks'].apply(paper_grading)
                
        return len(df), df
    except KeyError as e:
        raise KeyError(f'A mandatory co')
    except ValueError as e:
        raise ValueError(f'An error "{e}" was raise')
    

def iterate_through_subjects(subject_dir_path):
    if subject_dir_path ==None:
        raise ValueError('You have not passed a directory where the marks are stored')
    
    if not os.path.isdir(subject_dir_path):
        raise ValueError(f'The path you have passed {subject_dir_path} is not a directory')
    parsed_subjects = {}
    for file_name in os.listdir(subject_dir_path):
        if file_name.endswith('.xlsx'):
            subject_name = os.path.splitext(file_name)[0]
            parts = subject_name.split(' ')
            
            subject_name = parts[len(parts)-1]
            
            # if("A'" not in parts):
            #     raise NameError(f"The name of the file {file_name} is incorrect, it must include A' level to differentiate it from O level class")
            
            (student_number, subject_data_frame_senior_five) = subject_average_calculator(os.path.join(subject_dir_path, file_name), sheet_name='Senior Five', subject=subject_name)
            (student_number, subject_data_frame_senior_six) = subject_average_calculator(os.path.join(subject_dir_path, file_name), sheet_name='Senior Six', subject=subject_name)
            
            sub_dict = {
                'Senior Five': subject_data_frame_senior_five,
                'Senior Six': subject_data_frame_senior_six
            }
            
            
            parsed_subjects[subject_name] = sub_dict
        
    return parsed_subjects    


def combine_subjects_to_report_format(parsed_subjects, report_type='STUDENT_REPORT'):
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
    
    report_types = ['STUDENT_REPORT', 'MARKS_SUMMARY']
    
    if report_type not in report_types:
        raise ValueError(f'The report type {report_types} is not a valid report type')
    
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

    combined_data_frames = {
        'Senior Five': pd.DataFrame(columns=columns),
        'Senior Six': pd.DataFrame(columns=columns)
    }
    
    for subject, classes in parsed_subjects.items():
        
        for class_name, class_df in classes.items():
            for i, row in class_df.iterrows():
                if pd.notna(row['Subject Grade']):
                    student_index = combined_data_frames[class_name][combined_data_frames[class_name]['Name'] == row['Name']].index
                    if subject == 'GP':
                        if student_index.empty:
                            new_row = {
                                'Student ID': row['Student ID'],
                                'Name': row['Name'],
                            }
                            new_row.update({
                                'GP Marks': row['Paper 1'],
                                'GP Grade': row['Paper 1 Grade'],
                                'GP Comment': row['Comments']
                            })
                            combined_data_frames[class_name] = pd.concat([combined_data_frames[class_name], pd.DataFrame([new_row])], ignore_index=True)      
                        else:
                            combined_data_frames[class_name].loc[student_index, f'GP Marks'] = row['Paper 1']
                            combined_data_frames[class_name].loc[student_index, f'GP Comment'] = row['Comments']
                            combined_data_frames[class_name].loc[student_index, f'GP Grade'] = row['Paper 1 Grade']
                        continue
                                
                            
                    if subject == 'SUBMATH':
                        if student_index.empty:
                            new_row = {
                                'Student ID': row['Student ID'],
                                'Name': row['Name'],
                            }
                            new_row.update({
                                'Subsidiary Marks': row['Paper 1'],
                                'Subsidiary Grade': row['Paper 1 Grade'],
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
                    
                    if subject == 'ICT':
                        if student_index.empty:
                            new_row = {
                                'Student ID': row['Student ID'],
                                'Name': row['Name'],
                            }
                            new_row.update({
                                'Subsidiary Marks': row['Total Marks'],
                                'Subsidiary Grade': row['Subject Grade'],
                                'Subsidiary Subject': 'ICT',
                                'Subsidiary Comment': row['Comments']
                            })
                            combined_data_frames[class_name] = pd.concat([combined_data_frames[class_name], pd.DataFrame([new_row])], ignore_index=True)      
                        else:
                            combined_data_frames[class_name].loc[student_index, 'Subsidiary Marks'] = row['Total Marks']
                            combined_data_frames[class_name].loc[student_index, 'Subsidiary Comment'] = row['Comments']
                            combined_data_frames[class_name].loc[student_index, 'Subsidiary Grade'] = row['Subject Grade']
                            combined_data_frames[class_name].loc[student_index, 'Subsidiary Subject'] = 'ICT'
                        continue
                        
                    if student_index.empty:
                        new_row = {
                            'Student ID': row['Student ID'],
                            'Name': row['Name'],
                        }
                        for subject_num in range(1, 4):
                            new_row.update({
                                f'Subject {subject_num}': subject,
                                f'Subject {subject_num} Comment': row['Comments'],
                                f'Subject {subject_num} First Paper': class_df.columns[3],
                                f'Subject {subject_num} First Paper Marks': row[class_df.columns[3]],
                                f'Subject {subject_num} First Paper Grade': row[class_df.columns[3] +' Grade'],
                                f'Subject {subject_num} Second Paper': class_df.columns[4],
                                f'Subject {subject_num} Second Paper Marks': row[class_df.columns[4]],
                                f'Subject {subject_num} Second Paper Grade': row[class_df.columns[4] +' Grade'],
                                f'Subject {subject_num} Third Paper': class_df.columns[5] if len(class_df.columns) > 8 else None,
                                f'Subject {subject_num} Third Paper Marks': row[class_df.columns[5]] if len(class_df.columns) > 8 else None,
                                f'Subject {subject_num} Third Paper Grade': row[class_df.columns[5] +' Grade'] if len(class_df.columns) > 8 else None,
                                f'Subject {subject_num} Grade': row['Subject Grade']
                            })
                            break
                        combined_data_frames[class_name] = pd.concat([combined_data_frames[class_name], pd.DataFrame([new_row])], ignore_index=True)
                    else:
                        for subject_num in range(1, 4):
                            if pd.isna(combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Grade']).all():
                                combined_data_frames[class_name].loc[student_index, f'Subject {subject_num}'] = subject
                                combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Comment'] = row['Comments']
                                combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} First Paper'] = class_df.columns[3]
                                combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} First Paper Marks'] = row[class_df.columns[3]]
                                combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} First Paper Grade'] = row[class_df.columns[3] +' Grade']
                                combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Second Paper'] = class_df.columns[4]
                                combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Second Paper Marks'] = row[class_df.columns[4]]
                                combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Second Paper Grade'] = row[class_df.columns[4] +' Grade']
                                if len(class_df.columns) > 8:
                                    combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Third Paper'] = class_df.columns[5]
                                    combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Third Paper Marks'] = row[class_df.columns[5]]
                                    combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Third Paper Grade'] = row[class_df.columns[5] +' Grade']
                                combined_data_frames[class_name].loc[student_index, f'Subject {subject_num} Grade'] = row['Subject Grade']
                                break
    for class_name, df in combined_data_frames.items():
        combined_data_frames[class_name]['Total Points'] = df.apply(
            lambda row: calc_total_points(row['Subject 1 Grade'], row['Subject 2 Grade'], row['Subject 3 Grade'], row['Subsidiary Grade'], row['GP Grade']), axis=1
        )
    
    if report_type == 'STUDENT_REPORT':
        for class_name, df in combined_data_frames.items():
            columns = df.columns
            cols = {}
            for i in columns:
                cols.update({i: abbreviate_column_name(i)})
            
            df.rename(columns=cols, inplace=True)
    
    return combined_data_frames

def parse_advanced_level_marksheet(folder_path:str, report_type='STUDENT_REPORT'):
    try:
        parsed_subjects = iterate_through_subjects(folder_path)
        combined_subjects = combine_subjects_to_report_format(parsed_subjects=parsed_subjects, report_type=report_type)
        return combined_subjects
    except Exception as e:
        raise e


if __name__ == '__main__':
    parsed_subjects = iterate_through_subjects('./test_subjects/Marks Sheet A level')
    combined_subjects = combine_subjects_to_report_format(parsed_subjects=parsed_subjects)
    
    # with pd.ExcelWriter('combined_data_frames.xlsx') as writer:
    #     for class_name, df in combined_subjects.items():
    #         df.to_excel(writer, sheet_name=class_name, index=False)
    
    