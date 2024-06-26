import numpy as np
import pandas as pd


def calc_total_points(subject_1_grade, subject_2_grade, subject_3_grade, subsidiary_grade, GP_grade):
    valid_grades = {'A', 'B', 'C', 'D', 'E', 'O', 'F', 'D1', 'D2', 'C3', 'C4', 'C5', 'C6', 'P7', 'P8', 'F9', 'X'}
    
    valid_sub_gp_grades = {'D1', 'D2', 'C3', 'C4', 'C5', 'C6', 'P7', 'P8', 'F9', 'X'}
    
    for grade in [subject_1_grade, subject_2_grade, subject_3_grade]:
        if grade not in valid_grades:
            raise ValueError(f"Unrecognized grading for principal subjects: {grade}")
    
    for grade in [subsidiary_grade, GP_grade]:
        if grade not in valid_sub_gp_grades:
            raise ValueError(f"Unrecognized grading for Subsidiary subjects and General paper: {grade}")
    grades = [subject_1_grade, subject_2_grade, subject_3_grade, subsidiary_grade, GP_grade]
    if 'X' in grades:
        return 'X'
    
    points_dict = {
        'A': 6, 'B': 5, 'C': 4, 'D': 3, 'E': 2, 'O': 1, 'F': 0
    }
    
    passed = {'D1', 'D2', 'C3', 'C4', 'C5', 'C6'}
    
    def get_subject_points(grade):
        if grade in points_dict:
            return points_dict[grade]
        else:
            return 1 if grade in passed else 0
    
    total_points = 0
    for grade in grades:
        total_points += get_subject_points(grade)
    
    return total_points

def subject_grading_two_papers(paper_grade_one, paper_grade_two):
    if (pd.isna(paper_grade_two) or pd.isnull(paper_grade_two)) and (pd.isna(paper_grade_one) or pd.isnull(paper_grade_one)):
        return None
    
    if (pd.isna(paper_grade_one) or (pd.isnull(paper_grade_one)) or pd.isna(paper_grade_two) or pd.isnull(paper_grade_two)):
        return 'X'
    
    """
    Determine the subject grade for subjects with two papers.
    
    Grade               Paper grades
    Principal pass A    Both paper grades must be distinctions, eg (D1, D2), (D2, D1), (D1, D1)
    Principal pass B    Credit 3 both papers or one paper and better eg (C3, D1), (C3, C3)
    Principal pass C    Credit 4 in both papers or better in the other eg (C3, C4), (C4, C4), (C4, D1)
    Principal pass D    Credit 5 in both papers or better in the other eg (C3, C5), (C5, C5), (C5, D2)
    Principal pass E    Credit 6 in both papers or better in the other eg (C6, C5), (C6, C6), (C6, D1) or
                        paper grade involving a subject pass with (P7 or P8) whose aggregate sum does not
                        exceed 12 eg (P7, C5), (P8, D1)
    Subsidiary pass O   Paper grade involving a subject pass whose aggregate sum does not exceed 16
                        eg (C6, P7), (P7, P7), (P8, P8) or
                        paper grade involving a credit or better with and F9 in the other, eg (F9, C4), (F9, D1)
    Fail            F   An F9 in one of the papers and a Pass in or worse in the other eg (F9, P7), (F9, F9), (F9, P8)
    """
    
    distinctions = {'D1', 'D2'}
    credit_grades = {'C3', 'C4', 'C5', 'C6'}
    pass_grades = {'P7', 'P8'}
    fail_grade = 'F9'
    
    if paper_grade_one in distinctions and paper_grade_two in distinctions:
        return 'A'
    
    if (paper_grade_one == 'C3' and paper_grade_two in {'C3'} | distinctions) or \
       (paper_grade_two == 'C3' and paper_grade_one in {'C3'} | distinctions):
        return 'B'
    
    if (paper_grade_one == 'C4' and paper_grade_two in {'C4'} | distinctions | {'C3'}) or \
       (paper_grade_two == 'C4' and paper_grade_one in {'C4'} | distinctions | {'C3'}):
        return 'C'
    
    if (paper_grade_one == 'C5' and paper_grade_two in {'C5'} | distinctions | {'C3', 'C4'}) or \
       (paper_grade_two == 'C5' and paper_grade_one in {'C5'} | distinctions | {'C3', 'C4'}):
        return 'D'
    
    if (paper_grade_one == 'C6' and paper_grade_two in {'C6'} | distinctions | {'C3', 'C4', 'C5'}) or \
       (paper_grade_two == 'C6' and paper_grade_one in {'C6'} | distinctions | {'C3', 'C4', 'C5'}):
        return 'E'
    if (paper_grade_one in pass_grades and paper_grade_two in {'C5', 'C6', 'D1', 'D2'}) or \
       (paper_grade_two in pass_grades and paper_grade_one in {'C5', 'C6', 'D1', 'D2'}):
        if int(paper_grade_one[-1]) + int(paper_grade_two[-1]) <= 12:
            return 'E'
    if (paper_grade_one in pass_grades and paper_grade_two in {'C6', 'C5', 'C4', 'C3', 'D1', 'D2'}) or \
       (paper_grade_two in pass_grades and paper_grade_one in {'C6', 'C5', 'C4', 'C3', 'D1', 'D2'}):
        if int(paper_grade_one[-1]) + int(paper_grade_two[-1]) <= 16:
            return 'O'
    if (paper_grade_one == fail_grade and paper_grade_two in credit_grades | distinctions) or \
       (paper_grade_two == fail_grade and paper_grade_one in credit_grades | distinctions):
        return 'O'
    
    if fail_grade in {paper_grade_one, paper_grade_two}:
        return 'F'
    
    return 'F'


def subject_grading_three_papers(paper_grade_one, paper_grade_two, paper_grade_three):
    if (pd.isna(paper_grade_three) or pd.isnull(paper_grade_three)) and (pd.isna(paper_grade_two) or pd.isnull(paper_grade_two)) and (pd.isna(paper_grade_one) or pd.isnull(paper_grade_one)):
        return None
    
    if (pd.isna(paper_grade_one) or (pd.isnull(paper_grade_one)) or pd.isna(paper_grade_two) or pd.isnull(paper_grade_two)) or (pd.isna(paper_grade_three) or pd.isnull(paper_grade_three)):
        return 'X'

    distinctions = {'D1', 'D2'}
    credit_grades = {'C3', 'C4', 'C5', 'C6'}
    pass_grades = {'P7', 'P8'}
    fail_grade = 'F9'
    
    grades = [paper_grade_one, paper_grade_two, paper_grade_three]
    
    if 'C3' in grades and grades.count('C3') == 1 and all(grade in distinctions for grade in grades if grade != 'C3'):
        return 'A'
    
    if all(grade in distinctions for grade in grades):
        return 'A'
    
    if (('C4' in grades and grades.count('C4') == 1) or ('C3' in grades and grades.count('C3') > 1)) and all(grade in {'C3'} | distinctions for grade in grades if grade != 'C4'):
        return 'B'
    
    if (('C5' in grades and grades.count('C5') == 1 or ('C4' in grades and grades.count('C4') > 1)) and all(grade in {'C4', 'C3'} | distinctions for grade in grades if grade != 'C5')):
        return 'C'
    
    if (('C6' in grades and grades.count('C6') == 1) or ('C5' in grades and grades.count('C5') > 1)) and all(grade in {'C5', 'C4', 'C3'} | distinctions for grade in grades if grade != 'C6'):
        return 'D'
    
    if (((grades.count('P7') == 1 and grades.count('P8') == 0) or (grades.count('P8') == 1 and grades.count('P7') == 0)) or ('C6' in grades and grades.count('C6') > 1)) and all(grade in {'C6', 'C5', 'C4', 'C3'} | distinctions for grade in grades if grade not in {'P7', 'P8'}):
        return 'E'
    
    if ((grades.count(fail_grade) == 1 or len([grade for grade in grades if grade in pass_grades]) > 1) and all(grade in {'P8', 'P7','C6', 'C5', 'C4', 'C3', } | distinctions for grade in grades if grade != fail_grade)):
        return 'O'
    
    if grades.count(fail_grade) == 2 and all(grade in {'C3', 'C4', 'C5', 'C6', 'D1', 'D2'} for grade in grades if grade != fail_grade):
        return 'O'
    if all(grade in {'P8', 'P7'} for grade in grades):
        return 'O'
    
    
    return 'F'



def paper_grading(marks:np.number):
    if pd.isna(marks) or pd.isnull(marks):
        return None
    if marks >= 85:
        return 'D1'
    elif marks >= 80:
        return 'D2'
    elif marks >= 75:
        return 'C3'
    elif marks >= 70:
        return 'C4'
    elif marks >= 65:
        return 'C5'
    elif marks>= 60:
        return 'C6'
    elif marks >= 50:
        return 'P7'
    elif marks >= 40:
        return 'P8'
    else: return  'F9'

def calculate_ICT_total(paper_1, paper_2, paper_3):
    values = [paper_2, paper_1, paper_3]
    non_null_values = [value for value in values if pd.notnull(value) and pd.to_numeric(value, errors='coerce') == value]
    
    if len(non_null_values) == 3:
        return np.nanmean(non_null_values)
    if len(non_null_values) == 2:
        return np.nanmean(non_null_values)
    elif len(non_null_values) == 1:
        return non_null_values[0]   
    else:
        return np.nan

        
def abbreviate_column_name(column_name):
    column_name = column_name.replace('Subject 1', 's1')
    column_name = column_name.replace('Subject 2', 's2')
    column_name = column_name.replace('Subject 3', 's3')
    column_name = column_name.replace('First Paper', 'fp')
    column_name = column_name.replace('Second Paper', 'sp')
    column_name = column_name.replace('Third Paper', 'tp')
    column_name = column_name.replace('Marks', 'm')
    column_name = column_name.replace('Grade', 'g')
    column_name = column_name.replace('Comment', 'c')
    column_name = column_name.replace('GP', 'gp')
    column_name = column_name.replace('Name', 'name')
    column_name = column_name.replace('Subsidiary', 'sub')
    return column_name.replace(' ', '_')


if __name__ == '__main__':
    print('For two subjects')
    print(subject_grading_two_papers('D1', 'D2'))  # Output: A
    print(subject_grading_two_papers('P7', 'F9'))  # Output: B
    print(subject_grading_two_papers('C4', 'C4'))  # Output: C
    print(subject_grading_two_papers('C5', 'D2'))  # Output: D
    print(subject_grading_two_papers('P7', 'C5'))  # Output: E
    print(subject_grading_two_papers('F9', 'C4'))  # Output: O
    print(subject_grading_two_papers('F9', 'P7'))  # Output: F
    
    print('For three subjects')
    print('Printing A')
    print(subject_grading_three_papers('D1', 'D2', 'C3'))  
    print(subject_grading_three_papers('D2', 'C3', 'D1'))  
    print(subject_grading_three_papers('D1', 'D1', 'D2'))  
    print(subject_grading_three_papers('C3', 'D1', 'D1'))  
    print(subject_grading_three_papers('D2', 'D2', 'D2'))  
    print(subject_grading_three_papers('D1', 'C3', 'D2')) 
    print('Printing B')
    print(subject_grading_three_papers('D1', 'D2', 'C4'))  
    print(subject_grading_three_papers('C3', 'C3', 'D4'))  
    print(subject_grading_three_papers('C3', 'D1', 'C4'))  
    print(subject_grading_three_papers('C3', 'D1', 'C3'))  
    print(subject_grading_three_papers('D1', 'D1', 'C4'))  
    print(subject_grading_three_papers('C3', 'C3', 'C3'))  
    print(subject_grading_three_papers('D1', 'C3', 'C4')) 
    print('Printing C')
    print(subject_grading_three_papers('C4', 'C4', 'C4'))  
    print(subject_grading_three_papers('C3', 'C5', 'D1'))  
    print(subject_grading_three_papers('C3', 'C4', 'C4'))  
    print(subject_grading_three_papers('C3', 'D1', 'C5'))  
    print(subject_grading_three_papers('D1', 'C5', 'C4'))  
    print(subject_grading_three_papers('D1', 'C5', 'D2'))  
    print(subject_grading_three_papers('D1', 'C4', 'C4'))  

    print('Printing D')
    print(subject_grading_three_papers('C5', 'D2', 'C5'))  
    print(subject_grading_three_papers('C3', 'C6', 'D1'))  
    print(subject_grading_three_papers('C3', 'C6', 'C4'))  
    print(subject_grading_three_papers('C3', 'D1', 'C6'))  
    print(subject_grading_three_papers('C3', 'D1', 'C6'))  
    print(subject_grading_three_papers('C5', 'C5', 'C5'))  
    print(subject_grading_three_papers('C5', 'C6', 'C5'))  

    print('Printing E')
    print(subject_grading_three_papers('D1', 'P7', 'C3'))  
    print(subject_grading_three_papers('D2', 'P7', 'D1'))  
    print(subject_grading_three_papers('C3', 'C6', 'C6'))  
    print(subject_grading_three_papers('C3', 'D1', 'P7'))  
    print(subject_grading_three_papers('P8', 'C6', 'C6'))  
    print(subject_grading_three_papers('P8', 'D1', 'D1'))  
    print(subject_grading_three_papers('D2', 'C4', 'P8'))  
    print(subject_grading_three_papers('P8', 'C4', 'D1'))  
    print(subject_grading_three_papers('C5', 'C4', 'P7'))  

    print('Printing O')
    print(subject_grading_three_papers('F9', 'D2', 'C3'))  
    print(subject_grading_three_papers('F9', 'C3', 'D1'))  
    print(subject_grading_three_papers('F9', 'C4', 'C4'))  
    print(subject_grading_three_papers('F9', 'D1', 'C6'))  
    print(subject_grading_three_papers('P8', 'P8', 'C6'))  
    print(subject_grading_three_papers('C6', 'F9', 'C6'))  
    print(subject_grading_three_papers('P7', 'C6', 'P7'))  
    

    print('Printing F')
    print(subject_grading_three_papers('P7', 'F9', 'F9'))  
    print(subject_grading_three_papers('P8', 'F9', 'F9'))  
    print(subject_grading_three_papers('F9', 'F9', 'F9'))  
    
    