import numpy as np
import pandas as pd

def paper_grading(marks:np.number):
    print(marks)
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



def calc_total_points(subject_1_grade, subject_2_grade, subject_3_grade, subsidiary_grade, GP_grade):
    if pd.isna(subject_1_grade) or pd.isna(subject_2_grade) or pd.isna(subject_3_grade) or pd.isna(subsidiary_grade) or pd.isna(GP_grade):
        return 'X'
    
    valid_grades = {'A', 'B', 'C', 'D', 'E', 'O', 'F', 'D1', 'D2', 'C3', 'C4', 'C5', 'C6', 'P7', 'P8', 'F9', 'X'}
    
    valid_sub_gp_grades = {'D1', 'D2', 'C3', 'C4', 'C5', 'C6', 'P7', 'P8', 'F9', 'X'}
    for grade in [subject_1_grade, subject_2_grade, subject_3_grade]:
        if grade not in valid_grades:
            raise ValueError(f"Unrecognized grading for principal subjects: {grade}. This is normally because one or more of the students has missing marks in a subject not just a paper")
    
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
    if paper_grade_one in [None, ''] and paper_grade_two in [None, '']:
        return None
    
    
    if (pd.isna(paper_grade_one) or (pd.isnull(paper_grade_one)) or pd.isna(paper_grade_two) or pd.isnull(paper_grade_two)):
        return 'X'
    if paper_grade_one in [None, ''] or paper_grade_two in [None, '']:
        return 'X'
    
    grade_map = {
        'D1': 1, 'D2': 2, 'C3': 3, 'C4': 4, 'C5': 5,
        'C6': 6, 'P7': 7, 'P8': 8, 'F9': 9
    }
    
    grades = sorted([grade_map[paper_grade_one], grade_map[paper_grade_two]])
    
    if grades[0] <= 2 and grades[1] <= 2:
        return "A"
    
    if grades[0] <= 3 and grades[1] <= 3:
        return "B"
    
    if grades[0] <= 4 and grades[1] <= 4:
        return "C"
    
    if grades[0] <= 5 and grades[1] <= 5:
        return "D"
    
    if grades[1] <= 6 or (grades[0] in [7, 8] and sum(grades) <= 12):
        return "E"
    
    if grades[1] <= 8 or (grades[1] == 9 and grades[0] < 7):
        return "O"
    if grades[1] == 9 and grades[0] >=7:
        return "F"
    
    return "Invalid"



def subject_grading_three_papers(paper_grade_one:str, paper_grade_two:str, paper_grade_three:str):
    if (pd.isna(paper_grade_two) or pd.isnull(paper_grade_two)) and (pd.isna(paper_grade_one) or pd.isnull(paper_grade_one)) and (pd.isna(paper_grade_three) or pd.isnull(paper_grade_three)):
        return None
    if paper_grade_one in [None, ''] and paper_grade_two in [None, ''] and paper_grade_three in [None, '']:
        return None
    
    
    if (pd.isna(paper_grade_one) or pd.isnull(paper_grade_one)) or (pd.isna(paper_grade_two) or pd.isnull(paper_grade_two)) or (pd.isna(paper_grade_three) or pd.isnull(paper_grade_three)):
        return 'X'
    if paper_grade_one in [None, ''] or paper_grade_two in [None, ''] or paper_grade_three in [None, '']:
        return 'X'
    
    grade_map = {
        'D1': 1, 'D2': 2, 'C3': 3, 'C4': 4, 'C5': 5,
        'C6': 6, 'P7': 7, 'P8': 8, 'F9': 9
    }
    
    grades = sorted([grade_map[paper_grade_one], grade_map[paper_grade_two], grade_map[paper_grade_three]])

    if grades[0] < 3 and grades[1] < 3 and grades[2] <= 3:
        return "A"
    
    if grades[0] < 4 and grades[1] < 4 and grades[2] <= 4:
        return "B"
    
    if grades[0] < 5 and grades[1] < 5 and grades[2] <= 5:
        return "C"
    
    if grades[0] < 6 and grades[1] < 6 and grades[2] <= 6:
        return "D"
    
    if grades[0] <=6 and grades[1] <=6 and grades[2] <= 8:
        return "E"
    
    if (grades[0] <=8 and grades[1] <=8 and grades[2] <= 9) or \
        (grades[0] <=6 and grades[1] ==9 and grades[2] == 9):
        return "O"
    
    if (grades[0] > 6 and grades[1] == 9 and grades[2] == 9):
        return "F"

    return "Invalid"

def subject_grading_four_papers(paper_grade_one, paper_grade_two, paper_grade_three, paper_grade_four):
    
    if (pd.isna(paper_grade_two) or pd.isnull(paper_grade_two)) and (pd.isna(paper_grade_one) or pd.isnull(paper_grade_one)) and (pd.isna(paper_grade_three) or pd.isnull(paper_grade_three))  and (pd.isna(paper_grade_four) or pd.isnull(paper_grade_four)):
        return None
    if paper_grade_one in [None, ''] and paper_grade_two in [None, ''] and paper_grade_three in [None, ''] and paper_grade_four in [None, '']:
        return None
    
    
    if (pd.isna(paper_grade_one) or (pd.isnull(paper_grade_one)) or pd.isna(paper_grade_two) or pd.isnull(paper_grade_two)) or (pd.isna(paper_grade_three) or pd.isnull(paper_grade_three)) or (pd.isna(paper_grade_four) or pd.isnull(paper_grade_four)):
        return 'X'
    if paper_grade_one in [None, ''] or paper_grade_two in [None, ''] or paper_grade_three in [None, ''] or paper_grade_four in [None, '']:
        return 'X'
    
    
    
    grade_map = {
        'D1': 1, 'D2': 2, 'C3': 3, 'C4': 4, 'C5': 5,
        'C6': 6, 'P7': 7, 'P8': 8, 'F9': 9
    }
    
    grades = sorted([grade_map[paper_grade_one], grade_map[paper_grade_two], grade_map[paper_grade_three], grade_map[paper_grade_four]])
    
    if grades[0] < 3 and grades[1] < 3 and grades[2] < 3 and grades[3] <= 3:
        return "A"
    
    if grades[0] < 4 and grades[1] < 4 and grades[2] < 4 and grades[2] <= 4:
        return "B"
    
    if grades[0] < 5 and grades[1] < 5 and grades[2] <= 5 and grades[3] <= 5:
        return "C"
    
    if grades[0] < 6 and grades[1] < 6 and grades[2] <= 6 and grades[3] <= 6:
        return "D"
    
    if grades[0] <=6 and grades[1] <=6 and grades[2] <=6 and grades[3] <= 8:
        return "E"
    
    if (grades[0] <=8 and grades[1] <=8 and grades[2] <=8 and grades[3] <= 9) or \
        (grades[0] <=6 and grades[1] ==9 and grades[2] == 9):
        return "O"
    
    if (grades[0] > 6 and grades[1] > 6 and grades[2] == 9 and grades[3] == 9):
        return "F"

    return "Invalid"



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
    print(subject_grading_three_papers('C3', 'C3', 'C4'))  
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
    
    