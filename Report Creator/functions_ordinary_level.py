import numpy as np
import pandas as pd

def calculate_average(r1, r2, r3):
    values = [r1, r2, r3]
    non_null_values = [value for value in values if pd.notnull(value) and pd.to_numeric(value, errors='coerce') == value]
    
    if len(non_null_values) == 3:
        return np.nanmean(non_null_values) * 0.2
    elif len(non_null_values) == 2:
        return np.nanmean(non_null_values) * 0.2
    elif len(non_null_values) == 1:
        return non_null_values[0] * 0.2  
    else:
        return np.nan

def calc_grading(total):
    if(np.isnan(total)):
        return ''
    if total >= 90:
        return 'A*'
    elif total >= 80:
        return 'A'
    elif total >= 70:
        return 'B'
    elif total >= 60:
        return 'C'
    elif total >= 50:
        return 'D'
    elif total >= 40:
        return 'E'
    elif total >= 30:
        return 'F'
    else:
        return 'G'
    
def abbreviate_column_name(column_name):
    column_name = column_name.replace('Formative Score', 'fs')
    column_name = column_name.replace('EOT Score', 'es')
    column_name = column_name.replace('Total Score', 'to')
    column_name = column_name.replace('Grade', 'grade')
    return column_name.replace(' ', '_')