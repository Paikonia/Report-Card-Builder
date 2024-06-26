import os
from mailmerge import MailMerge
from pandas import ExcelWriter
from parser_logic_advanced_level import parse_advanced_level_marksheet
from parser_logic_ordinary_level import parse_ordinary_level_marksheets

from subject_report_generator_a_level import make_subject_report_a_level
from subject_report_generator_o_level import make_subject_report_o_level


def run_mail_merge(data_frame, template_path, output_path, class_name, academic_year, academic_term):
    n_path = os.path.join(output_path, class_name)
    os.makedirs(n_path, exist_ok=True)
    non_numeric_columns = data_frame.select_dtypes(exclude=['int', 'float']).columns
    data_frame[non_numeric_columns] = data_frame[non_numeric_columns].fillna('')
    
    numeric_columns = data_frame.select_dtypes(include=['int', 'float']).columns
    data_frame[numeric_columns] = data_frame[numeric_columns].fillna('').astype(str)
    
    records = data_frame.astype(str).to_dict('records')
    
    for record in records:
        record = {k: str(v) for k, v in record.items()}
        
        record.update({
            "_class": class_name,
            "term": academic_term,
            "year": academic_year
        })
        if 'Name' not in record and 'name' not in record:
            raise KeyError("'Name' field is missing in the record.")
        
        document = MailMerge(template_path)
        
        document.merge(**record)
        if 'Name' in record:
            output_file = os.path.join(n_path, f'Report Card {record["Name"]}.docx')
        if 'name' in record:
            output_file = os.path.join(n_path, f'Report Card {record["name"]}.docx')
        document.write(output_file)
        document.close()

def make_reports(unprocessed_marks_o_level, unprocessed_marks_a_level, template_path, output_path, term , year, report_type='MARKS_SUMMARY_REPORT'):
    error = {}
    caught_error = False
    classes = {
        'Senior One': 'S.1',
        'Senior Two': 'S.2',
        'Senior Three': 'S.3',
        'Senior Four': 'S.4',
        'Senior Five': 'S.5',
        'Senior Six': 'S.6'
    }
    
    if report_type == 'SUBJECT_SUMMARY_REPORT':
        a_level_per_subject_report = make_subject_report_a_level(unprocessed_marks_a_level)
        o_level_per_subject_report = make_subject_report_o_level(unprocessed_marks_o_level)
        file_path_a_level = os.path.join(output_path, 'Subject Reports A level.xlsx')
        file_path_o_level = os.path.join(output_path, 'Subject Reports O level.xlsx')
        with ExcelWriter(file_path_a_level) as writer:
            for class_name, df in a_level_per_subject_report.items():
                df.to_excel(writer, sheet_name=class_name, index=False)
            
        with ExcelWriter(file_path_o_level) as writer:
            for class_name, df in o_level_per_subject_report.items():
                df.to_excel(writer, sheet_name=class_name, index=False)
        return
    
    
    
    if report_type == 'END_OF_TERM_REPORT':
        (end_of_term_a_level, _) = parse_advanced_level_marksheet(folder_path=unprocessed_marks_a_level, report_type=report_type)
        (end_of_term_o_level, _) = parse_ordinary_level_marksheets(folder_path=unprocessed_marks_o_level, report_type=report_type)
    
        a_level_template_name = "A' level Report Card Template.docx"
        a_level_template_path = os.path.join(template_path, a_level_template_name)
        if not os.path.exists(path=a_level_template_path):
            raise NameError(f'The folder {template_path} does not include a template for A level.\nThe template should be named {a_level_template_name}')    
        
        o_level_template_name = "O' level Report Card Template.docx"
        o_level_template_path = os.path.join(template_path, o_level_template_name)
        if not os.path.exists(path=o_level_template_path):
            raise NameError(f'The folder {template_path} does not include a template for O level.\nThe template should be named {o_level_template_name}')
        for c, df in end_of_term_a_level.items():
            try:
                run_mail_merge(data_frame=df, template_path=a_level_template_path, output_path=output_path, class_name=classes[c], academic_year=year, academic_term=term)
            except Exception as e:
                caught_error = True
                if 'Caught Error A level' not in error:
                    error['Caught Error A level'] = []
                error['Caught Error A level'].append(f'{e}')
        for c, df in end_of_term_o_level.items():
            try:
                run_mail_merge(data_frame=df, template_path=o_level_template_path, output_path=output_path, class_name=classes[c], academic_year=year, academic_term=term)
            except Exception as e:
                caught_error = True
                if 'Caught Error O level' not in error:
                    error['Caught Error O level'] = []
                error['Caught Error O level'].append(f'{e}')
        if caught_error:
            raise Exception(error)
        
        
    if report_type == 'MID_TERM_REPORT':
        (_, mid_term_results_a_level) = parse_advanced_level_marksheet(folder_path=unprocessed_marks_a_level, report_type=report_type)
        (_, mid_term_results_o_level) = parse_ordinary_level_marksheets(folder_path=unprocessed_marks_o_level, report_type=report_type)
    
        a_level_template_name = "A' level Mid Term Report Card Template.docx"
        a_level_template_path = os.path.join(template_path, a_level_template_name)
        if not os.path.exists(path=a_level_template_path):
            raise NameError(f'The folder {template_path} does not include a template for A level Mid Term.\nThe template should be named {a_level_template_name}')    
        
        o_level_template_name = "O' level Mid Term Report Card Template.docx"
        o_level_template_path = os.path.join(template_path, o_level_template_name)
        if not os.path.exists(path=o_level_template_path):
            raise NameError(f'The folder {template_path} does not include a template for O levelMid Term.\nThe template should be named {o_level_template_name}')
        for c, df in mid_term_results_a_level.items():
            try:
                run_mail_merge(data_frame=df, template_path=a_level_template_path, output_path=output_path, class_name=classes[c], academic_year=year, academic_term=term)
            except Exception as e:
                caught_error = True
                if 'Caught Error A level' not in error:
                    error['Caught Error A level'] = []
                error['Caught Error A level'].append(f'{e}')
        for c, df in mid_term_results_o_level.items():
            try:
                run_mail_merge(data_frame=df, template_path=o_level_template_path, output_path=output_path, class_name=classes[c], academic_year=year, academic_term=term)
            except Exception as e:
                caught_error = True
                if 'Caught Error O level' not in error:
                    error['Caught Error O level'] = []
                error['Caught Error O level'].append(f'{e}')
        if caught_error:
            raise Exception(error)
    
    if report_type == 'MARKS_SUMMARY_REPORT':
        (end_of_term_a_level, mid_term_a_level) = parse_advanced_level_marksheet(folder_path=unprocessed_marks_a_level, report_type=report_type)
        (end_of_term_o_level, mid_term_o_level) = parse_ordinary_level_marksheets(folder_path=unprocessed_marks_o_level, report_type=report_type)
    
        file_path = os.path.join(output_path, 'Marks Summary.xlsx')
        with ExcelWriter(file_path) as writer:
            for class_name, df in end_of_term_a_level.items():
                df.to_excel(writer, sheet_name=class_name, index=False)
            for class_name, df in mid_term_a_level.items():
                df.to_excel(writer, sheet_name=f'Mid Term {class_name}', index=False)
            for class_name, df in end_of_term_o_level.items():
                df.to_excel(writer, sheet_name=class_name, index=False)
            for class_name, df in mid_term_o_level.items():
                df.to_excel(writer, sheet_name=f'Mid Term {class_name}', index=False)
            

    
if __name__ == '__main__':
    make_reports('./test_subjects/Marks Sheet O level', './test_subjects/Marks Sheet A level', './test_subjects/templates', 'test_subjects/Reports', '1', '2019')
    
    
    