import os
from mailmerge import MailMerge
from parser_logic_advanced_level import parse_advanced_level_marksheet
from parser_logic_ordinary_level import parse_ordinary_level_marksheets

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
        print(record)
        if 'Name' not in record:
            raise KeyError("'Name' field is missing in the record.")
        
        document = MailMerge(template_path)
        document.merge(**record)
        
        output_file = os.path.join(n_path, f'Report Card {record["Name"]}.docx')
        
        document.write(output_file)
        document.close()

def make_reports(unprocessed_marks, templete_path, output_path, term , year):
    error = {}
    classes = {
        'Senior One': 'S.1',
        'Senior Two': 'S.2',
        'Senior Three': 'S.3',
        'Senior Four': 'S.4',
        'Senior Five': 'S.5',
        'Senior Six': 'S.6'
    }
    a_level_dataframe = parse_advanced_level_marksheet(folder_path=unprocessed_marks)
    o_level_dataframe = parse_ordinary_level_marksheets(folder_path=unprocessed_marks)
    
    a_level_template_name = "A' level Report Card Template.docx"
    a_level_template_path = os.path.join(templete_path, a_level_template_name)
    if os.path.exists(path=a_level_template_path):
        raise NameError(f'The folder {templete_path} does not include a template for A level.\nThe template should be named {a_level_template_name}')
    
    o_level_template_name = "A' level Report Card Template.docx"
    o_level_template_path = os.path.join(templete_path, o_level_template_name)
    if os.path.exists(path=o_level_template_path):
        raise NameError(f'The folder {templete_path} does not include a template for A level.\nThe template should be named {o_level_template_name}')
    
    
    for c, df in a_level_dataframe.items():
        try:
            run_mail_merge(data_frame=df, template_path=a_level_template_path, output_path=output_path, class_name=classes[c], academic_year=year, academic_term=term)
        except Exception as e:
            error['Caught Error A level'] = f'Other errors: {e}'
    

    for c, df in o_level_dataframe.items():
        try:
            run_mail_merge(data_frame=df, template_path=o_level_template_path, output_path=output_path, class_name=classes[c], academic_year=year, academic_term=term)
        except Exception as e:
            error['Caught Error O level'] = f'Other errors: {e}'



if __name__ == '__main__':
    make_reports('./test_subjects', './G', './')
    
    
    