import os
from mailmerge import MailMerge
from parser_logic import iterate_through_subjects, combine_subjects
def run_mail_merge(data_frame, template_path, output_path, class_name):
    n_path = os.path.join(output_path, class_name)
    os.makedirs(n_path, exist_ok=True)
    non_numeric_columns = data_frame.select_dtypes(exclude=['int', 'float']).columns
    data_frame[non_numeric_columns] = data_frame[non_numeric_columns].fillna('')
    
    numeric_columns = data_frame.select_dtypes(include=['int', 'float']).columns
    data_frame[numeric_columns] = data_frame[numeric_columns].fillna('').astype(str)
    
    records = data_frame.astype(str).to_dict('records')
    
    for record in records:
        record = {k: str(v) for k, v in record.items()}
        
        if 'Name' not in record:
            raise KeyError("'Name' field is missing in the record.")
        
        document = MailMerge(template_path)
        print(document.get_merge_fields())
        document.merge(**record)
        
        output_file = os.path.join(n_path, f'Report Card {record["Name"]}.docx')
        
        document.write(output_file)
        document.close()

def make_reports(unprocessed_marks, templete_path, output_path):
    
    dfs_end_of_term = combine_subjects(parsed_subjects=iterate_through_subjects(unprocessed_marks))
    
    for c in dfs_end_of_term.keys():
        file_name = f'Report Card Template - {c}.docx'
        path = os.path.join(templete_path, file_name)
        if os.path.exists(path=path):
            run_mail_merge(data_frame=dfs_end_of_term[c], template_path=path, output_path=output_path, class_name=c)
        else:
            raise FileExistsError(f'Template for {c} does not exist.')




if __name__ == '__main__':
    make_reports('./test_subjects', './G', './')
    
    
    