import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from application_logic import make_reports

def main(marksheet_folder_o_path, marksheet_folder_a_level, template_path, output_path, year, term, report_type):
    make_reports(unprocessed_marks_o_level=marksheet_folder_o_path,unprocessed_marks_a_level= marksheet_folder_a_level, template_path=template_path, output_path=output_path, term=term, year=year, report_type=report_type)

def select_folder(entry_widget):
    folder_path = filedialog.askdirectory()
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, folder_path)

def select_file(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def update_template_requirement(*args):
    report_type = report_type_var.get()
    if report_type in ["End of Term", "Mid Term"]:
        template_label.grid(row=3, column=0, padx=10, pady=10)
        template_entry.grid(row=3, column=1, padx=10, pady=10)
        template_browse_button.grid(row=3, column=2, padx=10, pady=10)
        year_label.grid(row=5, column=0, padx=10, pady=10)
        year_entry.grid(row=5, column=1, padx=10, pady=10)
        term_label.grid(row=6, column=0, padx=10, pady=10)
        term_entry.grid(row=6, column=1, padx=10, pady=10)
    else:
        template_label.grid_remove()
        template_entry.grid_remove()
        template_browse_button.grid_remove()
        year_label.grid_remove()
        year_entry.grid_remove()
        term_label.grid_remove()
        term_entry.grid_remove()

def process_reports():
    report_types={
        "End of Term":'END_OF_TERM_REPORT', 
        "Mid Term":'MID_TERM_REPORT', 
        "Subject Summary Report":'SUBJECT_SUMMARY_REPORT', 
        "Marks Summary":'MARKS_SUMMARY_REPORT'
    }
    report_type = report_type_var.get()
    marksheet_folder_a_path = marksheet_folder_a_level.get()
    marksheet_folder_o_path = marksheet_folder_o_level.get()
    template_path = template_entry.get()
    output_path = output_entry.get()
    year = year_entry.get()
    term = term_entry.get()
    
    if not marksheet_folder_a_path or not marksheet_folder_o_path or not output_path:
        messagebox.showerror("Error", "Marks sheet folder and output folder must be selected.")
        return

    if report_type in ["End of Term", "Mid Term"]:
        if not template_path or not year or not term:
            messagebox.showerror("Error", "Template folder, year, and term must be selected for End of Term and Mid Term reports.")
            return

    try:
        main(marksheet_folder_o_path, marksheet_folder_a_path, template_path, output_path, year, term, report_type=report_types[report_type])
        messagebox.showinfo("Success", "Report cards generated successfully.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

app = tk.Tk()
app.title("Report Card Generator")

tk.Label(app, text="Marks Sheet Folder O level").grid(row=0, column=0, padx=10, pady=10)
marksheet_folder_o_level = tk.Entry(app, width=50)
marksheet_folder_o_level.grid(row=0, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=lambda: select_folder(marksheet_folder_o_level)).grid(row=0, column=2, padx=10, pady=10)

tk.Label(app, text="Marks Sheet Folder A level").grid(row=1, column=0, padx=10, pady=10)
marksheet_folder_a_level = tk.Entry(app, width=50)
marksheet_folder_a_level.grid(row=1, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=lambda: select_folder(marksheet_folder_a_level)).grid(row=1, column=2, padx=10, pady=10)


tk.Label(app, text="Report Type").grid(row=2, column=0, padx=10, pady=10)
report_type_var = tk.StringVar()
report_type_selector = ttk.Combobox(app, textvariable=report_type_var)
report_type_selector['values'] = ("End of Term", "Mid Term", "Subject Summary Report", "Marks Summary")
report_type_selector.grid(row=2, column=1, padx=10, pady=10)
report_type_selector.current(3)
report_type_selector.bind("<<ComboboxSelected>>", update_template_requirement)

template_label = tk.Label(app, text="Template Folder")
template_entry = tk.Entry(app, width=50)
template_browse_button = tk.Button(app, text="Browse", command=lambda: select_folder(template_entry))

tk.Label(app, text="Output Folder").grid(row=4, column=0, padx=10, pady=10)
output_entry = tk.Entry(app, width=50)
output_entry.grid(row=4, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=lambda: select_folder(output_entry)).grid(row=4, column=2, padx=10, pady=10)

year_label = tk.Label(app, text="Year")
year_entry = tk.Entry(app, width=50)

term_label = tk.Label(app, text="Term")
term_entry = tk.Entry(app, width=50)

process_button = tk.Button(app, text="Generate Report Cards", command=process_reports)
process_button.grid(row=7, column=1, padx=10, pady=10)

update_template_requirement()  

app.mainloop()


