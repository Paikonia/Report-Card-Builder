import tkinter as tk
from tkinter import filedialog, messagebox

from application_logic import make_reports


def main(end_of_term_path, mid_term_path, template_path, output_path):
    make_reports(end_of_term_path, templete_path=template_path, output_path=output_path)    

def select_folder(entry_widget):
    folder_path = filedialog.askdirectory()
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, folder_path)

def select_file(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def process_reports():
    end_of_term_path = end_of_term_entry.get()
    mid_term_path = mid_term_entry.get()
    template_path = template_entry.get()
    output_path = output_entry.get()
    
    if not end_of_term_path or not mid_term_path or not template_path or not output_path:
        messagebox.showerror("Error", "All paths must be selected.")
        return
    
    try:
        main(end_of_term_path, mid_term_path, template_path, output_path)
        messagebox.showinfo("Success", "Report cards generated successfully.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

app = tk.Tk()
app.title("Report Card Generator")

tk.Label(app, text="End of Term Folder").grid(row=0, column=0, padx=10, pady=10)
end_of_term_entry = tk.Entry(app, width=50)
end_of_term_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=lambda: select_folder(end_of_term_entry)).grid(row=0, column=2, padx=10, pady=10)

tk.Label(app, text="Mid Term Folder").grid(row=1, column=0, padx=10, pady=10)
mid_term_entry = tk.Entry(app, width=50)
mid_term_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=lambda: select_folder(mid_term_entry)).grid(row=1, column=2, padx=10, pady=10)

tk.Label(app, text="Template Folder").grid(row=2, column=0, padx=10, pady=10)
template_entry = tk.Entry(app, width=50)
template_entry.grid(row=2, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=lambda: select_folder(template_entry)).grid(row=2, column=2, padx=10, pady=10)

tk.Label(app, text="Output Folder").grid(row=3, column=0, padx=10, pady=10)
output_entry = tk.Entry(app, width=50)
output_entry.grid(row=3, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=lambda: select_folder(output_entry)).grid(row=3, column=2, padx=10, pady=10)

process_button = tk.Button(app, text="Generate Report Cards", command=process_reports)
process_button.grid(row=4, column=1, padx=10, pady=10)

app.mainloop()
