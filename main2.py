import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from fpdf import FPDF
import re
import os

def extract_class_info(class_string):
    try:
        # Split by underscore
        parts = class_string.strip().split("_")

        if len(parts) < 4:
            return {"bimester": "Unknown", "branch": "Unknown", "course": "Unknown", "schedule": "Unknown"}

        bimester_raw = parts[0]
        bimester = bimester_raw[-2:]  # e.g., 254A -> 4A
        branch = parts[1]
        course = parts[2]
        schedule = parts[3]

        return {
            "bimester": bimester,
            "branch": branch,
            "course": course,
            "schedule": schedule
        }

    except Exception:
        return {"bimester": "Unknown", "branch": "Unknown", "course": "Unknown", "schedule": "Unknown"}

def clean_student_name(raw):
    match = re.search(r"\((.*?)\)", str(raw))
    return match.group(1) if match else str(raw).strip()

def extract_score(text):
    match = re.match(r"(\d+)/(\d+)", str(text))
    if match:
        return int(match.group(1)), int(match.group(2))
    match = re.match(r"(\d+)", str(text))
    if match:
        return int(match.group(1)), 100
    return 0, 0

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '', str(name))

def read_flexible_excel(file_path):
    try:
        df = pd.read_html(file_path)[0]
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        raise ValueError(f"Could not read the file. Details:\n{str(e)}")

def detect_report_type(df):
    if "Category" in df.columns:
        return "Test"
    elif "Content" in df.columns and "Lesson" in df.columns:
        return "Assignment"
    else:
        raise ValueError("Unknown report format. Make sure your Excel has the correct structure.")

def extract_available_unit_options(df):
    unit_list = sorted(df['Unit'].dropna().unique())
    unit_list = sorted([unit.strip() for unit in unit_list if unit.startswith("Unit")])
    available_options = list(unit_list)

    pairs = [("Unit 2", "Unit 3"), ("Unit 4", "Unit 5"), ("Unit 6", "Unit 7"), ("Unit 9", "Unit 10")]
    for u1, u2 in pairs:
        if u1 in unit_list and u2 in unit_list:
            available_options.append(f"{u1} & {u2}")

    return available_options

def parse_units_from_label(label):
    if "&" in label:
        return [u.strip() for u in label.split("&")]
    else:
        return [label.strip()]

def process_assignment(file_path_or_df, selected_units, skip_read=False):
    df = file_path_or_df if skip_read else read_flexible_excel(file_path_or_df)
    df['Student'] = df['Student'].apply(clean_student_name)
    df[['Earned', 'Total']] = df['Result (DateSubmitted)'].apply(lambda x: pd.Series(extract_score(x)))
    all_students = sorted(df['Student'].unique())
    skills = ['Listening', 'Grammar', 'Vocabulary', 'Reading']
    course = df['Class'].iloc[0].strip() if 'Class' in df.columns else "Unknown"

    df_filtered = df[df['Unit'].isin(selected_units)].copy()
    grouped = df_filtered.groupby(['Student', 'Content']).agg({'Earned': 'sum', 'Total': 'sum'}).reset_index()

    results = []

    for student in all_students:
        row = {'Student': student}
        for skill in skills:
            student_skill = grouped[(grouped['Student'] == student) & (grouped['Content'] == skill)]
            earned = int(student_skill['Earned'].values[0]) if not student_skill.empty else 0

            relevant_tasks = df_filtered[df_filtered['Content'] == skill]
            assignment_keys = relevant_tasks[['Title', 'Unit', 'Lesson']].drop_duplicates()
            expected_total = 0
            for _, task in assignment_keys.iterrows():
                matching = relevant_tasks[
                    (relevant_tasks['Title'] == task['Title']) &
                    (relevant_tasks['Unit'] == task['Unit']) &
                    (relevant_tasks['Lesson'] == task['Lesson'])
                ]
                expected_total += matching['Total'].max()

            row[f"{skill}"] = f"{earned}/{expected_total}"
            row[f"{skill} %"] = round((earned / expected_total * 100)) if expected_total else 0
        results.append(row)

    return pd.DataFrame(results), course

def process_test(file_path_or_df, selected_units, skip_read=False):
    df = file_path_or_df if skip_read else read_flexible_excel(file_path_or_df)
    df['Student'] = df['Student'].apply(clean_student_name)
    df['Score'] = df['Result (DateSubmitted)'].apply(lambda x: extract_score(x)[0])
    df_all = df[['Student', 'Score', 'Unit', 'Class']]

    course = df['Class'].dropna().astype(str).str.strip().iloc[0] if 'Class' in df.columns and not df['Class'].dropna().empty else "Unknown"

    all_students = df_all['Student'].unique()
    df_filtered = df_all[df_all['Unit'].isin(selected_units)]

    results = []
    for student in sorted(all_students):
        student_score = df_filtered[df_filtered['Student'] == student]['Score']
        score = int(student_score.iloc[0]) if not student_score.empty else 0
        results.append({'Student': student, 'Score': score})

    return pd.DataFrame(results), course

def export_pdf(df, path, teacher, course, selected_units_label):
    pdf = FPDF(orientation='L')
    pdf.add_page()
    pdf.set_font("Arial", style='B', size=12)
    class_info = extract_class_info(course)

    left_labels = ["Teacher:", "Course:", "Bimester:"]
    left_values = [teacher, course, class_info["bimester"]]
    right_labels = ["Branch:", "Schedule:", "Units:"]
    right_values = [class_info["branch"], class_info["schedule"], selected_units_label]

    label_width = 30
    value_width = 80
    gap = 40  # space between the two columns

    for i in range(3):
        pdf.set_font("Arial", style='B', size=11)
        pdf.cell(label_width, 10, left_labels[i], ln=0)
        pdf.set_font("Arial", size=11)
        pdf.cell(value_width, 10, str(left_values[i]), ln=0)

        pdf.set_x(label_width + value_width + gap)
        pdf.set_font("Arial", style='B', size=11)
        pdf.cell(label_width, 10, right_labels[i], ln=0)
        pdf.set_font("Arial", size=11)
        pdf.cell(value_width, 10, str(right_values[i]), ln=1)

    pdf.ln(5)

    pdf.set_font("Arial", size=12)
    col_width = 270 / len(df.columns)
    row_height = 10

    pdf.set_font(style="B")
    for col in df.columns:
        pdf.cell(col_width, row_height, col, border=1, align='C')
    pdf.ln(row_height)

    for _, row in df.iterrows():
        pdf.set_font("Arial", size=10)
        for item in row:
            pdf.cell(col_width, row_height, str(item), border=1, align='C')
        pdf.ln(row_height)

    pdf.output(path)

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CBA Surpass Report Generator")
        self.root.geometry("1000x600")
        self.root.configure(bg="#F7F9FC")

        self.df = None
        self.df_filtered = None
        self.course = None
        self.report_type = None
        self.selected_units = []
        self.tree = None
        self.download_btn = None
        self.full_df = None

        self.set_style()

        top_frame = tk.Frame(root, bg="#F7F9FC")
        top_frame.pack(pady=20)

        input_frame = tk.Frame(top_frame, bg="#F7F9FC")
        input_frame.pack(side="left", padx=10)

        tk.Label(input_frame, text="Teacher:", bg="#F7F9FC", font=("Segoe UI", 12)).grid(row=0, column=0, sticky='w')
        self.teacher_entry = tk.Entry(input_frame, width=30, font=("Segoe UI", 12))
        self.teacher_entry.grid(row=0, column=1, pady=5)

        tk.Label(input_frame, text="Unit:", bg="#F7F9FC", font=("Segoe UI", 12)).grid(row=1, column=0, sticky='w')
        self.unit_var = tk.StringVar()
        self.unit_dropdown = ttk.Combobox(input_frame, textvariable=self.unit_var, state="readonly", font=("Segoe UI", 12), width=28)
        self.unit_dropdown.grid(row=1, column=1, pady=5)
        self.unit_dropdown.bind("<<ComboboxSelected>>", self.update_table_from_unit)

        button_frame = tk.Frame(top_frame, bg="#F7F9FC")
        button_frame.pack(side="left", padx=20)

        self.open_file_btn = ttk.Button(button_frame, text="Open File", style="Rounded.TButton", command=self.open_file)
        self.open_file_btn.grid(row=0, column=0, padx=10)

        self.download_btn = ttk.Button(root, text="Download PDF", style="Rounded.TButton", command=self.save_pdf)
        self.download_btn.pack(pady=10)
        self.download_btn.config(state="disabled")

        self.teacher_entry.bind("<KeyRelease>", lambda e: self.check_pdf_button())
        self.unit_var.trace_add("write", lambda *args: self.check_pdf_button())

        self.table_frame = tk.Frame(root, bg="#F7F9FC")
        self.table_frame.pack(fill="both", expand=True)

        self.footer = tk.Label(root, text="Powered by Tony R. and ChatGPT", fg="gray", bg="#F7F9FC", font=("Helvetica", 8))
        self.footer.pack(side="bottom", pady=5)

    def set_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except:
            pass
        style.configure("Rounded.TButton",
                        font=("Segoe UI", 11),
                        padding=(12, 8),
                        borderwidth=0,
                        relief="flat",
                        background="#4A90E2",
                        foreground="#ffffff")
        style.map("Rounded.TButton",
                  background=[("active", "#357ABD"), ("pressed", "#2C5E9E")],
                  relief=[("pressed", "sunken")])

    def get_selected_units(self):
        selected_label = self.unit_var.get()
        return parse_units_from_label(selected_label)

    def check_pdf_button(self):
        if self.full_df is not None and self.teacher_entry.get().strip() and self.unit_var.get().strip():
            self.download_btn.config(state="normal")
        else:
            self.download_btn.config(state="disabled")

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if not file_path:
            return

        try:
            df = read_flexible_excel(file_path)
            self.report_type = detect_report_type(df)
            self.full_df = df  # Store for reuse
            unit_options = extract_available_unit_options(df)

            if not unit_options:
                messagebox.showwarning("No Units Found", "No valid units were found in this Excel file.")
                return

            self.unit_dropdown['values'] = unit_options
            self.unit_dropdown.set(unit_options[0])
            self.selected_units = parse_units_from_label(unit_options[0])

            if self.report_type == "Assignment":
                self.df_filtered, self.course = process_assignment(df, self.selected_units, skip_read=True)
            else:
                self.df_filtered, self.course = process_test(df, self.selected_units, skip_read=True)

            self.df = self.df_filtered  # Store filtered table for export
            self.show_table(self.df_filtered)
            self.check_pdf_button()

        except Exception as e:
            messagebox.showerror("Error", str(e))


    def update_table_from_unit(self, event=None):
        if self.full_df is not None and self.unit_var.get():
            selected_units = self.get_selected_units()
            if self.report_type == "Assignment":
                filtered_df, _ = process_assignment(self.full_df, selected_units, skip_read=True)
            else:
                filtered_df, _ = process_test(self.full_df, selected_units, skip_read=True)
            self.df_filtered = filtered_df
            self.show_table(filtered_df)

    def show_table(self, df):
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        columns = list(df.columns)
        self.tree = ttk.Treeview(self.table_frame, columns=columns, show="headings", height=16)
        
        # Scrollbars
        x_scroll = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        y_scroll = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        y_scroll.grid(row=0, column=1, sticky='ns')
        x_scroll.grid(row=1, column=0, sticky='ew')

        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        # Set column widths
        for col in columns:
            if col.lower() == "student":
                self.tree.column(col, width=200, anchor='w')
            else:
                self.tree.column(col, width=90, anchor='center')
            self.tree.heading(col, text=col)

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

        self.enable_copy_on_tree(self.tree)

    def enable_copy_on_tree(self, tree_widget):
        def copy_selection(event):
            selected_items = tree_widget.selection()
            if not selected_items:
                return
            rows = []
            for item in selected_items:
                values = tree_widget.item(item, "values")
                rows.append("\t".join(str(value) for value in values))
            if rows:
                self.root.clipboard_clear()
                self.root.clipboard_append("\n".join(rows))
                self.root.update()

        tree_widget.bind("<Control-c>", copy_selection)
        tree_widget.bind("<Control-C>", copy_selection)


    def copy_selection_to_clipboard(self, event):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        rows = []
        for item in selected_items:
            row = self.tree.item(item)['values']
            rows.append("\t".join(map(str, row)))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(rows))
        self.root.update()

    def save_pdf(self):
        if self.df_filtered is None:
            messagebox.showerror("Error", "No data to export.")
            return
        teacher = self.teacher_entry.get().strip()
        course = self.course if self.course else "Unknown"
        unit_label = self.unit_var.get()
        filename = f"{sanitize_filename(teacher)} - {sanitize_filename(course)} - {sanitize_filename(unit_label)}.pdf"
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=filename,
                                                filetypes=[("PDF files", "*.pdf")])
        if filepath:
            export_pdf(self.df_filtered, filepath, teacher, course, unit_label)

if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()
