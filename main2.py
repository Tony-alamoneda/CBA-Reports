import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from fpdf import FPDF
import re
import os

STUDENT_COLUMN = "STUDENT"
PERCENT_DISPLAY_ALIASES = {
    "List %": "%",
    "Gram%": "%",
    "Voca %": "%",
    "Read %": "%",
}

DEFAULT_HEADER_COLOR = "#2A3F54"
HEADER_COLOR_MAP = {
    "List": "#6EDBE7",
    "List %": "#6EDBE7",
    "Gram": "#FF8A8A",
    "Gram%": "#FF8A8A",
    "Voca": "#FFB870",
    "Voca %": "#FFB870",
    "Read": "#76E1C4",
    "Read %": "#76E1C4",
}
COLUMN_ROW_COLORS = {
    "List": ("#f3feff", "#e1fcff"),
    "List %": ("#f3feff", "#e1fcff"),
    "Gram": ("#fffbfb", "#ffe9e9"),
    "Gram%": ("#fffbfb", "#ffe9e9"),
    "Voca": ("#fffaf5", "#fff1e3"),
    "Voca %": ("#fffaf5", "#fff1e3"),
    "Read": ("#f4fffc", "#dffff6"),
    "Read %": ("#f4fffc", "#dffff6"),
}
DEFAULT_ROW_COLORS = ("#FFFFFF", "#F7FAFC")
SELECTION_OUTLINE = "#4A90E2"

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
def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))

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

def style_assignment_table(df):
    """Return a pandas Styler configured with the academic table design.

    The styling complies with the requested palette, typography, alternating
    rows, hover colour, and subtle skill-based accent lines while keeping the
    table printer-friendly and compatible with downstream exports.
    """
    base_font = '"Inter", "Lato", "Segoe UI", sans-serif'
    accent_map = {**HEADER_COLOR_MAP, "Total": "#D18DF0"}

    styler = df.style.hide(axis="index")
    styler = styler.set_properties(**{
        "color": "#000000",
        "font-family": base_font,
        "font-size": "0.92rem",
    })
    if STUDENT_COLUMN in df.columns:
        styler = styler.set_properties(subset=pd.IndexSlice[:, [STUDENT_COLUMN]], **{
            "text-align": "left",
            "font-weight": "500",
        })
    numeric_columns = [col for col in df.columns if col != STUDENT_COLUMN]
    if numeric_columns:
        styler = styler.set_properties(subset=pd.IndexSlice[:, numeric_columns], **{
            "text-align": "center",
            "min-width": "72px",
        })

    table_styles = [
        {
            "selector": "table",
            "props": (
                "border-collapse: separate;"
                "border-spacing: 0;"
                "width: 100%;"
                "background-color: #FFFFFF;"
                "border: 1px solid #DCE3EA;"
                "border-radius: 6px;"
                "overflow: hidden;"
                f"font-family: {base_font};"
            ),
        },
        {
            "selector": "thead",
            "props": (
                "background-color: #2A3F54;"
                "color: #F7FBFF;"
                "text-transform: uppercase;"
                "letter-spacing: 0.5px;"
            ),
        },
        {
            "selector": "thead th",
            "props": (
                "padding: 12px 16px;"
                "font-weight: 600;"
                "border-right: 1px solid #DCE3EA;"
                "border-bottom: 1px solid #DCE3EA;"
                "text-align: center;"
            ),
        },
        {"selector": "thead th:first-child", "props": "text-align: left;"},
        {
            "selector": "tbody tr",
            "props": (
                "background-color: #FFFFFF;"
                "border-bottom: 1px solid #DCE3EA;"
            ),
        },
        {"selector": "tbody tr:nth-child(even)", "props": "background-color: #F7FAFC;"},
        {
            "selector": "tbody tr:hover",
            "props": "background-color: #E8F3FF;",
        },
        {
            "selector": "tbody td",
            "props": (
                "padding: 10px 14px;"
                "border-right: 1px solid #DCE3EA;"
                "vertical-align: middle;"
            ),
        },
        {
            "selector": "tbody td:first-child",
            "props": "font-weight: 500;",
        },
        {"selector": "tbody td:last-child", "props": "border-right: 0;"},
        {"selector": "thead th:last-child", "props": "border-right: 0;"},
        {"selector": "tbody tr:last-child td", "props": "border-bottom: 0;"},
    ]
    styler = styler.set_table_styles(table_styles)

    accent_styles = []
    for position, column in enumerate(df.columns, start=1):
        accent_color = accent_map.get(column)
        if accent_color:
            accent_styles.append({
                "selector": f"thead th:nth-child({position})",
                "props": f"border-bottom: 2px solid {accent_color};",
            })
    if accent_styles:
        styler = styler.set_table_styles(accent_styles, overwrite=False)

    if "Total" in df.columns:
        styler = styler.set_properties(subset=pd.IndexSlice[:, ["Total"]], **{"font-weight": "600"})

    return styler


def process_assignment(file_path_or_df, selected_units, skip_read=False):
    df = file_path_or_df if skip_read else read_flexible_excel(file_path_or_df)
    df['Student'] = df['Student'].apply(clean_student_name)
    df[['Earned', 'Total']] = df['Result (DateSubmitted)'].apply(lambda x: pd.Series(extract_score(x)))
    all_students = sorted(df['Student'].unique())
    skills = ['Listening', 'Grammar', 'Vocabulary', 'Reading']
    course = df['Class'].iloc[0].strip() if 'Class' in df.columns else "Unknown"

    df_filtered = df[df['Unit'].isin(selected_units)].copy()
    grouped = df_filtered.groupby(['Student', 'Content']).agg({'Earned': 'sum', 'Total': 'sum'}).reset_index()

    expected_totals_by_skill = {}
    for skill in skills:
        relevant_tasks = df_filtered[df_filtered['Content'] == skill]
        assignment_keys = relevant_tasks[['Title', 'Unit', 'Lesson']].drop_duplicates()
        expected_total = 0
        for _, task in assignment_keys.iterrows():
            matching = relevant_tasks[
                (relevant_tasks['Title'] == task['Title']) &
                (relevant_tasks['Unit'] == task['Unit']) &
                (relevant_tasks['Lesson'] == task['Lesson'])
            ]
            if not matching.empty:
                expected_total += matching['Total'].max()
        expected_totals_by_skill[skill] = expected_total

    earned_lookup = {}
    if not grouped.empty:
        earned_lookup = grouped.set_index(['Student', 'Content'])['Earned'].to_dict()

    results = []

    for student in all_students:
        row = {'Student': student}
        total_earned = 0
        total_expected = 0
        for skill in skills:
            earned = int(earned_lookup.get((student, skill), 0))

            expected_total = expected_totals_by_skill.get(skill, 0)

            row[f"{skill}"] = f"{earned}/{expected_total}" if expected_total else "0/0"
            row[f"{skill} %"] = f"{round((earned / expected_total * 100))}%" if expected_total else "0%"
            total_earned += earned
            total_expected += expected_total

        overall_percentage = round((total_earned / total_expected * 100)) if total_expected else 0
        row["Total"] = f"{overall_percentage}%"
        results.append(row)

    df_results = pd.DataFrame(results)
    column_renames = {
        "Student": STUDENT_COLUMN,
        "Listening": "List",
        "Listening %": "List %",
        "Grammar": "Gram",
        "Grammar %": "Gram%",
        "Vocabulary": "Voca",
        "Vocabulary %": "Voca %",
        "Reading": "Read",
        "Reading %": "Read %",
    }
    df_results.rename(columns=column_renames, inplace=True)

    styled = style_assignment_table(df_results)

    return df_results, course, styled

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

    df_results = pd.DataFrame(results)
    df_results.rename(columns={"Student": STUDENT_COLUMN}, inplace=True)

    return df_results, course

def try_configure_roboto(pdf: FPDF) -> str:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    regular_path = os.path.join(base_dir, "Roboto-Regular.ttf")
    bold_path = os.path.join(base_dir, "Roboto-Bold.ttf")
    if os.path.exists(regular_path) and os.path.exists(bold_path):
        try:
            pdf.add_font("Roboto", "", regular_path, uni=True)
            pdf.add_font("Roboto", "B", bold_path, uni=True)
            return "Roboto"
        except Exception:
            pass
    return "Arial"


def export_pdf(df, path, teacher, course, selected_units_label):
    pdf = FPDF(orientation='L')
    pdf.add_page()
    font_family = try_configure_roboto(pdf)
    pdf.set_font(font_family, style='B', size=12)
    class_info = extract_class_info(course)

    left_labels = ["Teacher:", "Course:", "Bimester:"]
    left_values = [teacher, course, class_info["bimester"]]
    right_labels = ["Branch:", "Schedule:", "Units:"]
    right_values = [class_info["branch"], class_info["schedule"], selected_units_label]

    label_width = 30
    value_width = 80
    gap = 36  # space between the two columns

    for i in range(3):
        pdf.set_font(font_family, style='B', size=11)
        pdf.cell(label_width, 8, left_labels[i], ln=0)
        pdf.set_font(font_family, size=11)
        pdf.cell(value_width, 8, str(left_values[i]), ln=0)

        pdf.set_x(label_width + value_width + gap)
        pdf.set_font(font_family, style='B', size=11)
        pdf.cell(label_width, 8, right_labels[i], ln=0)
        pdf.set_font(font_family, size=11)
        pdf.cell(value_width, 8, str(right_values[i]), ln=1)

    pdf.ln(3)

    pdf.set_font(font_family, size=11)
    total_table_width = 270
    student_col_width = 140
    remaining_columns = len(df.columns) - 1
    if remaining_columns > 0:
        numeric_col_width = (total_table_width - student_col_width) / remaining_columns
    else:
        numeric_col_width = total_table_width
    column_widths = []
    for col in df.columns:
        if col.lower() == "student":
            column_widths.append(student_col_width)
        else:
            column_widths.append(numeric_col_width)
    row_height = 10

    base_text_rgb = hex_to_rgb("#000000")
    header_text_rgb = hex_to_rgb("#FFFFFF")
    border_rgb = hex_to_rgb("#DCE3EA")
    stripe_rgb = hex_to_rgb(DEFAULT_ROW_COLORS[1])
    white_rgb = hex_to_rgb(DEFAULT_ROW_COLORS[0])
    default_header_bg = hex_to_rgb(DEFAULT_HEADER_COLOR)

    pdf.set_draw_color(*border_rgb)
    table_x = pdf.get_x()

    pdf.set_font(font_family, style='B', size=11)

    for col, width in zip(df.columns, column_widths):
        align = 'L' if col == STUDENT_COLUMN else 'C'
        header_hex = HEADER_COLOR_MAP.get(col) if col != STUDENT_COLUMN else None
        if header_hex:
            pdf.set_fill_color(*hex_to_rgb(header_hex))
        else:
            pdf.set_fill_color(*default_header_bg)
        pdf.set_text_color(*header_text_rgb)
        display_text = PERCENT_DISPLAY_ALIASES.get(col, col)
        if display_text != STUDENT_COLUMN:
            display_text = display_text.upper()
        pdf.cell(width, row_height + 2, display_text, border=1, align=align, fill=True)
    pdf.ln(row_height + 2)

    pdf.set_font(font_family, size=10)
    for row_index, row in enumerate(df.itertuples(index=False)):
        pdf.set_x(table_x)
        row_variant = row_index % 2
        for (col, width), value in zip(zip(df.columns, column_widths), row):
            if col == STUDENT_COLUMN:
                fill_rgb = white_rgb if row_variant == 0 else stripe_rgb
            else:
                palette = COLUMN_ROW_COLORS.get(col)
                if palette:
                    fill_rgb = hex_to_rgb(palette[row_variant])
                else:
                    fill_rgb = white_rgb if row_variant == 0 else stripe_rgb
            pdf.set_fill_color(*fill_rgb)
            align = 'L' if col == STUDENT_COLUMN else 'C'
            if col == "Total":
                pdf.set_font(font_family, style='B', size=10)
            else:
                pdf.set_font(font_family, size=10)
            pdf.set_text_color(*base_text_rgb)
            pdf.cell(width, row_height + 1, str(value), border=1, align=align, fill=True)
        pdf.ln(row_height + 1)

    pdf.set_text_color(0, 0, 0)
    pdf.output(path)

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CBA Surpass Report Generator")
        self.window_height = 520
        self.window_min_width = 900
        self.root.geometry(f"{self.window_min_width}x{self.window_height}")
        self.root.configure(bg="#F7F9FC")

        self.df = None
        self.df_internal = None
        self.df_filtered = None
        self.course = None
        self.report_type = None
        self.selected_units = []
        self.download_btn = None
        self.full_df = None
        self.assignment_styler = None
        self.table_canvas = None
        self.table_inner = None
        self.table_rows = []
        self.selected_row_indices = set()
        self.selection_anchor = None
        self.selection_active = False

        style_method = getattr(self, "set_style", None)
        if callable(style_method):
            style_method()

        top_frame = tk.Frame(root, bg="#F7F9FC")
        top_frame.pack(pady=6)

        input_frame = tk.Frame(top_frame, bg="#F7F9FC")
        input_frame.pack(side="left", padx=6)

        tk.Label(input_frame, text="Teacher:", bg="#F7F9FC", font=("Segoe UI", 12)).grid(row=0, column=0, sticky='w')
        self.teacher_entry = tk.Entry(input_frame, width=26, font=("Segoe UI", 12))
        self.teacher_entry.grid(row=0, column=1, pady=5)

        tk.Label(input_frame, text="Unit:", bg="#F7F9FC", font=("Segoe UI", 12)).grid(row=1, column=0, sticky='w')
        self.unit_var = tk.StringVar()
        self.unit_dropdown = ttk.Combobox(input_frame, textvariable=self.unit_var, state="readonly", font=("Segoe UI", 12), width=24)
        self.unit_dropdown.grid(row=1, column=1, pady=5)
        self.unit_dropdown.bind("<<ComboboxSelected>>", self.update_table_from_unit)

        button_frame = tk.Frame(top_frame, bg="#F7F9FC")
        button_frame.pack(side="left", padx=6)

        self.open_file_btn = ttk.Button(button_frame, text="Open File", style="Rounded.TButton", command=self.open_file)
        self.open_file_btn.grid(row=0, column=0, padx=(0, 8))

        self.download_btn = ttk.Button(button_frame, text="Download PDF", style="Rounded.TButton", command=self.save_pdf)
        self.download_btn.grid(row=0, column=1, padx=(0, 8))
        self.download_btn.config(state="disabled")

        self.teacher_entry.bind("<KeyRelease>", lambda e: self.check_pdf_button())
        self.unit_var.trace_add("write", lambda *args: self.check_pdf_button())

        self.table_frame = tk.Frame(root, bg="#F7F9FC")
        self.table_frame.pack(fill="both", expand=True, padx=8, pady=(4, 6))

        self.root.bind_all("<Control-c>", self.copy_selected_row)
        self.root.bind_all("<Control-C>", self.copy_selected_row)

        self.footer = tk.Label(root, text="Developed by Antonio Romano and A.I.", fg="gray", bg="#F7F9FC", font=("Helvetica", 8))
        self.footer.pack(side="bottom", pady=4)

    def set_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except:
            pass
        style.configure("Rounded.TButton",
                        font=("Segoe UI", 11),
                        padding=(10, 6),
                        borderwidth=0,
                        relief="flat",
                        background=DEFAULT_HEADER_COLOR,
                        foreground="#ffffff")
        style.map("Rounded.TButton",
                  background=[("active", "#354d67"), ("pressed", "#1f3041")],
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
                self.df_internal, self.course, self.assignment_styler = process_assignment(df, self.selected_units, skip_read=True)
            else:
                self.df_internal, self.course = process_test(df, self.selected_units, skip_read=True)
                self.assignment_styler = None

            self.df_filtered = self.df_internal.copy()
            self.df = self.df_filtered  # Store filtered table for export
            self.show_table(self.df_filtered)
            self.check_pdf_button()

        except Exception as e:
            messagebox.showerror("Error", str(e))


    def update_table_from_unit(self, event=None):
        if self.full_df is not None and self.unit_var.get():
            selected_units = self.get_selected_units()
            if self.report_type == "Assignment":
                self.df_internal, self.course, self.assignment_styler = process_assignment(self.full_df, selected_units, skip_read=True)
            else:
                self.df_internal, self.course = process_test(self.full_df, selected_units, skip_read=True)
                self.assignment_styler = None
            self.df_filtered = self.df_internal.copy()
            self.df = self.df_filtered
            self.show_table(self.df_filtered)

    def show_table(self, df):
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        self.table_rows = []
        self.selected_row_indices.clear()
        self.selection_anchor = None
        self.selection_active = False

        columns = list(df.columns)
        column_widths = []
        student_width = 260
        numeric_width = 90
        for col in columns:
            if col == STUDENT_COLUMN:
                column_widths.append(student_width)
            else:
                column_widths.append(numeric_width)

        container = tk.Frame(self.table_frame, bg="#F7F9FC")
        container.pack(fill="both", expand=True)

        self.table_canvas = tk.Canvas(container, bg="#F7F9FC", highlightthickness=0)
        self.table_canvas.pack(side="left", fill="both", expand=True)

        v_scroll = ttk.Scrollbar(container, orient="vertical", command=self.table_canvas.yview)
        v_scroll.pack(side="right", fill="y")
        self.table_canvas.configure(yscrollcommand=v_scroll.set)

        self.table_inner = tk.Frame(self.table_canvas, bg="#DCE3EA")
        self.table_canvas.create_window((0, 0), window=self.table_inner, anchor="nw")

        def _configure_canvas(event):
            self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))

        self.table_inner.bind("<Configure>", _configure_canvas)

        body_font = ("Roboto", 11)
        total_font = ("Roboto", 11, "bold")
        header_font = ("Roboto", 11, "bold")

        for idx, width in enumerate(column_widths):
            self.table_inner.grid_columnconfigure(idx, minsize=width)

        for col_index, col in enumerate(columns):
            header_bg = HEADER_COLOR_MAP.get(col, DEFAULT_HEADER_COLOR)
            display_text = PERCENT_DISPLAY_ALIASES.get(col, col)
            if display_text != STUDENT_COLUMN:
                display_text = display_text.upper()
            anchor = "w" if col == STUDENT_COLUMN else "center"
            label = tk.Label(
                self.table_inner,
                text=display_text,
                bg=header_bg,
                fg="#FFFFFF",
                font=header_font,
                padx=12,
                pady=8,
                bd=0,
                relief="flat",
                anchor=anchor,
            )
            label.grid(row=0, column=col_index, sticky="nsew", padx=(0 if col_index == 0 else 1), pady=0)

        for row_index, row in enumerate(df.itertuples(index=False), start=1):
            cell_widgets = []
            for col_index, (col, value) in enumerate(zip(columns, row)):
                if col == STUDENT_COLUMN:
                    palette = DEFAULT_ROW_COLORS
                else:
                    palette = COLUMN_ROW_COLORS.get(col, DEFAULT_ROW_COLORS)
                bg_color = palette[(row_index - 1) % 2]
                anchor = "w" if col == STUDENT_COLUMN else "center"
                font = total_font if col == "Total" else body_font
                cell = tk.Label(
                    self.table_inner,
                    text=str(value),
                    bg=bg_color,
                    fg="#000000",
                    font=font,
                    padx=12,
                    pady=6,
                    bd=0,
                    relief="flat",
                    anchor=anchor,
                )
                cell.grid(row=row_index, column=col_index, sticky="nsew", padx=(0 if col_index == 0 else 1), pady=(0, 1))
                cell.bind("<Button-1>", lambda e, idx=row_index - 1: self.on_row_press(e, idx))
                cell.bind("<B1-Motion>", lambda e, idx=row_index - 1: self.on_row_drag(e, idx))
                cell.bind("<Enter>", lambda e, idx=row_index - 1: self.on_row_enter(e, idx))
                cell.bind("<ButtonRelease-1>", self.on_row_release)
                cell_widgets.append({"widget": cell, "bg": bg_color})
            self.table_rows.append(cell_widgets)

        self.table_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.table_inner.bind("<MouseWheel>", self._on_mousewheel)
        self.table_canvas.bind("<Button-4>", self._on_mousewheel)
        self.table_canvas.bind("<Button-5>", self._on_mousewheel)
        self.table_inner.bind("<Button-4>", self._on_mousewheel)
        self.table_inner.bind("<Button-5>", self._on_mousewheel)

        total_table_width = sum(column_widths)
        required_width = total_table_width + 120
        self.root.update_idletasks()
        current_height = self.root.winfo_height()
        if current_height <= 1:
            current_height = self.window_height
        target_width = max(self.window_min_width, int(required_width))
        self.root.geometry(f"{target_width}x{int(current_height)}")

    def _on_mousewheel(self, event):
        if self.table_canvas is None:
            return
        if hasattr(event, "delta") and event.delta:
            self.table_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif getattr(event, "num", None) in (4, 5):
            direction = -1 if event.num == 4 else 1
            self.table_canvas.yview_scroll(direction, "units")

    def on_row_press(self, event, index):
        if index < 0 or index >= len(self.table_rows):
            return
        ctrl_pressed = bool(event.state & 0x0004)
        if ctrl_pressed:
            if index in self.selected_row_indices:
                new_selection = set(self.selected_row_indices)
                new_selection.remove(index)
            else:
                new_selection = set(self.selected_row_indices)
                new_selection.add(index)
            self._set_selected_rows(new_selection)
            self.selection_active = False
            self.selection_anchor = None
            return

        self.selection_active = True
        self.selection_anchor = index
        self._set_selected_rows({index})

    def on_row_drag(self, event, index):
        if not self.selection_active or self.selection_anchor is None:
            return
        index = max(0, min(index, len(self.table_rows) - 1))
        start = min(self.selection_anchor, index)
        end = max(self.selection_anchor, index)
        self._set_selected_rows(set(range(start, end + 1)))

    def on_row_enter(self, event, index):
        if self.selection_active and self.selection_anchor is not None and event.state & 0x0100:
            # 0x0100 corresponds to the left mouse button being pressed
            self.on_row_drag(event, index)

    def on_row_release(self, event):
        self.selection_active = False

    def _set_selected_rows(self, indices):
        indices = {idx for idx in indices if 0 <= idx < len(self.table_rows)}
        removed = self.selected_row_indices - indices
        added = indices - self.selected_row_indices

        for idx in removed:
            for cell in self.table_rows[idx]:
                widget = cell["widget"]
                widget.configure(highlightthickness=0)

        for idx in added:
            for cell in self.table_rows[idx]:
                widget = cell["widget"]
                widget.configure(highlightbackground=SELECTION_OUTLINE, highlightcolor=SELECTION_OUTLINE, highlightthickness=2)

        self.selected_row_indices = indices

    def copy_selected_row(self, event=None):
        if not self.selected_row_indices or self.df_filtered is None:
            return
        rows = []
        for idx in sorted(self.selected_row_indices):
            row_series = self.df_filtered.iloc[idx]
            rows.append("\t".join(str(row_series[col]) for col in self.df_filtered.columns))
        row_text = "\n".join(rows)
        self.root.clipboard_clear()
        self.root.clipboard_append(row_text)
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
            export_pdf(self.df_internal if self.df_internal is not None else self.df_filtered, filepath, teacher, course, unit_label)
            messagebox.showinfo("Export Complete", f"PDF saved to:\n{filepath}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()
