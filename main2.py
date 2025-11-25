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
        parts = str(class_string).strip().split("_")
        if len(parts) < 4:
            return {
                "bimester": "Unknown",
                "branch": "Unknown",
                "course": "Unknown",
                "schedule": "Unknown",
            }

        bimester_raw = parts[0].strip()
        bimester = bimester_raw[-2:] if bimester_raw else "Unknown"
        branch = parts[1].strip() or "Unknown"
        course_block = parts[2].strip()
        schedule_block = parts[3].strip()

        course_match = re.search(r"(SPS\s*\d+\s*-\s*\d+)", course_block)
        course = course_block
        prefix_extra = ""
        if course_match:
            course_raw = course_match.group(1)
            course = re.sub(r"\s*-\s*", "-", course_raw)
            course = re.sub(r"\s+", " ", course).strip()
            prefix_extra = course_block[course_match.end():].strip()
        else:
            course = course_block.strip() or "Unknown"

        prefix_extra = re.sub(r"\s+", " ", prefix_extra).strip()
        schedule_parts = []
        if prefix_extra:
            schedule_parts.append(prefix_extra.replace(" ", ""))
        if schedule_block:
            schedule_parts.append(schedule_block)
        schedule = "_".join(schedule_parts) if schedule_parts else "Unknown"

        return {
            "bimester": bimester or "Unknown",
            "branch": branch or "Unknown",
            "course": course or "Unknown",
            "schedule": schedule or "Unknown",
        }

    except Exception:
        return {
            "bimester": "Unknown",
            "branch": "Unknown",
            "course": "Unknown",
            "schedule": "Unknown",
        }


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
    """Read Surpass exports from either the original download or Google Sheets.

    We try HTML, Excel, and CSV loaders in succession so both the native
    Surpass export (often an HTML table inside an `.xls`) and Google Sheets
    re-exports are supported automatically. Column names are normalised to
    strings, stripped of whitespace, and canonicalised (e.g., ``Result (Date
    Submitted)`` → ``Result (DateSubmitted)``) to match downstream logic.
    """

    def _normalize_columns(df):
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                " ".join(str(part) for part in col if str(part).strip())
                for col in df.columns
            ]
        else:
            df.columns = df.columns.astype(str)
        df.columns = df.columns.str.strip()
        return df

    def _canonicalize_columns(df):
        """Rename common Surpass headers to their canonical spelling.

        Google Sheets downloads may inject spaces within parentheses or change
        casing; this helper aligns those variants so downstream lookups remain
        stable. We intentionally match loosely on any header that contains the
        tokens "result", "date", and "submit" to catch variants like
        "Result ( Date Submitted )", "Result - date submitted", or
        "result_dateSubmitted".
        """

        def key(name: str) -> str:
            return re.sub(r"\W+", "", name).lower()

        def is_result_column(name: str) -> bool:
            norm = key(name)
            return "result" in norm and "date" in norm and "submit" in norm

        rename_map = {}
        for col in df.columns:
            if is_result_column(col):
                rename_map[col] = "Result (DateSubmitted)"

        if rename_map:
            df = df.rename(columns=rename_map)
        return df

    errors = []

    def _attempt(loader_name, loader_fn):
        try:
            df_local = loader_fn()
            df_local = _normalize_columns(df_local)
            return _canonicalize_columns(df_local)
        except Exception as exc:  # pragma: no cover - best-effort fallbacks
            errors.append(f"{loader_name}: {exc}")
            return None

    # 1) HTML first (many native Surpass `.xls` files are HTML tables)
    html_tables = _attempt("HTML", lambda: pd.read_html(file_path)[0])
    if html_tables is not None:
        return html_tables

    # 2) Excel engines (covers genuine Excel files, including Sheets exports)
    for engine in (None, "openpyxl", "xlrd"):
        excel_df = _attempt(f"Excel ({engine or 'auto'})", lambda eng=engine: pd.read_excel(file_path, engine=eng))
        if excel_df is not None:
            return excel_df

    # 3) CSV fallbacks (in case the file is saved as CSV or mislabelled)
    for csv_kwargs in ({}, {"sep": None, "engine": "python"}):
        csv_df = _attempt(f"CSV ({csv_kwargs or 'default'})", lambda opts=csv_kwargs: pd.read_csv(file_path, **opts))
        if csv_df is not None:
            return csv_df

    error_detail = "\n".join(errors)
    raise ValueError(
        "Could not read the file. Supported formats: Excel (.xls/.xlsx), CSV, or HTML. "
        f"Details:\n{error_detail}"
    )


def detect_report_type(df):
    if "Category" in df.columns:
        return "Test"
    elif "Content" in df.columns and "Lesson" in df.columns:
        return "Assignment"
    else:
        raise ValueError("Unknown report format. Make sure your Excel has the correct structure.")


def natural_unit_sort_key(unit_label):
    match = re.search(r"(\d+)", unit_label or "")
    return int(match.group(1)) if match else float("inf")


def infer_course_suffix(class_string):
    if not class_string:
        return None

    direct_match = re.search(r"SPS\s*\d\s*-\s*([123])\b", class_string)
    if direct_match:
        return direct_match.group(1)

    details = extract_class_info(class_string)
    course_label = details.get("course", "")
    if course_label:
        structured_match = re.search(r"SPS\s*\d\s*-\s*([123])\b", course_label)
        if structured_match:
            return structured_match.group(1)

    return None


def derive_course_name(df):
    if 'Class' not in df.columns:
        return ""
    class_series = df['Class'].dropna()
    if class_series.empty:
        return ""
    return str(class_series.iloc[0]).strip()


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
    if 'Result (DateSubmitted)' not in df.columns:
        raise ValueError(
            "The file is missing the 'Result (DateSubmitted)' column. If you downloaded the file "
            "from Google Sheets, export it as Excel/CSV without altering the header names and try again."
        )
    df['Student'] = df['Student'].apply(clean_student_name)
    df[['Earned', 'Total']] = df['Result (DateSubmitted)'].apply(lambda x: pd.Series(extract_score(x)))
    all_students = sorted(df['Student'].unique())
    skills = ['Listening', 'Grammar', 'Vocabulary', 'Reading']
    course_source = df['Class'].iloc[0].strip() if 'Class' in df.columns else ""
    course_info = extract_class_info(course_source)
    course = course_info["course"]

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
            if matching.empty:
                continue
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
    if 'Result (DateSubmitted)' not in df.columns:
        raise ValueError(
            "The file is missing the 'Result (DateSubmitted)' column. If you downloaded the file "
            "from Google Sheets, export it as Excel/CSV without altering the header names and try again."
        )
    df['Student'] = df['Student'].apply(clean_student_name)
    df['Score'] = df['Result (DateSubmitted)'].apply(lambda x: extract_score(x)[0])
    df_all = df[['Student', 'Score', 'Unit', 'Class']]

    raw_course = df['Class'].dropna().astype(str).str.strip().iloc[0] if 'Class' in df.columns and not df['Class'].dropna().empty else ""
    course_info = extract_class_info(raw_course)
    course = course_info["course"]

    all_students = sorted(df_all['Student'].unique())
    selected_units_sorted = sorted(selected_units, key=natural_unit_sort_key)
    df_filtered = df_all[df_all['Unit'].isin(selected_units_sorted)] if selected_units_sorted else df_all.iloc[0:0]

    grouped_scores = {}
    if not df_filtered.empty:
        grouped_scores = df_filtered.groupby(['Student', 'Unit'])['Score'].sum().to_dict()

    unit_columns = [(unit, f"SCORE {unit.upper()}") for unit in selected_units_sorted]

    results = []
    for student in all_students:
        row = {STUDENT_COLUMN: student}
        for unit, column_label in unit_columns:
            row[column_label] = int(grouped_scores.get((student, unit), 0))
        results.append(row)

    df_results = pd.DataFrame(results)
    if df_results.empty:
        df_results = pd.DataFrame(columns=[STUDENT_COLUMN] + [label for _, label in unit_columns])

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


def export_pdf(df, path, teacher, course_display, class_string, selected_units_label):
    pdf = FPDF(orientation='L')
    pdf.add_page()
    font_family = try_configure_roboto(pdf)
    pdf.set_font(font_family, style='B', size=12)
    class_info = extract_class_info(class_string or course_display)

    left_labels = ["Teacher:", "Course:", "Bimester:"]
    left_values = [teacher, course_display, class_info["bimester"]]
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
        self.full_df = None
        self.assignment_styler = None
        self.table_canvas = None
        self.table_inner = None
        self.table_rows = []
        self.selected_row_indices = set()
        self.selection_anchor = None
        self.selection_active = False
        self.units_sorted = []
        self.quick_groups = {"eval1": [], "eval2": []}
        self.course_hint = ""
        self.quick_selection_var = tk.StringVar(value="eval1")
        self.quick_buttons = []
        self.selection_mode = tk.StringVar(value="quick")
        self.unit_vars = {}
        self.unit_checkbuttons = {}
        self.current_units_label = ""
        self.download_btn = None

        style_method = getattr(self, "set_style", None)
        if callable(style_method):
            style_method()

        top_frame = tk.Frame(root, bg="#F7F9FC")
        top_frame.pack(pady=6, fill="x")

        teacher_column = tk.Frame(top_frame, bg="#F7F9FC")
        teacher_column.pack(side="left", anchor="n")

        tk.Label(teacher_column, text="Teacher:", bg="#F7F9FC", font=("Segoe UI", 12)).pack(side="left")
        self.teacher_entry = tk.Entry(teacher_column, width=26, font=("Segoe UI", 12))
        self.teacher_entry.pack(side="left", padx=(6, 0))

        self.open_file_btn = ttk.Button(top_frame, text="Open File", style="Rounded.TButton", command=self.open_file)
        self.open_file_btn.pack(side="left", anchor="n", padx=(12, 10))

        self.selection_container = tk.Frame(top_frame, bg="#F7F9FC")
        self.selection_container.pack(side="left", anchor="n")

        header_frame = tk.Frame(self.selection_container, bg="#F7F9FC")
        header_frame.pack(fill="x")

        self.mode_buttons = []
        for value, text in (("quick", "Quick"), ("advanced", "Advanced")):
            btn = tk.Radiobutton(
                header_frame,
                text=text,
                variable=self.selection_mode,
                value=value,
                command=self.on_selection_mode_change,
                bg="#F7F9FC",
                activebackground="#F7F9FC",
                font=("Segoe UI", 10, "bold") if value == "quick" else ("Segoe UI", 10),
                indicatoron=False,
                relief="ridge",
                bd=1,
                padx=8,
                pady=2,
            )
            btn.pack(side="left", padx=(0, 6))
            self.mode_buttons.append(btn)

        self.selection_body = tk.Frame(self.selection_container, bd=2, relief="groove", bg="#FFFFFF")
        self.selection_body.pack(fill="both", expand=True, pady=(2, 0))

        self.quick_content = tk.Frame(self.selection_body, bg="#F7F9FC")
        self.advanced_content = tk.Frame(self.selection_body, bg="#F7F9FC")

        quick_button_specs = [("eval1", "1st Eval"), ("eval2", "2nd Eval")]
        for value, text in quick_button_specs:
            btn = tk.Radiobutton(
                self.quick_content,
                text=text,
                variable=self.quick_selection_var,
                value=value,
                command=self.on_quick_selection,
                bg="#F7F9FC",
                activebackground="#F7F9FC",
                font=("Segoe UI", 11),
                anchor="w"
            )
            btn.pack(anchor="w", padx=6, pady=1)
            btn.configure(highlightthickness=0)
            self.quick_buttons.append(btn)

        pdf_style = ttk.Style()
        pdf_style.configure("Pdf.TButton", font=("Segoe UI", 11, "bold"), padding=(18, 12),
                            background=DEFAULT_HEADER_COLOR, foreground="#ffffff")
        pdf_style.map("Pdf.TButton",
                      background=[("active", "#354d67"), ("pressed", "#1f3041")],
                      relief=[("pressed", "sunken")])

        self.download_btn = ttk.Button(top_frame, text="Export PDF", style="Pdf.TButton", command=self.save_pdf)
        self.download_btn.pack(side="left", anchor="n", padx=(12, 0))
        self.download_btn.config(state="disabled")

        # Style for the CSV export button (slightly different color)
        csv_style = ttk.Style()
        csv_style.configure(
            "Csv.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(18, 12),
            background="#1B9E77",     # green-ish tone
            foreground="#ffffff"
        )
        csv_style.map(
            "Csv.TButton",
            background=[("active", "#147558"), ("pressed", "#0E4C3A")],
            relief=[("pressed", "sunken")]
        )

        # New button: Export CSV
        self.csv_btn = ttk.Button(
            top_frame,
            text="Export CSV",
            style="Csv.TButton",
            command=self.save_csv
        )
        self.csv_btn.pack(side="left", anchor="n", padx=(8, 0))
        self.csv_btn.config(state="disabled")

        self.render_selection_body()
        self.update_mode_button_styles()
        if self.mode_buttons:
            self.mode_buttons[1].configure(state="disabled")

        self.teacher_entry.bind("<KeyRelease>", lambda e: self.check_pdf_button())
        self.set_quick_controls_enabled(False)

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

    def check_pdf_button(self):
        has_data = self.df_filtered is not None
        has_teacher = bool(self.teacher_entry.get().strip())

        # Both exports require data + teacher name
        if has_data and has_teacher:
            self.download_btn.config(state="normal")
            self.csv_btn.config(state="normal")
        else:
            self.download_btn.config(state="disabled")
            self.csv_btn.config(state="disabled")

    def open_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Excel/CSV/HTML", "*.xls *.xlsx *.csv *.html"),
                ("All files", "*.*"),
            ]
        )
        if not file_path:
            return

        try:
            df = read_flexible_excel(file_path)
            self.report_type = detect_report_type(df)
            self.full_df = df
            self.course_hint = derive_course_name(df)

            class_series = pd.Series(dtype=str)
            if 'Class' in df.columns:
                class_series = df['Class'].dropna().astype(str).str.strip()

            class_infos = [extract_class_info(value) for value in class_series if value]
            required_keys = ("bimester", "branch", "course", "schedule")
            missing_class_info = (
                not class_infos
                or all(any(info.get(key) == "Unknown" for key in required_keys) for info in class_infos)
            )

            if missing_class_info:
                messagebox.showwarning(
                    "Incomplete Class Information",
                    (
                        "It seems that this course wasn't created with the official code provided by the TC. "
                        "Please make sure your course was created with the correct code or copy the code to the "
                        "Class column of the downloaded Surpass excel file and retry. If not solved, the report "
                        "will not have the necessary information."
                    ),
                )
            units = [str(unit).strip() for unit in df['Unit'].dropna().unique() if str(unit).strip().startswith("Unit")]
            units = sorted(set(units), key=natural_unit_sort_key)

            if not units:
                messagebox.showwarning("No Units Found", "No valid units were found in this Excel file.")
                return

            self.units_sorted = units
            self.configure_selection_controls()
            self.update_table_by_selection()

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def configure_selection_controls(self):
        self.quick_groups = self.build_quick_groups(self.units_sorted, self.course_hint)
        has_quick = bool(self.quick_groups.get("eval1") or self.quick_groups.get("eval2"))
        if self.quick_groups.get("eval1"):
            self.quick_selection_var.set("eval1")
        elif self.quick_groups.get("eval2"):
            self.quick_selection_var.set("eval2")
        else:
            self.quick_selection_var.set("eval1")
        self.refresh_quick_buttons()
        self.build_advanced_checkboxes(self.units_sorted)
        self.selection_mode.set("quick" if has_quick else "advanced")
        self.set_quick_controls_enabled(has_quick)
        self.render_selection_body()
        self.update_mode_button_styles()

    def build_quick_groups(self, units, course_name):
        base_groups = self.build_even_split_groups(units)
        if not course_name:
            return base_groups

        suffix = infer_course_suffix(course_name)
        if not suffix:
            return base_groups
        course_mapping = {
            "1": {"eval1": ["Unit 1"], "eval2": ["Unit 2", "Unit 3"]},
            "2": {"eval1": ["Unit 4", "Unit 5"], "eval2": ["Unit 6", "Unit 7"]},
            "3": {"eval1": ["Unit 8"], "eval2": ["Unit 9", "Unit 10"]},
        }

        config = course_mapping.get(suffix)
        if not config:
            return base_groups

        available = set(units)
        mapped_groups = {}
        for key in ("eval1", "eval2"):
            required_units = config.get(key, [])
            if required_units and all(unit in available for unit in required_units):
                mapped_groups[key] = list(required_units)
            else:
                mapped_groups[key] = []

        if not mapped_groups["eval1"] and not mapped_groups["eval2"]:
            return base_groups
        return mapped_groups

    def build_even_split_groups(self, units):
        if not units:
            return {"eval1": [], "eval2": []}
        midpoint = max(1, (len(units) + 1) // 2)
        first_group = units[:midpoint]
        second_group = units[midpoint:]
        return {"eval1": first_group, "eval2": second_group}

    def refresh_quick_buttons(self):
        groups = self.quick_groups
        for btn, key in zip(self.quick_buttons, ("eval1", "eval2")):
            has_units = bool(groups.get(key))
            btn.configure(state="normal" if has_units else "disabled")

    def set_quick_controls_enabled(self, enabled):
        if enabled:
            self.refresh_quick_buttons()
        else:
            for btn in self.quick_buttons:
                btn.configure(state="disabled")
        quick_header_state = "normal" if enabled else "disabled"
        if self.mode_buttons:
            self.mode_buttons[0].configure(state=quick_header_state)

    def build_advanced_checkboxes(self, units):
        for child in self.advanced_content.winfo_children():
            child.destroy()
        self.unit_vars = {}
        self.unit_checkbuttons = {}
        if not units:
            tk.Label(self.advanced_content, text="No units available", bg="#F7F9FC", font=("Segoe UI", 11)).pack(anchor="w", padx=8, pady=2)
            return
        for index, unit in enumerate(units):
            var = tk.BooleanVar(value=index == 0)
            chk = tk.Checkbutton(
                self.advanced_content,
                text=unit,
                variable=var,
                command=lambda u=unit: self.on_unit_checkbox_change(u),
                bg="#F7F9FC",
                activebackground="#F7F9FC",
                font=("Segoe UI", 11),
                anchor="w"
            )
            chk.pack(anchor="w", padx=8, pady=2)
            chk.configure(highlightthickness=0)
            self.unit_vars[unit] = var
            self.unit_checkbuttons[unit] = chk
        has_advanced = bool(self.unit_vars)
        self.mode_buttons[1].configure(state="normal" if has_advanced else "disabled")
        if self.selection_mode.get() == "advanced" and not has_advanced:
            self.selection_mode.set("quick")
        self.render_selection_body()
        self.update_mode_button_styles()

    def is_advanced_mode(self):
        return self.selection_mode.get() == "advanced"

    def on_quick_selection(self):
        if self.is_advanced_mode():
            return
        self.update_table_by_selection()

    def on_selection_mode_change(self):
        advanced = self.is_advanced_mode()
        if advanced and not self.unit_vars:
            self.selection_mode.set("quick")
            advanced = False
        self.render_selection_body()
        self.update_mode_button_styles()
        if advanced and self.unit_vars and not any(var.get() for var in self.unit_vars.values()):
            first_unit = next(iter(self.unit_vars))
            self.unit_vars[first_unit].set(True)
        self.update_table_by_selection()

    def on_unit_checkbox_change(self, unit):
        if not self.is_advanced_mode():
            return
        if unit not in self.unit_vars:
            return
        var = self.unit_vars[unit]
        if not var.get():
            if sum(v.get() for v in self.unit_vars.values()) == 0:
                var.set(True)
                return
        self.update_table_by_selection()

    def get_current_units(self):
        if self.is_advanced_mode():
            return [unit for unit, var in self.unit_vars.items() if var.get()]
        return self.quick_groups.get(self.quick_selection_var.get(), [])

    def get_selection_label(self, units):
        if self.is_advanced_mode():
            unit_list = ", ".join(units) if units else "None"
            return f"Advanced ({unit_list})"
        label = "1st Eval" if self.quick_selection_var.get() == "eval1" else "2nd Eval"
        if units:
            return f"{label} ({', '.join(units)})"
        return label

    def update_table_by_selection(self):
        if self.full_df is None or self.report_type is None:
            return
        selected_units = self.get_current_units()
        if not selected_units and not self.is_advanced_mode():
            for key in ("eval1", "eval2"):
                fallback_units = self.quick_groups.get(key, [])
                if fallback_units:
                    self.quick_selection_var.set(key)
                    selected_units = fallback_units
                    break
        self.current_units_label = self.get_selection_label(selected_units)
        if self.full_df is not None:
            self.course_hint = derive_course_name(self.full_df)
        if self.report_type == "Assignment":
            self.df_internal, self.course, self.assignment_styler = process_assignment(self.full_df, selected_units, skip_read=True)
        else:
            self.df_internal, self.course = process_test(self.full_df, selected_units, skip_read=True)
            self.assignment_styler = None
        self.df_filtered = self.df_internal.copy()
        self.df = self.df_filtered
        self.show_table(self.df_filtered)
        self.check_pdf_button()

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
        unit_label = self.current_units_label or "Units"
        filename = f"{sanitize_filename(teacher)} - {sanitize_filename(course)} - {sanitize_filename(unit_label)}.pdf"
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=filename,
                                                filetypes=[("PDF files", "*.pdf")])
        if filepath:
            export_pdf(
                self.df_internal if self.df_internal is not None else self.df_filtered,
                filepath,
                teacher,
                course,
                self.course_hint,
                unit_label,
            )
            messagebox.showinfo("Export Complete", f"PDF saved to:\n{filepath}")

    def save_csv(self):
        """Export the current results table to a CSV file."""
        if self.df_filtered is None:
            messagebox.showerror("Error", "No data to export.")
            return

        # Use teacher name if available; otherwise fall back to a generic label
        teacher = self.teacher_entry.get().strip() or "Teacher"
        course = self.course if self.course else "Unknown"
        unit_label = self.current_units_label or "Units"

        filename = f"{sanitize_filename(teacher)} - {sanitize_filename(course)} - {sanitize_filename(unit_label)}.csv"

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=filename,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if not filepath:
            return

        try:
            # Use the same internal DataFrame used for PDF export
            df_to_export = self.df_internal if self.df_internal is not None else self.df_filtered

            # Export as UTF-8 with BOM so Excel opens it nicely (especially with accents)
            df_to_export.to_csv(filepath, index=False, encoding="utf-8-sig")

            messagebox.showinfo("Export Complete", f"CSV saved to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save CSV:\n{e}")

    def render_selection_body(self):
        self.quick_content.pack_forget()
        self.advanced_content.pack_forget()
        if self.is_advanced_mode():
            if not self.advanced_content.winfo_children():
                tk.Label(self.advanced_content, text="No units available", bg="#F7F9FC", font=("Segoe UI", 11)).pack(anchor="w", padx=8, pady=2)
            self.advanced_content.pack(fill="both", expand=True)
        else:
            self.quick_content.pack(fill="both", expand=True)

    def update_mode_button_styles(self):
        current = self.selection_mode.get()
        for btn in self.mode_buttons:
            value = btn.cget("value")
            font = ("Segoe UI", 10, "bold") if value == current else ("Segoe UI", 10)
            relief = "sunken" if value == current else "ridge"
            btn.configure(font=font, relief=relief)


if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()
