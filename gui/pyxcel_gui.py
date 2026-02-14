import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os
import pandas as pd

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

final_output_file = None
config_values = {}


from backend import data_cleaner
from backend import merger
from backend import sorter_filter
from backend import visualizer 
from backend import calculation

import pandas as pd

# ---------------- WINDOW ----------------
root = tk.Tk()
root.title("PyXcel – Excel Automation Tool")
root.geometry("800x450")
root.resizable(True, True)

# ---------------- CONTAINER ----------------
container = tk.Frame(root)
container.pack(fill="both", expand=True)

def show_frame(frame):
    frame.tkraise()

# ---------------- DATA ----------------
selected_files = []
selected_operations = {}

# ---------------- WELCOME SCREEN ----------------
welcome = tk.Frame(container, bg="#dbeeff")
welcome.place(relwidth=1, relheight=1)

tk.Label(
    welcome, text="Welcome to PyXcel",
    font=("Segoe UI", 28, "bold"),
    bg="#dbeeff", fg="#1f3c88"
).pack(pady=(80, 10))

tk.Label(
    welcome,
    text="Excel Automation Tool\nAutomate your Excel tasks with ease!",
    font=("Segoe UI", 14),
    bg="#dbeeff", fg="#333333",
    justify="center"
).pack(pady=10)

tk.Button(
    welcome, text="Start",
    font=("Segoe UI", 13, "bold"),
    bg="#2e8b57", fg="white",
    width=16, height=2,
    relief="flat",
    command=lambda: show_frame(upload)
).pack(pady=30)

# ---------------- UPLOAD SCREEN ----------------
upload = tk.Frame(container, bg="#dbeeff")
upload.place(relwidth=1, relheight=1)

tk.Label(
    upload, text="Upload Your Excel Files",
    font=("Segoe UI", 22, "bold"),
    bg="#dbeeff", fg="#1f3c88"
).pack(pady=15)

info_label = tk.Label(upload, text="No files selected",
                      bg="#dbeeff", fg="#444444")
info_label.pack()

# ---- LISTBOX AREA ----
list_frame = tk.Frame(upload)
list_frame.pack(pady=10)

scrollbar = tk.Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")

file_listbox = tk.Listbox(
    list_frame,
    width=55,
    height=6,
    yscrollcommand=scrollbar.set
)
file_listbox.pack(side="left")
scrollbar.config(command=file_listbox.yview)

def update_info():
    count = len(selected_files)
    info_label.config(
        text="No files selected" if count == 0 else f"{count} file(s) selected"
    )

def browse_files():
    files = filedialog.askopenfilenames(
        title="Select Excel Files",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    for f in files:
        if f not in selected_files:
            selected_files.append(f)
            file_listbox.insert("end", f.split("/")[-1])
    update_info()

tk.Button(upload, text="Browse Files",
          width=15, command=browse_files).pack(pady=5)

def remove_selected():
    selected = list(file_listbox.curselection())
    for i in reversed(selected):
        selected_files.pop(i)
        file_listbox.delete(i)
    update_info()

def clear_all():
    selected_files.clear()
    file_listbox.delete(0, "end")
    update_info()

btn_frame = tk.Frame(upload, bg="#dbeeff")
btn_frame.pack(pady=5)

tk.Button(btn_frame, text="Remove Selected",
          width=15, command=remove_selected).grid(row=0, column=0, padx=5)

tk.Button(btn_frame, text="Clear All",
          width=12, command=clear_all).grid(row=0, column=1, padx=5)

def next_clicked():
    if not selected_files:
        messagebox.showwarning(
            "No File Selected",
            "Please select at least one Excel file to continue."
        )
    else:
        print("Selected files:")
        for f in selected_files:
            print(f)
        show_frame(operations)


tk.Button(upload, text="Next",
          width=12, command=next_clicked).pack(pady=15)

tk.Button(upload, text="Exit",
          width=12, command=root.quit).pack()

# ---------------- OPERATIONS SCREEN ----------------
operations = tk.Frame(container, bg="#dbeeff")
operations.place(relwidth=1, relheight=1)

tk.Label(
    operations, text="Select Operations",
    font=("Segoe UI", 22, "bold"),
    bg="#dbeeff", fg="#1f3c88"
).pack(pady=20)

tk.Label(
    operations,
    text="Choose the operations you want to apply:",
    bg="#dbeeff", fg="#333333"
).pack(pady=5)

vlookup_var = tk.BooleanVar()




# ---- CHECKBOX VARIABLES ----
clean_var = tk.BooleanVar()
merge_var = tk.BooleanVar()
sort_filter_var = tk.BooleanVar()

tk.Checkbutton(
    operations, text="Clean Data",
    variable=clean_var,
    bg="#dbeeff"
).pack(anchor="w", padx=250, pady=5)

tk.Checkbutton(
    operations, text="Merge Files",
    variable=merge_var,
    bg="#dbeeff"
).pack(anchor="w", padx=250, pady=5)

tk.Checkbutton(
    operations, text="Sort / Filter Data",
    variable=sort_filter_var,
    bg="#dbeeff"
).pack(anchor="w", padx=250, pady=5)

def next_from_operations():
    selected_operations.clear()

    if clean_var.get():
        selected_operations["clean"] = True
    if merge_var.get():
        selected_operations["merge"] = True
    if sort_filter_var.get():
        selected_operations["sort_filter"] = True
    if vlookup_var.get():
        selected_operations["vlookup"] = True

    if not selected_operations:
        messagebox.showwarning(
            "No Operation Selected",
            "Please select at least one operation to continue."
        )
    else:
        print("Selected operations:", selected_operations)
        # Next screen will be configuration
        # show_frame(config)
        build_config_screen()
        show_frame(config)



tk.Button(
    operations, text="Next",
    width=12, command=next_from_operations
).pack(pady=20)

tk.Button(
    operations, text="Back",
    width=12, command=lambda: show_frame(upload)
).pack()

# ---------------- CONFIGURATION SCREEN ----------------
config = tk.Frame(container, bg="#dbeeff")
config.place(relwidth=1, relheight=1)

def build_config_screen():
    # Clear previous widgets
    for widget in config.winfo_children():
        widget.destroy()

    tk.Label(
        config, text="Configure Selected Operations",
        font=("Segoe UI", 22, "bold"),
        bg="#dbeeff", fg="#1f3c88"
    ).pack(pady=15)

    tk.Label(
        config,
        text="Set parameters for the operations you selected.",
        bg="#dbeeff", fg="#333333"
    ).pack(pady=5)

    # -------- CLEAN DATA CONFIG --------
    if "clean" in selected_operations:
        section = tk.LabelFrame(
            config, text="Clean Data Settings",
            bg="#dbeeff", padx=10, pady=10
        )
        section.pack(padx=150, pady=10, fill="x")

        tk.Label(section, text="Remove duplicate rows:",
                 bg="#dbeeff").pack(anchor="w")
        tk.Checkbutton(section, text="Enable",
                       bg="#dbeeff").pack(anchor="w")

        tk.Label(section, text="Handle empty cells:",
                 bg="#dbeeff").pack(anchor="w", pady=(10, 0))
        tk.Entry(section, width=30).pack(anchor="w")

    # -------- MERGE FILES CONFIG --------
    if "merge" in selected_operations:
        section = tk.LabelFrame(
            config, text="Merge Files Settings",
            bg="#dbeeff", padx=10, pady=10
        )
        section.pack(padx=150, pady=10, fill="x")

        tk.Label(section, text="Merge type:",
                 bg="#dbeeff").pack(anchor="w")
        tk.Entry(section, width=30).pack(anchor="w")

    # -------- SORT / FILTER CONFIG --------
    if "sort_filter" in selected_operations:
        section = tk.LabelFrame(
            config, text="Sort / Filter Settings",
            bg="#dbeeff", padx=10, pady=10
        )
        section.pack(padx=150, pady=10, fill="x")

        tk.Label(section, text="Column name:",
                 bg="#dbeeff").pack(anchor="w")
        tk.Entry(section, width=30).pack(anchor="w")

        tk.Label(section, text="Condition (e.g. > 50):",
                 bg="#dbeeff").pack(anchor="w", pady=(10, 0))
        tk.Entry(section, width=30).pack(anchor="w")
    
    

 

    # -------- NAV BUTTONS --------
    nav = tk.Frame(config, bg="#dbeeff")
    nav.pack(pady=20)

    tk.Button(
        nav, text="Back",
        width=12,
        command=lambda: show_frame(operations)
    ).grid(row=0, column=0, padx=10)

    tk.Button(
        nav, text="Next",
        width=12,
        command=start_processing
    ).grid(row=0, column=1, padx=10)


# ---------------- PROCESSING SCREEN ----------------
processing = tk.Frame(container, bg="#dbeeff")
processing.place(relwidth=1, relheight=1)

progress_value = tk.IntVar()
status_text = tk.StringVar()

def start_processing():
    progress_value.set(0)
    status_text.set("Initializing...")

    show_frame(processing)

    root.after(500, run_backend)
    process_steps(0)


def process_steps(step):
    steps = [
        "Loading files...",
        "Applying selected operations...",
        "Optimizing data...",
        "Finalizing output..."
    ]

    if step < len(steps):
        status_text.set(steps[step])
        progress_value.set((step + 1) * 25)
        root.after(800, lambda: process_steps(step + 1))
    else:
        status_text.set("Processing completed successfully!")
        root.after(800, lambda: show_frame(done))

tk.Label(
    processing,
    text="Processing",
    font=("Segoe UI", 22, "bold"),
    bg="#dbeeff", fg="#1f3c88"
).pack(pady=20)

tk.Label(
    processing,
    textvariable=status_text,
    font=("Segoe UI", 12),
    bg="#dbeeff", fg="#333333"
).pack(pady=10)

ttk.Progressbar(
    processing,
    length=400,
    variable=progress_value,
    maximum=100
).pack(pady=20)

# ---------------- VISUALIZATION SCREEN ----------------
visualize = tk.Frame(container, bg="#dbeeff")
visualize.place(relwidth=1, relheight=1)

# ---------------- COMPLETION SCREEN ----------------
done = tk.Frame(container, bg="#dbeeff")
done.place(relwidth=1, relheight=1)

tk.Label(
    done,
    text="Processing Complete!",
    font=("Segoe UI", 24, "bold"),
    bg="#dbeeff",
    fg="#1f3c88"
).pack(pady=(80, 10))

tk.Label(
    done,
    text="Your Excel file has been processed successfully.",
    font=("Segoe UI", 13),
    bg="#dbeeff",
    fg="#333333"
).pack(pady=10)



def download_file():
    if not final_output_file:
        messagebox.showwarning(
            "No File",
            "No processed file available to download."
        )
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Save Processed File As"
    )

    if save_path:
        shutil.copy(final_output_file, save_path)
        messagebox.showinfo(
            "Download Complete",
            "File downloaded successfully!"
        )


tk.Button(
    done,
    text="Download File",
    font=("Segoe UI", 12, "bold"),
    bg="#2e8b57",
    fg="white",
    width=18,
    height=2,
    relief="flat",
    command=download_file
).pack(pady=25)

tk.Button(
    done,
    text="Start New Process",
    width=18,
    command=lambda: show_frame(welcome)
).pack(pady=5)

tk.Button(
    done,
    text="Exit",
    width=18,
    command=root.quit
).pack(pady=5)

tk.Button(
    done,
    text="Visualize Data",
    width=18,
    command=lambda: (populate_columns(), show_frame(visualize))
).pack(pady=5)



def run_backend():
    try:
        print("Backend processing started")

        current_files = selected_files.copy()

        # ---------- CLEAN ----------
        if "clean" in selected_operations:
            cleaned = []
            for file in current_files:
                out = os.path.join(
                    OUTPUT_DIR, f"cleaned_{os.path.basename(file)}"
                )
                data_cleaner.clean_excel(file, out)
                cleaned.append(out)

            current_files = cleaned
            print("Cleaning done")

        # ---------- MERGE ----------
        if "merge" in selected_operations:
            if len(current_files) < 2:
                raise ValueError("Merge requires at least 2 files")

            merged_out = os.path.join(OUTPUT_DIR, "merged.xlsx")

            merger.merge_excel(
                current_files[0],
                current_files[1],
                merged_out,
                common_column="id",   # TEMP default
                how="inner"
            )

            current_files = [merged_out]
            print("Merging done")

       
       



        # ---------- SORT / FILTER ----------
        if "sort_filter" in selected_operations:
            final_out = os.path.join(OUTPUT_DIR, "final.xlsx")

            sorter_filter.sort_and_filter(
                input_file=current_files[0],
                output_file=final_out,
                sort_column="name",   # TEMP default
                city_filter=None
            )

            current_files = [final_out]
            print("Sorting/filtering done")

        print("Backend processing finished")

    except Exception as e:
        messagebox.showerror("Processing Error", str(e))
    global final_output_file
    final_output_file = current_files[0]

tk.Label(
    visualize, text="Data Visualization",
    font=("Segoe UI", 22, "bold"),
    bg="#dbeeff", fg="#1f3c88"
).pack(pady=20)

chart_type = tk.StringVar(value="Histogram")
col1_var = tk.StringVar()
col2_var = tk.StringVar()

tk.Label(visualize, text="Chart Type", bg="#dbeeff").pack()
chart_menu = ttk.Combobox(
    visualize,
    textvariable=chart_type,
    values=["Histogram", "Pie Chart", "Scatter Plot", "Correlation Heatmap"],
    state="readonly",
    width=30
)
chart_menu.pack(pady=5)

tk.Label(visualize, text="Column 1", bg="#dbeeff").pack()
col1_menu = ttk.Combobox(visualize, textvariable=col1_var, width=30)
col1_menu.pack(pady=5)

tk.Label(visualize, text="Column 2 (Scatter only)", bg="#dbeeff").pack()
col2_menu = ttk.Combobox(visualize, textvariable=col2_var, width=30)
col2_menu.pack(pady=5)

def generate_chart():
    try:
        if not final_output_file:
            messagebox.showerror("Error", "No processed file available. Please complete processing first.")
            return
        
        chart_selected = chart_type.get()
        col1_selected = col1_var.get()
        col2_selected = col2_var.get()
        
        if not col1_selected:
            messagebox.showerror("Error", "Please select Column 1")
            return
        
        if chart_selected == "Scatter Plot" and not col2_selected:
            messagebox.showerror("Error", "Scatter Plot requires Column 2")
            return
        
        df, numeric_cols, categorical_cols = visualizer.get_data(final_output_file)
        
        if chart_selected == "Histogram":
            visualizer.show_histogram(df, col1_selected)
        elif chart_selected == "Pie Chart":
            visualizer.show_pie_chart(df, col1_selected)
        elif chart_selected == "Scatter Plot":
            visualizer.show_scatter(df, col1_selected, col2_selected)
        elif chart_selected == "Correlation Heatmap":
            visualizer.show_correlation_heatmap(df, numeric_cols)
        
        messagebox.showinfo("Success", "Chart generated successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate chart: {str(e)}")

def populate_columns():
    try:
        if not final_output_file:
            return
        
        df, numeric_cols, categorical_cols = visualizer.get_data(final_output_file)
        all_cols = df.columns.tolist()
        
        col1_menu['values'] = all_cols
        col2_menu['values'] = numeric_cols  # Scatter plot typically uses numeric columns
    except Exception as e:
        print(f"Error populating columns: {e}")

tk.Button(
    visualize,
    text="Next",
    font=("Segoe UI", 12, "bold"),
    bg="#2e8b57",
    fg="white",
    width=16,
    relief="flat",
    command=generate_chart
).pack(pady=20)

# Back button to return to the download (done) page
tk.Button(
    visualize,
    text="Back",
    font=("Segoe UI", 10),
    width=16,
    command=lambda: show_frame(done)
).pack(pady=5)

# ---------------- CALCULATIONS SCREEN ----------------
calculations = tk.Frame(container, bg="#dbeeff")
calculations.place(relwidth=1, relheight=1)

def build_calculations_screen():
    for w in calculations.winfo_children():
        w.destroy()

    tk.Label(
        calculations, text="Calculations",
        font=("Segoe UI", 22, "bold"),
        bg="#dbeeff", fg="#1f3c88"
    ).pack(pady=15)

    tk.Label(calculations, text="Select input file:", bg="#dbeeff").pack()
    calc_file_var = tk.StringVar()
    calc_file_menu = ttk.Combobox(calculations, textvariable=calc_file_var, values=[os.path.basename(f) for f in selected_files], width=50, state="readonly")
    calc_file_menu.pack(pady=5)

    tk.Label(calculations, text="Select columns (multi-select):", bg="#dbeeff").pack()
    cols_listbox = tk.Listbox(calculations, selectmode="multiple", width=50, height=6)
    cols_listbox.pack(pady=5)

    def load_columns(event=None):
        cols_listbox.delete(0, "end")
        sel = calc_file_var.get()
        if not sel:
            return
        # find full path
        full = next((f for f in selected_files if os.path.basename(f) == sel), None)
        if not full:
            return
        try:
            df = pd.read_excel(full)
            for i, c in enumerate(df.columns):
                cols_listbox.insert("end", f"{i}: {c}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")

    calc_file_menu.bind("<<ComboboxSelected>>", load_columns)

    tk.Label(calculations, text="Operation:", bg="#dbeeff").pack(pady=(10,0))
    op_var = tk.StringVar(value="add")
    op_menu = ttk.Combobox(calculations, textvariable=op_var, values=["add", "multiply", "subtract", "divide"], state="readonly", width=30)
    op_menu.pack(pady=5)

    tk.Label(calculations, text="New column name:", bg="#dbeeff").pack(pady=(10,0))
    new_col_entry = tk.Entry(calculations, width=40)
    new_col_entry.pack(pady=5)

    def get_selected_column_indices():
        sel = cols_listbox.curselection()
        return [int(cols_listbox.get(i).split(":", 1)[0]) for i in sel]

    def do_validate():
        selfile = calc_file_var.get()
        if not selfile:
            messagebox.showwarning("Input required", "Select an input file first.")
            return
        full = next((f for f in selected_files if os.path.basename(f) == selfile), None)
        if not full:
            messagebox.showerror("Error", "Selected file not found.")
            return
        try:
            df = pd.read_excel(full)
            idxs = get_selected_column_indices()
            calculation.validate_rule(df, idxs, op_var.get())
            messagebox.showinfo("Valid", "Validation passed.")
        except Exception as e:
            messagebox.showerror("Validation Error", str(e))

    def do_apply_in_memory():
        selfile = calc_file_var.get()
        if not selfile:
            messagebox.showwarning("Input required", "Select an input file first.")
            return
        full = next((f for f in selected_files if os.path.basename(f) == selfile), None)
        try:
            df = pd.read_excel(full)
            idxs = get_selected_column_indices()
            new_name = new_col_entry.get().strip()
            if not new_name:
                messagebox.showwarning("Input required", "Enter a new column name.")
                return
            df2 = calculation.apply_calculated_column(df, idxs, op_var.get(), new_name)
            out_path = os.path.join(OUTPUT_DIR, f"{os.path.splitext(os.path.basename(full))[0]}_calc.xlsx")
            df2.to_excel(out_path, index=False)
            global final_output_file
            final_output_file = out_path
            messagebox.showinfo("Done", f"Calculation applied and saved to {out_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def do_process_file():
        selfile = calc_file_var.get()
        if not selfile:
            messagebox.showwarning("Input required", "Select an input file first.")
            return
        full = next((f for f in selected_files if os.path.basename(f) == selfile), None)
        try:
            idxs = get_selected_column_indices()
            new_name = new_col_entry.get().strip()
            if not new_name:
                messagebox.showwarning("Input required", "Enter a new column name.")
                return
            out = calculation.process_excel_file(full, OUTPUT_DIR, idxs, op_var.get(), new_name)
            global final_output_file
            final_output_file = out
            messagebox.showinfo("Done", f"File processed and saved to {out}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    btn_frame = tk.Frame(calculations, bg="#dbeeff")
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="Validate Rule", width=15, command=do_validate).grid(row=0, column=0, padx=5)
    tk.Button(btn_frame, text="Apply (in-memory)", width=15, command=do_apply_in_memory).grid(row=0, column=1, padx=5)
    tk.Button(btn_frame, text="Process Excel File", width=18, command=do_process_file).grid(row=0, column=2, padx=5)

    tk.Button(calculations, text="Back", width=12, command=lambda: show_frame(operations)).pack(pady=8)

# add a button on operations screen to open calculations
tk.Button(
    operations,
    text="Calculations...",
    width=12,
    command=lambda: (build_calculations_screen(), show_frame(calculations))
).pack(pady=8)

# ---------------- START ----------------
show_frame(welcome)
root.mainloop()