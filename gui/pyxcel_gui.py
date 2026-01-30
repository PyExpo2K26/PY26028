import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

final_output_file = None
config_values = {}


import data_cleaner
import merger
import sorter_filter
import vlookup

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

        
tk.Checkbutton(
    operations,
    text="VLOOKUP",
    variable=vlookup_var,
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
    
    # -------- VLOOKUP CONFIG --------
    if "vlookup" in selected_operations:
        section = tk.LabelFrame(
            config, text="VLOOKUP Settings",
            bg="#dbeeff", padx=10, pady=10
        )

    section.pack(padx=150, pady=10, fill="x")

    tk.Label(section, text="Lookup Column:", bg="#dbeeff").pack(anchor="w")
    lookup_col_entry = tk.Entry(section, width=30)
    lookup_col_entry.pack(anchor="w")

    tk.Label(section, text="Return Column:", bg="#dbeeff").pack(anchor="w", pady=(5, 0))
    return_col_entry = tk.Entry(section, width=30)
    return_col_entry.pack(anchor="w")

    # store entries for backend use
    config_values["lookup_col"] = lookup_col_entry
    config_values["return_col"] = return_col_entry



    section.pack(padx=150, pady=10, fill="x")


    section.pack(padx=150, pady=10, fill="x")

    tk.Label(section, text="Lookup Column:", bg="#dbeeff").pack(anchor="w")
    lookup_col_entry = tk.Entry(section, width=30)
    lookup_col_entry.pack(anchor="w")

    tk.Label(section, text="Return Column:", bg="#dbeeff").pack(anchor="w", pady=(5, 0))
    return_col_entry = tk.Entry(section, width=30)
    return_col_entry.pack(anchor="w")

        # store entries for backend use

    config_values["lookup_col"] = lookup_col_entry
    config_values["return_col"] = return_col_entry

    config_values["lookup_col"] = lookup_col_entry
    config_values["return_col"] = return_col_entry




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

        if "vlookup" in selected_operations:
            status_text.set("Applying VLOOKUP...")
            lookup_col = config_values["lookup_col"].get()
            return_col = config_values["return_col"].get()

        if not lookup_col or not return_col:
            raise ValueError("VLOOKUP requires both lookup and return columns.")

        vlookup_out = os.path.join(OUTPUT_DIR, "vlookup_output.xlsx")
        #-----------VLOOKUP ----------
        # Requires two files: base file + lookup table
        if len(current_files) < 2:
            raise ValueError("VLOOKUP requires two Excel files.")

        vlookup.vlookup_excel(
            main_file=current_files[0],
            lookup_file=current_files[1],
            output_file=vlookup_out,
            key_column=lookup_col,
            lookup_column=return_col
        )


        current_files = [vlookup_out]

        #-----------VLOOKUP ----------
        # Requires two files: base file + lookup table
        
        if "vlookup" in selected_operations:
            status_text.set("Applying VLOOKUP...")

            lookup_col = config_values["lookup_col"].get()
            return_col = config_values["return_col"].get()

            if not lookup_col or not return_col:
                raise ValueError("VLOOKUP requires both lookup and return columns.")

            vlookup_out = os.path.join(OUTPUT_DIR, "vlookup_output.xlsx")

            if len(current_files) < 2:
                raise ValueError("VLOOKUP requires two Excel files.")

            vlookup.vlookup_excel(
                main_file=current_files[0],
                lookup_file=current_files[1],
                output_file=vlookup_out,
                key_column=lookup_col,
                lookup_column=return_col
            )

            current_files = [vlookup_out]
        



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



# ---------------- START ----------------
show_frame(welcome)
root.mainloop()
