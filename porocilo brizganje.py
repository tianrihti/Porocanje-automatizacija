import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
import io
import logging
import warnings
import json

class TextRedirector(io.TextIOBase):
    def __init__(self, text_widget):
        self.text_widget = text_widget
    def write(self, s):
        try:
            self.text_widget.insert(tk.END, s)
            self.text_widget.see(tk.END)
        except tk.TclError:
            pass
    def flush(self):
        pass

def run_automation_with_terminal(pregled_file, porocanje_file, plan_file, terminal_win, terminal_text):
    warnings.filterwarnings("ignore")
    sys.path.append(os.path.dirname(__file__))
    from automate_process import ExcelAutomation

    sys.stdout = TextRedirector(terminal_text)
    sys.stderr = TextRedirector(terminal_text)

    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logging.basicConfig(stream=sys.stdout, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    log_path = os.path.join(os.path.dirname(__file__), "excel_changes.log")
    with open(log_path, "w", encoding="utf-8") as f:
        pass

    automation = ExcelAutomation()
    try:
        automation.pregled_file = pregled_file
        automation.porocanje_file = porocanje_file
        automation.plan_file = plan_file

        automation._validate_files()
        automation.kill_excel_processes()

        pregled_data = automation.step1_copy_pregled_data()
        automation.step2_paste_to_porocanje(pregled_data)
        target_col = automation.step3_find_date_in_plan()
        plan_range_data = automation.step4_copy_plan_range(target_col)
        automation.step5_paste_to_brizganje(plan_range_data)
        automation.recalc_excel()
        automation.scan_brizganje_errors()
        saved_texts = automation.step6_analyze_brizganje()
        automation.step7_process_saved_texts(saved_texts)

        messagebox.showinfo("Uspeh", "Skripta je bila uspešno izvedena!")
    except Exception as e:
        messagebox.showerror("Napaka", f"Napaka med izvajanjem: {e}")
    finally:
        automation.kill_excel_processes()
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        terminal_win.destroy()

def select_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls*")])
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)

def start_script_with_terminal(pregled_entry, porocanje_entry, plan_entry, run_btn):
    pregled_file = pregled_entry.get()
    porocanje_file = porocanje_entry.get()
    plan_file = plan_entry.get()
    if not all([pregled_file, porocanje_file, plan_file]):
        messagebox.showwarning("Manjkajoče datoteke", "Izberite vse tri datoteke!")
        return

    run_btn.config(state="disabled")

    terminal_win = tk.Toplevel()
    terminal_win.title("Terminal Output")
    terminal_win.geometry("780x400")
    terminal_text = tk.Text(terminal_win, font=("Consolas", 10), bg="#222", fg="#eee")
    terminal_text.pack(fill="both", expand=True)
    terminal_text.insert(tk.END, "Skripta se izvaja...\n")

    def run_and_enable():
        run_automation_with_terminal(pregled_file, porocanje_file, plan_file, terminal_win, terminal_text)
        run_btn.config(state="normal")

    threading.Thread(target=run_and_enable).start()

# --- Modern GUI setup with tabs ---
root = tk.Tk()
root.title("Excel Avtomatizacija")
root.geometry("620x340")
style = ttk.Style(root)
style.theme_use("clam")
style.configure("TLabel", font=("Segoe UI", 12))
style.configure("TButton", font=("Segoe UI", 11), padding=6)
style.configure("TEntry", font=("Segoe UI", 11))
style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#2E8B57")
style.configure("Run.TButton", font=("Segoe UI", 13, "bold"), background="#90EE90")

bg_color = "#d9d9d9"  # This matches the default ttk Frame background

root.configure(bg=bg_color)
style.configure("TNotebook", background=bg_color, borderwidth=0)
style.configure("TNotebook.Tab", background=bg_color)

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=5, pady=5)

style.layout("TNotebook.Tab", [
    ('Notebook.tab', {
        'children': [
            ('Notebook.padding', {
                'children': [
                    ('Notebook.label', {'side': 'left', 'expand': 1})
                ]
            })
        ]
    })
])

# --- User Tab ---
user_frame = ttk.Frame(notebook, padding=25)
notebook.add(user_frame, text="Uporabnik")

ttk.Label(user_frame, text="Excel Avtomatizacija", style="Title.TLabel").grid(row=0, column=0, columnspan=3, pady=(0, 18))

ttk.Label(user_frame, text="GoSoft data excel:").grid(row=1, column=0, sticky="e", pady=7)
pregled_entry = ttk.Entry(user_frame, width=40)
pregled_entry.grid(row=1, column=1, pady=7)
ttk.Button(user_frame, text="Izberi...", command=lambda: select_file(pregled_entry)).grid(row=1, column=2, padx=8, pady=7)

ttk.Label(user_frame, text="Poročanje proizvodnje excel:").grid(row=2, column=0, sticky="e", pady=7)
porocanje_entry = ttk.Entry(user_frame, width=40)
porocanje_entry.grid(row=2, column=1, pady=7)
ttk.Button(user_frame, text="Izberi...", command=lambda: select_file(porocanje_entry)).grid(row=2, column=2, padx=8, pady=7)

ttk.Label(user_frame, text="Plan brizganja mesečni excel:").grid(row=3, column=0, sticky="e", pady=7)
plan_entry = ttk.Entry(user_frame, width=40)
plan_entry.grid(row=3, column=1, pady=7)
ttk.Button(user_frame, text="Izberi...", command=lambda: select_file(plan_entry)).grid(row=3, column=2, padx=8, pady=7)

run_btn = ttk.Button(
    user_frame,
    text="Zaženi skripto",
    style="Run.TButton",
    command=lambda: start_script_with_terminal(pregled_entry, porocanje_entry, plan_entry, run_btn)
)
run_btn.grid(row=5, column=0, columnspan=3, pady=(10, 0))

# --- Admin Tab ---
admin_frame = ttk.Frame(notebook, padding=25)
notebook.add(admin_frame, text="Admin")

# Legend OUTSIDE the bordered frame, top left
ttk.Label(
    admin_frame,
    text="Legenda:",
    font=("Segoe UI", 11, "bold"),
    foreground="#1a4d2e"
).grid(row=0, column=0, sticky="w", pady=(0, 2), padx=(0, 0))

ttk.Label(
    admin_frame,
    text="• Plan: Datoteka Plan brizganja mesečni excel\n"
         "• List2, Izbor, Brizganje izračun: Datoteka Poročanje proizvodnje excel",
    font=("Segoe UI", 10),
    foreground="#333"
).grid(row=1, column=0, sticky="w", pady=(0, 12), padx=(0, 0))

# Modern look: add a bordered inner frame for settings
admin_inner = ttk.Frame(admin_frame, padding=18, relief="groove", borderwidth=2)
admin_inner.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)

# Title
ttk.Label(admin_inner, text="Napredne nastavitve", font=("Segoe UI", 18, "bold"), foreground="#2E8B57").grid(
    row=0, column=0, columnspan=6, pady=(0, 18)
)

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_config(config):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)

admin_vars = [
    ("PLAN_ROW_DATE", "Plan: vrstica z datumi (npr. 4)", 4),
    ("PLAN_ROW_FIKSNO", "Plan: vrstica z Fiksno/V pripravi (npr. 5)", 5),
    ("PLAN_ROW_START", "Plan: začetna vrstica za kopiranje (npr. 6)", 6),
    ("PLAN_ROW_END", "Plan: končna vrstica za kopiranje (npr. 44)", 44),
    ("PLAN_COLS_TO_COPY", "Plan: število stolpcev za kopiranje", 3),
    ("BRIZGANJE_ROW_4", "Brizganje izračun: vrstica z datumom (npr. 4)", 4),
    ("BRIZGANJE_ROW_7", "Brizganje izračun: začetna vrstica za analizo, slike in pregled napak (npr. 7)", 7),
    ("BRIZGANJE_ROW_46", "Brizganje izračun: končna vrstica za analizo analizo, slike in pregled napak (npr. 46)", 46),
    #("BRIZGANJE_ROW_44", "Brizganje izračun: končna vrstica za slike (npr. 44)", 44),
    #("BRIZGANJE_ROW_45", "Brizganje izračun: končna vrstica za pregled napak (npr. 45)", 45),
    ("BRIZGANJE_COL_A", "Brizganje izračun: stolpec z siframi (npr. 1)", 1),
    ("BRIZGANJE_COL_L", "Brizganje izračun: stolpec z procenti izmeta(npr. 12)", 12),
    ("BRIZGANJE_COL_M", "Brizganje izračun: stolpec z denarjem (npr. 13)", 13),
    ("BRIZGANJE_COL_H", "Brizganje izračun: stolpec za pregled napak (npr. 8)", 8),
    ("BRIZGANJE_COL_I", "Brizganje izračun: stolpec za preverjanje (npr. 9)", 9),
    ("BRIZGANJE_COL_E", "Brizganje izračun: stolpec E za brisanje in kopiranje (npr. 5)", 5),
    ("BRIZGANJE_COL_C", "Brizganje izračun: stolpec C za kopiranje (npr. 3)", 3),
    ("BRIZGANJE_COL_D", "Brizganje izračun: stolpec D za kopiranje (npr. 4)", 4),
    ("IZBOR_COL_AA", "Izbor: stolpec AA za filtriranje (npr. 27)", 27),
    ("IZBOR_COL_F", "Izbor: začetni stolpec za kopiranje (npr. 6)", 6),
    ("IZBOR_COL_M", "Izbor: končni stolpec za kopiranje (npr. 13)", 13),
    ("IZBOR_COL_S", "Izbor: stolpec za vrednost v izboru (npr. 19)", 19),
    ("LIST2_ROW_2", "List2: začetna vrstica za čiščenje (npr. 2)", 2),
    ("LIST2_COL_T", "List2: začetni stolpec za čiščenje/kopiranje (npr. 20)", 20),
    ("LIST2_COL_AA", "List2: končni stolpec za čiščenje (npr. 27)", 27),
    ("LIST2_ROW_8", "List2: končna vrstica za kopiranje (npr. 8)", 8),
    ("LIST2_ROW_9", "List2: začetna vrstica za iskanje (npr. 9)", 9),
    ("LIST2_ROW_27", "List2: končna vrstica za iskanje (npr. 27)", 27),
    ("LIST2_COL_C", "List2: stolpec za iskanje (npr. 3)", 3),
    ("LIST2_COL_B", "List2: začetni stolpec za kopiranje (črka, npr. 'B')", "B"),
    ("LIST2_COL_L", "List2: končni stolpec za kopiranje (črka, npr. 'L')", "L"),
    ("LIST2_ROW_1", "List2: začetna vrstica za kopiranje (npr. 1)", 1),
]

config = load_config()
admin_entries = {}

columns_per_row = 3  # Change to 4 or 5 if you want even more compact

for idx, (var, label, default) in enumerate(admin_vars):
    row = (idx // columns_per_row) + 3 
    col = (idx % columns_per_row) * 2  

    ttk.Label(admin_inner, text=label + ":", font=("Segoe UI", 9, "bold")).grid(
        row=row, column=col, sticky="e", padx=4, pady=4
    )
    entry = ttk.Entry(admin_inner, width=8, font=("Segoe UI", 9))
    entry.grid(row=row, column=col + 1, padx=4, pady=4)
    entry.insert(0, str(config.get(var, default)))
    admin_entries[var] = entry

def save_admin_config():
    new_config = {}
    for var, entry in admin_entries.items():
        value = entry.get()
        try:
            value = int(value)
        except ValueError:
            pass
        new_config[var] = value
    save_config(new_config)
    messagebox.showinfo("Shranjeno", "Nastavitve so bile shranjene!")

total_cols = columns_per_row * 2
ttk.Button(
    admin_inner,
    text="Shrani nastavitve",
    command=save_admin_config,
    style="Run.TButton"
).grid(
    row=(len(admin_vars) - 1) // columns_per_row + 4,
    column=0,
    columnspan=total_cols,
    pady=30
)

def on_tab_changed(event):
    selected_tab = event.widget.select()
    tab_text = event.widget.tab(selected_tab, "text")
    if tab_text == "Uporabnik":
        root.geometry("620x340")  # Original size for user tab
    elif tab_text == "Admin":
        root.geometry("1490x650")  # Bigger size for admin tab

notebook.bind("<<NotebookTabChanged>>", on_tab_changed)

root.mainloop()