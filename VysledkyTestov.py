import customtkinter
import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime
import re
import pymysql
from PIL import Image, ImageTk
from tkinter import Text, ttk, Listbox
import pandas as pd
import os
import sys

customtkinter.set_appearance_mode("white")
customtkinter.set_default_color_theme("dark-blue")

def initialize_database_connection():
    try:
        selected_friendly_name = selected_db_name_var.get()
        selected_db_name = db_name_mapping[selected_friendly_name]
        connection = pymysql.connect(host=db_host, user=db_user, password=db_password, database=selected_db_name)
        cursor = connection.cursor()
        database_connected = True
        return connection, cursor, database_connected
    except Exception as e:
        print(f"Database connection error: {e}")
        database_connected = False
        tk.messagebox.showerror("Database Connection Error", f"Unable to connect to the database '{selected_db_name}'. Check your credentials.")
        root.destroy()
        return None, None, False

# Function to update the message label when a different database is selected
def update_message_label(event):
    global connection, cursor, database_connected
    connection, cursor, database_connected = initialize_database_connection()
    if database_connected:
        message_label.config(text=f"Connected to database: {selected_db_name_var.get()}")
    else:
        message_label.config(text="Database connection failed")

# Modify the find_button_callback function
def find_button_callback():
    global total_results_count  

    if not database_connected:
        tk.messagebox.showerror("Database Error", "Unable to connect to the database.")
        return

    cell_values = cell_listbox.curselection()
    selected_cells = [cell_listbox.get(index) for index in cell_values]

    start_date_value = start_date_var.get()
    end_date_value = end_date_var.get()
    produkt_index_value = produkt_index.get()
    limit_index_value = limit_index.get()

    if any(not re.match(r'^\d{2}-\d{2}-\d{4}$', date_value) for date_value in [start_date_value, end_date_value]):
        tk.messagebox.showerror("Error", "Invalid date format. Please use dd-mm-yyyy")
        return

    try:
        total_results_count = 0 

        results_text.delete(1.0, tk.END)

        for selected_cell in selected_cells:
            query = f"SELECT * FROM test_results WHERE cell = '{selected_cell}' " \
                    f"AND DATE(test_date_time) BETWEEN STR_TO_DATE('{start_date_value}', '%d-%m-%Y') " \
                    f"AND STR_TO_DATE('{end_date_value}', '%d-%m-%Y') " \
                    f"AND produkt_index = '{produkt_index_value}' "

            query += "ORDER BY test_date_time DESC "

            if limit_index_value and (not limit_index_value.isdigit() or int(limit_index_value) < 0):
                tk.messagebox.showerror("Error", "Chyba: Číslo musí byť kladné !")
                return

            if limit_index_value:
                query += f"LIMIT {limit_index_value}"

            print("Constructed Query:", query)  
            cursor.execute(query)
            results = cursor.fetchall()

            print(f"Number of results for {selected_cell}: {len(results)}")  

            total_results_count += len(results)

            char = selected_cell  
            formatted_labels = [f"{char}{row[0]}" for row in results]
            results_text.insert(tk.END, "\n".join(formatted_labels) + "\n")

        result_count_label.configure(text=f"Počet nájdených výsledkov: {total_results_count}")
        results_button.configure(state="normal")

    except Exception as e:
        tk.messagebox.showerror("Database Error", f"An error occurred while interacting with the database: {str(e)}")

# Function to handle the "Results" button click
def results_button_callback():
    global total_results_count 

    labels_text = results_text.get("1.0", tk.END)
    labels = [line.strip() for line in labels_text.split('\n') if line]

    if not labels:
        tk.messagebox.showerror("Error", "Žiadné výsledky neboli nájdené.")
        return

    query = f"SELECT tr.*, tri.* " \
            f"FROM test_results tr " \
            f"JOIN test_result_items tri ON tr.ID = tri.test_result_id " \
            f"WHERE tr.label IN ({', '.join([f'\'{label}\'' for label in labels])}) "

    try:
        cursor.execute(query)
        results = cursor.fetchall()
        rows = [dict(zip([description[0] for description in cursor.description], row)) for row in results]
        display_results_in_treeview(rows)
        result_count_label.configure(text=f"Počet nájdených výsledkov: {len(results)}")

    except Exception as e:
        tk.messagebox.showerror("Database Error", f"An error occurred while interacting with the database: {str(e)}")

treeview = None
reset_button = None

# Function to display results in Treeview
def display_results_in_treeview(results, limit=None):
    global treeview, reset_button

    if treeview:
        treeview.destroy()
        treeview = None

    if reset_button:
        reset_button.destroy()
        reset_button = None

    columns = list(results[0].keys())
    treeview = ttk.Treeview(root, columns=columns, show='headings')

    for column in columns:
        treeview.heading(column, text=column)
        treeview.column(column, width=100, stretch=tk.YES)

    for row in reversed(results[:limit]):
        treeview.insert("", "end", values=list(row.values()))

    scrollbar_y = ttk.Scrollbar(root, orient="vertical", command=treeview.yview)
    scrollbar_y.pack(side="right", fill="y")
    treeview.configure(yscrollcommand=scrollbar_y.set)

    scrollbar_x = ttk.Scrollbar(root, orient="horizontal", command=treeview.xview)
    scrollbar_x.pack(side="bottom", fill="x")
    treeview.configure(xscrollcommand=scrollbar_x.set)

    if reset_button:
        reset_button.destroy()

    reset_button = customtkinter.CTkButton(root, text="Reset", command=reset_treeview)
    reset_button.pack(pady=10)

    treeview.pack(expand=True, fill="both")

# Function to reset Treeview
def reset_treeview():
    global treeview, reset_button

    if treeview:
        treeview.destroy()
        treeview = None

    if reset_button:
        reset_button.destroy()
        reset_button = None

# Function to handle the "Export to Excel" button click
def export_to_excel():
    labels_text = results_text.get("1.0", tk.END)
    labels = [line.strip() for line in labels_text.split('\n') if line]

    if not labels:
        tk.messagebox.showerror("Error", "Žiadné výsledky na exportovanie.")
        return

    query = f"SELECT tr.*, tri.* " \
            f"FROM test_results tr " \
            f"JOIN test_result_items tri ON tr.ID = tri.test_result_id " \
            f"WHERE tr.label IN ({', '.join([f'\'{label}\'' for label in labels])}) "

    try:
        cursor.execute(query)
        results = cursor.fetchall()

        df = pd.DataFrame(results, columns=[description[0] for description in cursor.description])

        current_datetime = datetime.now().strftime("%d-%m-%Y-%H-%M")
        file_name = f"{current_datetime}-vysledky.xlsx"

        desktop_path = tk.filedialog.askdirectory(title="Select Desktop Folder")
        excel_file_path = f"{desktop_path}/{file_name}"

        df.to_excel(excel_file_path, index=False)
        tk.messagebox.showinfo("Export Successful", f"Results exported to {excel_file_path}")

    except Exception as e:
        tk.messagebox.showerror("Database Error", f"An error occurred while exporting to Excel: {str(e)}")

# Main Application
root = customtkinter.CTk()
root.wm_title('SQL Results')
script_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
# Load icon
icon_path = os.path.join(script_dir, 'assets', 'icon.ico')
print(f"Icon path: {icon_path}")

try:
    root.iconbitmap(icon_path)
except tk.TclError as e:
    print(f"Icon loading error: {e}")

root.resizable(width=True, height=True)
root.geometry("1000x600")

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
root.geometry(f"{screen_width}x{screen_height}+0+0")


db_host = '172.19.0.202'
db_user = 'spaghetti'
db_password = 'JkbCEva9EadMqEXV'
db_name_mapping = {
    'Pass 1 a 3': 'pass_spagety',
    'Pass 2': 'pass_spaghetti',
}
possible_db_names = list(db_name_mapping.keys())

selected_db_name_var = tk.StringVar()
selected_db_name_var.set(possible_db_names[0])

database_combobox = ttk.Combobox(root, textvariable=selected_db_name_var, values=possible_db_names)
database_combobox.pack()

message_label = tk.Label(root, text="", fg="green")
message_label.pack()

connection, cursor, database_connected = initialize_database_connection()

database_combobox.bind("<<ComboboxSelected>>", update_message_label)

def on_closing():
    if database_connected:
        connection.close()
    root.destroy()

def check_root():
    if root.winfo_exists():
        root.protocol("WM_DELETE_WINDOW", on_closing)
    else:
        root.after(100, check_root)

if database_connected:
    success_label = tk.Label(root, text="Successfully connected to the database!", font=("Arial", 12), fg="green")
    success_label.pack(pady=20)

    def hide_success_message():
        success_label.pack_forget()

    root.after(5000, hide_success_message)

frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=10, padx=60, fill="both", expand=True)

label = customtkinter.CTkLabel(master=frame, text="Výsledky Testov", font=("Arial", 24, "bold"))
label.pack(pady=12, padx=10)

# Load logo image
logo_path = os.path.join(script_dir, 'assets', 'logo.png')
print(f"Logo path: {logo_path}")

try:
    logo_image = Image.open(logo_path)
    logo_image = logo_image.convert("RGBA")
    logo_image.thumbnail((150, 150))
    logo_image = ImageTk.PhotoImage(logo_image)
except FileNotFoundError as e:
    print(f"Logo file not found: {e}")
    logo_image = None  # Handle missing logo gracefully
except Exception as e:
    print(f"Error loading logo: {e}")
    logo_image = None  # Handle other errors

print(f"Logo image: {logo_image}")

logo_label = tk.Label(master=frame, image=logo_image)
logo_label.pack(pady=5, padx=10)

combo_frame = customtkinter.CTkFrame(master=frame)
combo_frame.pack(pady=12, padx=10)

cell_listbox = Listbox(combo_frame, selectmode=tk.MULTIPLE, height=4)
for cell in ["A", "B", "C", "D"]:
    cell_listbox.insert(tk.END, cell)

cell_listbox.pack(side=tk.LEFT, padx=5)

date_frame = customtkinter.CTkFrame(master=frame)
date_frame.pack(pady=12, padx=10)

start_date_var = tk.StringVar()
start_date_picker = DateEntry(date_frame, textvariable=start_date_var, date_pattern='dd-mm-yyyy', show_week_numbers=False)
start_date_picker.pack(side=tk.LEFT, padx=5)

end_date_var = tk.StringVar()
end_date_picker = DateEntry(date_frame, textvariable=end_date_var, date_pattern='dd-mm-yyyy', show_week_numbers=False)
end_date_picker.pack(side=tk.LEFT, padx=5)

cal_var = customtkinter.StringVar()

produkt_index = customtkinter.CTkEntry(master=frame, placeholder_text="produkt_index")
produkt_index.pack(pady=12, padx=20)

limit_index = customtkinter.CTkEntry(master=frame, placeholder_text="Počet výsledkov")
limit_index.pack(pady=12, padx=20)

include_all_results_var = tk.IntVar()
include_all_results_checkbox = customtkinter.CTkCheckBox(
    master=frame, text="Zahrnúť všetky výsledky", variable=include_all_results_var)
include_all_results_checkbox.pack(pady=5, padx=10)

find_button = customtkinter.CTkButton(master=frame, text="Hľadať", command=find_button_callback)
find_button.pack(pady=12, padx=10)

results_button = customtkinter.CTkButton(master=frame, text="Výsledky", command=results_button_callback, state=tk.DISABLED)
results_button.pack(pady=12, padx=10)

result_count_label = customtkinter.CTkLabel(master=frame, text="")
result_count_label.pack(pady=5, padx=10)

results_text = Text(master=frame, height=5, width=50)
results_text.pack(pady=10, padx=10)

export_button = customtkinter.CTkButton(master=frame, text="Export to Excel", command=export_to_excel)
export_button.pack(pady=5, padx=7)

paned_window = tk.PanedWindow(root, orient="horizontal", sashwidth=8, sashrelief="raised")
paned_window.pack(side="bottom", expand=True, fill="both")

check_root()
root.mainloop()
