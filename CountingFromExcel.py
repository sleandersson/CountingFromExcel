import pandas as pd
from tkinter import Tk, Label, Button, messagebox, filedialog, StringVar
from tkinter import ttk
from tkcalendar import DateEntry
import threading
import time
import logging
from datetime import datetime

# Setup logging with timestamps
logging.basicConfig(
    filename='error.log',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Global variables
global stop_timer, stop_process
stop_timer = False
stop_process = False

def select_file1():
    file_path = filedialog.askopenfilename(title="Select the first Excel file")
    file1_var.set(file_path)

def select_file2():
    file_path = filedialog.askopenfilename(title="Select the second Excel file (optional)")
    file2_var.set(file_path)

def count_testnames():
    progress_bar.start()
    elapsed_time_var.set("Elapsed time: 0.00 seconds")

    global stop_timer, stop_process
    stop_timer = False
    stop_process = False

    threading.Thread(target=update_elapsed_time).start()
    threading.Thread(target=process_files).start()

    count_button.config(state='disabled')
    close_button.config(text="Stop", command=stop_search)

def update_elapsed_time():
    start_time = time.time()
    while not stop_timer:
        elapsed_time = time.time() - start_time
        elapsed_time_var.set(f"Elapsed time: {elapsed_time:.2f} seconds")
        time.sleep(0.1)

def stop_search():
    global stop_process
    stop_process = True

def cleanup_and_exit(stopped=False):
    global stop_timer
    stop_timer = True

    progress_bar.stop()

    count_button.config(state='normal')
    close_button.config(text='Close', command=root.quit)

    if stopped:
        messagebox.showinfo("Process Stopped", "The search process was stopped by the user.")

def process_files():
    try:
        start_time = time.time()

        # Parse date range
        start_date = pd.to_datetime(start_date_entry.get_date().strftime('%Y-%m-%d'))
        end_date = pd.to_datetime(end_date_entry.get_date().strftime('%Y-%m-%d'))

        file1 = file1_var.get()
        file2 = file2_var.get()

        if not file1:
            messagebox.showerror("Error", "Please select at least one Excel file.")
            cleanup_and_exit()
            return

        columns = ['Indatum', 'Testnamn']
        filtered_rows = []

        def filter_excel(file):
            df = pd.read_excel(file, engine='openpyxl', usecols=columns)
            df['Indatum'] = pd.to_datetime(df['Indatum'], errors='coerce')
            return df[(df['Indatum'] >= start_date) & (df['Indatum'] <= end_date)]

        # Filter the files during reading
        dfs = []
        dfs.append(filter_excel(file1))
        if file2:
            dfs.append(filter_excel(file2))

        # Combine filtered data into a single DataFrame
        filtered_df = pd.concat(dfs) if dfs else pd.DataFrame(columns=columns)

        if stop_process:
            cleanup_and_exit(stopped=True)
            return

        # Process DataFrame in chunks of 10 rows
        total_count = 0
        adm_count = 0
        testname_counts = pd.Series(dtype=int)
        for i in range(0, len(filtered_df), 10):
            if stop_process:
                cleanup_and_exit(stopped=True)
                return

            chunk = filtered_df.iloc[i:i+10]
            total_count += chunk['Testnamn'].count()
            adm_count += chunk[chunk['Testnamn'] == 'Adm'].shape[0]
            testname_counts = testname_counts.add(chunk['Testnamn'].value_counts(), fill_value=0)

        if stop_process:
            cleanup_and_exit(stopped=True)
            return

        # Adjusted total excluding 'Adm'
        adjusted_total = total_count - adm_count

        # Write results to a file
        with open('SumTests_SearchLog.txt', 'a') as f:
            f.write(f"Date Range: {start_date.date()} to {end_date.date()}\n")
            f.write(f"Total count of Testnamn: {total_count}\n")
            f.write(f"Count of 'Adm': {adm_count}\n")
            f.write(f"Adjusted total (excluding 'Adm'): {adjusted_total}\n")
            f.write("\nTestnamn Counts:\n")
            f.write(testname_counts.to_string())
            f.write("\n\n")

        # Complete the process
        cleanup_and_exit()
        elapsed_time = time.time() - start_time
        elapsed_time_var.set(f"Elapsed time: {elapsed_time:.2f} seconds")
        messagebox.showinfo(
            "Results",
            f"Period: {start_date.date()} to {end_date.date()}\n"
            f"Adjusted total (excluding 'Adm'): {adjusted_total}"
        )

    except Exception as e:
        # Log error with timestamp
        logging.error("An error occurred during processing", exc_info=True)
        messagebox.showerror("Error", f"An error occurred: {e}")
        cleanup_and_exit()

# Create the main window
root = Tk()
root.title("Excel Testname Counter")

# Center the window on the screen
root.update_idletasks()
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 370
window_height = 400
x_offset = (screen_width // 2) - (window_width // 2)
y_offset = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{x_offset}+{y_offset}")

file1_var = StringVar()
file2_var = StringVar()
elapsed_time_var = StringVar(value="Elapsed time: 0.00 seconds")

Label(root, text="Start Date:").grid(row=0, column=0, padx=10, pady=10)
start_date_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
start_date_entry.grid(row=0, column=1, padx=10, pady=10)

Label(root, text="End Date:").grid(row=1, column=0, padx=10, pady=10)
end_date_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
end_date_entry.grid(row=1, column=1, padx=10, pady=10)

Button(root, text="Select First Excel File", command=select_file1).grid(row=2, column=0, padx=10, pady=10)
Label(root, textvariable=file1_var, wraplength=140, justify='left').grid(row=2, column=1, padx=10, pady=10)

Button(root, text="Select Second Excel File (optional)", command=select_file2).grid(row=3, column=0, padx=10, pady=10)
Label(root, textvariable=file2_var, wraplength=140, justify='left').grid(row=3, column=1, padx=10, pady=10)

count_button = Button(root, text="Count Testnames", command=count_testnames)
count_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

progress_bar = ttk.Progressbar(root, mode='indeterminate')
progress_bar.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

Label(root, textvariable=elapsed_time_var).grid(row=6, column=0, columnspan=2, padx=10, pady=10)

close_button = Button(root, text="Close", command=root.quit)
close_button.grid(row=7, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()
