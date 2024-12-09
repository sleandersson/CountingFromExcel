import pandas as pd
from tkinter import Tk, Label, Button, messagebox, filedialog, StringVar
from tkinter import ttk
from tkcalendar import DateEntry
import threading
import time

global stop_timer

def select_file1():
    file_path = filedialog.askopenfilename(title="Select the first Excel file")
    file1_var.set(file_path)

def select_file2():
    file_path = filedialog.askopenfilename(title="Select the second Excel file (optional)")
    file2_var.set(file_path)

def count_testnames():
    # Show progress bar
    progress_bar.start()

    # Clear previous elapsed time
    elapsed_time_var.set("Elapsed time: 0.00 seconds")

    # Start the elapsed time update thread
    global stop_timer
    stop_timer = False
    threading.Thread(target=update_elapsed_time).start()

    # Run the counting process in a separate thread to avoid freezing the GUI
    threading.Thread(target=process_files).start()

def update_elapsed_time():
    start_time = time.time()
    while not stop_timer:
        elapsed_time = time.time() - start_time
        elapsed_time_var.set(f"Elapsed time: {elapsed_time:.2f} seconds")
        time.sleep(0.1)  # Update every 0.1 seconds for smoother updates

def process_files():
    start_time = time.time()
    
    # Get the date range from the user inputs
    start_date = start_date_entry.get_date().strftime('%Y-%m-%d')
    end_date = end_date_entry.get_date().strftime('%Y-%m-%d')
    
    # Get the file paths
    file1 = file1_var.get()
    file2 = file2_var.get()
    
    # Check if at least one file path is provided
    if not file1:
        messagebox.showerror("Error", "Please select at least one Excel file.")
        progress_bar.stop()
        global stop_timer
        stop_timer = True
        return
    
    # Columns to read from the Excel files
    columns = ['Indatum', 'Testnamn']
    
    # Load the first Excel file into a DataFrame
    df1 = pd.read_excel(file1, engine='openpyxl', usecols=columns)

    if file2:
        # Load the second Excel file into a DataFrame
        df2 = pd.read_excel(file2, engine='openpyxl', usecols=columns)
        # Combine the two DataFrames into one
        df = pd.concat([df1, df2])
    else:
        df = df1

    # Filter based on date strings first
    filtered_df = df[(df['Indatum'] >= start_date) & (df['Indatum'] <= end_date)]

    # Convert the filtered date column to datetime
    filtered_df['Indatum'] = pd.to_datetime(filtered_df['Indatum'], errors='coerce')

    # Debug prints to check filtering results
    print(f"Filtered DataFrame:\n{filtered_df.head()}")

    # Count the occurrences of 'Testnamn'
    testname_counts = filtered_df['Testnamn'].value_counts()
    total_count = filtered_df['Testnamn'].count()
    adm_count = filtered_df[filtered_df['Testnamn'] == 'Adm'].shape[0]

    end_time = time.time()
    elapsed_time = end_time - start_time

    # Append the results to a text file
    with open('SumTests_SearchLog.txt', 'a') as f:
        f.write(f"Date Range: {start_date} to {end_date}\n")
        f.write(f"Total count of Testnamn: {total_count}\n")
        f.write(f"Count of 'Adm': {adm_count}\n")
        f.write(f"Search time: {elapsed_time:.2f} seconds\n")
        f.write("\nTestnamn Counts:\n")
        f.write(testname_counts.to_string())
        f.write("\n\n")

    # Stop the elapsed time update thread
    stop_timer = True

    # Update elapsed time display one last time
    elapsed_time_var.set(f"Elapsed time: {elapsed_time:.2f} seconds")

    # Show a messagebox with the results
    messagebox.showinfo("Results", f"Period: {start_date} to {end_date}\nTotal count: {total_count}")

    # Stop progress bar
    progress_bar.stop()

# Create the main window
root = Tk()
root.title("Excel Testname Counter")

# Variables to store file paths and elapsed time
file1_var = StringVar()
file2_var = StringVar()
elapsed_time_var = StringVar(value="Elapsed time: 0.00 seconds")
stop_timer = False

# Create and place the widgets
Label(root, text="Start Date:").grid(row=0, column=0, padx=10, pady=10)
start_date_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
start_date_entry.grid(row=0, column=1, padx=10, pady=10)

Label(root, text="End Date:").grid(row=1, column=0, padx=10, pady=10)
end_date_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
end_date_entry.grid(row=1, column=1, padx=10, pady=10)

Button(root, text="Select First Excel File", command=select_file1).grid(row=2, column=0, padx=10, pady=10)
Label(root, textvariable=file1_var).grid(row=2, column=1, padx=10, pady=10)

Button(root, text="Select Second Excel File (optional)", command=select_file2).grid(row=3, column=0, padx=10, pady=10)
Label(root, textvariable=file2_var).grid(row=3, column=1, padx=10, pady=10)

count_button = Button(root, text="Count Testnames", command=count_testnames)
count_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

progress_bar = ttk.Progressbar(root, mode='indeterminate')
progress_bar.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

Label(root, textvariable=elapsed_time_var).grid(row=6, column=0, columnspan=2, padx=10, pady=10)

close_button = Button(root, text="Close", command=root.quit)
close_button.grid(row=7, column=0, columnspan=2, padx=10, pady=10)

# Run the main loop
root.mainloop()

