import configparser
from datetime import datetime
import os
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import concurrent.futures
import signal
import pandas as pd
import time 

from embassy import Embassies
from VisaScheduler import VisaScheduler

# Load user details from Excel file
excel_file = 'user_details.xlsx'
df = pd.read_excel(excel_file)

config_file = 'config.ini'
config = configparser.ConfigParser()
config.read(config_file)

MAX_THREADS = config['RUNTIME'].getfloat('MAX_THREADS')

# Define a function that performs the visa scheduling process for a single row
def run_visa_scheduling(username, password, schedule_id, period_start, period_end, your_embassy, embassies):
    while 1:
        try:
            visa_scheduler = VisaScheduler(username, password, schedule_id, period_start, period_end, your_embassy, embassies)
            visa_scheduler.run()
            print(f"Visa scheduling for {username} completed.")
        except Exception as e:
            print(f"An error occurred for {username}: {str(e)}")
        time.sleep(5)

# Define a cleanup function to terminate all processes
def cleanup(signum, frame):
    print("Received Ctrl+C. Shutting down...")
    executor.shutdown(wait=False)
    os._exit(1)

# Function to run the script
def start_script():
    try:
        # Initialize the ThreadPoolExecutor
        with concurrent.futures.ThreadPoolExecutor(MAX_THREADS) as executor:
            # Register the cleanup function to handle Ctrl+C
            signal.signal(signal.SIGINT, cleanup)

            # Iterate over each row in the DataFrame and submit it for concurrent execution
            for index, row in df.iterrows():
                username = row['USERNAME']
                password = row['PASSWORD']
                schedule_id = row['SCHEDULE_ID']
                period_start = row['PRIOD_START'].strftime('%Y-%m-%d')
                period_end = row['PRIOD_END'].strftime('%Y-%m-%d')
                your_embassy = row['PREFERRED_EMBASSY']
                executor.submit(run_visa_scheduling, username, password, schedule_id, period_start, period_end, your_embassy, Embassies)

        # This code will properly handle Ctrl+C and terminate all processes.

        messagebox.showinfo("Script Complete", "The script has completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
# Function to stop the script
def stop_script():
    global terminate_script
    terminate_script.set()  # Set the termination event

# Function to gracefully terminate the script on Ctrl+C
def cleanup(signum, frame):
    print("Received Ctrl+C. Shutting down...")
    executor.shutdown(wait=False)
    os._exit(1)

start_script()

# # Create a Tkinter window
# root = tk.Tk()
# root.title("Visa Scheduler")
# root.geometry("500x200")

# # Create and configure GUI elements
# frame = ttk.Frame(root)
# frame.pack(padx=20, pady=20)

# # Define the allowed date (15th October 2023)
# allowed_date = datetime(2999, 11, 15)

# # Check the current date
# current_date = datetime.now()

# # Function to hide the "Click here to start" label
# def hide_start_label():
#     start_label.pack_forget()

# # Create a label widget to display the message
# message_label = ttk.Label(frame, text="")
# message_label.pack()
# # Function to update the message label and button states
# def update_message_and_buttons():
#     if current_date > allowed_date:
#         message_label.config(text="This application is no longer available for use after " + allowed_date.strftime("%dth %B %Y") + ".")
#         # Hide the "Start" and "Stop" buttons when the date has passed
#         start_button.pack_forget()
#         stop_button.pack_forget()
#         # Hide the "Click here to start" label as well
#         hide_start_label()
#     else:
#         message_label.config(text="Click the Start button to run the script.")
#         # Show the buttons if the date is not reached
#         start_button.pack()
#         stop_button.pack()
#         # Show the "Click here to start" label
#         start_label.pack()

# # Create and configure GUI elements
# start_label = ttk.Label(frame, text="")
# start_button = ttk.Button(frame, text="Start", command=start_script)
# stop_button = ttk.Button(frame, text="Stop", command=stop_script)

# # Initialize the ThreadPoolExecutor
# with concurrent.futures.ThreadPoolExecutor(1) as executor:
#     # Register the cleanup function to handle Ctrl+C
#     signal.signal(signal.SIGINT, cleanup)

#     # Update the message label and buttons initially
#     update_message_and_buttons()

#     # Start the Tkinter main loop
#     root.mainloop()