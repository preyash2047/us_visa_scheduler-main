import subprocess
import logging
import openpyxl
import time
import psutil
import pandas as pd

# Initialize the logging system
filename='visa_scheduler.log'
logging.basicConfig(filename=filename, level=logging.INFO)

# Define a function to start a new process
def start_process(username, password, schedule_id, period_start, period_end, your_embassy):
    cmd = f"python visa_v.4.py {username} {password} {schedule_id} {period_start} {period_end} {your_embassy}"
    subprocess.Popen(cmd, shell=True)

# Define a function to stop a process by username
def stop_process(username):
    # Iterate over all running processes
    for process in psutil.process_iter(['pid', 'name', 'username']):
        try:
            process_info = process.info()
            process_username = process_info.get('username')
            
            # Check if the process username matches the target username
            if process_username == username:
                pid = process_info['pid']
                
                # Terminate the process
                process = psutil.Process(pid)
                process.terminate()
                
                print(f"Process with username '{username}' (PID {pid}) has been terminated.")
        
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

def monitor_logs(username):
    # Initialize a variable to keep track of the last read position in the log file
    last_position = 0
    
    while True:
        try:
            # Open the log file in read mode
            with open(filename, 'r') as log_file:
                # Move to the last read position
                log_file.seek(last_position)
                
                # Read and display new log entries
                new_logs = log_file.read()
                if new_logs:
                    print(new_logs, end='')
                    
                    # Update the last read position
                    last_position = log_file.tell()
            
            # Sleep for a while before checking for new logs again (e.g., every 1 second)
            time.sleep(1)
        
        except FileNotFoundError:
            # Handle the case where the log file does not exist
            print(f"Log file not found: {filename}")
            break
        except KeyboardInterrupt:
            # Handle Ctrl+C to gracefully exit the log monitoring
            print("Log monitoring stopped.")
            break

def update_from_excel(filename, existing_dataframe):
    try:
        # Read the updated Excel file into a new DataFrame
        updated_dataframe = pd.read_excel(filename)

        # Merge the updated DataFrame with the existing DataFrame to add new entries
        merged_dataframe = pd.concat([existing_dataframe, updated_dataframe], ignore_index=True)

        # Optionally, you can drop duplicates if needed
        merged_dataframe.drop_duplicates(inplace=True)

        # Return the updated DataFrame
        return merged_dataframe

    except FileNotFoundError:
        print(f"Excel file not found: {filename}")
        return existing_dataframe
    except Exception as e:
        print(f"Error updating from Excel: {str(e)}")
        return existing_dataframe

# Main menu
while True:
    print("\n=== Visa Scheduler Console ===")
    print("1. Start Process")
    print("2. Stop Process")
    print("3. Monitor Logs")
    print("4. Update from Excel")
    print("5. Exit")
    choice = input("Select an option: ")

    if choice == "1":
        username = input("Username: ")
        password = input("Password: ")
        schedule_id = input("Schedule ID: ")
        period_start = input("Period Start: ")
        period_end = input("Period End: ")
        your_embassy = input("Your Embassy: ")
        start_process(username, password, schedule_id, period_start, period_end, your_embassy)

    elif choice == "2":
        username = input("Enter the username to stop the process: ")
        stop_process(username)

    elif choice == "3":
        username = input("Enter the username to monitor logs: ")
        monitor_logs(username)

    elif choice == "4":
        filename = input("Enter the Excel filename: ")
        update_from_excel(filename)

    elif choice == "5":
        break
