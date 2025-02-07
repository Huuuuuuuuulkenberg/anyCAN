import os
import can
import sys
import time
import signal
import keyboard
import openpyxl
import threading
import pandas as pd
import tkinter as tk
from tkinter import PhotoImage
from tkinter import messagebox
from tkinter import filedialog, ttk
from datetime import datetime, timedelta

# Global flags
running = True
paused = False
automatic_mode = False
read_messages = pd.DataFrame()
current_test_case_index = 0
test_case_files = []

# Function to initialize CAN interface
def init_can_interface(channel, bitrate):
    bus = can.interface.Bus(channel=channel, interface='ixxat', bitrate=bitrate)
    return bus

# Function to log CAN messages into an Excel file with absolute time
def log_to_excel(messages, filename):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "CAN Messages"

    # Set up column headers    
    headers = ['Timestamp', 'ID', 'DLC', 'Data', 'Delay (ms)']
    sheet.append(headers)
    
    # Capture start time for absolute time calculation
    start_time = datetime.now()
    
    for msg in messages:
        # Compute the absolute time
        timestamp_seconds = msg.timestamp
        absolute_time = start_time + timedelta(seconds=timestamp_seconds)
        formatted_time = absolute_time.strftime('%d:%H:%M:%S')
        
        msg_id = hex(msg.arbitration_id)
        dlc = msg.dlc
        data = ' '.join(format(byte, '02x') for byte in msg.data)
        delay = msg.delay if hasattr(msg, 'delay') else 0
        row = [formatted_time, msg_id, dlc, data, delay]
        sheet.append(row)
    
    # Save the Excel file
    workbook.save(filename)
    print(f"Data successfully logged to {filename}")

# Function to load folder containing Test Cases
def select_test_cases_folder():
    global test_case_files, current_test_case_index
    folder_path = filedialog.askdirectory(title="Select Folder Containing Test Cases")
    if not folder_path:
        return False
    
    test_case_files = [
        os.path.join(folder_path, f) for f in os.listdir(folder_path)
        if f.endswith(('.xlsx', '.xls'))
    ]
    test_case_files.sort()
    
    if not test_case_files:
        messagebox.showwarning("Warning", "No Excel files found in the selected folder")
        return False
    
    current_test_case_index = 0
    print(f"Found {len(test_case_files)} test case files")
    return True

# Function to load test cases onto GUI fields
def load_test_case(entries, file_path=None):
    global current_test_case_index
    
    if not file_path and current_test_case_index >= len(test_case_files):
        messagebox.showinfo("Complete", "All test cases have been completed!")
        return False
    
    try:
        if not file_path:
            file_path = test_case_files[current_test_case_index]
        
        df = pd.read_excel(file_path)
        
        # Clear existing entries
        for entry in entries:
            for field in entry[:4]:
                field.delete(0, tk.END)
        
        write_messages = df[df['Read/Write'].str.lower() == 'write']
        
        row_counter = 0
        for i, row in write_messages.iterrows():
            if row_counter >= len(entries):
                break
            entries[row_counter][0].insert(0, row['ID'])
            entries[row_counter][2].insert(0, row['Data'])
            
            delay_value = int(row['Delay']) if pd.notna(row['Delay']) else 0
            entries[row_counter][3].insert(0, str(delay_value))
            
            data = row['Data']
            if isinstance(data, str):
                data = data.replace(" ", "")
                dlc = len(data) // 2
            else:
                dlc = 0
            
            entries[row_counter][1].insert(0, dlc)
            row_counter += 1
        
        print(f"Test case loaded successfully: {os.path.basename(file_path)}")
        return True
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load test case: {e}")
        return False

# Function to send a single message onto the CANbus
def send_single_message(bus, msg_id, dlc, data):
    try:
        msg = can.Message(
            arbitration_id=int(msg_id, 16),
            dlc=int(dlc),
            data=[int(byte, 16) for byte in data.split()],
            is_extended_id=False
        )
        bus.send(msg)
        print(f"Sent message with ID: {msg_id}, DLC: {dlc}, Data: {data}")
        return True
    except Exception as e:
        print(f"Error sending message: {e}")
        return False

# Function to send all selected CAN messages according to cycle settings
def send_all_messages(bus, entries, cycle_count, cycle_delay, window):
    """    
    Args:
        bus: CAN bus instance
        entries: List of message entry tuples
        cycle_count: Number of cycles to send messages
        cycle_delay: Delay between cycles in milliseconds
        window: Main window instance for showing error messages
    """
    global current_test_case_index, automatic_mode
    
    if automatic_mode:
        threading.Thread(target=run_automatic_mode, 
                       args=(window, bus, entries, cycle_count, cycle_delay),
                       daemon=True).start()
        return
        
    try:
        cycle_count = int(cycle_count)
        cycle_delay = int(cycle_delay)
    except ValueError:
        messagebox.showerror("Error", "Invalid input for cycle count or cycle delay.")
        return

    # Manual mode execution
    for cycle in range(cycle_count):
        for i in range(10):
            msg_id = entries[i][0].get()
            dlc = entries[i][1].get()
            data = entries[i][2].get()
            delay = entries[i][3].get()
            selected = entries[i][4].get()

            if selected and msg_id and dlc and data:
                while paused:
                    time.sleep(0.1)

                if not send_single_message(bus, msg_id, dlc, data):
                    messagebox.showerror("Error", f"Failed to send message {i+1}")
                    return
                
                if delay:
                    try:
                        time.sleep(int(delay) / 1000.0)
                    except ValueError:
                        messagebox.showerror("Error", f"Invalid delay value for message {i+1}.")
                        return
        
        time.sleep(cycle_delay / 1000.0)
        print(f"Cycle {cycle + 1}/{cycle_count} completed.")
    
    current_test_case_index += 1
    
    if current_test_case_index < len(test_case_files):
        if messagebox.askyesno("Test Case Complete", 
                            f"Current test case completed. Load next test case? ({current_test_case_index + 1}/{len(test_case_files)})"):
            load_test_case(entries)
        else:
            current_test_case_index = len(test_case_files)
    else:
        messagebox.showinfo("Complete", "All test cases have been completed!")

# Run automatic mode for sending CAN messages
def run_automatic_mode(window, bus, entries, cycle_count, cycle_delay):
    global automatic_mode, current_test_case_index
    
    while automatic_mode and current_test_case_index < len(test_case_files):
        try:
            cycle_count_val = int(cycle_count)
            cycle_delay_val = int(cycle_delay)
            
            for cycle in range(cycle_count_val):
                if not automatic_mode:
                    return
                
                for i in range(10):
                    msg_id = entries[i][0].get()
                    dlc = entries[i][1].get()
                    data = entries[i][2].get()
                    delay = entries[i][3].get()
                    selected = entries[i][4].get()

                    if selected and msg_id and dlc and data:
                        while paused:
                            time.sleep(0.1)
                            
                        if not send_single_message(bus, msg_id, dlc, data):
                            messagebox.showerror("Error", f"Failed to send message {i+1}")
                            automatic_mode = False
                            return
                        
                        if delay:
                            try:
                                time.sleep(int(delay) / 1000.0)
                            except ValueError:
                                messagebox.showerror("Error", f"Invalid delay value for message {i+1}")
                                automatic_mode = False
                                return
                
                time.sleep(cycle_delay_val / 1000.0)
                print(f"Cycle {cycle + 1}/{cycle_count_val} completed.")
                
            current_test_case_index += 1
            
            if current_test_case_index < len(test_case_files):
                load_test_case(entries)
                time.sleep(1)  # Add delay between test cases
            
        except Exception as e:
            messagebox.showerror("Error", f"Error in automatic mode: {str(e)}")
            automatic_mode = False
            return
    
    if automatic_mode:
        response = messagebox.askyesno("Complete", "All test cases completed! Would you like to select a new folder?")
        if response:
            current_test_case_index = 0
            if select_test_cases_folder():
                load_test_case(entries)
                run_automatic_mode(window, bus, entries, cycle_count, cycle_delay)
            else:
                automatic_mode = False
        else:
            automatic_mode = False

# Function to auto format Data bytes
def auto_format_data(event, data_entry, dlc_entry):
    data = data_entry.get().replace(" ", "").upper()
    formatted_data = ' '.join([data[i:i+2] for i in range(0, len(data), 2)])
    
    data_entry.delete(0, tk.END)
    data_entry.insert(0, formatted_data)
    
    dlc_value = len(formatted_data.split())
    dlc_entry.delete(0, tk.END)
    dlc_entry.insert(0, str(dlc_value))

# Function to create GUI for Test case Tx 
def create_gui(bus):
    window = tk.Tk()
    window.title("CAN Tx")
    
    # Mode selection frame
    mode_frame = ttk.LabelFrame(window, text="Operation Mode")
    mode_frame.grid(row=0, column=5, rowspan=2, padx=10, pady=5, sticky='nsew')
    
    # Add folder selection button
    folder_button = ttk.Button(
        mode_frame, 
        text="Upload Folder", 
        command=lambda: select_test_cases_folder() and load_test_case(entries)
    )
    folder_button.pack(pady=5, padx=5, fill='x')

    # Add automatic mode toggle button
    automatic_button = ttk.Button(
        mode_frame,
        text="Choose Mode!",
        command=lambda: toggle_automatic_mode(automatic_button)
    )
    automatic_button.pack(pady=5, padx=5, fill='x')
    
    # Main grid headers
    headers = ["Select", "ID (hex)", "DLC", "Data (hex)", "Delay (ms)"]
    for i, header in enumerate(headers):
        ttk.Label(window, text=header).grid(row=0, column=i, padx=10, pady=10)

    entries = []
    for i in range(10):
        selected_var = tk.BooleanVar(value=True)
        select_checkbox = ttk.Checkbutton(window, variable=selected_var)
        select_checkbox.grid(row=i+1, column=0, padx=10, pady=5)

        can_id_entry = ttk.Entry(window, width=10)
        can_id_entry.grid(row=i+1, column=1, padx=10, pady=5)

        dlc_entry = ttk.Entry(window, width=5)
        dlc_entry.grid(row=i+1, column=2, padx=10, pady=5)

        data_entry = ttk.Entry(window, width=30)
        data_entry.grid(row=i+1, column=3, padx=10, pady=5)
        data_entry.bind("<KeyRelease>", 
            lambda event, de=data_entry, dl=dlc_entry: auto_format_data(event, de, dl))

        delay_entry = ttk.Entry(window, width=10)
        delay_entry.grid(row=i+1, column=4, padx=10, pady=5)

        entries.append((can_id_entry, dlc_entry, data_entry, delay_entry, selected_var))

    # Control frame for cycle settings
    control_frame = ttk.LabelFrame(window, text="Cycle Settings")
    control_frame.grid(row=12, column=0, columnspan=5, padx=10, pady=5, sticky='ew')

    ttk.Label(control_frame, text="Cycle Count:").grid(row=0, column=0, padx=10, pady=5)
    cycle_count_entry = ttk.Entry(control_frame, width=10)
    cycle_count_entry.grid(row=0, column=1, padx=10, pady=5)
    cycle_count_entry.insert(0, "1")

    ttk.Label(control_frame, text="Cycle Delay (ms):").grid(row=0, column=2, padx=10, pady=5)
    cycle_delay_entry = ttk.Entry(control_frame, width=10)
    cycle_delay_entry.grid(row=0, column=3, padx=10, pady=5)
    cycle_delay_entry.insert(0, "0")

    # Send button
    send_button = ttk.Button(
        window, 
        text="Send All", 
        command=lambda: send_all_messages(
            bus, 
            entries, 
            cycle_count_entry.get(), 
            cycle_delay_entry.get(),
            window
        )
    )
    send_button.grid(row=13, column=0, columnspan=5, pady=10)

    keyboard.add_hotkey('ctrl+p', toggle_pause)
    window.mainloop()
    
# Continuously capture CAN log 
def capture_can_messages(bus, messages):
    global running
    capturing = True
    while running:
        if keyboard.is_pressed('esc'):
            capturing = not capturing  # Toggle capture state
            if capturing:
                print("Resuming CAN message capture...")
            else:
                print("Pausing CAN message capture...")
            time.sleep(1)  # Debounce to prevent multiple toggles from one press

        if capturing:
            msg = bus.recv(timeout=1)  # Capture message
            if msg:
                print(f"Received: {msg}")
                messages.append(msg)

# Gracefully disconnect and exit CAN logging
def handle_exit(signal, frame, messages):
    global running
    running = False
    print("\nExiting CAN message capture...")

    if messages:
        log_to_excel(messages, 'can_messages.xlsx')
    else:
        print("No CAN messages captured.")

    sys.exit(0)

# Function to launch the GUI when 'Alt + S' is pressed
def monitor_keyboard_for_popup(bus):
    while running:
        if keyboard.is_pressed('alt+s'):
            print("Opening message sender window...")
            threading.Thread(target=create_gui, args=(bus,), daemon=True).start()
            time.sleep(1)  # To debounce the 'Alt + S' key press

# Function for pausing and resuming Tx mid-process.
def toggle_pause():
    global paused
    paused = not paused
    print("Paused" if paused else "Resumed")

# Function to toggle between automatic and manual modes
def toggle_automatic_mode(button):
    global automatic_mode
    automatic_mode = not automatic_mode
    button.config(text="Automatic" if automatic_mode else "Manual")
    print("Switched to", "automatic" if automatic_mode else "manual", "mode")

# Main function to configure baud rate and capture CAN messages
def main():
    # Initialize CAN interface
    channel = '0'
    bitrate = 500000

    try:
        bus = init_can_interface(channel, bitrate)
        print(f"CAN interface initialized on channel {channel} with baud rate {bitrate} bps")
    except Exception as e:
        print(f"Failed to initialize CAN interface: {e}")
        return

    # List to store CAN messages
    messages = []

    # Register the signal handler for Ctrl + C (SIGINT)
    signal.signal(signal.SIGINT, lambda s, f: handle_exit(s, f, messages))

    # Start CAN message capture in a separate thread
    capture_thread = threading.Thread(target=capture_can_messages, args=(bus, messages), daemon=True)
    capture_thread.start()

    # Start GUI on detecting Alt + S
    monitor_keyboard_for_popup(bus)

    # Keep the main thread alive
    try:
        while running:
            time.sleep(0.1)
    except KeyboardInterrupt:
        handle_exit(None, None, messages)

if __name__ == "__main__":
    main()