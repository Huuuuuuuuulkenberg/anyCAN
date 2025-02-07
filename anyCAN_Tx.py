import can
import sys
import time
import signal
import keyboard
import openpyxl
import threading
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import PhotoImage
from tkinter import messagebox

# Global flags
running = True
paused = False
global read_messages
read_messages = pd.DataFrame()

# Function to initialize CAN interface
def init_can_interface(channel, bitrate):
    bus = can.interface.Bus(channel=channel, interface='ixxat', bitrate=bitrate)
    return bus

from datetime import datetime, timedelta

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

# Function to load Excel file and populate GUI with Write messages
def load_test_case(entries):
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path)  # Read the Excel file

        # Clear existing GUI entries
        for entry in entries:
            for field in entry[:3]:  # Clear ID, DLC, and Data fields
                field.delete(0, tk.END)

        # Filter and process write messages only
        write_messages = df[df['Read/Write'].str.lower() == 'write']
        
        row_counter = 0  # Track the current row to populate
        for i, row in write_messages.iterrows():
            if row_counter >= len(entries):
                break  # Prevent overflow in the GUI rows
            entries[row_counter][0].insert(0, row['ID'])  # CAN ID
            entries[row_counter][2].insert(0, row['Data'])  # Data

            # Ensure Delay is an integer, without decimals
            delay_value = int(row['Delay']) if pd.notna(row['Delay']) else 0
            entries[row_counter][3].insert(0, str(delay_value))  # Delay (integer)

            # Auto-calculate DLC
            data = row['Data']
            if isinstance(data, str):  # Ensure Data is a string
                data = data.replace(" ", "")  # Remove any spaces
                dlc = len(data) // 2  # Each byte is 2 hex characters
            else:
                dlc = 0  # Default to 0 if no valid data

            entries[row_counter][1].insert(0, dlc)  # Populate DLC field

            row_counter += 1  # Move to the next row for GUI

        print("Test case loaded successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load test case: {e}")

# Function to send a single CAN message
def send_single_message(bus, msg_id, dlc, data):
    try:
        # Create and send the message
        msg = can.Message(
            arbitration_id=int(msg_id, 16),
            dlc=int(dlc),
            data=[int(byte, 16) for byte in data.split()],
            is_extended_id=False
        )
        bus.send(msg)
        print(f"Sent message with ID: {msg_id}, DLC: {dlc}, Data: {data}")
    except Exception as e:
        print(f"Error sending message: {e}")

# Function to send all CAN messages entered in the GUI table in sequence with delays
def send_all_messages(bus, entries, cycle_count, cycle_delay):
    global paused
    try:
        cycle_count = int(cycle_count)
        cycle_delay = int(cycle_delay)
    except ValueError:
        messagebox.showerror("Error", "Invalid input for cycle count or cycle delay.")
        return

    for cycle in range(cycle_count):
        for i in range(10):
            msg_id = entries[i][0].get()
            dlc = entries[i][1].get()
            data = entries[i][2].get()
            delay = entries[i][3].get()
            selected = entries[i][4].get()  # Check if the message is selected

            if selected and msg_id and dlc and data:
                # Pause the transmission if paused is True
                while paused:
                    time.sleep(0.1)  # Check every 100 ms if we are still paused

                # Send the message
                send_single_message(bus, msg_id, dlc, data)
                
                # If delay is specified, wait before sending the next message
                if delay:
                    try:
                        time.sleep(int(delay) / 1000.0)
                    except ValueError:
                        messagebox.showerror("Error", f"Invalid delay value for message {i+1}.")
                        return
        
        # After each cycle, wait for cycle delay
        time.sleep(cycle_delay / 1000.0)
        print(f"Cycle {cycle + 1}/{cycle_count} completed.")

    messagebox.showinfo("Success", "All messages sent successfully!")

# Function to toggle pause/resume
def toggle_pause():
    global paused
    paused = not paused
    print("Paused" if paused else "Resumed")

# Function to auto-format Data(hex) and auto-update DLC
def auto_format_data(event, data_entry, dlc_entry):
    data = data_entry.get().replace(" ", "").upper()  # Remove existing spaces and make uppercase
    formatted_data = ' '.join([data[i:i+2] for i in range(0, len(data), 2)])  # Add space every 2 characters
    
    # Update the data_entry field
    data_entry.delete(0, tk.END)
    data_entry.insert(0, formatted_data)

    # Calculate DLC
    dlc_value = len(formatted_data.split())  # Count the number of bytes (space-separated hex values)
    dlc_entry.delete(0, tk.END)
    dlc_entry.insert(0, str(dlc_value))

# Function to create the GUI for entering up to 10 messages
def create_gui(bus):
    window = tk.Tk()
    window.title("CAN Tx")
    window.iconphoto(True, PhotoImage(file="anyCAN.png")) #window icon

    # Column labels
    tk.Label(window, text="Select").grid(row=0, column=0, padx=10, pady=10)
    tk.Label(window, text="ID (hex)").grid(row=0, column=1, padx=10, pady=10)
    tk.Label(window, text="DLC").grid(row=0, column=2, padx=10, pady=10)
    tk.Label(window, text="Data (hex)").grid(row=0, column=3, padx=10, pady=10)
    tk.Label(window, text="Delay (ms)").grid(row=0, column=4, padx=10, pady=10)

    # List to hold the entries for the table
    entries = []
    for i in range(10):
        # Select checkbox
        selected_var = tk.BooleanVar(value=True)
        select_checkbox = tk.Checkbutton(window, variable=selected_var)
        select_checkbox.grid(row=i+1, column=0, padx=10, pady=5)

        # CAN ID entry
        can_id_entry = tk.Entry(window, width=10)
        can_id_entry.grid(row=i+1, column=1, padx=10, pady=5)

        # DLC entry
        dlc_entry = tk.Entry(window, width=5)
        dlc_entry.grid(row=i+1, column=2, padx=10, pady=5)

        # Data entry (hex space-separated)
        data_entry = tk.Entry(window, width=30)
        data_entry.grid(row=i+1, column=3, padx=10, pady=5)

        # Bind the auto_format_data function to the key release event for the data entry
        data_entry.bind("<KeyRelease>", lambda event, de=data_entry, dl=dlc_entry: auto_format_data(event, de, dl))

        # Delay entry (in ms)
        delay_entry = tk.Entry(window, width=10)
        delay_entry.grid(row=i+1, column=4, padx=10, pady=5)

        # Store entries in a list
        entries.append((can_id_entry, dlc_entry, data_entry, delay_entry, selected_var))

    # Cycle count field
    tk.Label(window, text="Cycle Count:").grid(row=12, column=0, padx=10, pady=10)
    cycle_count_entry = tk.Entry(window, width=10)
    cycle_count_entry.grid(row=12, column=1, padx=10, pady=10)
    cycle_count_entry.insert(0, "1")  # Set default cycle count to 1

    # Cycle delay field (in ms)
    tk.Label(window, text="Cycle Delay (ms):").grid(row=12, column=2, padx=10, pady=10)
    cycle_delay_entry = tk.Entry(window, width=10)
    cycle_delay_entry.grid(row=12, column=3, padx=10, pady=10)
    cycle_delay_entry.insert(0, "0")  # Set default cycle count to 0

    # Send button to send all messages in sequence with cycle count and delay
    send_button = tk.Button(window, text="Send All", command=lambda: send_all_messages(bus, entries, cycle_count_entry.get(), cycle_delay_entry.get()))
    send_button.grid(row=13, column=0, columnspan=4, pady=10)

    # Button to load test case
    load_button = tk.Button(window, text="Load Test Case", command=lambda: load_test_case(entries))
    load_button.grid(row=14, column=0, columnspan=4, pady=10)

    # Start a separate thread to listen for the Ctrl+P shortcut
    keyboard.add_hotkey('ctrl+p', toggle_pause)

    # Start the GUI loop
    window.mainloop()

# Function to capture CAN messages and display them
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

# Function to launch the GUI when 'S' is pressed
def monitor_keyboard_for_popup(bus):
    while running:
        if keyboard.is_pressed('alt+s'):
            print("Opening message sender window...")
            threading.Thread(target=create_gui, args=(bus,), daemon=True).start()
            time.sleep(1)  # To debounce the 'Alt+S' key press

# Function to handle graceful exit when Ctrl+C is pressed
def handle_exit(signal, frame, messages):
    global running
    running = False
    print("\nExiting CAN message capture...")

    # Log messages to Excel when exiting
    if messages:
        log_to_excel(messages, 'can_messages.xlsx')
    else:
        print("No CAN messages captured.")

    sys.exit(0)

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

    # Register the signal handler for Ctrl+C (SIGINT)
    signal.signal(signal.SIGINT, lambda s, f: handle_exit(s, f, messages))

    # Start CAN message capture in a separate thread
    capture_thread = threading.Thread(target=capture_can_messages, args=(bus, messages), daemon=True)
    capture_thread.start()

    # Monitor for 'S' key press to open the GUI
    monitor_keyboard_for_popup(bus)

    # Keep the main thread alive
    capture_thread.join()

if __name__ == "__main__":
    main()
