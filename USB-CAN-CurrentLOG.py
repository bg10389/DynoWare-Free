import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from gs_usb.gs_usb import GsUsb
from gs_usb.gs_usb_frame import GsUsbFrame
from gs_usb.constants import (
    CAN_EFF_FLAG,
    CAN_ERR_FLAG,
    CAN_RTR_FLAG,
)
import struct
import threading
import time
import matplotlib
import csv
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from datetime import datetime

# Additional imports for COM ports
import serial
import serial.tools.list_ports

# Use the TkAgg backend
matplotlib.use('TkAgg')

# Constants
GS_USB_NONE_ECHO_ID = 0xFFFFFFFF

GS_CAN_MODE_NORMAL = 0
GS_CAN_MODE_LISTEN_ONLY = (1 << 0)
GS_CAN_MODE_LOOP_BACK = (1 << 1)
GS_CAN_MODE_ONE_SHOT = (1 << 3)
GS_CAN_MODE_HW_TIMESTAMP = (1 << 4)

# Define VESC ID (Set this to your VESC's CAN ID configured in VESC Tool)
VESC_ID = 1  # Example: 1

# Define Command IDs for status messages
CAN_PACKET_STATUS = 9
CAN_PACKET_STATUS_5 = 27

# Calculate the full CAN ID with command ID and VESC ID
def get_can_id(command_id, vesc_id):
    """
    Constructs the 29-bit extended CAN ID based on command ID and VESC ID.
    """
    return (command_id << 8) | vesc_id

CAN_ID_STATUS = get_can_id(CAN_PACKET_STATUS, VESC_ID)      # Example: 0x0901
CAN_ID_STATUS_5 = get_can_id(CAN_PACKET_STATUS_5, VESC_ID)  # Example: 0x1B01

# Define available baud rates (bitrates) for CAN
AVAILABLE_BITRATES = [
    125000,  # 125 kbps
    250000,  # 250 kbps
    500000,  # 500 kbps
    1000000  # 1 Mbps
]

class VESCReader:
    def __init__(self, root):
        self.root = root
        self.root.title("VESC Real-Time Data (Current, Voltage, RPM)")
        self.root.geometry("900x700")

        # Initialize variables
        self.device = None
        self.active = False
        self.logging = False
        self.log_data = []
        self.plot_data = {'time': [], 'current': [], 'voltage': []}
        self.start_time = None
        self.lock = threading.Lock()

        # Setup GUI
        self.setup_gui()

    def setup_gui(self):
        # Top Frame for Controls
        control_frame = tk.Frame(self.root)
        control_frame.pack(pady=10)

        # COM Port Selection (Newly added)
        com_label = tk.Label(control_frame, text="Select COM Port:")
        com_label.pack(side=tk.LEFT, padx=5)

        self.com_var = tk.StringVar()
        ports = serial.tools.list_ports.comports()
        port_list = [p.device for p in ports] if ports else []
        self.com_combobox = ttk.Combobox(control_frame, textvariable=self.com_var, state="readonly", width=10)
        if port_list:
            self.com_combobox['values'] = port_list
            self.com_combobox.current(0)
        else:
            self.com_combobox['values'] = ["No COM ports found"]
            self.com_combobox.current(0)
        self.com_combobox.pack(side=tk.LEFT, padx=5)

        # Baud Rate Selection
        baud_label = tk.Label(control_frame, text="Select CAN Bitrate (bps):")
        baud_label.pack(side=tk.LEFT, padx=5)

        self.baud_var = tk.StringVar()
        self.baud_combobox = ttk.Combobox(control_frame, textvariable=self.baud_var, state="readonly", width=10)
        self.baud_combobox['values'] = AVAILABLE_BITRATES
        self.baud_combobox.current(3)  # Default to 1 Mbps
        self.baud_combobox.pack(side=tk.LEFT, padx=5)

        # Connect Button
        self.connect_button = tk.Button(control_frame, text="Connect", command=self.connect)
        self.connect_button.pack(side=tk.LEFT, padx=5)

        # Disconnect Button
        self.disconnect_button = tk.Button(control_frame, text="Disconnect", command=self.disconnect, state=tk.DISABLED)
        self.disconnect_button.pack(side=tk.LEFT, padx=5)

        # Start Log Button
        self.start_log_button = tk.Button(control_frame, text="Start Log", command=self.start_logging, state=tk.DISABLED)
        self.start_log_button.pack(side=tk.LEFT, padx=5)

        # Stop Log Button
        self.stop_log_button = tk.Button(control_frame, text="Stop Log", command=self.stop_logging, state=tk.DISABLED)
        self.stop_log_button.pack(side=tk.LEFT, padx=5)

        # Save to CSV Button
        self.save_button = tk.Button(control_frame, text="Save to CSV", command=self.save_to_csv, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)

        # Status Label
        self.status_label = tk.Label(self.root, text="Not connected.", font=("Arial", 12), fg="red")
        self.status_label.pack(pady=5)

        # Data Display Frame with Separate Boxes
        data_display_frame = tk.Frame(self.root)
        data_display_frame.pack(pady=10)

        # Current Box
        current_box = tk.LabelFrame(data_display_frame, text="Current (A)", padx=10, pady=10)
        current_box.pack(side=tk.LEFT, padx=20, pady=10, fill="both", expand=True)
        self.current_label = tk.Label(current_box, text="-- A", font=("Arial", 16))
        self.current_label.pack()

        # Voltage Box
        voltage_box = tk.LabelFrame(data_display_frame, text="Voltage (V)", padx=10, pady=10)
        voltage_box.pack(side=tk.LEFT, padx=20, pady=10, fill="both", expand=True)
        self.voltage_label = tk.Label(voltage_box, text="-- V", font=("Arial", 16))
        self.voltage_label.pack()

        # RPM Box
        rpm_box = tk.LabelFrame(data_display_frame, text="RPM", padx=10, pady=10)
        rpm_box.pack(side=tk.LEFT, padx=20, pady=10, fill="both", expand=True)
        self.rpm_label = tk.Label(rpm_box, text="-- RPM", font=("Arial", 16))
        self.rpm_label.pack()

        # Matplotlib Figure
        self.figure = Figure(figsize=(8, 4), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_title("VESC Current and Voltage Over Time")
        self.ax.set_xlabel("Time (s)")
        self.ax.set_ylabel("Value")
        self.ax.grid(True)
        self.ax.legend(['Current (A)', 'Voltage (V)'])

        self.current_line, = self.ax.plot([], [], label='Current (A)', color='blue')
        self.voltage_line, = self.ax.plot([], [], label='Voltage (V)', color='red')

        self.ax.legend()

        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(pady=10)

    def connect(self):
        """
        Handles the Connect button click.
        Initializes the CAN device with the selected bitrate and starts reading and sending threads.
        """
        if self.active:
            messagebox.showinfo("Info", "Already connected.")
            return

        # Get selected bitrate
        try:
            bitrate = int(self.baud_var.get())
        except ValueError:
            messagebox.showerror("Error", "Please select a valid bitrate.")
            return

        # Initialize CAN device
        try:
            devs = GsUsb.scan()
            if len(devs) == 0:
                self.status_label.config(text="No gs_usb device detected.", fg="red")
                messagebox.showerror("Error", "No gs_usb device detected.")
                return

            self.device = devs[0]
            print(f"Connected to device: {self.device}")

            # Set the bitrate
            if not self.device.set_bitrate(bitrate):
                self.status_label.config(text="Failed to set CAN bitrate.", fg="red")
                messagebox.showerror("Error", "Failed to set CAN bitrate.")
                self.device = None
                return

            # Start the device in NORMAL mode
            self.device.start(GS_CAN_MODE_NORMAL)
            # For testing without a CAN bus, use loopback mode:
            # self.device.start(GS_CAN_MODE_LOOP_BACK)

            self.status_label.config(text="Connected and CAN device initialized successfully.", fg="green")
            self.connect_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.NORMAL)
            self.start_log_button.config(state=tk.NORMAL)

            self.active = True

            # Start reading and sending threads
            self.read_thread = threading.Thread(target=self.read_can_data, daemon=True)
            self.read_thread.start()

            self.send_thread = threading.Thread(target=self.send_keep_alive, daemon=True)
            self.send_thread.start()

        except Exception as e:
            self.status_label.config(text=f"Connection Error: {e}", fg="red")
            messagebox.showerror("Error", f"Connection Error: {e}")

    def disconnect(self):
        """
        Handles the Disconnect button click.
        Stops the reading and sending threads and closes the CAN device.
        """
        if not self.active:
            messagebox.showinfo("Info", "Not connected.")
            return

        self.active = False

        # Stop logging if active
        if self.logging:
            self.stop_logging()

        # Stop the device
        try:
            if self.device:
                self.device.stop()
                print("CAN device stopped.")
        except Exception as e:
            print(f"Stop Error: {e}")

        self.device = None

        self.status_label.config(text="Disconnected.", fg="red")
        self.connect_button.config(state=tk.NORMAL)
        self.disconnect_button.config(state=tk.DISABLED)
        self.start_log_button.config(state=tk.DISABLED)
        self.stop_log_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        self.baud_combobox.config(state=tk.NORMAL)

    def send_keep_alive(self):
        """
        Periodically sends a keep-alive command to prevent VESC timeout.
        Uses CAN_PACKET_SET_CURRENT_REL with zero current.
        Sends at 50 Hz (every 20 ms).
        """
        while self.active and self.device:
            try:
                # Construct CAN ID for CAN_PACKET_SET_CURRENT_REL
                command_id = 10  # CAN_PACKET_SET_CURRENT_REL
                can_id = get_can_id(command_id, VESC_ID)

                # Prepare data: float32 scaled by 100000, set to 0.0 (no current)
                current_rel = 0.0  # % / 100
                scaled_current_rel = int(current_rel * 100000.0)
                data = scaled_current_rel.to_bytes(4, byteorder='big', signed=True)

                # Create CAN frame
                frame = GsUsbFrame()
                frame.can_id = can_id
                frame.is_extended = True  # 29-bit
                frame.is_remote = False
                frame.data = data
                frame.dlc = 4

                # Send the frame
                self.device.send(frame)
                print(f"Sent keep-alive: CAN ID=0x{can_id:04X}, Data={data.hex()}")

            except Exception as e:
                print(f"Send Error: {e}")
                self.status_label.config(text=f"Send Error: {e}", fg="red")
                self.disconnect()
                return

            # Wait for 20 ms
            time.sleep(0.02)

    def read_can_data(self):
        """
        Continuously reads CAN frames and updates the GUI with Current, Voltage, and RPM.
        """
        while self.active and self.device:
            try:
                frame = GsUsbFrame()
                # Attempt to read with a short timeout (e.g., 10 ms)
                if self.device.read(frame, 0.01):
                    # Check if the frame is extended (29-bit) and not an error frame
                    if frame.is_extended and frame.echo_id == GS_USB_NONE_ECHO_ID and not (frame.can_id & CAN_ERR_FLAG):
                        # Check for CAN_PACKET_STATUS
                        if frame.can_id == CAN_ID_STATUS:
                            # Parse CAN_PACKET_STATUS
                            # Bytes 0-3: ERPM (RPM), unsigned 32-bit, scale 1
                            # Bytes 4-5: Current (A), signed 16-bit, scale 10
                            # Bytes 6-7: Duty Cycle (% / 100), unsigned 16-bit, scale 1000

                            if len(frame.data) >= 8:
                                erpm = struct.unpack('>I', frame.data[0:4])[0]
                                current_raw = struct.unpack('>h', frame.data[4:6])[0]
                                # duty_cycle_raw = struct.unpack('>H', frame.data[6:8])[0]  # Not used

                                current = current_raw / 10.0  # Scale
                                rpm = erpm  # Scale is 1

                                # Update GUI labels
                                self.current_label.config(text=f"{current:.2f} A")
                                self.rpm_label.config(text=f"{rpm} RPM")
                                print(f"Received Current: {current:.2f} A, RPM: {rpm} RPM")

                                # Log data if logging is active
                                if self.logging:
                                    with self.lock:
                                        timestamp = time.time() - self.start_time
                                        self.log_data.append({
                                            'timestamp': timestamp,
                                            'current': current,
                                            'voltage': self.plot_data['voltage'][-1] if self.plot_data['voltage'] else 0,
                                            'rpm': rpm
                                        })
                                        # Update plot data
                                        self.plot_data['time'].append(timestamp)
                                        self.plot_data['current'].append(current)

                            else:
                                print("CAN_PACKET_STATUS frame has insufficient data.")

                        # Check for CAN_PACKET_STATUS_5
                        elif frame.can_id == CAN_ID_STATUS_5:
                            # Parse CAN_PACKET_STATUS_5
                            # Bytes 0-3: Tachometer (EREV), unsigned 32-bit, scale 6 (Not used)
                            # Bytes 4-5: Voltage In (V), unsigned 16-bit, scale 10

                            if len(frame.data) >= 6:
                                # tach_raw = struct.unpack('>I', frame.data[0:4])[0]  # Not used
                                voltage_raw = struct.unpack('>H', frame.data[4:6])[0]

                                voltage = voltage_raw / 10.0  # Scale

                                # Update GUI label
                                self.voltage_label.config(text=f"{voltage:.2f} V")
                                print(f"Received Voltage: {voltage:.2f} V")

                                # Log data if logging is active
                                if self.logging:
                                    with self.lock:
                                        timestamp = time.time() - self.start_time
                                        self.log_data.append({
                                            'timestamp': timestamp,
                                            'current': self.plot_data['current'][-1] if self.plot_data['current'] else 0,
                                            'voltage': voltage,
                                            'rpm': self.rpm_label.cget("text").split()[0]  # Extract RPM value
                                        })
                                        # Update plot data
                                        self.plot_data['time'].append(timestamp)
                                        self.plot_data['voltage'].append(voltage)

                            else:
                                print("CAN_PACKET_STATUS_5 frame has insufficient data.")

            except Exception as e:
                print(f"Read Error: {e}")
                self.status_label.config(text=f"Read Error: {e}", fg="red")
                self.disconnect()
                return

    def start_logging(self):
        """
        Starts logging the data by clearing previous logs and enabling logging flags.
        """
        if not self.active:
            messagebox.showerror("Error", "Device not connected.")
            return

        if self.logging:
            messagebox.showinfo("Info", "Already logging.")
            return

        self.logging = True
        self.log_data = []
        self.start_time = time.time()
        self.plot_data = {'time': [], 'current': [], 'voltage': []}
        self.ax.cla()
        self.ax.set_title("VESC Current and Voltage Over Time")
        self.ax.set_xlabel("Time (s)")
        self.ax.set_ylabel("Value")
        self.ax.grid(True)
        self.ax.legend(['Current (A)', 'Voltage (V)'])

        self.current_line, = self.ax.plot([], [], label='Current (A)', color='blue')
        self.voltage_line, = self.ax.plot([], [], label='Voltage (V)', color='red')

        self.ax.legend()
        self.canvas.draw()

        self.status_label.config(text="Logging started.", fg="blue")
        self.start_log_button.config(state=tk.DISABLED)
        self.stop_log_button.config(state=tk.NORMAL)
        self.save_button.config(state=tk.DISABLED)

    def stop_logging(self):
        """
        Stops logging the data.
        """
        if not self.logging:
            messagebox.showinfo("Info", "Logging is not active.")
            return

        self.logging = False
        self.status_label.config(text="Logging stopped.", fg="green")
        self.start_log_button.config(state=tk.NORMAL)
        self.stop_log_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.NORMAL)

        # Update the plot with the final data
        with self.lock:
            self.current_line.set_data(self.plot_data['time'], self.plot_data['current'])
            self.voltage_line.set_data(self.plot_data['time'], self.plot_data['voltage'])

            if self.plot_data['time']:
                self.ax.set_xlim(0, max(self.plot_data['time']))
                y_min = min(min(self.plot_data['current'], default=0), min(self.plot_data['voltage'], default=0)) * 0.9
                y_max = max(max(self.plot_data['current'], default=0), max(self.plot_data['voltage'], default=0)) * 1.2
                self.ax.set_ylim(y_min, y_max)

            self.canvas.draw()

    def save_to_csv(self):
        """
        Saves the logged data to a CSV file.
        """
        if not self.log_data:
            messagebox.showinfo("Info", "No data to save.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")],
                                                 title="Save Log Data")
        if not file_path:
            return

        try:
            with open(file_path, mode='w', newline='') as csv_file:
                fieldnames = ['Timestamp (s)', 'Current (A)', 'Voltage (V)', 'RPM']
                writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

                writer.writeheader()
                for entry in self.log_data:
                    writer.writerow({
                        'Timestamp (s)': f"{entry['timestamp']:.2f}",
                        'Current (A)': f"{entry['current']:.2f}",
                        'Voltage (V)': f"{entry['voltage']:.2f}",
                        'RPM': entry['rpm']
                    })
            messagebox.showinfo("Success", f"Data saved to {file_path}")
            print(f"Data saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {e}")
            print(f"Failed to save data: {e}")

    def stop(self):
        """
        Stops the CAN device and any running threads.
        """
        self.active = False
        self.logging = False
        self.disconnect()

def main():
    root = tk.Tk()
    vesc_reader = VESCReader(root)

    def on_closing():
        vesc_reader.stop()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
