import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import serial
import serial.tools.list_ports
import struct
import time
import csv
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import threading
from gs_usb.gs_usb import GsUsb
from gs_usb.gs_usb_frame import GsUsbFrame
from gs_usb.constants import CAN_EFF_FLAG, CAN_ERR_FLAG, CAN_RTR_FLAG

# --------------------
# Helper Functions
# --------------------
def get_can_id(command_id, vesc_id):
    """
    Constructs the 29-bit extended CAN ID based on command ID and VESC ID.
    """
    return (command_id << 8) | vesc_id

# --------------------
# Main Application Class
# --------------------
class CombinedLoggerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Modbus RTU & CAN Bus Logger")
        self.master.geometry("1200x800")
        
        # Initialize variables for Modbus RTU
        self.ser = None
        self.logging_modbus = False
        self.modbus_log_data = []
        self.modbus_time_data = []
        self.modbus_torque_data = []
        self.modbus_speed_data = []
        self.modbus_start_time = None
        self.modbus_job_id = None
        
        # Initialize variables for CAN Bus
        self.can_device = None
        self.can_active = False
        self.can_logging = False
        self.can_log_data = []
        self.can_plot_data = {'time': [], 'current': [], 'voltage': []}
        self.can_start_time = None
        self.can_read_thread = None
        self.can_send_thread = None
        self.can_lock = threading.Lock()
        
        # Setup GUI
        self.setup_gui()
        
    def setup_gui(self):
        # --------------------
        # FRAME: Modbus RTU Configuration
        # --------------------
        modbus_frame = tk.LabelFrame(self.master, text="Modbus RTU Configuration")
        modbus_frame.pack(padx=10, pady=10, fill=tk.X)
        
        # COM Port Selection
        tk.Label(modbus_frame, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.modbus_port_var = tk.StringVar()
        self.modbus_port_dropdown = ttk.Combobox(modbus_frame, textvariable=self.modbus_port_var, state='readonly', width=12)
        self.modbus_port_dropdown.grid(row=0, column=1, padx=5, pady=5)
        self.refresh_modbus_ports()
        self.modbus_refresh_button = tk.Button(modbus_frame, text="Refresh", command=self.refresh_modbus_ports)
        self.modbus_refresh_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Baud Rate Selection
        tk.Label(modbus_frame, text="Baud Rate:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.modbus_baud_var = tk.StringVar(value="38400")
        self.modbus_baud_dropdown = ttk.Combobox(
            modbus_frame,
            textvariable=self.modbus_baud_var,
            values=["9600", "19200", "38400", "57600", "115200"],
            state='readonly',
            width=12
        )
        self.modbus_baud_dropdown.grid(row=1, column=1, padx=5, pady=5)
        
        # Poll Interval
        tk.Label(modbus_frame, text="Poll Interval (ms):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.modbus_poll_interval_var = tk.StringVar(value="100")  # default 100ms
        tk.Entry(modbus_frame, textvariable=self.modbus_poll_interval_var, width=14).grid(row=2, column=1, padx=5, pady=5)
        
        # Start/Stop Logging Buttons
        self.modbus_start_button = tk.Button(modbus_frame, text="Start Logging", command=self.start_modbus_logging)
        self.modbus_start_button.grid(row=3, column=0, padx=5, pady=5)
        
        self.modbus_stop_button = tk.Button(modbus_frame, text="Stop Logging", command=self.stop_modbus_logging, state=tk.DISABLED)
        self.modbus_stop_button.grid(row=3, column=1, padx=5, pady=5)
        
        # --------------------
        # FRAME: CAN Bus Configuration
        # --------------------
        can_frame = tk.LabelFrame(self.master, text="CAN Bus Configuration")
        can_frame.pack(padx=10, pady=10, fill=tk.X)
        
        # CAN Bitrate Selection
        tk.Label(can_frame, text="CAN Bitrate (bps):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.can_bitrate_var = tk.StringVar()
        self.can_bitrate_dropdown = ttk.Combobox(can_frame, textvariable=self.can_bitrate_var, state='readonly', width=12)
        self.can_bitrate_dropdown['values'] = [125000, 250000, 500000, 1000000]  # Common CAN bitrates
        self.can_bitrate_dropdown.current(3)  # Default to 1 Mbps
        self.can_bitrate_dropdown.grid(row=0, column=1, padx=5, pady=5)
        
        # Connect/Disconnect Buttons
        self.can_connect_button = tk.Button(can_frame, text="Connect", command=self.connect_can)
        self.can_connect_button.grid(row=1, column=0, padx=5, pady=5)
        
        self.can_disconnect_button = tk.Button(can_frame, text="Disconnect", command=self.disconnect_can, state=tk.DISABLED)
        self.can_disconnect_button.grid(row=1, column=1, padx=5, pady=5)
        
        # Start/Stop Logging Buttons
        self.can_start_log_button = tk.Button(can_frame, text="Start Logging", command=self.start_can_logging, state=tk.DISABLED)
        self.can_start_log_button.grid(row=2, column=0, padx=5, pady=5)
        
        self.can_stop_log_button = tk.Button(can_frame, text="Stop Logging", command=self.stop_can_logging, state=tk.DISABLED)
        self.can_stop_log_button.grid(row=2, column=1, padx=5, pady=5)
        
        # --------------------
        # FRAME: Live Readings
        # --------------------
        readings_frame = tk.LabelFrame(self.master, text="Live Readings")
        readings_frame.pack(padx=10, pady=10, fill=tk.X)
        
        # Modbus Readings
        modbus_readings = tk.LabelFrame(readings_frame, text="Modbus RTU (Dynamometer)", padx=10, pady=10)
        modbus_readings.pack(side=tk.LEFT, padx=20, pady=10, fill="both", expand=True)
        
        tk.Label(modbus_readings, text="Torque (Nm):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.modbus_torque_label = tk.Label(modbus_readings, text="-- Nm", width=12, anchor="w")
        self.modbus_torque_label.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(modbus_readings, text="Speed (RPM):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.modbus_speed_label = tk.Label(modbus_readings, text="-- RPM", width=12, anchor="w")
        self.modbus_speed_label.grid(row=1, column=1, padx=5, pady=5)
        
        # CAN Readings
        can_readings = tk.LabelFrame(readings_frame, text="CAN Bus (VESC)", padx=10, pady=10)
        can_readings.pack(side=tk.LEFT, padx=20, pady=10, fill="both", expand=True)
        
        tk.Label(can_readings, text="Current (A):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.can_current_label = tk.Label(can_readings, text="-- A", width=12, anchor="w")
        self.can_current_label.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(can_readings, text="Voltage (V):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.can_voltage_label = tk.Label(can_readings, text="-- V", width=12, anchor="w")
        self.can_voltage_label.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(can_readings, text="RPM:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.can_rpm_label = tk.Label(can_readings, text="-- RPM", width=12, anchor="w")
        self.can_rpm_label.grid(row=2, column=1, padx=5, pady=5)
        
        # --------------------
        # FRAME: Plotting
        # --------------------
        plot_frame = tk.LabelFrame(self.master, text="Real-Time Plots")
        plot_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Matplotlib Figure
        self.figure = Figure(figsize=(10, 6), dpi=100)
        self.ax_modbus = self.figure.add_subplot(211)
        self.ax_modbus.set_title("Modbus RTU Data Over Time")
        self.ax_modbus.set_xlabel("Time (s)")
        self.ax_modbus.set_ylabel("Torque (Nm) / Speed (RPM)")
        self.ax_modbus.grid(True)
        self.ax_modbus.legend(['Torque (Nm)', 'Speed (RPM)'])
        
        self.line_modbus_torque, = self.ax_modbus.plot([], [], label='Torque (Nm)', color='blue')
        self.line_modbus_speed, = self.ax_modbus.plot([], [], label='Speed (RPM)', color='green')
        
        self.ax_can = self.figure.add_subplot(212)
        self.ax_can.set_title("CAN Bus Data Over Time")
        self.ax_can.set_xlabel("Time (s)")
        self.ax_can.set_ylabel("Current (A) / Voltage (V)")
        self.ax_can.grid(True)
        self.ax_can.legend(['Current (A)', 'Voltage (V)'])
        
        self.line_can_current, = self.ax_can.plot([], [], label='Current (A)', color='red')
        self.line_can_voltage, = self.ax_can.plot([], [], label='Voltage (V)', color='orange')
        
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # --------------------
        # FRAME: Logging Controls
        # --------------------
        logging_frame = tk.LabelFrame(self.master, text="Data Logging Controls")
        logging_frame.pack(padx=10, pady=10, fill=tk.X)
        
        # Save Buttons
        self.save_modbus_button = tk.Button(logging_frame, text="Save Modbus Data to CSV", command=self.save_modbus_csv, state=tk.DISABLED)
        self.save_modbus_button.pack(side=tk.LEFT, padx=20, pady=5)
        
        self.save_can_button = tk.Button(logging_frame, text="Save CAN Data to CSV", command=self.save_can_csv, state=tk.DISABLED)
        self.save_can_button.pack(side=tk.LEFT, padx=20, pady=5)
        
        # --------------------
        # FRAME: Output Text
        # --------------------
        output_frame = tk.LabelFrame(self.master, text="Last Raw Responses")
        output_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Modbus Output
        modbus_output = tk.LabelFrame(output_frame, text="Modbus RTU (Dynamometer)", padx=10, pady=10)
        modbus_output.pack(side=tk.LEFT, padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        self.modbus_output_text = tk.Text(modbus_output, height=10, width=50)
        self.modbus_output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        modbus_scrollbar = ttk.Scrollbar(modbus_output, command=self.modbus_output_text.yview)
        modbus_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.modbus_output_text['yscrollcommand'] = modbus_scrollbar.set
        
        # CAN Output
        can_output = tk.LabelFrame(output_frame, text="CAN Bus (VESC)", padx=10, pady=10)
        can_output.pack(side=tk.LEFT, padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        self.can_output_text = tk.Text(can_output, height=10, width=50)
        self.can_output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        can_scrollbar = ttk.Scrollbar(can_output, command=self.can_output_text.yview)
        can_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.can_output_text['yscrollcommand'] = can_scrollbar.set
        
        # --------------------
        # Status Label
        # --------------------
        self.status_label = tk.Label(self.master, text="Ready", fg="green", wraplength=1100)
        self.status_label.pack(pady=5)
    
    # --------------------
    # Modbus RTU Functions
    # --------------------
    def refresh_modbus_ports(self):
        """Populate COM ports into the dropdown for Modbus RTU."""
        ports = serial.tools.list_ports.comports()
        port_names = [p.device for p in ports]
        self.modbus_port_dropdown['values'] = port_names
        if port_names:
            self.modbus_port_var.set(port_names[0])
            self.status_label.config(text=f"Available Modbus COM Ports: {port_names}")
        else:
            self.modbus_port_var.set("")
            self.status_label.config(text="No Modbus COM ports found.")
    
    def start_modbus_logging(self):
        """Open serial port if needed and begin polling at user-defined interval."""
        if self.logging_modbus:
            self.status_label.config(text="Modbus logging already active.", fg="orange")
            return
    
        port = self.modbus_port_var.get()
        baud = int(self.modbus_baud_var.get())
        poll_str = self.modbus_poll_interval_var.get()
        try:
            poll_interval = int(poll_str)
        except ValueError:
            poll_interval = 100  # fallback to default
    
        # Convert poll_interval to a minimum threshold if too small
        if poll_interval < 10:
            poll_interval = 10
    
        self.modbus_polling_interval = poll_interval / 1000.0  # Convert ms to seconds
    
        if not self.ser or not self.ser.is_open:
            try:
                self.ser = serial.Serial(
                    port=port,
                    baudrate=baud,
                    bytesize=8,
                    parity=serial.PARITY_NONE,
                    stopbits=1,
                    timeout=0.05
                )
                self.status_label.config(text=f"Modbus connected to {port} at {baud} baud.", fg="green")
                self.modbus_output_text.delete("1.0", tk.END)
            except Exception as e:
                self.status_label.config(text=f"Modbus connection error: {e}", fg="red")
                self.ser = None
                return
    
        # Clear previous data
        self.modbus_time_data.clear()
        self.modbus_torque_data.clear()
        self.modbus_speed_data.clear()
        self.modbus_log_data.clear()
        self.modbus_start_time = time.time()
    
        self.logging_modbus = True
        self.modbus_start_button.config(state=tk.DISABLED)
        self.modbus_stop_button.config(state=tk.NORMAL)
        self.save_modbus_button.config(state=tk.DISABLED)
    
        self.status_label.config(text="Modbus logging started.", fg="blue")
        self.schedule_modbus_update()
    
    def stop_modbus_logging(self):
        """Stop active Modbus polling."""
        if self.logging_modbus:
            self.logging_modbus = False
            if self.modbus_job_id is not None:
                self.master.after_cancel(self.modbus_job_id)
                self.modbus_job_id = None
            self.modbus_start_button.config(state=tk.NORMAL)
            self.modbus_stop_button.config(state=tk.DISABLED)
            self.save_modbus_button.config(state=tk.NORMAL)
            self.status_label.config(text="Modbus logging stopped.", fg="green")
    
    def schedule_modbus_update(self):
        """Repeatedly poll torque & speed, update the plot, if logging is active."""
        if not self.logging_modbus:
            return
        self.update_modbus_readings()
        self.modbus_job_id = self.master.after(int(self.modbus_polling_interval * 1000), self.schedule_modbus_update)
    
    def update_modbus_readings(self):
        """Send Modbus request frames, parse responses, update labels and plot."""
        current_time = time.time() - self.modbus_start_time
        
        # 1) Read Torque
        # Using request frame: 01 03 00 00 00 02 C4 0B
        torque_frame_hex = "01 03 00 00 00 02 C4 0B"
        torque_val = self.send_modbus_frame(torque_frame_hex, parse_scale=100.0, is_speed=False)
        if torque_val is not None:
            torque_nm = torque_val / 100.0
            self.modbus_torque_label.config(text=f"{torque_nm:.2f} Nm")
        else:
            torque_nm = None
    
        # 2) Read Speed
        # Using request frame: 01 03 00 02 00 02 65 CB
        speed_frame_hex = "01 03 00 02 00 02 65 CB"
        speed_val = self.send_modbus_frame(speed_frame_hex, parse_scale=10.0, is_speed=True)
        if speed_val is not None:
            speed_rpm = speed_val / 10.0
            self.modbus_speed_label.config(text=f"{speed_rpm:.1f} RPM")
        else:
            speed_rpm = None
    
        # If both readings are valid, log & update plot
        if torque_nm is not None and speed_rpm is not None:
            self.modbus_time_data.append(current_time)
            self.modbus_torque_data.append(torque_nm)
            self.modbus_speed_data.append(speed_rpm)
            self.modbus_log_data.append({
                'timestamp': current_time,
                'torque': torque_nm,
                'speed': speed_rpm
            })
            
            # Update matplotlib plot
            self.line_modbus_torque.set_data(self.modbus_time_data, self.modbus_torque_data)
            self.line_modbus_speed.set_data(self.modbus_time_data, self.modbus_speed_data)
    
            # Adjust axes
            self.ax_modbus.set_xlim(0, max(self.modbus_time_data) + 1)
            y_min = min(min(self.modbus_torque_data, default=0), min(self.modbus_speed_data, default=0)) * 0.9
            y_max = max(max(self.modbus_torque_data, default=0), max(self.modbus_speed_data, default=0)) * 1.1
            self.ax_modbus.set_ylim(y_min, y_max)
    
            self.canvas.draw()
    
    def send_modbus_frame(self, frame_hex, parse_scale=1.0, is_speed=False):
        """
        Writes the raw Modbus frame, reads the response, parses it.
        Returns the parsed value or None.
        """
        if not self.ser or not self.ser.is_open:
            return None
    
        frame_bytes = bytes.fromhex(frame_hex.replace(" ", ""))
        try:
            self.ser.reset_input_buffer()
            self.ser.write(frame_bytes)
            self.ser.flush()
    
            start_time = time.time()
            response_buffer = bytearray()
    
            while (time.time() - start_time) < 0.05:
                chunk = self.ser.read(64)
                if chunk:
                    response_buffer.extend(chunk)
                else:
                    time.sleep(0.005)
    
            if response_buffer:
                hex_resp = response_buffer.hex(sep=' ')
                if is_speed:
                    self.modbus_output_text.delete("1.0", tk.END)
                    self.modbus_output_text.insert(tk.END, f"Modbus Speed Response ({len(response_buffer)} bytes):\n{hex_resp}\n")
                else:
                    self.modbus_output_text.delete("1.0", tk.END)
                    self.modbus_output_text.insert(tk.END, f"Modbus Torque Response ({len(response_buffer)} bytes):\n{hex_resp}\n")
    
                if len(response_buffer) >= 9:
                    byte_count = response_buffer[2]
                    if byte_count == 4:
                        data_bytes = response_buffer[3:7]
                        val_16 = struct.unpack('>H', data_bytes[2:4])[0]
                        return val_16
                    else:
                        if is_speed:
                            self.modbus_output_text.insert(tk.END, f"Invalid byte count for Speed: {byte_count}\n")
                        else:
                            self.modbus_output_text.insert(tk.END, f"Invalid byte count for Torque: {byte_count}\n")
                else:
                    if is_speed:
                        self.modbus_output_text.insert(tk.END, "Speed Response too short.\n")
                    else:
                        self.modbus_output_text.insert(tk.END, "Torque Response too short.\n")
            else:
                if is_speed:
                    self.modbus_output_text.delete("1.0", tk.END)
                    self.modbus_output_text.insert(tk.END, "No Speed response.\n")
                else:
                    self.modbus_output_text.delete("1.0", tk.END)
                    self.modbus_output_text.insert(tk.END, "No Torque response.\n")
    
        except Exception as e:
            if is_speed:
                self.modbus_output_text.delete("1.0", tk.END)
                self.modbus_output_text.insert(tk.END, f"Error sending Speed frame: {e}\n")
            else:
                self.modbus_output_text.delete("1.0", tk.END)
                self.modbus_output_text.insert(tk.END, f"Error sending Torque frame: {e}\n")
    
        return None
    
    def save_modbus_csv(self):
        """Save the Modbus log data to a CSV file."""
        if not self.modbus_log_data:
            messagebox.showinfo("Info", "No Modbus data to save.")
            return
    
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")],
                                                 title="Save Modbus Data")
        if not file_path:
            return
    
        try:
            with open(file_path, mode='w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Timestamp (s)", "Torque (Nm)", "Speed (RPM)"])
                for entry in self.modbus_log_data:
                    writer.writerow([f"{entry['timestamp']:.3f}", f"{entry['torque']:.3f}", f"{entry['speed']:.3f}"])
            self.status_label.config(text=f"Modbus data saved to {file_path}", fg="green")
        except Exception as e:
            self.status_label.config(text=f"Error saving Modbus CSV: {e}", fg="red")
    
    # --------------------
    # CAN Bus Functions
    # --------------------
    def connect_can(self):
        """Initialize and start CAN device."""
        if self.can_active:
            messagebox.showinfo("Info", "CAN device already connected.")
            return
        
        try:
            self.can_device = GsUsb.scan()
            if len(self.can_device) == 0:
                self.status_label.config(text="No CAN device detected.", fg="red")
                messagebox.showerror("Error", "No CAN device detected.")
                return
            self.can_device = self.can_device[0]
            print(f"Connected to CAN device: {self.can_device}")
            
            # Set bitrate
            bitrate = int(self.can_bitrate_var.get())
            if not self.can_device.set_bitrate(bitrate):
                self.status_label.config(text="Failed to set CAN bitrate.", fg="red")
                messagebox.showerror("Error", "Failed to set CAN bitrate.")
                self.can_device = None
                return
            
            # Start CAN device in NORMAL mode
            self.can_device.start(0)  # NORMAL mode
            # For testing without a CAN bus, use loopback mode:
            # self.can_device.start(GS_CAN_MODE_LOOP_BACK)
            
            self.can_active = True
            self.status_label.config(text="CAN device connected and initialized.", fg="green")
            self.can_connect_button.config(state=tk.DISABLED)
            self.can_disconnect_button.config(state=tk.NORMAL)
            self.can_start_log_button.config(state=tk.NORMAL)
            
            # Initialize CAN IDs
            self.CAN_PACKET_STATUS = 9
            self.CAN_PACKET_STATUS_5 = 27
            self.VESC_ID = 1  # Ensure this matches your VESC's CAN ID
            self.CAN_ID_STATUS = get_can_id(self.CAN_PACKET_STATUS, self.VESC_ID)
            self.CAN_ID_STATUS_5 = get_can_id(self.CAN_PACKET_STATUS_5, self.VESC_ID)
            
        except Exception as e:
            self.status_label.config(text=f"CAN Connection Error: {e}", fg="red")
            messagebox.showerror("Error", f"CAN Connection Error: {e}")
    
    def disconnect_can(self):
        """Stop CAN device."""
        if not self.can_active:
            messagebox.showinfo("Info", "CAN device not connected.")
            return
        
        self.can_active = False
        self.can_logging = False
        self.can_start_log_button.config(state=tk.NORMAL)
        self.can_stop_log_button.config(state=tk.DISABLED)
        self.save_can_button.config(state=tk.NORMAL)
        
        try:
            if self.can_device:
                self.can_device.stop()
                print("CAN device stopped.")
        except Exception as e:
            print(f"Error stopping CAN device: {e}")
        
        self.can_device = None
        self.status_label.config(text="CAN device disconnected.", fg="red")
        self.can_connect_button.config(state=tk.NORMAL)
        self.can_disconnect_button.config(state=tk.DISABLED)
    
    def start_can_logging(self):
        """Start logging CAN data."""
        if not self.can_active:
            messagebox.showerror("Error", "CAN device not connected.")
            return
        
        if self.can_logging:
            self.status_label.config(text="CAN logging already active.", fg="orange")
            return
        
        self.can_logging = True
        self.can_start_log_button.config(state=tk.DISABLED)
        self.can_stop_log_button.config(state=tk.NORMAL)
        self.save_can_button.config(state=tk.DISABLED)
        self.can_log_data.clear()
        self.can_plot_data = {'time': [], 'current': [], 'voltage': []}
        self.can_start_time = time.time()
        
        # Start reading thread
        self.can_read_thread = threading.Thread(target=self.read_can_data, daemon=True)
        self.can_read_thread.start()
        
        # Start keep-alive thread
        self.can_send_thread = threading.Thread(target=self.send_keep_alive, daemon=True)
        self.can_send_thread.start()
        
        self.status_label.config(text="CAN logging started.", fg="blue")
    
    def stop_can_logging(self):
        """Stop logging CAN data."""
        if self.can_logging:
            self.can_logging = False
            self.can_start_log_button.config(state=tk.NORMAL)
            self.can_stop_log_button.config(state=tk.DISABLED)
            self.save_can_button.config(state=tk.NORMAL)
            self.status_label.config(text="CAN logging stopped.", fg="green")
    
    def send_keep_alive(self):
        """Periodically send keep-alive commands to VESC to prevent timeout."""
        while self.can_logging and self.can_active and self.can_device:
            try:
                # Construct CAN ID for CAN_PACKET_SET_CURRENT_REL
                command_id = 10  # CAN_PACKET_SET_CURRENT_REL
                can_id = get_can_id(command_id, self.VESC_ID)
                
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
                self.can_device.send(frame)
                print(f"Sent keep-alive: CAN ID=0x{can_id:04X}, Data={data.hex()}")
            except Exception as e:
                self.status_label.config(text=f"CAN Send Error: {e}", fg="red")
                self.stop_can_logging()
                break
            time.sleep(0.02)  # 20 ms for 50 Hz
    
    def read_can_data(self):
        """Continuously read CAN frames and update GUI."""
        while self.can_logging and self.can_active and self.can_device:
            try:
                frame = GsUsbFrame()
                if self.can_device.read(frame, 0.01):  # 10 ms timeout
                    if frame.is_extended and frame.echo_id == 0xFFFFFFFF and not (frame.can_id & CAN_ERR_FLAG):
                        if frame.can_id == self.CAN_ID_STATUS:
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
                                
                                # Update GUI
                                self.can_current_label.config(text=f"{current:.2f} A")
                                self.can_rpm_label.config(text=f"{rpm} RPM")
                                
                                # Log data
                                timestamp = time.time() - self.can_start_time
                                with self.can_lock:
                                    self.can_log_data.append({
                                        'timestamp': timestamp,
                                        'current': current,
                                        'voltage': self.can_plot_data['voltage'][-1] if self.can_plot_data['voltage'] else 0,
                                        'rpm': rpm
                                    })
                                    self.can_plot_data['time'].append(timestamp)
                                    self.can_plot_data['current'].append(current)
                                
                            else:
                                print("CAN_PACKET_STATUS frame has insufficient data.")
                        
                        elif frame.can_id == self.CAN_ID_STATUS_5:
                            # Parse CAN_PACKET_STATUS_5
                            # Bytes 0-3: Tachometer (EREV), unsigned 32-bit, scale 6 (Not used)
                            # Bytes 4-5: Voltage In (V), unsigned 16-bit, scale 10
                            if len(frame.data) >= 6:
                                # tach_raw = struct.unpack('>I', frame.data[0:4])[0]  # Not used
                                voltage_raw = struct.unpack('>H', frame.data[4:6])[0]
                                
                                voltage = voltage_raw / 10.0  # Scale
                                
                                # Update GUI
                                self.can_voltage_label.config(text=f"{voltage:.2f} V")
                                
                                # Log data
                                timestamp = time.time() - self.can_start_time
                                with self.can_lock:
                                    self.can_log_data.append({
                                        'timestamp': timestamp,
                                        'current': self.can_plot_data['current'][-1] if self.can_plot_data['current'] else 0,
                                        'voltage': voltage,
                                        'rpm': self.can_rpm_label.cget("text").split()[0]  # Extract RPM value
                                    })
                                    self.can_plot_data['time'].append(timestamp)
                                    self.can_plot_data['voltage'].append(voltage)
                                
                            else:
                                print("CAN_PACKET_STATUS_5 frame has insufficient data.")
                
                # Update Plot
                self.update_can_plot()
                
            except Exception as e:
                self.status_label.config(text=f"CAN Read Error: {e}", fg="red")
                self.stop_can_logging()
                break
    
    def update_can_plot(self):
        """Update the CAN plot with new data."""
        with self.can_lock:
            time_data = self.can_plot_data['time']
            current_data = self.can_plot_data['current']
            voltage_data = self.can_plot_data['voltage']
        
        self.line_can_current.set_data(time_data, current_data)
        self.line_can_voltage.set_data(time_data, voltage_data)
        
        # Adjust axes
        if time_data:
            self.ax_can.set_xlim(0, max(time_data) + 1)
        if current_data or voltage_data:
            y_min = min(current_data + voltage_data, default=0) * 0.9
            y_max = max(current_data + voltage_data, default=0) * 1.1
            self.ax_can.set_ylim(y_min, y_max)
        
        self.canvas.draw()
    
    def save_can_csv(self):
        """Save the CAN log data to a CSV file."""
        if not self.can_log_data:
            messagebox.showinfo("Info", "No CAN data to save.")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")],
                                                 title="Save CAN Data")
        if not file_path:
            return
        
        try:
            with open(file_path, mode='w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Timestamp (s)", "Current (A)", "Voltage (V)", "RPM"])
                for entry in self.can_log_data:
                    writer.writerow([f"{entry['timestamp']:.3f}", f"{entry['current']:.3f}", f"{entry['voltage']:.3f}", f"{entry['rpm']:.1f}"])
            self.status_label.config(text=f"CAN data saved to {file_path}", fg="green")
        except Exception as e:
            self.status_label.config(text=f"Error saving CAN CSV: {e}", fg="red")
    
    # --------------------
    # Saving Functions
    # --------------------
    
    # (Already implemented separate save functions for Modbus and CAN)
    
    # --------------------
    # Cleanup on Exit
    # --------------------
    def on_closing(self):
        """Handle application closing."""
        if self.logging_modbus:
            self.stop_modbus_logging()
        if self.can_logging:
            self.stop_can_logging()
        if self.ser and self.ser.is_open:
            self.ser.close()
        if self.can_active and self.can_device:
            try:
                self.can_device.stop()
            except:
                pass
        self.master.destroy()
    
# --------------------
# Run Application
# --------------------
def main():
    root = tk.Tk()
    app = CombinedLoggerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
