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
        self.master.geometry("1800x1000")  # Increased width for better layout

        # Initialize variables for Modbus RTU
        self.ser_modbus = None
        self.logging_modbus = False
        self.modbus_log_data = []
        self.modbus_time_data = []
        self.modbus_torque_data = []
        self.modbus_speed_data = []
        self.modbus_start_time = None
        self.modbus_job_id = None

        # Initialize variables for CAN Bus
        self.ser_can = None
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
        # FRAME: Small Segments (Configuration & Live Readings)
        # --------------------
        small_segments_frame = tk.Frame(self.master)
        small_segments_frame.pack(padx=10, pady=10, fill=tk.X)

        # --------------------
        # FRAME: Modbus RTU Configuration
        # --------------------
        modbus_frame = tk.LabelFrame(small_segments_frame, text="Modbus RTU Configuration", padx=10, pady=10)
        modbus_frame.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # COM Port Selection
        tk.Label(modbus_frame, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.modbus_port_var = tk.StringVar()
        self.modbus_port_dropdown = ttk.Combobox(modbus_frame, textvariable=self.modbus_port_var, state='readonly', width=15)
        self.modbus_port_dropdown.grid(row=0, column=1, padx=5, pady=5)
        self.modbus_refresh_button = tk.Button(modbus_frame, text="Refresh", command=self.refresh_modbus_ports, width=10)
        self.modbus_refresh_button.grid(row=0, column=2, padx=5, pady=5)

        # Baud Rate Selection
        tk.Label(modbus_frame, text="Baud Rate:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.modbus_baud_var = tk.StringVar(value="38400")
        self.modbus_baud_dropdown = ttk.Combobox(
            modbus_frame,
            textvariable=self.modbus_baud_var,
            values=["9600", "19200", "38400", "57600", "115200"],
            state='readonly',
            width=15
        )
        self.modbus_baud_dropdown.grid(row=1, column=1, padx=5, pady=5)

        # Connect/Disconnect Buttons
        self.modbus_connect_button = tk.Button(modbus_frame, text="Connect", command=self.connect_modbus, width=10)
        self.modbus_connect_button.grid(row=2, column=0, padx=5, pady=10)

        self.modbus_disconnect_button = tk.Button(modbus_frame, text="Disconnect", command=self.disconnect_modbus, state=tk.DISABLED, width=10)
        self.modbus_disconnect_button.grid(row=2, column=1, padx=5, pady=10)

        # Poll Interval
        tk.Label(modbus_frame, text="Poll Interval (ms):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.modbus_poll_interval_var = tk.StringVar(value="100")  # default 100ms
        tk.Entry(modbus_frame, textvariable=self.modbus_poll_interval_var, width=17).grid(row=3, column=1, padx=5, pady=5)

        # Start/Stop Logging Buttons
        self.modbus_start_button = tk.Button(modbus_frame, text="Start Logging", command=self.start_modbus_logging, state=tk.DISABLED, width=15)
        self.modbus_start_button.grid(row=4, column=0, padx=5, pady=10)

        self.modbus_stop_button = tk.Button(modbus_frame, text="Stop Logging", command=self.stop_modbus_logging, state=tk.DISABLED, width=15)
        self.modbus_stop_button.grid(row=4, column=1, padx=5, pady=10)

        # --------------------
        # FRAME: CAN Bus Configuration
        # --------------------
        can_frame = tk.LabelFrame(small_segments_frame, text="CAN Bus Configuration", padx=10, pady=10)
        can_frame.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # COM Port Selection
        tk.Label(can_frame, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.can_port_var = tk.StringVar()
        self.can_port_dropdown = ttk.Combobox(can_frame, textvariable=self.can_port_var, state='readonly', width=15)
        self.can_port_dropdown.grid(row=0, column=1, padx=5, pady=5)
        self.can_refresh_button = tk.Button(can_frame, text="Refresh", command=self.refresh_can_ports, width=10)
        self.can_refresh_button.grid(row=0, column=2, padx=5, pady=5)

        # Baud Rate Selection
        tk.Label(can_frame, text="Baud Rate:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.can_baud_var = tk.StringVar(value="9600")
        self.can_baud_dropdown = ttk.Combobox(
            can_frame,
            textvariable=self.can_baud_var,
            values=["9600", "19200", "38400", "57600", "115200"],
            state='readonly',
            width=15
        )
        self.can_baud_dropdown.grid(row=1, column=1, padx=5, pady=5)

        # Connect/Disconnect Buttons
        self.can_connect_button = tk.Button(can_frame, text="Connect", command=self.connect_can, width=10)
        self.can_connect_button.grid(row=2, column=0, padx=5, pady=10)

        self.can_disconnect_button = tk.Button(can_frame, text="Disconnect", command=self.disconnect_can, state=tk.DISABLED, width=10)
        self.can_disconnect_button.grid(row=2, column=1, padx=5, pady=10)

        # Start/Stop Logging Buttons
        self.can_start_log_button = tk.Button(can_frame, text="Start Logging", command=self.start_can_logging, state=tk.DISABLED, width=15)
        self.can_start_log_button.grid(row=3, column=0, padx=5, pady=10)

        self.can_stop_log_button = tk.Button(can_frame, text="Stop Logging", command=self.stop_can_logging, state=tk.DISABLED, width=15)
        self.can_stop_log_button.grid(row=3, column=1, padx=5, pady=10)

        # --------------------
        # FRAME: Live Readings
        # --------------------
        live_readings_frame = tk.Frame(self.master)
        live_readings_frame.pack(padx=10, pady=10, fill=tk.X)

        # --------------------
        # FRAME: Modbus Live Readings
        # --------------------
        modbus_readings = tk.LabelFrame(live_readings_frame, text="Modbus RTU (Dynamometer)", padx=10, pady=10)
        modbus_readings.pack(side=tk.LEFT, padx=20, pady=10, fill="both", expand=True)

        tk.Label(modbus_readings, text="Torque (Nm):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.modbus_torque_label = tk.Label(modbus_readings, text="-- Nm", width=15, anchor="w", relief="sunken", bg="white")
        self.modbus_torque_label.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(modbus_readings, text="Speed (RPM):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.modbus_speed_label = tk.Label(modbus_readings, text="-- RPM", width=15, anchor="w", relief="sunken", bg="white")
        self.modbus_speed_label.grid(row=1, column=1, padx=5, pady=5)

        # --------------------
        # FRAME: CAN Bus Live Readings
        # --------------------
        can_readings = tk.LabelFrame(live_readings_frame, text="CAN Bus (VESC)", padx=10, pady=10)
        can_readings.pack(side=tk.LEFT, padx=20, pady=10, fill="both", expand=True)

        tk.Label(can_readings, text="Current (A):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.can_current_label = tk.Label(can_readings, text="-- A", width=15, anchor="w", relief="sunken", bg="white")
        self.can_current_label.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(can_readings, text="Voltage (V):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.can_voltage_label = tk.Label(can_readings, text="-- V", width=15, anchor="w", relief="sunken", bg="white")
        self.can_voltage_label.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(can_readings, text="RPM:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.can_rpm_label = tk.Label(can_readings, text="-- RPM", width=15, anchor="w", relief="sunken", bg="white")
        self.can_rpm_label.grid(row=2, column=1, padx=5, pady=5)

        # --------------------
        # FRAME: Plotting
        # --------------------
        plot_container = tk.Frame(self.master)
        plot_container.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Matplotlib Figure
        self.figure = Figure(figsize=(14, 7), dpi=100)

        # Modbus Plot
        self.ax_modbus = self.figure.add_subplot(121)
        self.ax_modbus.set_title("Modbus RTU Data Over Time")
        self.ax_modbus.set_xlabel("Time (s)")
        self.ax_modbus.set_ylabel("Torque (Nm) / Speed (RPM)")
        self.ax_modbus.grid(True)
        self.ax_modbus.legend(['Torque (Nm)', 'Speed (RPM)'])

        self.line_modbus_torque, = self.ax_modbus.plot([], [], label='Torque (Nm)', color='blue')
        self.line_modbus_speed, = self.ax_modbus.plot([], [], label='Speed (RPM)', color='green')

        # CAN Plot
        self.ax_can = self.figure.add_subplot(122)
        self.ax_can.set_title("CAN Bus Data Over Time")
        self.ax_can.set_xlabel("Time (s)")
        self.ax_can.set_ylabel("Current (A) / Voltage (V)")
        self.ax_can.grid(True)
        self.ax_can.legend(['Current (A)', 'Voltage (V)'])

        self.line_can_current, = self.ax_can.plot([], [], label='Current (A)', color='red')
        self.line_can_voltage, = self.ax_can.plot([], [], label='Voltage (V)', color='orange')

        # Embed Matplotlib Figure in Tkinter
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_container)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --------------------
        # FRAME: Logging Controls
        # --------------------
        logging_frame = tk.LabelFrame(self.master, text="Data Logging Controls", padx=10, pady=10)
        logging_frame.pack(padx=10, pady=10, fill=tk.X)

        # Save Buttons
        self.save_modbus_button = tk.Button(logging_frame, text="Save Modbus Data to CSV", command=self.save_modbus_csv, state=tk.DISABLED, width=25)
        self.save_modbus_button.pack(side=tk.LEFT, padx=20, pady=5)

        self.save_can_button = tk.Button(logging_frame, text="Save CAN Data to CSV", command=self.save_can_csv, state=tk.DISABLED, width=25)
        self.save_can_button.pack(side=tk.LEFT, padx=20, pady=5)

        # --------------------
        # FRAME: Output Text
        # --------------------
        output_frame = tk.LabelFrame(self.master, text="Last Raw Responses", padx=10, pady=10)
        output_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Modbus Output
        modbus_output = tk.LabelFrame(output_frame, text="Modbus RTU (Dynamometer)", padx=10, pady=10)
        modbus_output.pack(side=tk.LEFT, padx=20, pady=10, fill=tk.BOTH, expand=True)

        self.modbus_output_text = tk.Text(modbus_output, height=10, width=60, bg="#F0F0F0")
        self.modbus_output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        modbus_scrollbar = ttk.Scrollbar(modbus_output, command=self.modbus_output_text.yview)
        modbus_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.modbus_output_text['yscrollcommand'] = modbus_scrollbar.set

        # CAN Output
        can_output = tk.LabelFrame(output_frame, text="CAN Bus (VESC)", padx=10, pady=10)
        can_output.pack(side=tk.LEFT, padx=20, pady=10, fill=tk.BOTH, expand=True)

        self.can_output_text = tk.Text(can_output, height=10, width=60, bg="#F0F0F0")
        self.can_output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        can_scrollbar = ttk.Scrollbar(can_output, command=self.can_output_text.yview)
        can_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.can_output_text['yscrollcommand'] = can_scrollbar.set

        # --------------------
        # Status Label
        # --------------------
        self.status_label = tk.Label(self.master, text="Ready", fg="green", wraplength=1700)
        self.status_label.pack(pady=5)

        # Initial Port Refresh
        self.refresh_modbus_ports()
        self.refresh_can_ports()

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
            self.status_label.config(text=f"Available Modbus COM Ports: {port_names}", fg="green")
        else:
            self.modbus_port_var.set("")
            self.status_label.config(text="No Modbus COM ports found.", fg="red")

    def connect_modbus(self):
        """Connect to the selected Modbus RTU COM port."""
        if self.ser_modbus and self.ser_modbus.is_open:
            self.status_label.config(text="Modbus already connected.", fg="orange")
            return

        port = self.modbus_port_var.get()
        baud = int(self.modbus_baud_var.get())

        try:
            self.ser_modbus = serial.Serial(
                port=port,
                baudrate=baud,
                bytesize=8,
                parity=serial.PARITY_NONE,
                stopbits=1,
                timeout=0.05
            )
            self.status_label.config(text=f"Modbus connected to {port} at {baud} baud.", fg="green")
            self.modbus_output_text.delete("1.0", tk.END)
            self.modbus_start_button.config(state=tk.NORMAL)
            self.modbus_connect_button.config(state=tk.DISABLED)
            self.modbus_disconnect_button.config(state=tk.NORMAL)
        except Exception as e:
            self.status_label.config(text=f"Modbus connection error: {e}", fg="red")
            messagebox.showerror("Connection Error", f"Failed to connect Modbus RTU: {e}")
            self.ser_modbus = None

    def disconnect_modbus(self):
        """Disconnect from the Modbus RTU COM port."""
        if self.ser_modbus and self.ser_modbus.is_open:
            self.ser_modbus.close()
            self.status_label.config(text="Modbus disconnected.", fg="red")
            self.modbus_start_button.config(state=tk.DISABLED)
            self.modbus_stop_button.config(state=tk.DISABLED)
            self.save_modbus_button.config(state=tk.DISABLED)
            self.modbus_connect_button.config(state=tk.NORMAL)
            self.modbus_disconnect_button.config(state=tk.DISABLED)
        else:
            self.status_label.config(text="Modbus not connected.", fg="orange")

    def start_modbus_logging(self):
        """Begin polling Modbus RTU data."""
        if not self.ser_modbus or not self.ser_modbus.is_open:
            self.status_label.config(text="Modbus not connected.", fg="red")
            messagebox.showerror("Error", "Modbus RTU is not connected.")
            return

        if self.logging_modbus:
            self.status_label.config(text="Modbus logging already active.", fg="orange")
            return

        poll_str = self.modbus_poll_interval_var.get()
        try:
            poll_interval = int(poll_str)
        except ValueError:
            poll_interval = 100  # fallback to default

        # Convert poll_interval to a minimum threshold if too small
        if poll_interval < 10:
            poll_interval = 10

        self.modbus_polling_interval = poll_interval / 1000.0  # Convert ms to seconds

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
        """Stop polling Modbus RTU data."""
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
        """Schedule the next Modbus RTU data poll."""
        if not self.logging_modbus:
            return
        self.update_modbus_readings()
        self.modbus_job_id = self.master.after(int(self.modbus_polling_interval * 1000), self.schedule_modbus_update)

    def update_modbus_readings(self):
        """Poll Modbus RTU for Torque and Speed, update readings and plot."""
        current_time = time.time() - self.modbus_start_time

        # 1) Read Torque
        # Using request frame: 01 03 00 00 00 02 C4 0B
        torque_frame_hex = "01 03 00 00 00 02 C4 0B"
        torque_val = self.send_modbus_frame(torque_frame_hex, parse_scale=100.0, is_speed=False)
        if torque_val is not None:
            torque_nm = torque_val / 100.0

            # Apply Filtering: Cap torque at 100 Nm and prevent negative values
            if torque_nm > 100:
                torque_nm = 100
                self.modbus_output_text.insert(tk.END, "Torque value capped at 100 Nm.\n")
            elif torque_nm < 0:
                torque_nm = 0
                self.modbus_output_text.insert(tk.END, "Negative torque value set to 0 Nm.\n")

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
        Sends a raw Modbus RTU frame, reads the response, and parses it.
        Returns the parsed value or None.
        """
        if not self.ser_modbus or not self.ser_modbus.is_open:
            return None

        frame_bytes = bytes.fromhex(frame_hex.replace(" ", ""))
        try:
            self.ser_modbus.reset_input_buffer()
            self.ser_modbus.write(frame_bytes)
            self.ser_modbus.flush()

            start_time = time.time()
            response_buffer = bytearray()

            while (time.time() - start_time) < 0.05:
                chunk = self.ser_modbus.read(64)
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
        """Save the Modbus RTU log data to a CSV file."""
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
    def refresh_can_ports(self):
        """Populate COM ports into the dropdown for CAN Bus."""
        ports = serial.tools.list_ports.comports()
        port_names = [p.device for p in ports]
        self.can_port_dropdown['values'] = port_names
        if port_names:
            self.can_port_var.set(port_names[0])
            self.status_label.config(text=f"Available CAN COM Ports: {port_names}", fg="green")
        else:
            self.can_port_var.set("")
            self.status_label.config(text="No CAN COM ports found.", fg="red")

    def connect_can(self):
        """Connect to the selected CAN Bus COM port."""
        if self.ser_can and self.ser_can.is_open:
            self.status_label.config(text="CAN Bus already connected.", fg="orange")
            return

        port = self.can_port_var.get()
        baud = int(self.can_baud_var.get())

        try:
            self.ser_can = serial.Serial(
                port=port,
                baudrate=baud,
                bytesize=8,
                parity=serial.PARITY_NONE,
                stopbits=1,
                timeout=0.05
            )
            self.status_label.config(text=f"CAN Bus connected to {port} at {baud} baud.", fg="green")
            self.can_output_text.delete("1.0", tk.END)
            self.can_start_log_button.config(state=tk.NORMAL)
            self.can_connect_button.config(state=tk.DISABLED)
            self.can_disconnect_button.config(state=tk.NORMAL)
        except Exception as e:
            self.status_label.config(text=f"CAN Bus connection error: {e}", fg="red")
            messagebox.showerror("Connection Error", f"Failed to connect CAN Bus: {e}")
            self.ser_can = None

    def disconnect_can(self):
        """Disconnect from the CAN Bus COM port."""
        if self.ser_can and self.ser_can.is_open:
            self.ser_can.close()
            self.status_label.config(text="CAN Bus disconnected.", fg="red")
            self.can_start_log_button.config(state=tk.DISABLED)
            self.can_stop_log_button.config(state=tk.DISABLED)
            self.save_can_button.config(state=tk.DISABLED)
            self.can_connect_button.config(state=tk.NORMAL)
            self.can_disconnect_button.config(state=tk.DISABLED)
        else:
            self.status_label.config(text="CAN Bus not connected.", fg="orange")

    def start_can_logging(self):
        """Begin polling CAN Bus data."""
        if not self.ser_can or not self.ser_can.is_open:
            self.status_label.config(text="CAN Bus not connected.", fg="red")
            messagebox.showerror("Error", "CAN Bus is not connected.")
            return

        if self.can_logging:
            self.status_label.config(text="CAN Bus logging already active.", fg="orange")
            return

        # Clear previous data
        self.can_plot_data = {'time': [], 'current': [], 'voltage': []}
        self.can_log_data.clear()
        self.can_start_time = time.time()

        self.can_logging = True
        self.can_start_log_button.config(state=tk.DISABLED)
        self.can_stop_log_button.config(state=tk.NORMAL)
        self.save_can_button.config(state=tk.DISABLED)

        self.status_label.config(text="CAN Bus logging started.", fg="blue")
        self.can_read_thread = threading.Thread(target=self.read_can_data, daemon=True)
        self.can_read_thread.start()

    def stop_can_logging(self):
        """Stop polling CAN Bus data."""
        if self.can_logging:
            self.can_logging = False
            self.can_start_log_button.config(state=tk.NORMAL)
            self.can_stop_log_button.config(state=tk.DISABLED)
            self.save_can_button.config(state=tk.NORMAL)
            self.status_label.config(text="CAN Bus logging stopped.", fg="green")

    def read_can_data(self):
        """Continuously read CAN Bus frames and update GUI."""
        while self.can_logging and self.ser_can and self.ser_can.is_open:
            try:
                line = self.ser_can.readline().decode('utf-8').strip()
                if line:
                    # Assuming the CAN Bus sends data in a specific format, e.g., "CURRENT:1.23; VOLTAGE:4.56; RPM:789"
                    # Adjust the parsing based on your actual data format
                    parsed_data = self.parse_can_data(line)
                    if parsed_data:
                        current, voltage, rpm = parsed_data
                        timestamp = time.time() - self.can_start_time

                        # Update GUI
                        self.can_current_label.config(text=f"{current:.2f} A")
                        self.can_voltage_label.config(text=f"{voltage:.2f} V")
                        self.can_rpm_label.config(text=f"{rpm} RPM")

                        # Log data
                        self.can_log_data.append({
                            'timestamp': timestamp,
                            'current': current,
                            'voltage': voltage,
                            'rpm': rpm
                        })

                        # Update plot data
                        self.can_plot_data['time'].append(timestamp)
                        self.can_plot_data['current'].append(current)
                        self.can_plot_data['voltage'].append(voltage)

                        # Update matplotlib plot
                        self.update_can_plot()

                        # Update Output Text
                        self.can_output_text.insert(tk.END, f"{line}\n")
                        self.can_output_text.see(tk.END)
            except Exception as e:
                self.status_label.config(text=f"CAN Bus Read Error: {e}", fg="red")
                messagebox.showerror("Read Error", f"Error reading CAN Bus data: {e}")
                self.stop_can_logging()
                break

    def parse_can_data(self, data_line):
        """
        Parses a line of CAN Bus data.
        Expected format: "CURRENT:1.23; VOLTAGE:4.56; RPM:789"
        Adjust the parsing based on your actual data format.
        Returns a tuple: (current, voltage, rpm) or None if parsing fails.
        """
        try:
            parts = data_line.split(';')
            current = float(parts[0].split(':')[1])
            voltage = float(parts[1].split(':')[1])
            rpm = float(parts[2].split(':')[1])
            return current, voltage, rpm
        except (IndexError, ValueError):
            print(f"Failed to parse CAN Bus data: {data_line}")
            return None

    def update_can_plot(self):
        """Update the CAN Bus plot with new data."""
        time_data = self.can_plot_data['time']
        current_data = self.can_plot_data['current']
        voltage_data = self.can_plot_data['voltage']

        self.line_can_current.set_data(time_data, current_data)
        self.line_can_voltage.set_data(time_data, voltage_data)

        # Adjust axes
        if time_data:
            self.ax_can.set_xlim(0, max(time_data) + 1)
        if current_data or voltage_data:
            y_min = min(current_data + voltage_data) * 0.9
            y_max = max(current_data + voltage_data) * 1.1
            self.ax_can.set_ylim(y_min, y_max)

        self.canvas.draw()

    def save_can_csv(self):
        """Save the CAN Bus log data to a CSV file."""
        if not self.can_log_data:
            messagebox.showinfo("Info", "No CAN Bus data to save.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")],
                                                 title="Save CAN Bus Data")
        if not file_path:
            return

        try:
            with open(file_path, mode='w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Timestamp (s)", "Current (A)", "Voltage (V)", "RPM"])
                for entry in self.can_log_data:
                    writer.writerow([f"{entry['timestamp']:.3f}", f"{entry['current']:.3f}", f"{entry['voltage']:.3f}", f"{entry['rpm']:.1f}"])
            self.status_label.config(text=f"CAN Bus data saved to {file_path}", fg="green")
        except Exception as e:
            self.status_label.config(text=f"Error saving CAN Bus CSV: {e}", fg="red")

    # --------------------
    # Cleanup on Exit
    # --------------------
    def on_closing(self):
        """Handle application closing."""
        # Stop Modbus logging
        if self.logging_modbus:
            self.stop_modbus_logging()

        # Stop CAN logging
        if self.can_logging:
            self.stop_can_logging()

        # Close Modbus serial port
        if self.ser_modbus and self.ser_modbus.is_open:
            self.ser_modbus.close()

        # Close CAN serial port
        if self.ser_can and self.ser_can.is_open:
            self.ser_can.close()

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
