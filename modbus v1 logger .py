import tkinter as tk
from tkinter import ttk
import serial
import serial.tools.list_ports
import struct
import time
import csv
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

class ModbusLoggerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Modbus RTU Logger with Adjustable Interval & Start/Stop")

        # --------------------
        # FRAME: Serial Configuration
        # --------------------
        config_frame = tk.LabelFrame(self.master, text="Serial Configuration")
        config_frame.pack(padx=10, pady=10, fill=tk.X)

        tk.Label(config_frame, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.port_var = tk.StringVar()
        self.port_dropdown = ttk.Combobox(config_frame, textvariable=self.port_var, state='readonly', width=12)
        self.port_dropdown.grid(row=0, column=1, padx=5, pady=5)

        self.refresh_button = tk.Button(config_frame, text="Refresh", command=self.refresh_ports)
        self.refresh_button.grid(row=0, column=2, padx=5, pady=5)

        tk.Label(config_frame, text="Baud Rate:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.baud_var = tk.StringVar(value="38400")
        self.baud_dropdown = ttk.Combobox(
            config_frame,
            textvariable=self.baud_var,
            values=["9600", "19200", "38400", "57600", "115200"],
            state='readonly',
            width=12
        )
        self.baud_dropdown.grid(row=1, column=1, padx=5, pady=5)

        # Start/Stop and Poll Interval
        tk.Label(config_frame, text="Poll Interval (ms):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.poll_interval_var = tk.StringVar(value="100")  # default 100ms
        tk.Entry(config_frame, textvariable=self.poll_interval_var, width=10).grid(row=2, column=1, padx=5, pady=5)

        self.start_button = tk.Button(config_frame, text="Start Logging", command=self.start_logging)
        self.start_button.grid(row=3, column=0, padx=5, pady=5)

        self.stop_button = tk.Button(config_frame, text="Stop Logging", command=self.stop_logging)
        self.stop_button.grid(row=3, column=1, padx=5, pady=5)

        # --------------------
        # FRAME: Active Readings
        # --------------------
        read_frame = tk.LabelFrame(self.master, text="Live Readings")
        read_frame.pack(padx=10, pady=10, fill=tk.X)

        # Torque display
        tk.Label(read_frame, text="Torque (Nm):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.torque_label = tk.Label(read_frame, text="--", width=12, anchor="w")
        self.torque_label.grid(row=0, column=1, padx=5, pady=5)

        # Speed display
        tk.Label(read_frame, text="Speed (RPM):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.speed_label = tk.Label(read_frame, text="--", width=12, anchor="w")
        self.speed_label.grid(row=1, column=1, padx=5, pady=5)

        # --------------------
        # FRAME: Plot and Save CSV
        # --------------------
        plot_frame = tk.LabelFrame(self.master, text="Plot & Data Logging")
        plot_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Matplotlib figure
        self.figure = Figure(figsize=(5,4), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_title("Torque & Speed Over Time")
        self.ax.set_xlabel("Time (s)")
        self.ax.set_ylabel("Value")
        self.ax.grid(True)

        self.canvas = FigureCanvasTkAgg(self.figure, plot_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Lines for torque & speed
        self.time_data = []
        self.torque_data = []
        self.speed_data = []

        (self.line_torque,) = self.ax.plot([], [], label="Torque (Nm)")
        (self.line_speed,)  = self.ax.plot([], [], label="Speed (RPM)")
        self.ax.legend()

        # Button & frame for CSV
        button_frame = tk.Frame(plot_frame)
        button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10)

        self.save_button = tk.Button(button_frame, text="Save to CSV", command=self.save_csv)
        self.save_button.pack(pady=5, anchor="n")

        # FRAME: Last Raw Response
        output_frame = tk.LabelFrame(self.master, text="Last Raw Response")
        output_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.output_text = tk.Text(output_frame, height=6, width=60)
        self.output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(output_frame, command=self.output_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.output_text['yscrollcommand'] = scrollbar.set

        # Status label
        self.status_label = tk.Label(self.master, text="", fg="blue", wraplength=500)
        self.status_label.pack(pady=5)

        # Serial port object
        self.ser = None
        self.logging_active = False  # controls whether we are actively logging

        # Start time for logging
        self.start_time = None
        self.job_id = None  # handle for the scheduled after() call

        self.refresh_ports()

    def refresh_ports(self):
        """Populate COM ports into the dropdown."""
        ports = serial.tools.list_ports.comports()
        port_names = [p.device for p in ports]
        self.port_dropdown['values'] = port_names
        if port_names:
            self.port_var.set(port_names[0])
            self.status_label.config(text=f"Ports: {port_names}")
        else:
            self.port_var.set("")
            self.status_label.config(text="No COM ports found.")

    def start_logging(self):
        """Open serial port if needed and begin polling at user-defined interval."""
        if self.logging_active:
            self.status_label.config(text="Already logging.")
            return

        port = self.port_var.get()
        baud = int(self.baud_var.get())
        poll_str = self.poll_interval_var.get()
        try:
            poll_interval = int(poll_str)
        except ValueError:
            poll_interval = 100  # fallback to default

        # Convert poll_interval to a minimum threshold if too small
        if poll_interval < 10:
            poll_interval = 10

        self.polling_interval = poll_interval

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
                self.status_label.config(text=f"Connected to {port} at {baud} baud.")
                self.output_text.delete("1.0", tk.END)
            except Exception as e:
                self.status_label.config(text=f"Connection error: {e}")
                self.ser = None
                return

        # Clear previous data
        self.time_data.clear()
        self.torque_data.clear()
        self.speed_data.clear()
        self.start_time = time.time()

        self.logging_active = True
        self.schedule_update()

    def stop_logging(self):
        """Stop active polling."""
        if self.logging_active:
            self.logging_active = False
            if self.job_id is not None:
                self.master.after_cancel(self.job_id)
                self.job_id = None
            self.status_label.config(text="Logging stopped.")

    def schedule_update(self):
        """Repeatedly poll torque & speed, update the plot, if logging is active."""
        if not self.logging_active:
            return
        self.update_readings()
        self.job_id = self.master.after(self.polling_interval, self.schedule_update)

    def update_readings(self):
        if not self.ser or not self.ser.is_open:
            return

        current_time = time.time() - self.start_time

        # 1) Read Torque
        #   Using request frame: 01 03 00 00 00 02 C4 0B
        #   Interpret last 2 data bytes as 16-bit int, scale by /100 => Nm
        torque_frame_hex = "01 03 00 00 00 02 C4 0B"
        torque_val = self.send_raw_frame(torque_frame_hex, parse_scale=100.0)
        if torque_val is not None:
            torque_nm = torque_val / 100.0
            self.torque_label.config(text=f"{torque_nm:.2f}")
        else:
            torque_nm = None

        # 2) Read Speed
        #   Using request frame: 01 03 00 02 00 02 65 CB
        #   Interpret last 2 data bytes as 16-bit int, scale by /10 => RPM
        speed_frame_hex = "01 03 00 02 00 02 65 CB"
        speed_val = self.send_raw_frame(speed_frame_hex, parse_scale=10.0)
        if speed_val is not None:
            speed_rpm = speed_val / 10.0
            self.speed_label.config(text=f"{speed_rpm:.1f}")
        else:
            speed_rpm = None

        # If both readings are valid, log & update plot
        if torque_nm is not None and speed_rpm is not None:
            self.time_data.append(current_time)
            self.torque_data.append(torque_nm)
            self.speed_data.append(speed_rpm)

            # Update matplotlib plot
            self.line_torque.set_xdata(self.time_data)
            self.line_torque.set_ydata(self.torque_data)

            self.line_speed.set_xdata(self.time_data)
            self.line_speed.set_ydata(self.speed_data)

            # Adjust axes
            self.ax.set_xlim(left=0, right=max(self.time_data) if self.time_data else 10)
            all_vals = self.torque_data + self.speed_data
            if all_vals:
                y_min = min(all_vals)
                y_max = max(all_vals)
                if y_min == y_max:
                    self.ax.set_ylim(y_min - 1, y_max + 1)
                else:
                    self.ax.set_ylim(y_min * 0.95, y_max * 1.05)

            self.canvas.draw()

    def send_raw_frame(self, frame_hex, parse_scale=1.0):
        """
        Writes the raw frame, reads for ~50ms, 
        interprets the last 2 of 4 data bytes as a 16-bit integer big-endian.

        parse_scale=1.0 means raw=100 => 100.0 final
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
                self.output_text.delete("1.0", tk.END)
                self.output_text.insert(tk.END, f"Raw Response ({len(response_buffer)} bytes):\n{hex_resp}\n")

                if len(response_buffer) >= 9:
                    byte_count = response_buffer[2]
                    if byte_count == 4:
                        data_bytes = response_buffer[3:3+4]
                        val_16 = struct.unpack('>H', data_bytes[2:4])[0]
                        return val_16
                    else:
                        self.output_text.insert(tk.END, f"Invalid byte count: {byte_count}\n")
                else:
                    self.output_text.insert(tk.END, "Response too short.\n")
            else:
                self.output_text.delete("1.0", tk.END)
                self.output_text.insert(tk.END, "No response.\n")

        except Exception as e:
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert(tk.END, f"Error sending frame: {e}\n")

        return None

    def save_csv(self):
        """Save the current time_data, torque_data, speed_data to a CSV file."""
        filename = time.strftime("dyno_data_%Y%m%d_%H%M%S.csv")
        try:
            with open(filename, mode='w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Time (s)", "Torque (Nm)", "Speed (RPM)"])
                for t, tq, spd in zip(self.time_data, self.torque_data, self.speed_data):
                    writer.writerow([f"{t:.3f}", f"{tq:.3f}", f"{spd:.3f}"])
            self.status_label.config(text=f"Data saved to {filename}")
        except Exception as e:
            self.status_label.config(text=f"Error saving CSV: {e}")

def main():
    root = tk.Tk()
    app = ModbusLoggerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
