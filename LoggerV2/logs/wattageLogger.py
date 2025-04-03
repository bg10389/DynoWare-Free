import tkinter as tk
from tkinter import ttk
import serial
import serial.tools.list_ports
import struct
import time
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import csv
from datetime import datetime

class ModbusActiveApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Active Modbus RTU (Torque & Speed every 100ms)")
        
        # Initialize essential attributes first
        self.ser = None
        self.polling_interval = 100  # 100 ms (update cycle)
        self.last_torque = None
        self.last_speed = None
        self.last_watts = None
        self.logging_active = False
        self.logged_data = []

        # --------------------
        # FRAME: Serial Configuration
        # --------------------
        config_frame = tk.LabelFrame(self.master, text="Serial Configuration")
        config_frame.pack(padx=10, pady=10, fill=tk.X)

        tk.Label(config_frame, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.port_var = tk.StringVar()
        self.port_dropdown = ttk.Combobox(config_frame, textvariable=self.port_var, state='readonly', width=12)
        self.port_dropdown.grid(row=0, column=1, padx=5, pady=5)

        self.refresh_button = tk.Button(config_frame, text="Refresh Ports", command=self.refresh_ports)
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

        self.connect_button = tk.Button(config_frame, text="Connect", command=self.connect_serial)
        self.connect_button.grid(row=2, column=0, columnspan=3, pady=5)

        # --------------------
        # FRAME: Active Readings
        # --------------------
        read_frame = tk.LabelFrame(self.master, text="Active Readings (every 100ms)")
        read_frame.pack(padx=10, pady=10, fill=tk.X)

        tk.Label(read_frame, text="Torque (Nm):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.torque_label = tk.Label(read_frame, text="--", width=12, anchor="w")
        self.torque_label.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(read_frame, text="Speed (RPM):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.speed_label = tk.Label(read_frame, text="--", width=12, anchor="w")
        self.speed_label.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(read_frame, text="Watts:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.watts_label = tk.Label(read_frame, text="--", width=12, anchor="w")
        self.watts_label.grid(row=2, column=1, padx=5, pady=5)

        # --------------------
        # FRAME: Last Raw Response
        # --------------------
        output_frame = tk.LabelFrame(self.master, text="Last Raw Response")
        output_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.output_text = tk.Text(output_frame, height=8, width=60)
        self.output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(output_frame, command=self.output_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.output_text['yscrollcommand'] = scrollbar.set

        self.status_label = tk.Label(self.master, text="", fg="blue", wraplength=500)
        self.status_label.pack(pady=5)

        # --------------------
        # FRAME: Data Logging and Plotting
        # --------------------
        log_plot_frame = tk.LabelFrame(self.master, text="Data Logging and Plotting")
        log_plot_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Controls for logging and logging interval
        controls_frame = tk.Frame(log_plot_frame)
        controls_frame.pack(padx=5, pady=5, fill=tk.X)

        tk.Label(controls_frame, text="Logging Interval (ms):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.logging_interval_var = tk.StringVar(value="500")
        self.logging_interval_entry = tk.Entry(controls_frame, textvariable=self.logging_interval_var, width=8)
        self.logging_interval_entry.grid(row=0, column=1, padx=5, pady=5)

        self.start_log_button = tk.Button(controls_frame, text="Start Logging", command=self.start_logging)
        self.start_log_button.grid(row=0, column=2, padx=5, pady=5)

        self.stop_log_button = tk.Button(controls_frame, text="Stop Logging", command=self.stop_logging)
        self.stop_log_button.grid(row=0, column=3, padx=5, pady=5)

        self.save_csv_button = tk.Button(controls_frame, text="Save CSV", command=self.save_csv)
        self.save_csv_button.grid(row=0, column=4, padx=5, pady=5)

        # Create a matplotlib figure for realtime plotting
        self.figure = plt.Figure(figsize=(5, 3), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_title("Real-time Data")
        self.ax.set_xlabel("Time (s)")
        self.ax.set_ylabel("Values")
        self.ax.grid(True)

        self.canvas = FigureCanvasTkAgg(self.figure, master=log_plot_frame)
        self.canvas.get_tk_widget().pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

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

    def connect_serial(self):
        """Open the serial port with the selected baud rate, start active polling."""
        port = self.port_var.get()
        baud = int(self.baud_var.get())

        if self.ser and self.ser.is_open:
            self.ser.close()

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
            self.schedule_update()
        except Exception as e:
            self.status_label.config(text=f"Connection error: {e}")
            self.ser = None

    def schedule_update(self):
        """Repeatedly read torque & speed every 100 ms."""
        self.update_readings()
        self.master.after(self.polling_interval, self.schedule_update)

    def update_readings(self):
        """
        1) Read torque:
           - Interpreted as a signed 16-bit integer (big-endian).
           - If torque > 150 Nm, ignore torque reading (do not update label).
           - If torque < 0 Nm, add 0.05 Nm to it.
        2) Always read and update speed (RPM), regardless of torque value.
        """
        if not self.ser or not self.ser.is_open:
            return

        # 1) Read Torque 
        torque_frame_hex = "01 03 00 00 00 02 C4 0B"
        torque_val = self.send_raw_frame(torque_frame_hex, parse_scale=100.0, signed=True)
        if torque_val is not None:
            torque_nm = torque_val / 100.0
            self.output_text.delete("1.0", tk.END)

            if torque_nm > 150.0:
                self.torque_label.config(text="--")
                self.output_text.insert(tk.END, f"Torque {torque_nm:.2f} Nm exceeds threshold, ignoring torque.\n")
            else:
                if torque_nm < 0.0:
                    torque_nm += 0.05
                self.torque_label.config(text=f"{torque_nm:.2f}")
                self.output_text.insert(tk.END, f"Torque: {torque_nm:.2f} Nm\n")
        else:
            torque_nm = None

        # 2) Read Speed 
        speed_frame_hex = "01 03 00 02 00 02 65 CB"
        speed_val = self.send_raw_frame(speed_frame_hex, parse_scale=10.0, signed=False)
        if speed_val is not None:
            speed_rpm = speed_val / 10.0
            self.speed_label.config(text=f"{speed_rpm:.1f}")
            self.output_text.insert(tk.END, f"Speed: {speed_rpm:.1f} RPM\n")
        else:
            speed_rpm = None

        # 3) Read Watts (optional, if needed)
        watts_frame_hex = "01 03 00 04 00 02 85 CA" 
        watts_val = self.send_raw_frame(watts_frame_hex, parse_scale=1.0, signed=False)
        if watts_val is not None:
            watts = watts_val
            self.watts_label.config(text=f"{watts:.1f}")
            self.output_text.insert(tk.END, f"Watts: {watts:.1f}\n")
        else:
            watts = None
            self.watts_label.config(text="--")
            self.output_text.insert(tk.END, "Watts: --\n")

        # Update last read values for logging
        if torque_nm is not None:
            self.last_torque = torque_nm
        if speed_rpm is not None:
            self.last_speed = speed_rpm
        if watts is not None:
            self.last_watts = watts

    def send_raw_frame(self, frame_hex, parse_scale=1.0, signed=False):
        """
        Sends the raw frame and reads up to ~50ms.
        Extracts the last 2 data bytes as a 16-bit integer.
          - If signed=True, interpret as a signed 16-bit.
          - Otherwise, interpret as unsigned 16-bit.
        Returns the integer, or None if invalid/no response.
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
                hex_resp = " ".join(f"{b:02x}" for b in response_buffer)
                if len(response_buffer) >= 9:
                    byte_count = response_buffer[2]
                    if byte_count == 4:
                        data_bytes = response_buffer[3:3+4]
                        if signed:
                            val_16 = struct.unpack('>h', data_bytes[2:4])[0]
                        else:
                            val_16 = struct.unpack('>H', data_bytes[2:4])[0]
                        return val_16
            else:
                pass

        except Exception as e:
            pass

        return None

    def start_logging(self):
        self.logging_active = True
        self.logged_data = []  # Reset log data
        self.log_start_time = time.time()  # Record start time for relative time axis
        self.update_logging()  # Start logging loop
        self.status_label.config(text="Logging started.")

    def stop_logging(self):
        self.logging_active = False
        self.status_label.config(text="Logging stopped.")

    def update_logging(self):
        if not self.logging_active:
            return
        current_time = time.time()
        elapsed = current_time - self.log_start_time
        # Append current logged data; if a value is None, record as empty
        self.logged_data.append((
            elapsed,
            self.last_torque if self.last_torque is not None else "",
            self.last_speed if self.last_speed is not None else "",
            self.last_watts if self.last_watts is not None else ""
        ))
        # Update realtime plot with logged data
        self.update_plot()
        # Get logging interval from entry; default to 500 ms if invalid
        try:
            interval = int(self.logging_interval_var.get())
        except:
            interval = 500
        self.master.after(interval, self.update_logging)

    def update_plot(self):
        self.ax.cla()
        self.ax.set_title("Real-time Data")
        self.ax.set_xlabel("Time (s)")
        self.ax.set_ylabel("Values")
        self.ax.grid(True)
        if self.logged_data:
            times, torques, speeds, watts = zip(*self.logged_data)
            if any(t != "" for t in torques):
                self.ax.plot(times, torques, label="Torque (Nm)")
            if any(s != "" for s in speeds):
                self.ax.plot(times, speeds, label="Speed (RPM)")
            if any(w != "" for w in watts):
                self.ax.plot(times, watts, label="Watts")
            self.ax.legend()
        self.canvas.draw()

    def save_csv(self):
        if not self.logged_data:
            self.status_label.config(text="No data to save.")
            return
        if not os.path.exists("logs"):
            os.makedirs("logs")
        filename = datetime.now().strftime("logs/log_%Y%m%d_%H%M%S.csv")
        try:
            with open(filename, "w", newline="") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Time (s)", "Torque (Nm)", "Speed (RPM)", "Watts"])
                for row in self.logged_data:
                    writer.writerow(row)
            self.status_label.config(text=f"Data saved to {filename}")
        except Exception as e:
            self.status_label.config(text=f"Error saving CSV: {e}")

def main():
    root = tk.Tk()
    app = ModbusActiveApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
