import tkinter as tk
from tkinter import ttk
import serial
import serial.tools.list_ports
import struct
import time

class ModbusActiveApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Active Modbus RTU (Torque & Speed every 100ms)")

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

        # Serial port object
        self.ser = None
        self.polling_interval = 100  # 100 ms (update cycle)
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
        """Read torque and speed in sequence, parse the data, and update labels."""
        if not self.ser or not self.ser.is_open:
            return

        # 1) Read Torque 
        #    Using the corrected request frame: 01 03 00 00 00 02 C4 0B
        #    The device response is typically 01 03 04 00 00 00 00 FA 33 if torque=0
        torque_frame_hex = "01 03 00 00 00 02 C4 0B"
        torque_val = self.send_raw_frame(torque_frame_hex, parse_scale=100.0, is_torque=True)
        if torque_val is not None:
            # If torque is displayed in Nm, maybe we scale by 0.01, 0.1, etc. 
            # But user says device returns last two data bytes for torque. 
            # parse_scale=100.0 means we do val_16 / 100 => ex: raw "0100" => 1.00 Nm
            # Adjust as needed if the documentation specifies a different scaling.
            torque_nm = torque_val / 100.0  # example scale
            self.torque_label.config(text=f"{torque_nm:.2f}")

        # 2) Read Speed 
        #    We keep the same frame: 01 03 00 02 00 02 65 CB 
        #    If raw is 491 => 49.1 rpm
        speed_frame_hex = "01 03 00 02 00 02 65 CB"
        speed_val = self.send_raw_frame(speed_frame_hex, parse_scale=10.0, is_torque=False)
        if speed_val is not None:
            speed_rpm = speed_val / 10.0
            self.speed_label.config(text=f"{speed_rpm:.1f}")

    def send_raw_frame(self, frame_hex, parse_scale=1.0, is_torque=False):
        """
        Writes the raw frame, polls for up to ~50ms, 
        extracts the LAST 2 data bytes as a 16-bit big-endian integer.
        
        parse_scale: the factor to divide by for final value. 
                     e.g., if parse_scale=10.0, raw 491 => 49.1
                     if parse_scale=100.0, raw 1234 => 12.34
                     
        Return: integer from the last 2 data bytes, or None if invalid/no response.
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

                # Must have at least 9 bytes for a valid 2-reg read: address(1),func(1),byteCount(1=4?),data(4),crc(2)
                if len(response_buffer) >= 9:
                    byte_count = response_buffer[2]
                    if byte_count == 4:
                        data_bytes = response_buffer[3:3+4]
                        # The last two bytes store the raw 16-bit integer
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

def main():
    root = tk.Tk()
    app = ModbusActiveApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
