def calc_crc(message_bytes):
    crc = 0xFFFF
    for byte in message_bytes:
        crc ^= byte
        for _ in range(8):
            if crc & 0x0001:
                crc = (crc >> 1) ^ 0xA001
            else:
                crc >>= 1
    return crc

# Example usage:
frame_bytes = bytes.fromhex("01 03 00 04 00 02")
crc_value = calc_crc(frame_bytes)
# CRC is returned as an integer, but Modbus RTU frames use little-endian:
crc_low = crc_value & 0xFF
crc_high = (crc_value >> 8) & 0xFF
print(f"Calculated CRC: {crc_low:02X} {crc_high:02X}")