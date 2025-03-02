import serial
import time
import math
import psutil
import wmi

# Configuration
COM_PORT = 'COM4'
BAUD_RATE = 115200
WIDTH = 9
HEIGHT = 34
MAX_BRIGHT = 255
MIN_M = 10 / 255
SIGMA = 2.0
STEP_SCALE = 100.0
FADE_SPEED = 0.01
PULSE_SPEED_MODIFIER = 2.0
CMD_STAGE_COL = 0x07
CMD_FLUSH_COLS = 0x08
FPS = 10

ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
wmi_interface = wmi.WMI(namespace="root\\wmi")

def send_column(column_id, values):
    cmd = [0x32, 0xAC, CMD_STAGE_COL, column_id]
    for val in values:
        if val < 0 or val > 255:
            print(f"Error: Brightness value {val} out of range (0-255) for column {column_id}")
            val = max(0, min(255, val))
        cmd.append(val)
    ser.write(bytearray(cmd))

def send_flush():
    cmd = [0x32, 0xAC, CMD_FLUSH_COLS]
    ser.write(bytearray(cmd))

def compute_multiplier(r, c, sigma, min_m):
    if c is None:
        return 1.0
    dist = (r - c) / sigma
    return 1 - (1 - min_m) * math.exp(-dist**2)

def create_battery_frame(p, c, pulse_fade):
    columns = []
    fill_level = (p / 100.0) * 30  # 30 rows (2 to 32) for 0-100%
    full_rows = math.floor(fill_level)
    partial_fraction = fill_level - full_rows
    partial_row = 32 - full_rows if full_rows < 30 else None

    for col in range(WIDTH):
        column = [0] * HEIGHT
        # Borders
        if 3 <= col <= 5:
            column[0] = MAX_BRIGHT  # Top cap
        if col in (0, 1, 2, 6, 7, 8):
            column[1] = MAX_BRIGHT  # Top border
        for row in range(2, 33):
            if col in (0, 8):
                column[row] = MAX_BRIGHT  # Side borders
            elif 2 <= col <= 6:
                if row > 32 - full_rows:
                    column[row] = MAX_BRIGHT  # Full rows
                elif row == partial_row and partial_row is not None:
                    # Center-out fade for partial row
                    if col == 4:
                        fade_factor = min(1.0, partial_fraction / 0.33)  # 0 to 0.33
                    elif col in (3, 5):
                        fade_factor = max(0.0, (partial_fraction - 0.33) / 0.33)  # 0.33 to 0.66
                    elif col in (2, 6):
                        fade_factor = max(0.0, (partial_fraction - 0.66) / 0.34)  # 0.66 to 1.0
                    column[row] = int(round(MAX_BRIGHT * fade_factor))
                else:
                    column[row] = 0
        column[33] = MAX_BRIGHT  # Bottom border
        columns.append(column)

    # Apply pulse effect
    if c is not None:
        for col in range(2, 7):
            for row in range(2, 33):
                m = compute_multiplier(row, c, SIGMA, MIN_M)
                columns[col][row] = int(round(columns[col][row] * m))

    return columns

# Main loop
start_time = time.time()
frame_time = 1.0 / FPS
pulse_pos = None
pulse_fade = 0.0

while True:
    loop_start = time.time()
    battery = psutil.sensors_battery()
    if battery is None:
        break

    # Get precise battery percentage using WMI
    try:
        battery_status = wmi_interface.BatteryStatus()[0]
        remaining_capacity = battery_status.RemainingCapacity
        charge_rate = getattr(battery_status, 'ChargeRate', 0) or 0
        discharge_rate = getattr(battery_status, 'DischargeRate', 0) or 0
        battery_full = wmi_interface.BatteryFullChargedCapacity()[0]
        full_charge_capacity = battery_full.FullChargedCapacity
        if remaining_capacity is not None and full_charge_capacity is not None and full_charge_capacity > 0:
            p = (remaining_capacity / full_charge_capacity) * 100
        else:
            p = battery.percent  # Fallback to psutil if WMI data is invalid
    except Exception as e:
        print(f"WMI error: {e}. Falling back to psutil battery percentage.")
        p = battery.percent  # Fallback to psutil on error
        charge_rate = 0
        discharge_rate = 0

    print(f"Precise Battery Percentage: {p:.2f}%")
    charge_watts = charge_rate / 1000.0
    discharge_watts = discharge_rate / 1000.0

    mode = "charge" if charge_watts > 0 else "discharge" if discharge_watts > 0 else "idle"
    target_fade = 1.0 if mode != "idle" else 0.0
    pulse_fade += FADE_SPEED if pulse_fade < target_fade else -FADE_SPEED if pulse_fade > target_fade else 0
    pulse_fade = max(0.0, min(1.0, pulse_fade))

    fill_level = (p / 100.0) * 30
    full_rows = math.floor(fill_level)
    top_fill = 32 - full_rows  # Top of the filled area

    if mode == "charge" and charge_watts > 0:
        if pulse_pos is None:
            pulse_pos = 33  # Start at bottom
        else:
            pulse_pos -= (charge_watts / STEP_SCALE) * PULSE_SPEED_MODIFIER  # Move up
        if pulse_pos < top_fill:  # Reset when past the top of fill
            pulse_pos = 33
        c = pulse_pos if pulse_fade > 0 else None
    elif mode == "discharge" and discharge_watts > 0:
        if pulse_pos is None:
            pulse_pos = top_fill  # Start at top of fill
        else:
            pulse_pos += (discharge_watts / STEP_SCALE) * PULSE_SPEED_MODIFIER  # Move down
        if pulse_pos > 33:  # Reset when past bottom
            pulse_pos = top_fill
        c = pulse_pos if pulse_fade > 0 else None
    else:
        c = None

    columns = create_battery_frame(p, c, pulse_fade)
    for col in range(WIDTH):
        send_column(col, columns[col])
    send_flush()

    elapsed = time.time() - loop_start
    if elapsed < frame_time:
        time.sleep(frame_time - elapsed)

ser.close()