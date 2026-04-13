import serial
import serial.tools.list_ports
import serial.tools.list_ports_windows
import time
import math
import psutil
import wmi
import threading
import pystray
import pythoncom
import json
import os
import traceback
from PIL import Image, ImageDraw
import sys
import ctypes

# ============== CONSOLE ONLY ON DEMAND ==============
#   Double-click .exe      → completely hidden (only tray icon)
#   Run with --console     → creates a new console window + shows all prints
if "--console" in [arg.lower() for arg in sys.argv]:
    try:
        ctypes.windll.kernel32.AllocConsole()
        # Redirect Python output so all your print() statements appear
        sys.stdout = open("CONOUT$", "w", buffering=1)
        sys.stderr = open("CONOUT$", "w", buffering=1)
        print("=== FW16 Battery Matrix - DEBUG CONSOLE ENABLED ===")
        print("Close this window = close the app")
    except:
        pass
# ===================================================

# ==================== CONFIG & PERSISTENCE ====================
CONFIG_FILE = "batt4_config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except:
            pass
    return {"com_port": "COM3", "brightness": 70}

def save_config(config):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f, indent=2)
    except:
        pass

config = load_config()
COM_PORT = config.get("com_port", "COM3")
MATRIX_BRIGHTNESS = config.get("brightness", 70)

# ==================== OTHER CONFIG ====================
BAUD_RATE = 115200
WIDTH = 9
HEIGHT = 34
MAX_BRIGHT = 255
MIN_M = 10 / 255
SIGMA = 2.0
PULSE_SPEED_MODIFIER = 2.0
FADE_SPEED = 0.08
FPS = 12
RECONNECT_DELAY = 5.0
WMI_POLL_INTERVAL = 0.8

# ==================== STATE ====================
ser = None
last_reconnect_attempt = 0

last_wmi_remaining_mwh = None
last_raw_wmi_percent = None
last_wmi_time = None
estimated_remaining_mwh = None
full_charge_mwh = None
last_wmi_poll_time = 0

last_change_time = 0
previous_change_time = 0

pulse_position = None
pulse_opacity = 0.0
running = True

def save_current_config():
    global COM_PORT, MATRIX_BRIGHTNESS
    config["com_port"] = COM_PORT
    config["brightness"] = MATRIX_BRIGHTNESS
    save_config(config)

def send_column(column_id, values):
    cmd = [0x32, 0xAC, 0x07, column_id]
    for val in values:
        val = max(0, min(255, int(val)))
        cmd.append(val)
    ser.write(bytearray(cmd))

def send_flush():
    ser.write(bytearray([0x32, 0xAC, 0x08]))

def set_brightness(bright=None):
    global MATRIX_BRIGHTNESS
    if bright is not None:
        MATRIX_BRIGHTNESS = bright
        save_current_config()
    bright = max(0, min(255, MATRIX_BRIGHTNESS))
    cmd = [0x32, 0xAC, 0x00, bright]
    if ser and ser.is_open:
        try:
            ser.write(bytearray(cmd))
        except:
            pass

def compute_multiplier(r, c, sigma, min_m):
    if c is None:
        return 1.0
    dist = (r - c) / sigma
    return 1 - (1 - min_m) * math.exp(-dist**2)

def create_battery_frame(p, pulse_center, pulse_opacity):
    columns = []
    fill_level = (p / 100.0) * 30
    full_rows = math.floor(fill_level)
    partial_fraction = fill_level - full_rows
    partial_row = 32 - full_rows if full_rows < 30 else None

    for col in range(WIDTH):
        column = [0] * HEIGHT
        if 3 <= col <= 5:
            column[0] = MAX_BRIGHT
        if col in (0, 1, 2, 6, 7, 8):
            column[1] = MAX_BRIGHT
        for row in range(2, 33):
            if col in (0, 8):
                column[row] = MAX_BRIGHT
            elif 2 <= col <= 6:
                if row > 32 - full_rows:
                    column[row] = MAX_BRIGHT
                elif row == partial_row and partial_row is not None:
                    if col == 4:
                        fade_factor = min(1.0, partial_fraction / 0.33)
                    elif col in (3, 5):
                        fade_factor = max(0.0, min(1.0, (partial_fraction - 0.33) / 0.33))
                    elif col in (2, 6):
                        fade_factor = max(0.0, min(1.0, (partial_fraction - 0.66) / 0.34))
                    column[row] = int(round(MAX_BRIGHT * fade_factor))
                else:
                    column[row] = 0
        column[33] = MAX_BRIGHT
        columns.append(column)

    if pulse_center is not None and pulse_opacity > 0.05:
        for col in range(2, 7):
            for row in range(2, 33):
                m = compute_multiplier(row, pulse_center, SIGMA, MIN_M)
                columns[col][row] = int(round(columns[col][row] * m * pulse_opacity))

    return columns

def tray_icon():
    image = Image.new('RGB', (64, 64), (0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rectangle((8, 8, 56, 56), outline=(0, 255, 0), width=6)
    draw.text((22, 18), "🔋", fill=(0, 255, 0))
    return image

def on_exit(icon, item):
    global running
    running = False
    icon.stop()

def change_port(new_port):
    global COM_PORT, ser
    COM_PORT = new_port
    save_current_config()
    print(f"✅ Switching to {new_port}...")

    if ser and ser.is_open:
        try:
            ser.close()
        except:
            pass
        ser = None

def brightness_up(icon, item):
    global MATRIX_BRIGHTNESS
    MATRIX_BRIGHTNESS = min(255, MATRIX_BRIGHTNESS + 10)
    set_brightness()
    print(f"🔆 Brightness +10 → {MATRIX_BRIGHTNESS}")

def brightness_down(icon, item):
    global MATRIX_BRIGHTNESS
    MATRIX_BRIGHTNESS = max(0, MATRIX_BRIGHTNESS - 10)
    set_brightness()
    print(f"🔆 Brightness -10 → {MATRIX_BRIGHTNESS}")

def main_loop():
    global ser, last_reconnect_attempt, last_wmi_remaining_mwh, last_raw_wmi_percent
    global last_wmi_time, estimated_remaining_mwh, full_charge_mwh, pulse_position, pulse_opacity
    global last_wmi_poll_time, last_change_time, previous_change_time
    pythoncom.CoInitialize()
    # Create WMI object INSIDE the thread (this fixes the error)
    wmi_interface = wmi.WMI(namespace="root\\wmi")

    frame_time = 1.0 / FPS
    brightness_set = False

    while running:
        loop_start = time.time()
        current_time = time.time()

        battery = psutil.sensors_battery()
        if battery is None:
            print("No battery detected.")
            time.sleep(1)
            continue

        do_wmi_poll = (current_time - last_wmi_poll_time >= WMI_POLL_INTERVAL) or (last_wmi_poll_time == 0)

        if do_wmi_poll:
            try:
                # === EXACT WMI CODE FROM YOUR WORKING batt3.py ===
                battery_status = wmi_interface.BatteryStatus()[0]
                remaining_capacity = battery_status.RemainingCapacity
                charge_rate = getattr(battery_status, 'ChargeRate', 0) or 0
                discharge_rate = getattr(battery_status, 'DischargeRate', 0) or 0
                battery_full = wmi_interface.BatteryFullChargedCapacity()[0]
                full_charge_capacity = battery_full.FullChargedCapacity

                if remaining_capacity is not None and full_charge_capacity is not None and full_charge_capacity > 0:
                    full_charge_mwh = full_charge_capacity
                    if last_wmi_remaining_mwh is None or remaining_capacity != last_wmi_remaining_mwh:
                        estimated_remaining_mwh = remaining_capacity
                        previous_change_time = last_change_time
                        last_change_time = current_time
                        print(f"🔄 WMI SYNC → New real value: {remaining_capacity:6d} mWh (changed!)")
                    else:
                        time_between_changes = last_change_time - previous_change_time if previous_change_time > 0 else WMI_POLL_INTERVAL
                        failsafe_interval = 2.0 * time_between_changes
                        if current_time - last_change_time >= failsafe_interval:
                            print(f"🔄 WMI still same: {remaining_capacity:6d} mWh (using interpolation)")
                            last_change_time = current_time

                    last_wmi_remaining_mwh = remaining_capacity
                    last_raw_wmi_percent = (remaining_capacity / full_charge_capacity * 100)
                    last_wmi_time = current_time
                else:
                    charge_rate = discharge_rate = 0
            except Exception as e:
                print(f"⚠️ WMI error: {e}")
                traceback.print_exc()   # ← Full traceback for diagnosis
                charge_rate = discharge_rate = 0
            last_wmi_poll_time = current_time

        # Interpolation
        if full_charge_mwh and last_wmi_time and estimated_remaining_mwh is not None:
            now = current_time
            dt = now - last_wmi_time
            delta_mWh = (charge_rate - discharge_rate) * dt / 3600.0
            estimated_remaining_mwh = max(0, min(full_charge_mwh, estimated_remaining_mwh + delta_mWh))
            last_wmi_time = now
            est_percent = (estimated_remaining_mwh / full_charge_mwh) * 100
        else:
            est_percent = battery.percent

        wmi_percent = last_raw_wmi_percent if last_raw_wmi_percent is not None else est_percent

        net_mW = charge_rate - discharge_rate
        mode = "charge" if net_mW > 0 else "discharge" if net_mW < 0 else "idle"
        target_opacity = 1.0 if mode != "idle" else 0.0

        fill_level = (est_percent / 100.0) * 30
        full_rows = math.floor(fill_level)
        top_fill = 32 - full_rows

        if mode == "charge":
            if pulse_position is None or pulse_position < top_fill:
                pulse_position = 33.0
            pulse_position -= (net_mW / 1000.0 / 50.0) * PULSE_SPEED_MODIFIER
            if pulse_position < top_fill:
                pulse_position = 33.0
        elif mode == "discharge":
            if pulse_position is None or pulse_position > 33:
                pulse_position = top_fill
            pulse_position += (abs(net_mW) / 1000.0 / 50.0) * PULSE_SPEED_MODIFIER
            if pulse_position > 33:
                pulse_position = top_fill

        if pulse_opacity < target_opacity:
            pulse_opacity = min(1.0, pulse_opacity + FADE_SPEED)
        elif pulse_opacity > target_opacity:
            pulse_opacity = max(0.0, pulse_opacity - FADE_SPEED)

        pulse_center = pulse_position if pulse_opacity > 0.05 else None

        columns = create_battery_frame(est_percent, pulse_center, pulse_opacity)

        # Serial handling
        if ser is None or not ser.is_open:
            if current_time - last_reconnect_attempt >= RECONNECT_DELAY:
                try:
                    ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
                    print(f"✅ Connected to {COM_PORT}")
                    brightness_set = False
                except Exception:
                    pass
                last_reconnect_attempt = current_time
        else:
            if not brightness_set:
                set_brightness()
                brightness_set = True
                time.sleep(0.05)

            try:
                for col in range(WIDTH):
                    send_column(col, columns[col])
                send_flush()
            except (serial.SerialException, OSError):
                try:
                    ser.close()
                except:
                    pass
                ser = None
                brightness_set = False

        pulse_str = f"{pulse_center:.1f}" if pulse_center is not None else "N/A"
        est_mwh = estimated_remaining_mwh or 0
        print(f"Frame | WMI: {wmi_percent:6.3f}%  Est: {est_percent:6.3f}% "
              f"(mWh: {est_mwh:7.1f}/{full_charge_mwh}) | Net: {net_mW:+6.0f} mW | "
              f"Pulse: {pulse_str} | Op: {pulse_opacity:.2f} | Mode: {mode}")

        elapsed = time.time() - loop_start
        if elapsed < frame_time:
            time.sleep(frame_time - elapsed)

def main():
    global running

    loop_thread = threading.Thread(target=main_loop, daemon=True)
    loop_thread.start()

    def create_port_menu():
        ports = [p.device for p in serial.tools.list_ports.comports()]
        items = []
        for port in ports:
            def make_handler(p=port):
                def handler(icon, item):
                    change_port(p)
                return handler
            items.append(pystray.MenuItem(port, make_handler()))
        return items

    menu = pystray.Menu(
        pystray.MenuItem("Brightness +10", brightness_up),
        pystray.MenuItem("Brightness -10", brightness_down),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Select Port", pystray.Menu(*create_port_menu())),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Made by syber-labs.com", None),      # ← non-clickable credit
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Exit", on_exit)
    )

    icon = pystray.Icon(
        name="Framework Battery",
        icon=tray_icon(),
        title=f"FW16 Battery - {COM_PORT}",
        menu=menu
    )

    print("✅ Tray icon started - right-click the battery icon near the clock")
    icon.run()

if __name__ == "__main__":
    main()