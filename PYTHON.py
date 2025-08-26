# PYTHON.py (Windows host)
import sys, time, serial
from serial.tools import list_ports

from comtypes import CLSCTX_ALL, GUID
from comtypes.client import CreateObject
from ctypes import POINTER, cast
from pycaw.pycaw import IAudioEndpointVolume, IMMDeviceEnumerator

# ---- MMDevice constants
eRender, eCapture, eAll = 0, 1, 2
eConsole, eMultimedia, eCommunications = 0, 1, 2

CLSID_MMDeviceEnumerator = GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")

def _get_enumerator():
    try:
        return CreateObject(IMMDeviceEnumerator)
    except Exception:
        return CreateObject(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator)

def _endpoint_volume():
    enum = _get_enumerator()
    dev  = enum.GetDefaultAudioEndpoint(eCapture, eConsole)   # default microphone
    iface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(iface, POINTER(IAudioEndpointVolume))

def get_mic_muted() -> bool:
    return bool(_endpoint_volume().GetMute())

def set_mic_mute(mute: bool) -> None:
    _endpoint_volume().SetMute(1 if mute else 0, None)

def toggle_mic() -> bool:
    vol = _endpoint_volume()
    cur = bool(vol.GetMute())
    vol.SetMute(0 if cur else 1, None)
    return not cur  # True if now muted

def pick_port(prefer=None):
    if prefer: return prefer
    # Prefer the CircuitPython "data" interface
    for p in list_ports.comports():
        if "data" in ((p.interface or "") + (p.description or "")).lower():
            return p.device
    # Fallback: any CircuitPython/RP2-looking port
    for p in list_ports.comports():
        if any(k in (p.description or "") for k in ("CircuitPython","Adafruit","RP2","Pico")):
            return p.device
    raise RuntimeError("No serial ports found.")

def send_state(ser):
    try:
        muted = get_mic_muted()
        msg = b"MIC:STATE MUTED\n" if muted else b"MIC:STATE UNMUTED\n"
        ser.write(msg)
    except Exception as e:
        print("State send failed:", repr(e))

if __name__ == "__main__":
    PORT = pick_port(sys.argv[1] if len(sys.argv) > 1 else None)
    print("Using port:", PORT)

    ser = serial.Serial(PORT, 115200, timeout=1)
    time.sleep(0.3)
    ser.reset_input_buffer()
    print("Listening for MIC:TOGGLE ...")

    # Tell the Pico the current state on startup
    send_state(ser)

    try:
        while True:
            line = ser.read_until(b"\n", 64).decode(errors="ignore").strip()
            if not line:
                continue
            print("<<", line)

            if line == "MIC:STATE?":
                send_state(ser)

            elif line == "MIC:TOGGLE":
                try:
                    muted = toggle_mic()
                    # Echo the new state back
                    ser.write(b"MIC:STATE MUTED\n" if muted else b"MIC:STATE UNMUTED\n")
                    print("Mic is now", "MUTED" if muted else "UNMUTED")
                except Exception as e:
                    print("Toggle failed:", repr(e))
    except KeyboardInterrupt:
        pass
    finally:
        ser.close()
