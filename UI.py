
import sys, time, serial, threading, os, subprocess
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal, QTimer, QProcess
from serial.tools import list_ports

from comtypes import CLSCTX_ALL, GUID
from comtypes.client import CreateObject
from ctypes import POINTER, cast
from pycaw.pycaw import IAudioEndpointVolume, IMMDeviceEnumerator

# CoreAudio constants
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
    dev  = enum.GetDefaultAudioEndpoint(eCapture, eConsole)   # default mic
    iface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(iface, POINTER(IAudioEndpointVolume))

def get_mic_muted():
    return bool(_endpoint_volume().GetMute())

def toggle_mic():
    vol = _endpoint_volume()
    cur = bool(vol.GetMute())
    vol.SetMute(0 if cur else 1, None)
    return not cur

def pick_port(prefer="COM6"):
    ports = list(list_ports.comports())
    for p in ports:
        if p.device == prefer:
            return prefer
    for p in ports:
        text = ((p.interface or "") + (p.description or "")).lower()
        if "data" in text:
            return p.device
    for p in ports:
        hint = " ".join(filter(None, [p.description, p.manufacturer, p.product])).lower()
        if any(k in hint for k in ["circuitpython", "adafruit", "rp2", "pico", "raspberry"]):
            return p.device
    return None

class SerialWorker(QThread):
    line = pyqtSignal(str)
    status = pyqtSignal(str)
    def __init__(self, port, baud=115200, timeout=0.2):
        super().__init__()
        self.port, self.baud, self.timeout = port, baud, timeout
        self._stop = threading.Event()
        self._lock = threading.Lock()
        self.ser = None
    def run(self):
        try:
            self.ser = serial.Serial(self.port, self.baud, timeout=self.timeout)
            time.sleep(0.3); self.ser.reset_input_buffer()
            self.status.emit(f"Connected: {self.port}")
            while not self._stop.is_set():
                raw = self.ser.read_until(b"\n", 128)
                if raw:
                    self.line.emit(raw.decode("utf-8", "ignore").strip())
        except Exception:
            pass
        finally:
            try:
                if self.ser: self.ser.close()
            except: pass
            self.status.emit("Disconnected")
    def write_line(self, s: str):
        with self._lock:
            if self.ser:
                try:
                    self.ser.write((s + "\n").encode("utf-8"))
                except Exception:
                    pass
    def stop(self):
        self._stop.set()

class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        central = QWidget(self); self.setCentralWidget(central)
        lay = QVBoxLayout(central)

        self.status_lbl = QLabel("Disconnected"); lay.addWidget(self.status_lbl)
        self.mic_lbl = QLabel("Mic: ?"); lay.addWidget(self.mic_lbl)

        self.btn = QPushButton("Pick & launch an app…")
        self.btn.clicked.connect(self.pick_path)
        lay.addWidget(self.btn)

        self.app_path = ""  

        port = pick_port("COM6")
        if port:
            self.worker = SerialWorker(port)
            self.worker.status.connect(self.status_lbl.setText)
            self.worker.line.connect(self.on_line)
            self.worker.start()
            QTimer.singleShot(600, self.push_state_to_pico)
        else:
            self.worker = None
            self.status_lbl.setText("No serial port found")

    def pick_path(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose app/shortcut",
            "",
            "Executables/Shortcuts (*.exe *.lnk *.bat *.cmd *.url);;All files (*)"
        )
        if not path:
            return
        self.app_path = os.path.expandvars(os.path.expanduser(path))
        self.status_lbl.setText(f"App set: {self.app_path}")
        # launch immediately 
        self.launch_app(self.app_path)

    def launch_app(self, path: str):
        try:
            p = path
            # Let Windows resolve shortcuts/URLs
            if p.lower().endswith((".lnk", ".url")) or not os.path.exists(p):
                os.startfile(p)
            elif p.lower().endswith((".bat", ".cmd")):
                subprocess.Popen(["cmd", "/c", p])
            else:
                if not QProcess.startDetached(p):
                    subprocess.Popen([p], shell=False)
            self.status_lbl.setText(f"Launched: {p}")
        except Exception as e:
            self.status_lbl.setText(f"Launch failed: {e}")

    def push_state_to_pico(self):
        if not self.worker: return
        try:
            muted = get_mic_muted()
            self.mic_lbl.setText("Mic: " + ("MUTED" if muted else "UNMUTED"))
            self.worker.write_line("MIC:STATE MUTED" if muted else "MIC:STATE UNMUTED")
        except Exception as e:
            self.mic_lbl.setText(f"Mic: ? ({e})")

    def on_line(self, text: str):
        print("<<", text)
        if text == "MIC:TOGGLE":
            try:
                muted = toggle_mic()
                self.mic_lbl.setText("Mic: " + ("MUTED" if muted else "UNMUTED"))
                self.worker.write_line("MIC:STATE MUTED" if muted else "MIC:STATE UNMUTED")
            except Exception as e:
                self.status_lbl.setText(f"Toggle failed: {e}")
        elif text == "MIC:STATE?":
            self.push_state_to_pico()
        elif text.startswith("BTN:"):
            if not self.app_path:
                self.status_lbl.setText("No app path set. Click 'Pick & launch…' first.")
                return
            self.launch_app(self.app_path)

    def closeEvent(self, e):
        if self.worker:
            self.worker.stop()
            self.worker.wait(1000)
        super().closeEvent(e)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = Main(); w.resize(420, 180); w.show()
    sys.exit(app.exec_())
