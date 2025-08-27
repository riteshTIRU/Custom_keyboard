
import time, board, digitalio, usb_hid, usb_cdc
from adafruit_hid.keyboard import Keyboard
from adafruit_hid.keycode import Keycode


kbd = Keyboard(usb_hid.devices)
ser = usb_cdc.data

def write_line(s: str):
    if ser and ser.connected:
        try: ser.write((s+"\n").encode())
        except Exception: pass

def send_keys(*keys):
    kbd.send(*keys)

def mic_toggle():
    write_line("MIC:TOGGLE")
    
def launch_app():
    write_line("BTN:GP0")


BUTTONS_DEF = [
    (board.GP2, send_keys, (Keycode.LEFT_CONTROL, Keycode.LEFT_SHIFT, Keycode.ESCAPE)),  
    (board.GP3, mic_toggle, ()),                                                         
    (board.GP1, send_keys, (Keycode.UP_ARROW,)),
    (board.GP5, send_keys, (Keycode.DOWN_ARROW,)),
    (board.GP4, send_keys, (Keycode.LEFT_ARROW,)),
    (board.GP6, send_keys, (Keycode.RIGHT_ARROW,)),
    (board.GP7, send_keys, (Keycode.WINDOWS,Keycode.TAB)),
    #(board.GP0, send_keys, (Keycode.WINDOWS, Keycode.FIVE)),
    (board.GP0, launch_app,()),
    
]


buttons = []
for pin_id, action, args in BUTTONS_DEF:
    p = digitalio.DigitalInOut(pin_id)
    p.switch_to_input(pull=digitalio.Pull.UP)   
    buttons.append((p, action, args))

debounce = 0.03
last_state = [p.value for p,_,_ in buttons]
last_time  = [time.monotonic()] * len(buttons)

time.sleep(1.0)  


while True:
    now = time.monotonic()
    for i, (p, action, args) in enumerate(buttons):
        state = p.value
        if state != last_state[i] and (now - last_time[i]) > debounce:
            last_time[i]  = now
            last_state[i] = state
            if state is False:          
                action(*args)           
    time.sleep(0.003)  
