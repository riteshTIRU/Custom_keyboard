# boot.py
import usb_cdc, usb_hid
usb_cdc.enable(console=True, data=True)       # <-- creates a 2nd COM port
usb_hid.enable((usb_hid.Device.KEYBOARD,))    # optional
