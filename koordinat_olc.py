# Small helper to measure the screen coordinate of a button in Akinsoft (Windows only).
# Usage: open a command prompt on the shop PC and run  python koordinat_olc.py
# Move the mouse over the button you want to measure and keep it still for 3 seconds:
# the coordinate is printed. Press Ctrl+C to quit.
# Put the printed (x, y) value into the matching COORD_... constant in app.py.
import ctypes
import time


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


def cursor_position():
    p = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(p))
    return p.x, p.y


print("Move the mouse over the button to measure and keep it still for 3 s. Quit: Ctrl+C")
last = None
still_since = time.time()
try:
    while True:
        now = cursor_position()
        if now != last:
            last = now
            still_since = time.time()
        elif time.time() - still_since >= 3 and last is not None:
            print(f"-> ({last[0]}, {last[1]})")
            still_since = time.time() + 1e9   # do not print the same point again
        time.sleep(0.1)
except KeyboardInterrupt:
    print("Done.")
