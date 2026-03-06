# mic_tray_fixed.py
# Требует: pip install pystray pillow pycaw comtypes
import sys
import os
import threading
import ctypes
import tkinter as tk
from pystray import Icon, Menu, MenuItem
from PIL import Image, ImageDraw

# --- optional: pycaw для управления громкостью микрофона ---
try:
    from ctypes import POINTER, cast
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    devices = AudioUtilities.GetMicrophone()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
except Exception:
    volume = None

def set_volume(val):
    try:
        if volume is not None:
            volume.SetMasterVolumeLevelScalar(float(val)/100.0, None)
    except Exception as e:
        print("set_volume error:", e)

# Попытка скрыть консоль (лучше запускать через pythonw)
def hide_console_window():
    if sys.platform != "win32":
        return
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        pass

# Загрузка иконки, корректно и при запуске из PyInstaller
def resource_path(fname):
    # Если упаковано pyinstaller --onefile, ресурсы в sys._MEIPASS
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base, fname)

def load_tray_image(path='mic.ico', size=64):
    """Попытка загрузить .ico/.png. Возвращает PIL.Image."""
    try:
        full = resource_path(path)
        img = Image.open(full)
        # Убедимся, что изображение имеет альфу и подходящий размер.
        img = img.convert("RGBA")
        # pystray может принимать разные размеры; оставим как есть, но при необходимости ресайзим
        return img
    except Exception:
        # Фолбэк — простая рисованная иконка микрофона
        s = size
        img = Image.new('RGBA', (s, s), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        cx, cy = s//2, s//2 - 6
        r = s//6
        draw.ellipse((cx-r, cy-r, cx+r, cy+r), fill="white")
        draw.rectangle((cx-6, cy+r-2, cx+6, cy+r+18), fill="white")
        draw.rectangle((cx-14, cy+r+18, cx+14, cy+r+22), fill="white")
        return img

# UI-thread management
root_lock = threading.Lock()
root_ref = {"root": None}

def create_or_show_window():
    with root_lock:
        if root_ref["root"] is not None:
            try:
                r = root_ref["root"]
                r.deiconify()
                r.lift()
                return
            except Exception:
                root_ref["root"] = None

    def tk_thread():
        root = tk.Tk()
        root.title("Mic Volume")
        root.geometry("320x110")

        def on_close():
            root.withdraw()
        root.protocol("WM_DELETE_WINDOW", on_close)

        tk.Label(root, text="Microphone volume").pack(anchor="w", padx=12, pady=(8,0))

        scale = tk.Scale(root, from_=0, to=100, orient="horizontal", command=set_volume)
        try:
            if volume is not None:
                current = volume.GetMasterVolumeLevelScalar()
                scale.set(int(current * 100))
            else:
                scale.set(50)
        except Exception:
            scale.set(50)
        scale.pack(fill="x", padx=12, pady=8)

        with root_lock:
            root_ref["root"] = root

        root.mainloop()
        # Когда mainloop заканчивает — очистим ссылку
        with root_lock:
            root_ref["root"] = None

    t = threading.Thread(target=tk_thread, daemon=True)
    t.start()

def open_window(icon, item):
    create_or_show_window()

def quit_app(icon, item):
    """
    Безопасно закрываем Tk (через its thread) и затем останавливаем icon.
    Всё делаем в отдельном потоке, чтобы не блокировать pystray.
    """
    def do_quit():
        # 1) Скрыть саму иконку, чтобы пользователь видел, что оно уходит
        try:
            icon.visible = False
        except Exception:
            pass

        # 2) Попросить Tk корректно завершиться через его event loop
        with root_lock:
            r = root_ref["root"]
            if r is not None:
                try:
                    # schedule destroy in the Tk thread
                    r.after(0, r.destroy)
                except Exception:
                    try:
                        r.quit()
                    except Exception:
                        pass

        # 3) потом останавливаем pystray icon
        try:
            icon.stop()
        except Exception:
            pass

    threading.Thread(target=do_quit, daemon=True).start()

# Меню — делаем Open default, чтобы left-click работал
menu = Menu(
    MenuItem("Open", open_window, default=True),
    MenuItem("Quit", quit_app)
)

# Загружаем иконку (mic.ico), либо рисуем fallback
tray_image = load_tray_image('mic.ico')

icon = Icon("mic_volume", tray_image, "Mic Volume", menu=menu)

if __name__ == "__main__":
    hide_console_window()
    icon.run()