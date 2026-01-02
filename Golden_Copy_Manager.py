import json
import os
import queue
import re
import socket
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageEnhance, ImageFilter, ImageTk

# Optional deps (tray + cross-platform hotkeys)
try:
    import pystray  # type: ignore
    from pystray import MenuItem as TrayItem  # type: ignore
    PYSTRAY_AVAILABLE = True
except Exception:
    pystray = None  # type: ignore
    TrayItem = None  # type: ignore
    PYSTRAY_AVAILABLE = False

try:
    from pynput import keyboard as pynput_keyboard  # type: ignore
    PYNPUT_AVAILABLE = True
except Exception:
    pynput_keyboard = None  # type: ignore
    PYNPUT_AVAILABLE = False

PIL_AVAILABLE = True

MAX_ITEMS = 50
POLL_MS = 500

MAX_SEARCH_CHARS = 200
SEARCH_DEBOUNCE_MS = 150

APP_TITLE = "Golden Copy Manager"

def app_dir() -> Path:
    # Where the app "lives" (script folder now, exe folder later)
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def resource_dir() -> Path:
    # Where bundled resources are located (PyInstaller uses sys._MEIPASS)
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return app_dir()

APP_DIR = app_dir()
RES_DIR = resource_dir()

ASSETS_DIR = RES_DIR / "assets"
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

HISTORY_PATH    = DATA_DIR / "clipboard_history.json"
SETTINGS_PATH  = DATA_DIR / "clipboard_manager_settings.json"

ICON_ICO = ASSETS_DIR / "icon.ico"
ICON_PNG = ASSETS_DIR / "icon.png"
BG_PATH  = ASSETS_DIR / "background.png"


# Bigger previews so horizontal scrollbar is useful:
PREVIEW_CHARS = 2000  # show up to this many chars in the list line
# (If original is longer, we append " …")

# Timestamp shown next to each entry
TIMESTAMP_FORMAT = "%Y-%m-%d %H:%M:%S"
PIN_ICON = "📌"

def _exe_path():
    # When frozen by PyInstaller, sys.executable is the .exe path
    if getattr(sys, "frozen", False):
        return sys.executable
    # When running as .py
    return os.path.abspath(sys.argv[0])

def _has_startup_flag(argv=None) -> bool:
    """Return True if we should start hidden (e.g., when launched by Startup)."""
    argv = argv or sys.argv
    flags = {"--startup", "--start-hidden", "--hidden", "--minimized"}
    return any(str(a).lower() in flags for a in argv[1:])

def set_start_with_windows(enable: bool, app_name: str = "GoldenCopyManager"):
    """Enable/disable auto-start on Windows via HKCU Run key."""
    if sys.platform != "win32":
        return  # no-op on non-Windows

    import winreg

    run_key = r"Software\Microsoft\Windows\CurrentVersion\Run"
    exe = _exe_path()

    # Quote path for spaces. You can also add args after the closing quote if needed.
    cmd = f'"{exe}" --startup'

    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key, 0, winreg.KEY_SET_VALUE) as key:
        if enable:
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, cmd)
        else:
            try:
                winreg.DeleteValue(key, app_name)
            except FileNotFoundError:
                pass

def is_start_with_windows_enabled(app_name: str = "GoldenCopyManager") -> bool:
    if sys.platform != "win32":
        return False
    import winreg
    run_key = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key, 0, winreg.KEY_READ) as key:
            winreg.QueryValueEx(key, app_name)
        return True
    except FileNotFoundError:
        return False

# ---- Single-instance (run again -> show existing) ----
SINGLE_INSTANCE_HOST = "127.0.0.1"
SINGLE_INSTANCE_PORT = 50677  # change if you ever conflict with another app


class SingleInstanceServer:
    def __init__(self, host: str, port: int, event_queue: "queue.Queue[tuple]"):
        self.host = host
        self.port = port
        self.q = event_queue
        self._stop = threading.Event()
        self._sock = None
        self._thread = None

    def start(self) -> bool:
        """Try to bind and start server. Returns True if this is the first instance."""
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.bind((self.host, self.port))
            s.listen(5)
            s.settimeout(0.5)
            self._sock = s
        except OSError:
            return False

        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        return True

    def stop(self):
        self._stop.set()
        try:
            if self._sock:
                self._sock.close()
        except Exception:
            pass

    def _run(self):
        while not self._stop.is_set():
            try:
                conn, _ = self._sock.accept()
            except socket.timeout:
                continue
            except Exception:
                break

            try:
                data = conn.recv(1024) or b""
                if data.strip().upper().startswith(b"SHOW"):
                    self.q.put(("SHOW", None))
                conn.sendall(b"OK\n")
            except Exception:
                pass
            finally:
                try:
                    conn.close()
                except Exception:
                    pass

    @staticmethod
    def send_show(host: str, port: int) -> bool:
        """Ask an existing instance to show itself."""
        try:
            with socket.create_connection((host, port), timeout=0.25) as s:
                s.sendall(b"SHOW\n")
                s.recv(16)
            return True
        except Exception:
            return False


# ---- Windows global hotkey (works while hidden) ----
class WindowsHotkeyListener:
    WM_APP_SET_HOTKEY = 0x8000 + 1

    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    MOD_WIN = 0x0008
    MOD_NOREPEAT = 0x4000

    def __init__(self, event_queue):
        self.q = event_queue
        self._thread = None
        self._stop = threading.Event()

        self._hwnd = None
        self._hotkey_id = 1
        self._mods = self.MOD_CONTROL | self.MOD_ALT | self.MOD_NOREPEAT
        self._vk = ord("V")
        self._display = "Ctrl+Alt+V"

        self._pending_mods = None
        self._pending_vk = None
        self._pending_display = None

        self._ready = threading.Event()

    def start(self):
        if os.name != "nt":
            self.q.put(("INFO", "Hotkey only works on Windows."))
            return
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._ready.wait(timeout=2.0)

    def stop(self):
        if os.name != "nt":
            return
        try:
            import ctypes
            user32 = ctypes.windll.user32
            if self._hwnd:
                user32.PostMessageW(self._hwnd, 0x0010, 0, 0)  # WM_CLOSE
        except Exception:
            pass
        self._stop.set()

    def set_hotkey(self, mods: int, vk: int, display: str):
        if os.name != "nt":
            self.q.put(("ERROR", "Hotkey change only supported on Windows."))
            return
        self._pending_mods = mods
        self._pending_vk = vk
        self._pending_display = display
        try:
            import ctypes
            user32 = ctypes.windll.user32
            if self._hwnd:
                user32.PostMessageW(self._hwnd, self.WM_APP_SET_HOTKEY, 0, 0)
        except Exception as e:
            self.q.put(("ERROR", f"Failed to request hotkey update: {e}"))

    def _run(self):
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        # ---- Pointer-sized Win32 types (fixes OverflowError on 64-bit) ----
        if ctypes.sizeof(ctypes.c_void_p) == 8:
            WPARAM = ctypes.c_uint64
            LPARAM = ctypes.c_int64
            LRESULT = ctypes.c_int64
        else:
            WPARAM = ctypes.c_uint32
            LPARAM = ctypes.c_int32
            LRESULT = ctypes.c_int32

        WNDPROCTYPE = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, WPARAM, LPARAM)

        # Define MSG with pointer-sized wParam/lParam
        class POINT(ctypes.Structure):
            _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

        class MSG(ctypes.Structure):
            _fields_ = [
                ("hwnd", wintypes.HWND),
                ("message", wintypes.UINT),
                ("wParam", WPARAM),
                ("lParam", LPARAM),
                ("time", wintypes.DWORD),
                ("pt", POINT),
            ]

        class WNDCLASSW(ctypes.Structure):
            _fields_ = [
                ("style", wintypes.UINT),
                ("lpfnWndProc", WNDPROCTYPE),
                ("cbClsExtra", ctypes.c_int),
                ("cbWndExtra", ctypes.c_int),
                ("hInstance", wintypes.HINSTANCE),
                ("hIcon", wintypes.HICON),
                ("hCursor", wintypes.HCURSOR),
                ("hbrBackground", wintypes.HBRUSH),
                ("lpszMenuName", wintypes.LPCWSTR),
                ("lpszClassName", wintypes.LPCWSTR),
            ]

        # IMPORTANT: set DefWindowProcW prototype so ctypes won’t guess wrong
        user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, WPARAM, LPARAM]
        user32.DefWindowProcW.restype = LRESULT

        user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
        user32.RegisterClassW.restype = wintypes.ATOM

        user32.CreateWindowExW.restype = wintypes.HWND

        user32.GetMessageW.argtypes = [ctypes.POINTER(MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
        user32.GetMessageW.restype = wintypes.BOOL

        user32.TranslateMessage.argtypes = [ctypes.POINTER(MSG)]
        user32.DispatchMessageW.argtypes = [ctypes.POINTER(MSG)]

        WM_DESTROY = 0x0002
        WM_HOTKEY = 0x0312
        WM_CLOSE = 0x0010

        def register_hotkey(hwnd, mods, vk):
            return bool(user32.RegisterHotKey(hwnd, self._hotkey_id, mods, vk))

        def unregister_hotkey(hwnd):
            try:
                user32.UnregisterHotKey(hwnd, self._hotkey_id)
            except Exception:
                pass

        @WNDPROCTYPE
        def wndproc(hwnd, msg, wparam, lparam):
            if msg == WM_HOTKEY and int(wparam) == self._hotkey_id:
                self.q.put(("HOTKEY", None))
                return 0

            if msg == self.WM_APP_SET_HOTKEY:
                new_mods = self._pending_mods
                new_vk = self._pending_vk
                new_disp = self._pending_display
                if new_mods is None or new_vk is None or not new_disp:
                    return 0

                unregister_hotkey(hwnd)
                if register_hotkey(hwnd, new_mods, new_vk):
                    self._mods, self._vk, self._display = new_mods, new_vk, new_disp
                    self.q.put(("INFO", f"Hotkey set to: {new_disp}"))
                else:
                    # rollback
                    if register_hotkey(hwnd, self._mods, self._vk):
                        self.q.put(("ERROR", f"Hotkey '{new_disp}' unavailable (already used). Kept: {self._display}"))
                    else:
                        self.q.put(("ERROR", "Hotkey registration failed and rollback failed."))
                return 0

            if msg == WM_CLOSE:
                user32.DestroyWindow(hwnd)
                return 0

            if msg == WM_DESTROY:
                unregister_hotkey(hwnd)
                user32.PostQuitMessage(0)
                return 0

            return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        hInstance = kernel32.GetModuleHandleW(None)
        class_name = "ClipboardManagerHotkeyWindow"

        wc = WNDCLASSW()
        wc.style = 0
        wc.lpfnWndProc = wndproc
        wc.cbClsExtra = 0
        wc.cbWndExtra = 0
        wc.hInstance = hInstance
        wc.hIcon = 0
        wc.hCursor = 0
        wc.hbrBackground = 0
        wc.lpszMenuName = None
        wc.lpszClassName = class_name

        user32.RegisterClassW(ctypes.byref(wc))

        hwnd = user32.CreateWindowExW(
            0, class_name, "hotkey", 0,
            0, 0, 0, 0,
            0, 0, hInstance, None
        )
        self._hwnd = hwnd

        if register_hotkey(hwnd, self._mods, self._vk):
            self.q.put(("INFO", f"Hotkey ready: {self._display}"))
        else:
            self.q.put(("ERROR", f"Failed to register hotkey: {self._display} (already used?)"))

        self._ready.set()

        msg = MSG()
        while not self._stop.is_set():
            r = user32.GetMessageW(ctypes.byref(msg), 0, 0, 0)
            if r <= 0:
                break
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))



class PynputHotkeyListener:
    """Cross-platform global hotkey via pynput (macOS/Linux; optional dependency)."""

    def __init__(self, event_queue: "queue.Queue[tuple]"):
        self.q = event_queue
        self._listener = None
        self._thread = None
        self._lock = threading.Lock()
        self._combo = None
        self._display = None

    def start(self, combo: str, display: str):
        if not PYNPUT_AVAILABLE:
            self.q.put(("INFO", "Global hotkey on macOS/Linux requires 'pynput' (pip install pynput)."))
            return
        self.set_hotkey(combo, display)

    def stop(self):
        with self._lock:
            try:
                if self._listener is not None:
                    self._listener.stop()
            except Exception:
                pass
            self._listener = None

    def set_hotkey(self, combo: str, display: str):
        if not PYNPUT_AVAILABLE:
            self.q.put(("ERROR", "Cannot set hotkey: 'pynput' is not installed."))
            return

        with self._lock:
            # Replace any existing listener
            try:
                if self._listener is not None:
                    self._listener.stop()
            except Exception:
                pass

            self._combo = combo
            self._display = display

            try:
                self._listener = pynput_keyboard.GlobalHotKeys({combo: lambda: self.q.put(("HOTKEY", None))})
                self._thread = threading.Thread(target=self._listener.run, daemon=True)
                self._thread.start()
                self.q.put(("INFO", f"Hotkey ready: {display}"))
            except Exception as e:
                self._listener = None
                self.q.put(("ERROR", f"Failed to register hotkey '{display}': {e}"))


class TrayManager:
    """Windows tray icon (optional dependency: pystray)."""

    def __init__(self, event_queue: "queue.Queue[tuple]"):
        self.q = event_queue
        self._icon = None
        self._thread = None

    def start(self) -> bool:
        if os.name != "nt":
            return False
        if not PYSTRAY_AVAILABLE:
            self.q.put(("INFO", "Tray icon requires 'pystray' (pip install pystray)."))
            return False
        if self._thread and self._thread.is_alive():
            return True

        # Pick an icon image
        img = None
        try:
            if ICON_PNG.exists():
                img = Image.open(str(ICON_PNG)).convert("RGBA")
        except Exception:
            img = None

        if img is None:
            # fallback: simple generated icon
            img = Image.new("RGBA", (64, 64), (40, 40, 40, 255))

        def on_show_hide(icon, item):
            self.q.put(("TRAY_TOGGLE", None))

        def on_toggle_monitor(icon, item):
            self.q.put(("TRAY_TOGGLE_MONITOR", None))

        def on_exit(icon, item):
            self.q.put(("TRAY_EXIT", None))

        menu = pystray.Menu(
            TrayItem("Show / Hide", on_show_hide),
            TrayItem("Pause / Resume monitoring", on_toggle_monitor),
            TrayItem("Exit", on_exit),
        )

        self._icon = pystray.Icon("GoldenCopyManager", img, APP_TITLE, menu)

        def run():
            try:
                self._icon.run()
            except Exception as e:
                self.q.put(("ERROR", f"Tray icon failed: {e}"))

        self._thread = threading.Thread(target=run, daemon=True)
        self._thread.start()
        return True

    def stop(self):
        try:
            if self._icon is not None:
                self._icon.stop()
        except Exception:
            pass
        self._icon = None


class ClipboardManager(tk.Tk):
    def __init__(self, event_queue: "queue.Queue[tuple]", ipc_server: SingleInstanceServer | None, start_hidden: bool = False):
        super().__init__()
        if start_hidden:
            # Start hidden (e.g., when launched at Windows logon)
            self.withdraw()
        self.title(APP_TITLE)
        self.geometry("760x540")
        self.minsize(650, 470)

        self.q = event_queue
        self.ipc_server = ipc_server

        # History items are dicts:
        #   {"text": str, "fold": str, "ts": float, "pinned": bool}
        self.history: list[dict] = []
        self.last_clip: str | None = None
        self.filtered_items: list[dict] = []
        self._search_job = None

        # Hotkey defaults (platform-aware)
        if sys.platform == "darwin":
            # Cmd+Shift+V feels more "native"
            self.hk_ctrl = False
            self.hk_alt = False
            self.hk_shift = True
            self.hk_win = True   # Cmd on macOS
            self.hk_key = "V"
        else:
            self.hk_ctrl = True
            self.hk_alt = True
            self.hk_shift = False
            self.hk_win = False
            self.hk_key = "V"

        self.hk_display = self._format_hotkey_display(self.hk_ctrl, self.hk_alt, self.hk_shift, self.hk_win, self.hk_key)
        self.hotkey_listener = None

        # Pause/resume monitoring
        self.monitoring_paused_var = tk.BooleanVar(value=False)

        # Tray (Windows)
        self.tray = TrayManager(self.q)

        self.search_warn_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="")

        # Startup toggle
        self.startup_var = tk.BooleanVar(value=self.is_startup_enabled())

        # Background image holders
        self._bg_original = None
        self._bg_photo = None
        self._bg_item = None

        self._apply_style()
        self._setup_icon()
        self._build_ui()

        self._load_settings()
        self._load_history()

        # Close hides to background (keeps hotkey + single-instance server alive)
        self.protocol("WM_DELETE_WINDOW", self.on_close_hide)

        # Keys
        self.listbox.bind("<Delete>", lambda e: self.remove_selected())
        self.listbox.bind("<BackSpace>", lambda e: self.remove_selected())
        self.bind("<Escape>", lambda e: self._hide_window())

        # Start clipboard polling
        self.after(POLL_MS, self._poll_clipboard)

        # Start hotkey
        self._setup_hotkey()

        # Start tray icon (Windows; optional dependency)
        self.tray.start()

        # Process hotkey + show-from-second-run + tray events
        self.after(50, self._process_events)

    # ---------- Styling / Assets ----------
    def _apply_style(self):
        style = ttk.Style(self)

        # Prefer native theme; avoid forcing 'clam' (it tends to paint solid widget backgrounds).
        try:
            names = set(style.theme_names())
            if os.name == "nt":
                for t in ("vista", "xpnative"):
                    if t in names:
                        style.theme_use(t)
                        break
            elif sys.platform == "darwin":
                if "aqua" in names:
                    style.theme_use("aqua")
        except Exception:
            pass

        self.option_add("*Font", ("Segoe UI", 10))
        style.configure("TButton", padding=(10, 6))
        style.configure("TEntry", padding=(6, 4))

        # Labels used on top of the glass background:
        # Use a layout with only the text element so ttk doesn't draw a solid background box.
        style.layout("Glass.TLabel", [("Label.label", {"sticky": "nswe"})])
        style.layout("Title.TLabel", [("Label.label", {"sticky": "nswe"})])
        style.layout("Sub.TLabel", [("Label.label", {"sticky": "nswe"})])

        # Make frames on the glass background draw nothing (so the image shows through).
        style.layout("Glass.TFrame", [])

        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Sub.TLabel", font=("Segoe UI", 10))

    def _setup_icon(self):
        try:
            if os.name == "nt" and ICON_ICO.exists():
                self.iconbitmap(str(ICON_ICO))
            elif ICON_PNG.exists():
                img = tk.PhotoImage(file=str(ICON_PNG))
                self.iconphoto(True, img)
                self._icon_ref = img  # keep reference so it doesn't get garbage collected
        except Exception as e:
            print("Icon load failed:", e)

    # ---------- UI ----------
    def _build_ui(self):
        # Menu bar
        menubar = tk.Menu(self)

        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Show", command=self._show_window)
        filemenu.add_command(label="Hide", command=self._hide_window)
        filemenu.add_separator()
        filemenu.add_command(label="Hotkey…", command=self.open_hotkey_dialog)
        filemenu.add_command(label="Save As…", command=self.save_as)

        filemenu.add_separator()
        filemenu.add_checkbutton(
            label="Pause clipboard monitoring",
            variable=self.monitoring_paused_var,
            command=self._monitoring_menu_changed,
        )

        filemenu.add_separator()
        filemenu.add_checkbutton(
            label="Run at startup (Windows)",
            variable=self.startup_var,
            command=self.toggle_startup,
        )

        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.exit_app)
        menubar.add_cascade(label="Setting", menu=filemenu)
        self.config(menu=menubar)

        # Canvas draws the real background image (full window)
        self.bg_canvas = tk.Canvas(self, highlightthickness=0, bd=0)
        self.bg_canvas.pack(fill="both", expand=True)
        self.bg_canvas.bind("<Configure>", self._on_resize)

        # This canvas image is the "real" background
        self._bg_item = None
        self._bg_photo = None
        self._bg_base_rgba = None  # PIL Image (RGBA) at current window size

        # This Frame holds all widgets; a background Label provides the "glass" layer
        # (blurred background + white overlay). Using a Frame as the container avoids
        # rare redraw glitches on Windows where ttk widgets can momentarily disappear
        # while moving the window quickly.
        self.glass = tk.Frame(self.bg_canvas, bd=0, highlightthickness=0)
        self.glass_bg = tk.Label(self.glass, bd=0, highlightthickness=0)
        self.glass_bg.place(x=0, y=0, relwidth=1, relheight=1)
        self.glass_bg.lower()
        self.glass_id = self.bg_canvas.create_window(0, 0, anchor="nw", window=self.glass)

        # IMPORTANT: build widgets directly on the glass frame
        # Use grid for layout
        self.glass.grid_columnconfigure(0, weight=1)
        self.glass.grid_rowconfigure(1, weight=1)

        # Top controls
        top = ttk.Frame(self.glass, style="Glass.TFrame")
        top.grid(row=0, column=0, sticky="ew", padx=14, pady=(0, 10))
        top.grid_columnconfigure(1, weight=1)

        ttk.Label(top, text="Search:", style="Glass.TLabel").grid(row=0, column=0, sticky="w")

        self.search_var = tk.StringVar()
        vcmd = (self.register(self._validate_search), "%P")
        self.search_entry = ttk.Entry(top, textvariable=self.search_var, validate="key", validatecommand=vcmd)
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(6, 10))
        self.search_entry.bind("<KeyRelease>", self._on_search_key)

        ttk.Button(top, text="Copy", command=self.copy_selected).grid(row=0, column=2, padx=4)
        ttk.Button(top, text="Pin", command=self.toggle_pin_selected).grid(row=0, column=3, padx=4)
        ttk.Button(top, text="Remove", command=self.remove_selected).grid(row=0, column=4, padx=4)
        ttk.Button(top, text="Clear", command=self.clear_all).grid(row=0, column=5, padx=4)

        # List area
        mid = ttk.Frame(self.glass, style="Glass.TFrame")
        mid.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 10))
        mid.grid_rowconfigure(0, weight=1)
        mid.grid_columnconfigure(0, weight=1)
        self.listbox = tk.Listbox(mid, activestyle="dotbox", height=20)
        self.listbox.grid(row=0, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(mid, orient="vertical", command=self.listbox.yview)
        yscroll.grid(row=0, column=1, sticky="ns")

        xscroll = ttk.Scrollbar(mid, orient="horizontal", command=self.listbox.xview)
        xscroll.grid(row=1, column=0, sticky="ew")

        self.listbox.config(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.listbox.bind("<Double-Button-1>", lambda e: self.copy_selected())
        self.listbox.bind("<Return>", lambda e: self.copy_selected())

        # Right-click menu in app
        self._ctx = tk.Menu(self, tearoff=0)
        self._ctx.add_command(label="Copy", command=self.copy_selected)
        self._ctx.add_command(label="Pin", command=self.toggle_pin_selected)
        self._ctx.add_command(label="Remove", command=self.remove_selected)
        self._ctx.add_separator()
        self._ctx.add_command(label="Pause monitoring", command=self.toggle_monitoring)
        self._ctx.add_separator()
        self._ctx.add_command(label="Exit", command=self.exit_app)

        self.listbox.bind("<Button-3>", self._show_context_menu)
        self.listbox.bind("<Button-2>", self._show_context_menu)  # macOS trackpads often use Button-2

        # # Status + warnings
        # bottom = ttk.Frame(self.glass, style="Glass.TFrame")
        # bottom.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 12))
        # bottom.grid_columnconfigure(0, weight=1)

        # warn = ttk.Label(bottom, textvariable=self.search_warn_var, style="Glass.TLabel")
        # warn.grid(row=0, column=0, sticky="w")

        # status = ttk.Label(bottom, textvariable=self.status_var, style="Glass.TLabel")
        # status.grid(row=1, column=0, sticky="w")

        # Load background image now; draw on first resize and after idle
        self._load_background()
        self.after_idle(lambda: self._render_layers(
            self.bg_canvas.winfo_width(), self.bg_canvas.winfo_height()
        ))

    def _show_context_menu(self, event):
        try:
            idx = self.listbox.nearest(event.y)
            if idx >= 0:
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(idx)
                self.listbox.activate(idx)
        except Exception:
            pass

        self._update_context_menu_labels()

        try:
            self._ctx.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self._ctx.grab_release()
            except Exception:
                pass

    def _update_context_menu_labels(self):
        # Pin label
        try:
            item = self._selected_history_item()
            pin_label = "Unpin" if (item and item.get("pinned")) else "Pin"
            self._ctx.entryconfig(1, label=pin_label)  # 0:Copy, 1:Pin
        except Exception:
            pass

        # Monitoring label
        try:
            monitor_label = "Resume monitoring" if self.monitoring_paused_var.get() else "Pause monitoring"
            self._ctx.entryconfig(4, label=monitor_label)  # 4 is the monitoring action
        except Exception:
            pass

    def _load_background(self):
        self._bg_original = None
        try:
            if not BG_PATH.exists():
                print("Background not found:", BG_PATH)
                return
            img = Image.open(str(BG_PATH)).convert("RGBA")
            self._bg_original = img
        except Exception as e:
            print("Failed to load background:", e)

    def _cover_resize(self, img: Image.Image, w: int, h: int) -> Image.Image:
        """Resize preserving aspect ratio, then center-crop to exactly w x h."""
        iw, ih = img.size
        scale = max(w / iw, h / ih)
        nw, nh = int(iw * scale), int(ih * scale)
        resized = img.resize((max(1, nw), max(1, nh)), Image.LANCZOS)
        left = (nw - w) // 2
        top = (nh - h) // 2
        return resized.crop((left, top, left + w, top + h))

    def _render_layers(self, w: int, h: int):
        """Render background + 'glass' overlay (moderate transparency) to fill the window."""
        if not PIL_AVAILABLE or self._bg_original is None:
            return
        if w <= 2 or h <= 2:
            return

        try:
            # 1) Base background
            base = self._cover_resize(self._bg_original, w, h).convert("RGBA")
            self._bg_base_rgba = base

            self._bg_photo = ImageTk.PhotoImage(base)
            if self._bg_item is None:
                self._bg_item = self.bg_canvas.create_image(0, 0, anchor="nw", image=self._bg_photo)
            else:
                self.bg_canvas.itemconfigure(self._bg_item, image=self._bg_photo)
            self.bg_canvas.tag_lower(self._bg_item)

            # 2) Glass overlay (blur + white overlay = looks semi-transparent)
            glass = base.filter(ImageFilter.GaussianBlur(radius=6))

            # white overlay with alpha (0..255). 110~150 is "moderate".
            overlay_alpha = 130
            overlay = Image.new("RGBA", (w, h), (255, 255, 255, overlay_alpha))
            glass = Image.alpha_composite(glass, overlay)

            # slightly reduce brightness to keep text readable
            glass = ImageEnhance.Brightness(glass).enhance(0.97)

            self._glass_photo = ImageTk.PhotoImage(glass)
            self.glass_bg.configure(image=self._glass_photo)
            # keep a reference so it doesn't get GC'd
            self.glass_bg._img_ref = self._glass_photo
        except Exception as e:
            print("Render layers failed:", e)

    def _on_resize(self, event):
        """Handle canvas <Configure>.

        On Windows, <Configure> can fire repeatedly while the window is being moved (x/y changes)
        even when the size is unchanged. Re-rendering the PIL background/glass layers on every move
        is expensive and can cause transient redraw glitches where ttk widgets appear blank until the
        next repaint. We only re-render when (width, height) changes, and we debounce heavy renders.
        """
        # Keep the glass frame covering the whole canvas
        try:
            self.bg_canvas.coords(self.glass_id, 0, 0)
            self.bg_canvas.itemconfigure(self.glass_id, width=event.width, height=event.height)
        except Exception:
            pass

        w, h = int(event.width), int(event.height)
        if getattr(self, "_last_render_wh", None) == (w, h):
            # Lightweight repaint only
            try:
                self.bg_canvas.update_idletasks()
            except Exception:
                pass
            return

        self._last_render_wh = (w, h)

        # Debounce heavy PIL re-render during rapid resizes
        try:
            if getattr(self, "_render_job", None) is not None:
                self.after_cancel(self._render_job)
        except Exception:
            pass

        self._render_job = self.after(50, lambda: self._render_layers(w, h))

    def _validate_search(self, proposed: str) -> bool:
        if len(proposed) <= MAX_SEARCH_CHARS:
            self.search_warn_var.set("")
            return True
        self.bell()
        self.search_warn_var.set(f"Search limited to {MAX_SEARCH_CHARS} characters")
        return False

    def _on_search_key(self, _event=None):
        if self._search_job is not None:
            try:
                self.after_cancel(self._search_job)
            except Exception:
                pass
        self._search_job = self.after(SEARCH_DEBOUNCE_MS, self._refresh_list)

    def _set_status(self, msg: str, clear_ms: int = 2500):
        self.status_var.set(msg)
        if clear_ms:
            self.after(clear_ms, lambda: self.status_var.set(""))

    # ---------- Clipboard ----------
    def _get_clipboard_text(self):
        try:
            data = self.clipboard_get()
            if isinstance(data, str):
                return data
        except tk.TclError:
            return None
        return None

    def _set_clipboard_text(self, text: str):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update_idletasks()

    def _poll_clipboard(self):
        try:
            text = self._get_clipboard_text()
            if text is not None:
                normalized = text.strip()
                if normalized:
                    if self.monitoring_paused_var.get():
                        # Update last_clip so resuming doesn't immediately capture something copied while paused.
                        self.last_clip = normalized
                    elif normalized != self.last_clip:
                        self.last_clip = normalized
                        self._add_to_history(normalized)
        except Exception:
            pass
        self.after(POLL_MS, self._poll_clipboard)

    # ---------- History ----------
    def _format_ts(self, ts) -> str:
        try:
            return datetime.fromtimestamp(float(ts)).strftime(TIMESTAMP_FORMAT)
        except Exception:
            return ""

    def _trim_history(self):
        # Pinned items are never pushed out by MAX_ITEMS; we only cap the unpinned section.
        pinned = [h for h in self.history if h.get("pinned")]
        unpinned = [h for h in self.history if not h.get("pinned")]

        if len(unpinned) > MAX_ITEMS:
            unpinned = unpinned[:MAX_ITEMS]

        self.history = pinned + unpinned

    def _add_to_history(self, item: str):
        now = time.time()
        pinned = False

        for i, h in enumerate(self.history):
            if h.get("text") == item:
                pinned = bool(h.get("pinned", False))
                self.history.pop(i)
                break

        entry = {"text": item, "fold": item.casefold(), "ts": now, "pinned": pinned}

        if pinned:
            # Keep pinned items at the top
            self.history.insert(0, entry)
        else:
            pinned_count = sum(1 for h in self.history if h.get("pinned"))
            self.history.insert(pinned_count, entry)

        self._trim_history()
        self._refresh_list()
        self._auto_save_history()

    def _selected_history_item(self):
        sel = self.listbox.curselection()
        if not sel:
            return None
        idx = sel[0]
        if 0 <= idx < len(self.filtered_items):
            return self.filtered_items[idx]
        return None

    def _make_preview(self, h: dict) -> str:
        text = h.get("text", "")
        snippet = text[:PREVIEW_CHARS].replace("\n", " ⏎ ")
        if len(text) > PREVIEW_CHARS:
            snippet += " …"

        pin = f"{PIN_ICON} " if h.get("pinned") else ""
        ts = self._format_ts(h.get("ts"))
        ts = f"{ts}  " if ts else ""
        return f"{pin}{ts}{snippet}"

    def _refresh_list(self):
        selected = self._selected_history_item()
        selected_text = selected.get("text") if selected else None

        q = self.search_var.get().strip().casefold()
        self.listbox.delete(0, tk.END)
        self.filtered_items = []

        for h in self.history:
            if (not q) or (q in h.get("fold", "")):
                self.filtered_items.append(h)
                self.listbox.insert(tk.END, self._make_preview(h))

        if selected_text:
            for i, h in enumerate(self.filtered_items):
                if h.get("text") == selected_text:
                    self.listbox.selection_set(i)
                    self.listbox.activate(i)
                    self.listbox.see(i)
                    break

    def copy_selected(self):
        item = self._selected_history_item()
        if item is None:
            self._set_status("Select an item first.")
            return
        self._set_clipboard_text(item.get("text", ""))
        self._set_status("Copied back to clipboard!")

    def remove_selected(self):
        item = self._selected_history_item()
        if item is None:
            self._set_status("Select an item first.")
            return
        txt = item.get("text")
        self.history = [h for h in self.history if h.get("text") != txt]
        self._refresh_list()
        self._auto_save_history()
        self._set_status("Removed.")

    def toggle_pin_selected(self):
        item = self._selected_history_item()
        if item is None:
            self._set_status("Select an item first.")
            return

        txt = item.get("text", "")
        pinned = not bool(item.get("pinned", False))
        ts = float(item.get("ts", time.time()))

        # Remove and re-insert to keep pinned items on top
        self.history = [h for h in self.history if h.get("text") != txt]
        entry = {"text": txt, "fold": txt.casefold(), "ts": ts, "pinned": pinned}
        if pinned:
            self.history.insert(0, entry)
            self._set_status("Pinned.")
        else:
            pinned_count = sum(1 for h in self.history if h.get("pinned"))
            self.history.insert(pinned_count, entry)
            self._set_status("Unpinned.")

        self._trim_history()
        self._refresh_list()
        self._auto_save_history()

    def toggle_monitoring(self):
        # Used by right-click and tray: flips the state, then applies it.
        self.monitoring_paused_var.set(not bool(self.monitoring_paused_var.get()))
        self._monitoring_menu_changed()

    def _monitoring_menu_changed(self):
        # Used by the Settings menu checkbutton: reads the current state (no flipping).
        if self.monitoring_paused_var.get():
            self._set_status("Clipboard monitoring paused.")
        else:
            self._set_status("Clipboard monitoring resumed.")
        self._save_settings()

    def clear_all(self):
        if not self.history:
            return
        if messagebox.askyesno(APP_TITLE, "Clear all clipboard history?"):
            self.history.clear()
            self._refresh_list()
            self._auto_save_history()
            self._set_status("Cleared.")

    # ---------- Save/Load ----------
    def _load_history(self):
        if not os.path.exists(HISTORY_PATH):
            return
        try:
            with open(HISTORY_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)

            items: list[dict] = []
            now = time.time()

            if isinstance(data, list):
                if data and isinstance(data[0], dict):
                    # New format: list of dicts
                    for d in data:
                        try:
                            txt = str(d.get("text", "")).strip()
                        except Exception:
                            continue
                        if not txt:
                            continue
                        ts = d.get("ts", now)
                        pinned = bool(d.get("pinned", False))
                        items.append({"text": txt, "fold": txt.casefold(), "ts": float(ts), "pinned": pinned})
                else:
                    # Legacy format: list[str]
                    for i, t in enumerate(data):
                        if not isinstance(t, str):
                            continue
                        txt = t.strip()
                        if not txt:
                            continue
                        items.append({"text": txt, "fold": txt.casefold(), "ts": now - i, "pinned": False})

            # De-dupe (keep first occurrence)
            seen = set()
            uniq: list[dict] = []
            for h in items:
                txt = h.get("text")
                if not txt or txt in seen:
                    continue
                seen.add(txt)
                uniq.append(h)

            self.history = uniq
            self._trim_history()
            self._refresh_list()
        except Exception:
            pass

    def _auto_save_history(self):
        try:
            with open(HISTORY_PATH, "w", encoding="utf-8") as f:
                payload = [{"text": h.get("text", ""), "ts": h.get("ts"), "pinned": bool(h.get("pinned", False))} for h in self.history]
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def save_as(self):
        path = filedialog.asksaveasfilename(
            title="Save Clipboard History",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                payload = [{"text": h.get("text", ""), "ts": h.get("ts"), "pinned": bool(h.get("pinned", False))} for h in self.history]
                json.dump(payload, f, ensure_ascii=False, indent=2)
            self._set_status("Saved.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Save failed:\n{e}")

    def startup_folder(self) -> str | None:
        if os.name != "nt":
            return None
        appdata = os.environ.get("APPDATA")
        if not appdata:
            return None
        return os.path.join(appdata, r"Microsoft\Windows\Start Menu\Programs\Startup")

    def startup_script_path(self) -> str | None:
        folder = self.startup_folder()
        if not folder:
            return None
        return os.path.join(folder, "ClipboardManagerStartup.cmd")

    def is_startup_enabled(self) -> bool:
        p = self.startup_script_path()
        return bool(p and os.path.exists(p))

    def toggle_startup(self):
        if os.name != "nt":
            self._set_status("Startup option only works on Windows.", 4000)
            self.startup_var.set(False)
            return

        if self.startup_var.get():
            ok = self.enable_startup()
            if not ok:
                self.startup_var.set(False)
        else:
            ok = self.disable_startup()
            if not ok:
                self.startup_var.set(True)

    def enable_startup(self) -> bool:
        p = self.startup_script_path()
        if not p:
            self._set_status("Startup folder not found.", 4000)
            return False

        # If packaged later, sys.executable is the .exe. If running as script, use pythonw + script path.
        if getattr(sys, "frozen", False):
            target = f'"{sys.executable}" --startup'
        else:
            python_exe = sys.executable
            pythonw = os.path.join(os.path.dirname(python_exe), "pythonw.exe")
            if os.path.exists(pythonw):
                target = f'"{pythonw}" "{os.path.abspath(__file__)}" --startup'
            else:
                # fallback to python.exe (may open console)
                target = f'"{python_exe}" "{os.path.abspath(__file__)}" --startup'

        try:
            with open(p, "w", encoding="utf-8") as f:
                f.write("@echo off\n")
                f.write("start \"\" /min " + target + "\n")
            self._set_status("Enabled: Run at startup.")
            return True
        except Exception as e:
            self._set_status(f"Failed to enable startup: {e}", 5000)
            return False

    def disable_startup(self) -> bool:
        p = self.startup_script_path()
        if not p:
            self._set_status("Startup folder not found.", 4000)
            return False
        try:
            if os.path.exists(p):
                os.remove(p)
            self._set_status("Disabled: Run at startup.")
            return True
        except Exception as e:
            self._set_status(f"Failed to disable startup: {e}", 5000)
            return False

    # ---------- Settings / Hotkey ----------
    def _load_settings(self):
        if os.path.exists(SETTINGS_PATH):
            try:
                with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                    s = json.load(f)
                hk = s.get("hotkey", {})
                self.hk_ctrl = bool(hk.get("ctrl", self.hk_ctrl))
                self.hk_alt = bool(hk.get("alt", self.hk_alt))
                self.hk_shift = bool(hk.get("shift", self.hk_shift))
                self.hk_win = bool(hk.get("win", self.hk_win))
                self.hk_key = str(hk.get("key", self.hk_key))

                self.monitoring_paused_var.set(bool(s.get("monitoring_paused", False)))
            except Exception:
                pass

        self.hk_display = self._format_hotkey_display(self.hk_ctrl, self.hk_alt, self.hk_shift, self.hk_win, self.hk_key)

    def _save_settings(self):
        try:
            data = {
                "hotkey": {
                    "ctrl": self.hk_ctrl,
                    "alt": self.hk_alt,
                    "shift": self.hk_shift,
                    "win": self.hk_win,
                    "key": self.hk_key,
                },
                "monitoring_paused": bool(self.monitoring_paused_var.get()),
            }
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    def open_hotkey_dialog(self):
        win = tk.Toplevel(self)
        win.title("Set Hotkey")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        ctrl_var = tk.BooleanVar(value=self.hk_ctrl)
        alt_var = tk.BooleanVar(value=self.hk_alt)
        shift_var = tk.BooleanVar(value=self.hk_shift)
        win_var = tk.BooleanVar(value=self.hk_win)

        mod_label = "Cmd" if sys.platform == "darwin" else "Win"

        row1 = ttk.Frame(frm)
        row1.pack(fill="x")
        ttk.Checkbutton(row1, text="Ctrl", variable=ctrl_var).pack(side="left", padx=(0, 8))
        ttk.Checkbutton(row1, text="Alt", variable=alt_var).pack(side="left", padx=(0, 8))
        ttk.Checkbutton(row1, text="Shift", variable=shift_var).pack(side="left", padx=(0, 8))
        ttk.Checkbutton(row1, text=mod_label, variable=win_var).pack(side="left")

        ttk.Label(frm, text="Key (A-Z, 0-9, F1..F24, etc):").pack(anchor="w", pady=(12, 2))
        key_var = tk.StringVar(value=self.hk_key)
        key_entry = ttk.Entry(frm, textvariable=key_var, width=12)
        key_entry.pack(anchor="w")
        key_entry.focus_set()

        note = ""
        if os.name != "nt" and not PYNPUT_AVAILABLE:
            note = "Note: Global hotkey on macOS/Linux requires 'pynput' (pip install pynput)."

        preview_var = tk.StringVar(value=self.hk_display)
        ttk.Label(frm, textvariable=preview_var).pack(anchor="w", pady=(10, 0))
        if note:
            ttk.Label(frm, text=note).pack(anchor="w", pady=(6, 0))

        def update_preview(*_):
            display = self._format_hotkey_display(
                ctrl_var.get(), alt_var.get(), shift_var.get(), win_var.get(), key_var.get().strip() or "?"
            )
            preview_var.set(display)

        ctrl_var.trace_add("write", update_preview)
        alt_var.trace_add("write", update_preview)
        shift_var.trace_add("write", update_preview)
        win_var.trace_add("write", update_preview)
        key_var.trace_add("write", update_preview)

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(14, 0))

        def apply():
            key = key_var.get().strip()
            if not key:
                messagebox.showerror("Hotkey", "Please enter a key.")
                return

            ctrl = bool(ctrl_var.get())
            alt = bool(alt_var.get())
            shift = bool(shift_var.get())
            winm = bool(win_var.get())

            display = self._format_hotkey_display(ctrl, alt, shift, winm, key)

            # Validate / register
            if os.name == "nt":
                mods, vk = self._hotkey_to_win32(ctrl, alt, shift, winm, key)
                if vk is None:
                    messagebox.showerror("Hotkey", f"Unsupported key: {key}")
                    return
            else:
                combo = self._hotkey_to_pynput(ctrl, alt, shift, winm, key)
                if combo is None:
                    messagebox.showerror("Hotkey", f"Unsupported key: {key}")
                    return

            self.hk_ctrl, self.hk_alt, self.hk_shift, self.hk_win, self.hk_key = ctrl, alt, shift, winm, key
            self.hk_display = display
            self._save_settings()

            try:
                if os.name == "nt" and isinstance(self.hotkey_listener, WindowsHotkeyListener):
                    mods, vk = self._hotkey_to_win32(ctrl, alt, shift, winm, key)
                    if vk is not None:
                        self.hotkey_listener.set_hotkey(mods, vk, display)
                elif isinstance(self.hotkey_listener, PynputHotkeyListener):
                    combo = self._hotkey_to_pynput(ctrl, alt, shift, winm, key)
                    if combo is not None:
                        self.hotkey_listener.set_hotkey(combo, display)
                else:
                    self._setup_hotkey()
            except Exception:
                self._setup_hotkey()

            # self._set_status(f"Hotkey set: {display}")
            win.destroy()

        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="right")
        ttk.Button(btns, text="Apply", command=apply).pack(side="right", padx=(0, 8))
    def _format_hotkey_display(self, ctrl, alt, shift, win, key):
        parts = []
        if ctrl:
            parts.append("Ctrl")
        if alt:
            parts.append("Alt")
        if shift:
            parts.append("Shift")
        if win:
            parts.append("Cmd" if sys.platform == "darwin" else "Win")
        k = (key or "").strip()
        k = k.upper() if len(k) == 1 else k
        parts.append(k if k else "?")
        return "+".join(parts)

    def _hotkey_to_win32(self, ctrl, alt, shift, win, key):
        mods = 0
        if ctrl: mods |= WindowsHotkeyListener.MOD_CONTROL
        if alt: mods |= WindowsHotkeyListener.MOD_ALT
        if shift: mods |= WindowsHotkeyListener.MOD_SHIFT
        if win: mods |= WindowsHotkeyListener.MOD_WIN
        mods |= WindowsHotkeyListener.MOD_NOREPEAT
        return mods, self._key_to_vk(key)

    def _hotkey_to_pynput(self, ctrl, alt, shift, win, key):
        k = self._key_to_pynput(key)
        if not k:
            return None
        parts = []
        if ctrl:
            parts.append("<ctrl>")
        if alt:
            parts.append("<alt>")
        if shift:
            parts.append("<shift>")
        if win:
            # Command on macOS, Super/Win elsewhere
            parts.append("<cmd>")
        parts.append(k)
        return "+".join(parts)

    def _key_to_pynput(self, key):
        k = (key or "").strip()
        if not k:
            return None

        # Single character
        if len(k) == 1:
            return k.lower()

        ku = k.upper()

        # Function keys: F1..F24
        m = re.match(r"^F(\d{1,2})$", ku)
        if m:
            n = int(m.group(1))
            if 1 <= n <= 24:
                return f"<f{n}>"

        special = {
            "ENTER": "<enter>",
            "RETURN": "<enter>",
            "TAB": "<tab>",
            "ESC": "<esc>",
            "ESCAPE": "<esc>",
            "SPACE": "<space>",
            "BACKSPACE": "<backspace>",
            "DELETE": "<delete>",
            "UP": "<up>",
            "DOWN": "<down>",
            "LEFT": "<left>",
            "RIGHT": "<right>",
            "HOME": "<home>",
            "END": "<end>",
            "PAGEUP": "<page_up>",
            "PAGEDOWN": "<page_down>",
        }
        return special.get(ku)

    def _key_to_vk(self, key: str):
        k = key.strip()

        if len(k) == 1:
            ch = k.upper()
            if "A" <= ch <= "Z" or "0" <= ch <= "9":
                return ord(ch)

        if k.upper().startswith("F"):
            try:
                n = int(k[1:])
                if 1 <= n <= 24:
                    return 0x70 + (n - 1)
            except Exception:
                pass

        mapping = {
            "space": 0x20,
            "Return": 0x0D,
            "Tab": 0x09,
            "Escape": 0x1B,
        }
        return mapping.get(k, None)
    def _setup_hotkey(self):
        # Tear down any existing listener
        try:
            if self.hotkey_listener:
                self.hotkey_listener.stop()
        except Exception:
            pass
        self.hotkey_listener = None

        if os.name == "nt":
            self.hotkey_listener = WindowsHotkeyListener(self.q)
            self.hotkey_listener.start()
            mods, vk = self._hotkey_to_win32(self.hk_ctrl, self.hk_alt, self.hk_shift, self.hk_win, self.hk_key)
            if vk is not None:
                self.hotkey_listener.set_hotkey(mods, vk, self.hk_display)
            return

        combo = self._hotkey_to_pynput(self.hk_ctrl, self.hk_alt, self.hk_shift, self.hk_win, self.hk_key)
        if combo is None:
            if not PYNPUT_AVAILABLE:
                self.q.put(("INFO", "Hotkey on macOS/Linux requires 'pynput' (pip install pynput)."))
            return

        self.hotkey_listener = PynputHotkeyListener(self.q)
        self.hotkey_listener.start(combo, self.hk_display)


    # ---------- Event processing (hotkey + second-run SHOW + tray) ----------
    def _process_events(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind in ("HOTKEY", "SHOW", "TRAY_TOGGLE"):
                    self.toggle_window()
                elif kind == "TRAY_EXIT":
                    self.exit_app()
                elif kind == "TRAY_TOGGLE_MONITOR":
                    self.toggle_monitoring()
                elif kind == "INFO":
                    self._set_status(str(payload))
                elif kind == "ERROR":
                    self._set_status(str(payload), clear_ms=4500)
        except queue.Empty:
            pass
        self.after(50, self._process_events)


    # ---------- Show / Hide ----------
    def toggle_window(self):
        if self.winfo_viewable():
            self._hide_window()
        else:
            self._show_window()

    def _show_window(self):
        try:
            self.deiconify()
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            self.geometry(f"+{max(0, x - 60)}+{max(0, y - 60)}")
            self.lift()
            self.attributes("-topmost", True)
            self.after(120, lambda: self.attributes("-topmost", False))
            self.focus_force()
        except Exception:
            pass

    def _hide_window(self):
        try:
            self.withdraw()
        except Exception:
            pass

    # ---------- Close behavior ----------
    def on_close_hide(self):
        self._hide_window()
        if os.name == "nt" and PYSTRAY_AVAILABLE:
            self._set_status("Hidden to tray. Use tray icon, hotkey, or run again to show.", clear_ms=3500)
        else:
            self._set_status("Hidden to background. Use hotkey or run again to show.", clear_ms=3500)
    def exit_app(self):
        try:
            if self.hotkey_listener:
                self.hotkey_listener.stop()
        except Exception:
            pass
        try:
            if self.tray:
                self.tray.stop()
        except Exception:
            pass
        try:
            if self.ipc_server:
                self.ipc_server.stop()
        except Exception:
            pass
        self.destroy()



def main():
    event_q = queue.Queue()

    # If launched by Windows Startup (or with explicit flag), start hidden.
    start_hidden = _has_startup_flag()

    # Start single-instance server. If we can't bind, ask existing instance to show and exit.
    ipc = SingleInstanceServer(SINGLE_INSTANCE_HOST, SINGLE_INSTANCE_PORT, event_q)
    if not ipc.start():
        SingleInstanceServer.send_show(SINGLE_INSTANCE_HOST, SINGLE_INSTANCE_PORT)
        return

    app = ClipboardManager(event_q, ipc, start_hidden=start_hidden)
    app.mainloop()


if __name__ == "__main__":
    main()