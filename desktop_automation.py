import cv2
import numpy as np
import pyautogui
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk
import mss
import json
import os
from pynput.keyboard import Listener as KeyListener, Controller as KeyController, Key
import random
import hashlib
from datetime import datetime
import uuid
from pynput.mouse import Listener as MouseListener
import pyperclip
from cryptography.fernet import Fernet


SECRET_HMAC = b""  # TODO: set via environment variable or external config
FERNET_KEY = b""   # TODO: set via environment variable or external config
fernet = Fernet(FERNET_KEY)


def get_machine_id() -> str:
    """Return a short hash of the current machine's hardware address."""
    return hashlib.sha256(str(uuid.getnode()).encode()).hexdigest()[:16]


def encrypt_license(machine_id: str, expire_date: str) -> str:
    """Create an encrypted license token for a given machine and expiry.

    :param machine_id: a 16‑character identifier from :func:`get_machine_id`
    :param expire_date: a date string in ``YYYY-MM-DD`` format
    :returns: an encrypted token including a short signature
    """
    raw = f"{machine_id}|{expire_date}"
    sig = hashlib.sha256((raw + "|" + SECRET_HMAC.decode()).encode()).hexdigest()[:16]
    token = f"{machine_id}|{expire_date}|{sig}"
    return fernet.encrypt(token.encode()).decode()


def decrypt_and_verify_license(encrypted_key: str) -> tuple[bool, str]:
    """Attempt to decrypt and verify an encrypted license token.

    :param encrypted_key: the encrypted token returned by
        :func:`encrypt_license`
    :returns: a tuple ``(valid, message)`` where ``valid`` is ``True`` if
        the token is valid for this machine and not expired.  ``message``
        will either hold the expiry date or an error description.
    """
    try:
        decrypted = fernet.decrypt(encrypted_key.encode()).decode()
        machine_id, expire_date, signature = decrypted.split("|")
        if machine_id != get_machine_id():
            return False, "Wrong machine"

        raw = f"{machine_id}|{expire_date}"
        expected_sig = hashlib.sha256((raw + "|" + SECRET_HMAC.decode()).encode()).hexdigest()[:16]
        if signature != expected_sig:
            return False, "Invalid signature"

        if datetime.now().date() > datetime.strptime(expire_date, "%Y-%m-%d").date():
            return False, "Expired"

        return True, expire_date
    except Exception as exc:  # noqa: BLE001
        return False, f"Decryption error: {exc}"


CONFIG_FILE = "autoclicker_config.json"


class MacroAutomationTool:
    """A GUI based automation tool for repetitive on‑screen interactions.

    This class encapsulates all of the state and behaviour required to
    configure and run automated clicking and key pressing sequences based
    on simple visual cues.  It is designed to be extensible and avoids
    using any terminology tied to a particular game or application.
    """

    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        self.master.title("Macro Automation Tool")
        self.running = False
        self.active = False
        self.keyboard = KeyController()
        self.debug_img: np.ndarray | None = None
        self.hotkey_char = tk.StringVar()
        self.auto_move_var = tk.BooleanVar(value=True)
        self.edge_touched = False
        self.edge_armed = True
        self.last_click_coords = "—"
        self.click_history: list[str] = []
        self.require_edge_touch_var = tk.BooleanVar(value=True)
        self.assist_click_var = tk.BooleanVar(value=False)
        self.assist_key_var = tk.StringVar(value='f')
        self.assist_key_enabled_var = tk.BooleanVar(value=False)
        self.debug_status_var = tk.BooleanVar(value=False)
        self.timed_press_enabled_var = tk.BooleanVar(value=False)
        self.timed_press_direction_var = tk.StringVar(value='right')
        self.keypress_lock = threading.RLock()
        self.force_keypress_interrupt = threading.Event()
        self.global_pause_event = threading.Event()
        self.global_pause_event.set()
        self.partial_pause_event = threading.Event()
        self.partial_pause_event.set()

        # Default values for region detection and timing behaviour
        self.defaults = {
            'x': 519, 'y': 838, 'w': 886, 'h': 70,
            'bar_width_min': 3, 'bar_width_max': 15,
            'bar_height_ratio': 0.7,
            'threshold': 50,
            'brightness_offset': 10,
            'click_tolerance': 30,
            'cooldown': 0.05,
            'hotkey': 'p',
            'debug_view': True
        }

        # Auto sell settings (timers and coordinates)
        self.auto_sell_enabled = True
        self.auto_sell_interval = 600  # seconds
        self.auto_sell_point: tuple[int, int] = (500, 500)

        # Build the user interface and register listeners
        self.setup_ui()
        self.listen_for_termination_hotkey()
        self.load_config()

    # ------------------------------------------------------------------
    #  Configuration handling
    #
    #  Persisting and restoring user preferences from JSON allows you to
    #  restart the tool without re‑entering all values.  Only values
    #  defined in ``self.defaults`` and the advanced options are
    #  persisted.

    def load_config(self) -> None:
        """Load saved configuration values from ``CONFIG_FILE`` if present."""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # Update defaults in place and populate UI entries
                    self.defaults.update(data)
                    # Primary region
                    self.x_entry.delete(0, tk.END)
                    self.x_entry.insert(0, str(self.defaults["x"]))
                    self.y_entry.delete(0, tk.END)
                    self.y_entry.insert(0, str(self.defaults["y"]))
                    self.w_entry.delete(0, tk.END)
                    self.w_entry.insert(0, str(self.defaults["w"]))
                    self.h_entry.delete(0, tk.END)
                    self.h_entry.insert(0, str(self.defaults["h"]))

                    # Boolean flags
                    self.require_edge_touch_var.set(data.get('require_edge_touch', True))
                    self.assist_click_var.set(data.get('assist_click_enabled', False))
                    self.assist_key_enabled_var.set(data.get('assist_key_enabled', False))
                    self.assist_key_var.set(data.get('assist_key', 'f'))
                    self.timed_press_enabled_var.set(data.get('timed_press_enabled', False))
                    self.timed_press_direction_var.set(data.get('timed_press_direction', 'right'))

                    for key, entry in self.adv_entries.items():
                        if key in data:
                            entry.delete(0, tk.END)
                            entry.insert(0, str(data[key]))
            except Exception as exc:  # noqa: BLE001
                print(f"[load_config] error: {exc}")

    def save_config(self) -> None:
        """Persist the current configuration to ``CONFIG_FILE``."""
        config: dict[str, object] = {
            'x': int(self.x_entry.get()),
            'y': int(self.y_entry.get()),
            'w': int(self.w_entry.get()),
            'h': int(self.h_entry.get()),
            'hotkey': self.hotkey_char.get(),
            'debug_view': self.debug_var.get(),
            'require_edge_touch': self.require_edge_touch_var.get(),
            'assist_click_enabled': self.assist_click_var.get(),
            'assist_key_enabled': self.assist_key_enabled_var.get(),
            'assist_key': self.assist_key_var.get(),
            'timed_press_enabled': self.timed_press_enabled_var.get(),
            'timed_press_direction': self.timed_press_direction_var.get()
        }
        # Advanced entries stored as strings; convert types when reading
        for key in self.adv_entries:
            config[key] = self.adv_entries[key].get()
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
        messagebox.showinfo("Saved", "Configuration saved successfully!")

    # ------------------------------------------------------------------
    #  User interface construction
    #
    #  A fairly simple Tk based UI is defined here.  It exposes
    #  configuration inputs for the region of interest, debug options
    #  and advanced behaviour.  Additional elements can be added as
    #  needed.

    def setup_ui(self) -> None:
        """Construct the main window and advanced settings panes."""
        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(padx=10, pady=10)

        # Primary region configuration
        self.x_entry = self._create_labeled_entry("X:", self.defaults['x'], 0)
        self.y_entry = self._create_labeled_entry("Y:", self.defaults['y'], 1)
        self.w_entry = self._create_labeled_entry("Width:", self.defaults['w'], 2)
        self.h_entry = self._create_labeled_entry("Height:", self.defaults['h'], 3)

        # Debug toggles
        self.debug_var = tk.BooleanVar(value=self.defaults['debug_view'])
        tk.Checkbutton(self.main_frame, text="Show Debug View", variable=self.debug_var).grid(row=4, column=0, columnspan=2, pady=4)
        tk.Checkbutton(self.main_frame, text="Show Status Window", variable=self.debug_status_var, command=self.toggle_status_window).grid(row=5, column=0, columnspan=2, pady=4)

        # Region selection and start/stop buttons
        ttk.Button(self.main_frame, text="Click to change region", command=self.define_region_by_click).grid(row=6, column=0, columnspan=2, pady=4)
        ttk.Button(self.main_frame, text="Start", command=self.start).grid(row=7, column=0, pady=4)
        ttk.Button(self.main_frame, text="Stop", command=self.stop).grid(row=7, column=1, pady=4)

        # Movement toggle
        ttk.Checkbutton(self.main_frame, text="Enable Auto Move + Click", variable=self.auto_move_var, command=self.update_active_state).grid(row=8, column=0, columnspan=2, pady=4)

        # Collapsible frames for additional settings
        self.farm_toggle_btn = ttk.Button(self.master, text="▼ Automation Settings", command=self.toggle_secondary_settings)
        self.farm_toggle_btn.pack()
        self.secondary_frame = ttk.LabelFrame(self.master, text="Automation Options")

        # Auto sell controls
        self.auto_sell_var = tk.BooleanVar(value=True)
        self.auto_sell_interval_var = tk.StringVar(value="600")  # default 10 minutes
        self.auto_sell_point_var = tk.StringVar(value="(500, 500)")

        ttk.Checkbutton(
            self.secondary_frame,
            text="Enable Auto Sell",
            variable=self.auto_sell_var
        ).pack(anchor="w", pady=(10, 2))
        tk.Label(self.secondary_frame, text="Interval (s):").pack(anchor="w")
        tk.Entry(self.secondary_frame, textvariable=self.auto_sell_interval_var).pack(fill="x", padx=2)

        ttk.Button(self.secondary_frame, text="Set Auto Sell Point", command=self.define_auto_sell_point).pack(anchor="w", pady=2)
        tk.Label(self.secondary_frame, textvariable=self.auto_sell_point_var, fg="blue").pack(anchor="w", padx=10)

        # Advanced settings
        self.toggle_btn = ttk.Button(self.master, text="▼ Advanced Settings", command=self.toggle_advanced_settings)
        self.toggle_btn.pack()
        self.adv_frame = ttk.LabelFrame(self.master, text="Advanced Settings")
        self.adv_entries: dict[str, tk.Entry] = {}
        self._populate_advanced_settings()

    def _create_labeled_entry(self, label: str, default: int | float, row: int) -> tk.Entry:
        """Helper to build a labeled entry and insert a default value."""
        tk.Label(self.main_frame, text=label).grid(row=row, column=0)
        entry = tk.Entry(self.main_frame)
        entry.insert(0, str(default))
        entry.grid(row=row, column=1)
        return entry

    def _populate_advanced_settings(self) -> None:
        """Populate the advanced settings frame with configurable fields."""
        row = 0
        for key in ['bar_width_min', 'bar_width_max', 'bar_height_ratio', 'threshold', 'brightness_offset', 'click_tolerance', 'cooldown']:
            tk.Label(self.adv_frame, text=key.replace('_', ' ').title()+":").grid(row=row, column=0)
            entry = tk.Entry(self.adv_frame)
            entry.insert(0, str(self.defaults[key]))
            entry.grid(row=row, column=1)
            self.adv_entries[key] = entry
            row += 1

        # Interval for switching movement direction
        tk.Label(self.adv_frame, text="Move Direction Interval (s):").grid(row=row, column=0)
        move_entry = tk.Entry(self.adv_frame)
        move_entry.insert(0, "0.8")
        move_entry.grid(row=row, column=1)
        self.adv_entries['move_interval'] = move_entry
        row += 1

        # Edge touch requirement
        ttk.Checkbutton(
            self.adv_frame,
            text="Require edge touch for click", 
            variable=self.require_edge_touch_var
        ).grid(row=row, column=0, columnspan=2, pady=2)
        row += 1

        # Emergency hotkey
        tk.Label(self.adv_frame, text="Emergency Hotkey:").grid(row=row, column=0)
        self.hotkey_char.set(self.defaults['hotkey'])
        tk.Entry(self.adv_frame, textvariable=self.hotkey_char).grid(row=row, column=1)
        row += 1

        ttk.Separator(self.adv_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=5)
        row += 1

        # Save settings button
        ttk.Button(self.adv_frame, text="Save Settings", command=self.save_config).grid(
            row=row, column=0, columnspan=2, pady=5
        )
        row += 1

        # ------------------------------------------------------------------
        # Secondary frame (assist options)
        #
        #  These options control additional clicking and key pressing behaviour
        #  once the main clicker has been triggered.  They are grouped
        #  together under the secondary frame rather than the advanced frame
        #  to avoid cluttering the primary advanced settings.

        ttk.Checkbutton(
            self.secondary_frame,
            text="Enable auto click after trigger",
            variable=self.assist_click_var
        ).pack(anchor="w", pady=2)

        tk.Label(self.secondary_frame, text="Click Delay (s):").pack(anchor="w")
        assist_click_delay = tk.Entry(self.secondary_frame)
        assist_click_delay.insert(0, "0.5")
        assist_click_delay.pack(fill="x")
        self.adv_entries["assist_click_delay"] = assist_click_delay

        ttk.Checkbutton(
            self.secondary_frame,
            text="Enable auto key press",
            variable=self.assist_key_enabled_var
        ).pack(anchor="w", pady=4)

        tk.Label(self.secondary_frame, text="Key (char):").pack(anchor="w")
        tk.Entry(self.secondary_frame, textvariable=self.assist_key_var).pack(fill="x")

        tk.Label(self.secondary_frame, text="Key Delay (s):").pack(anchor="w")
        assist_key_delay = tk.Entry(self.secondary_frame)
        assist_key_delay.insert(0, "0.5")
        assist_key_delay.pack(fill="x")
        self.adv_entries["assist_key_delay"] = assist_key_delay

        tk.Label(self.secondary_frame, text="Key Hold (s):").pack(anchor="w")
        assist_key_hold = tk.Entry(self.secondary_frame)
        assist_key_hold.insert(0, "0.3")
        assist_key_hold.pack(fill="x")
        self.adv_entries["assist_key_hold"] = assist_key_hold

        ttk.Checkbutton(
            self.secondary_frame,
            text="Enable periodic key press",
            variable=self.timed_press_enabled_var
        ).pack(anchor="w", pady=4)

        tk.Label(self.secondary_frame, text="Direction (left/right):").pack(anchor="w")
        tk.Entry(self.secondary_frame, textvariable=self.timed_press_direction_var).pack(fill="x")

        tk.Label(self.secondary_frame, text="Press Every (s):").pack(anchor="w")
        entry = tk.Entry(self.secondary_frame)
        entry.insert(0, "60.0")
        entry.pack(fill="x")
        self.adv_entries["timed_press_interval"] = entry

        tk.Label(self.secondary_frame, text="Press Duration (s):").pack(anchor="w")
        entry = tk.Entry(self.secondary_frame)
        entry.insert(0, "1.0")
        entry.pack(fill="x")
        self.adv_entries["timed_press_duration"] = entry

    # ------------------------------------------------------------------
    #  UI toggles

    def toggle_advanced_settings(self) -> None:
        """Show or hide the advanced settings frame."""
        if self.adv_frame.winfo_ismapped():
            self.adv_frame.pack_forget()
            self.toggle_btn.config(text="▼ Advanced Settings")
            self.master.after(100, lambda: self.master.geometry(""))
        else:
            self.adv_frame.pack()
            self.toggle_btn.config(text="▲ Hide Advanced Settings")
            self.master.update_idletasks()
            width = self.master.winfo_reqwidth()
            height = self.master.winfo_reqheight()
            self.master.geometry(f"{width}x{height}")

    def toggle_secondary_settings(self) -> None:
        """Show or hide the secondary (automation) settings."""
        if self.secondary_frame.winfo_ismapped():
            self.secondary_frame.pack_forget()
            self.farm_toggle_btn.config(text="▼ Automation Settings")
            self.master.update_idletasks()
            self.master.geometry("")
        else:
            self.secondary_frame.pack()
            self.farm_toggle_btn.config(text="▲ Hide Automation Settings")
            self.master.update_idletasks()
            width = self.master.winfo_reqwidth()
            height = self.master.winfo_reqheight()
            self.master.geometry(f"{width}x{height}")

    def update_active_state(self) -> None:
        """Update whether the movement/click loop is active."""
        self.active = self.auto_move_var.get()

    # ------------------------------------------------------------------
    #  Region definition helpers

    def define_region_by_click(self) -> None:
        """Prompt the user to define the region of interest by clicking twice."""
        self.master.withdraw()
        clicks: list[tuple[int, int]] = []

        def on_click(x: int, y: int, button, pressed: bool) -> bool | None:
            if pressed:
                clicks.append((x, y))
                if len(clicks) == 2:
                    x1, y1 = clicks[0]
                    x2, y2 = clicks[1]
                    self.x_entry.delete(0, tk.END)
                    self.x_entry.insert(0, str(min(x1, x2)))
                    self.y_entry.delete(0, tk.END)
                    self.y_entry.insert(0, str(min(y1, y2)))
                    self.w_entry.delete(0, tk.END)
                    self.w_entry.insert(0, str(abs(x2 - x1)))
                    self.h_entry.delete(0, tk.END)
                    self.h_entry.insert(0, str(abs(y2 - y1)))
                    self.master.deiconify()
                    return False
            return None

        threading.Thread(target=lambda: MouseListener(on_click=on_click).start(), daemon=True).start()

    def define_auto_sell_point(self) -> None:
        """Prompt the user to pick an on‑screen point used during auto sell."""
        self.master.withdraw()

        def on_click(x: int, y: int, button, pressed: bool) -> bool | None:
            if pressed:
                self.auto_sell_point_var.set(f"({x}, {y})")
                self.master.deiconify()
                return False
            return None

        threading.Thread(target=lambda: MouseListener(on_click=on_click).start(), daemon=True).start()

    # ------------------------------------------------------------------
    #  Hotkey listening

    def listen_for_termination_hotkey(self) -> None:
        """Listen for a single hotkey that will stop the automation."""
        def on_press(key: Key) -> None:
            try:
                if hasattr(key, 'char') and key.char == self.hotkey_char.get():
                    self.stop()
            except Exception:
                pass

        listener = KeyListener(on_press=on_press)
        listener.daemon = True
        listener.start()

    # ------------------------------------------------------------------
    #  Configuration getters

    def get_region(self) -> tuple[int, int, int, int] | None:
        """Return the currently configured region as a 4‑tuple."""
        try:
            return tuple(map(int, [
                self.x_entry.get(),
                self.y_entry.get(),
                self.w_entry.get(),
                self.h_entry.get()
            ]))
        except Exception:
            return None

    def get_configuration(self) -> dict[str, object]:
        """Gather numeric and boolean settings into a dictionary."""
        cfg: dict[str, object] = {
            k: float(self.adv_entries[k].get()) if '.' in self.adv_entries[k].get() else int(self.adv_entries[k].get())
            for k in self.adv_entries
        }
        cfg['debug_view'] = self.debug_var.get()
        cfg['require_edge_touch'] = self.require_edge_touch_var.get()
        cfg['assist_click_enabled'] = self.assist_click_var.get()
        cfg['assist_key_enabled'] = self.assist_key_enabled_var.get()
        cfg['assist_key'] = self.assist_key_var.get()
        cfg['timed_press_enabled'] = self.timed_press_enabled_var.get()
        cfg['timed_press_direction'] = self.timed_press_direction_var.get()
        cfg['auto_sell_enabled'] = self.auto_sell_var.get()
        cfg['auto_sell_interval'] = int(self.auto_sell_interval_var.get())
        cfg['auto_sell_point'] = eval(self.auto_sell_point_var.get())
        return cfg

    # ------------------------------------------------------------------
    #  Core automation control

    def start(self) -> None:
        """Start the automation threads if not already running."""
        if self.running:
            return

        region = self.get_region()
        if not region:
            return

        self.config = self.get_configuration()
        self.running = True
        self.active = self.auto_move_var.get()
        self.threads: list[threading.Thread] = []

        def launch(target, *args) -> None:
            t = threading.Thread(target=target, args=args, daemon=True)
            t.start()
            self.threads.append(t)

        launch(self.detection_and_click_loop, region)
        if self.config['debug_view']:
            launch(self.display_debug_window)
        launch(self.movement_click_loop, region)
        launch(self.handle_assist_click, region)
        launch(self.handle_assist_key_press, region)
        launch(self.handle_periodic_key_press)

        if self.config.get("auto_sell_enabled", False):
            launch(self.perform_auto_sell)

    def stop(self) -> None:
        """Stop all running automation threads and reset state."""
        self.running = False
        self.active = False
        self.force_keypress_interrupt.set()
        if hasattr(self, 'threads'):
            for t in self.threads:
                if t.is_alive():
                    t.join(timeout=1)
            self.threads.clear()
        self.force_keypress_interrupt.clear()
        # Release any keys that might still be held down
        try:
            self.keyboard.release('w')
            self.keyboard.release(Key.left)
            self.keyboard.release(Key.right)
            assist_key = self.config.get("assist_key", "f")
            self.keyboard.release(assist_key)
        except Exception as exc:  # noqa: BLE001
            print(f"[stop] Key release error: {exc}")

    # ------------------------------------------------------------------
    #  Status window

    def toggle_status_window(self) -> None:
        """Show or hide the status window with runtime indicators."""
        if self.debug_status_var.get():
            self.open_status_window()
        else:
            if hasattr(self, 'status_window') and self.status_window.winfo_exists():
                self.status_window.destroy()

    def open_status_window(self) -> None:
        """Open a separate window showing internal state like last click."""
        self.status_window = tk.Toplevel(self.master)
        self.status_window.title("Macro Status")
        self.status_labels: dict[str, tk.Label] = {}

        for i, name in enumerate(["bar_box", "target_x", "edge_touched", "can_click"]):
            tk.Label(self.status_window, text=name).grid(row=i, column=0, sticky="w")
            lbl = tk.Label(self.status_window, text="False", fg="red", font=("Segoe UI", 10, "bold"))
            lbl.grid(row=i, column=1, padx=10)
            self.status_labels[name] = lbl

        tk.Label(self.status_window, text="Click History").grid(row=4, column=0, columnspan=2, pady=(10, 0))
        self.history_listbox = tk.Listbox(self.status_window, height=7, width=25)
        self.history_listbox.grid(row=5, column=0, columnspan=2, pady=(0, 10))

        if not hasattr(self, 'click_history'):
            self.click_history = []
        tk.Label(self.status_window, text="Require Edge Touch:").grid(row=6, column=0, sticky="w")
        self.status_labels["require_edge"] = tk.Label(
            self.status_window,
            text=str(self.require_edge_touch_var.get()),
            fg="blue",
            font=("Segoe UI", 10)
        )
        self.status_labels["require_edge"].grid(row=6, column=1, padx=10)

    # ------------------------------------------------------------------
    #  Assist click/key handlers

    def handle_assist_click(self, region: tuple[int, int, int, int]) -> None:
        """Optionally click in the centre of the region after a delay when idle."""
        x, y, w, h = region
        cfg = self.config
        delay = float(cfg.get("assist_click_delay", 0.5))
        cooldown = cfg.get("cooldown", 0.05)

        while self.running:
            self.global_pause_event.wait()
            self.partial_pause_event.wait()
            if cfg.get("assist_click_enabled", False) and not self.is_event_active(region):
                start = time.time()
                while time.time() - start < delay:
                    if self.is_event_active(region) or not self.running or self.force_keypress_interrupt.is_set():
                        break
                    time.sleep(0.01)
                else:
                    while self.running and not self.is_event_active(region):
                        if self.force_keypress_interrupt.is_set():
                            break
                        with self.keypress_lock:
                            pyautogui.click(x + w // 2, y + h // 2)
                        time.sleep(cooldown)
            time.sleep(0.05)

    def handle_assist_key_press(self, region: tuple[int, int, int, int]) -> None:
        """Optionally press a user configured key when idle for a period."""
        while self.running:
            self.global_pause_event.wait()
            self.partial_pause_event.wait()
            cfg = self.config
            if cfg.get("assist_key_enabled", False) and not self.is_event_active(region):
                key_delay = float(cfg.get("assist_key_delay", 0.5))
                key_hold = float(cfg.get("assist_key_hold", 0.3))
                key_char = cfg.get("assist_key", "f")
                # Wait initial delay
                start = time.time()
                while time.time() - start < key_delay:
                    if self.is_event_active(region) or not self.running or self.force_keypress_interrupt.is_set():
                        break
                    time.sleep(0.01)
                else:
                    if self.force_keypress_interrupt.is_set():
                        continue
                    with self.keypress_lock:
                        if self.force_keypress_interrupt.is_set():
                            continue
                        try:
                            self.keyboard.press(key_char)
                            time.sleep(key_hold)
                            self.keyboard.release(key_char)
                        except Exception as exc:
                            print(f"[assist_key] error: {exc}")
            time.sleep(0.05)

    def handle_periodic_key_press(self) -> None:
        """Press the left or right key at a fixed interval if enabled."""
        if not self.config.get("timed_press_enabled", False):
            return
        while self.running:
            self.partial_pause_event.wait()
            self.global_pause_event.wait()
            cfg = self.config
            interval = float(cfg.get("timed_press_interval", 60.0))
            duration = float(cfg.get("timed_press_duration", 1.0))
            direction = cfg.get("timed_press_direction", "right").strip().lower()
            key = Key.right if direction == "right" else Key.left
            time.sleep(max(0, interval - 0.5))
            self.force_keypress_interrupt.set()
            time.sleep(0.5)
            with self.keypress_lock:
                try:
                    self.keyboard.press(key)
                    time.sleep(duration)
                except Exception as exc:
                    print(f"[timed_key] press error: {exc}")
                finally:
                    try:
                        self.keyboard.release(key)
                    except Exception as exc:
                        print(f"[timed_key] release error: {exc}")
            self.force_keypress_interrupt.clear()

    # ------------------------------------------------------------------
    #  Main movement and click loop

    def movement_click_loop(self, region: tuple[int, int, int, int]) -> None:
        """Control movement and clicking based on visual cues."""
        x, y, w, h = region
        cooldown = self.config['cooldown']
        move_interval = float(self.config.get('move_interval', 0.8))
        direction: Key = Key.right
        phase = "start"
        while self.running:
            in_event = self.is_event_active(region)
            if self.active:
                if phase == "start" and not in_event:
                    self.keyboard.press('w')
                    self.keyboard.press(direction)
                    while self.running and not self.is_event_active(region):
                        pyautogui.click(x + w // 2, y + h // 2)
                        time.sleep(cooldown)
                    self.keyboard.release('w')
                    self.keyboard.release(direction)
                    phase = "post-event"
                elif phase == "post-event":
                    while self.running and self.is_event_active(region):
                        time.sleep(0.1)
                    direction = Key.left if direction == Key.right else Key.right
                    self.keyboard.press('w')
                    self.keyboard.press(direction)
                    start_time = time.time()
                    while self.running and time.time() - start_time < move_interval:
                        if self.is_event_active(region):
                            break
                        pyautogui.click(x + w // 2, y + h // 2)
                        time.sleep(cooldown)
                    self.keyboard.release('w')
                    self.keyboard.release(direction)
                    phase = "start"
            time.sleep(cooldown)

    # ------------------------------------------------------------------
    #  Event detection and clicking loop

    def is_event_active(self, region: tuple[int, int, int, int]) -> bool:
        """Return True if a bar and bright region are detected in the ROI."""
        x, y, w, h = region
        cfg = self.config
        bar_min_h = int(h * cfg['bar_height_ratio'])
        monitor = {"top": y, "left": x, "width": w, "height": h}
        with mss.mss() as sct:
            img = np.array(sct.grab(monitor))
        gray = cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)
        _, thresh = cv2.threshold(gray, cfg['threshold'], 255, cv2.THRESH_BINARY_INV)
        cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        bar_box = any(cfg['bar_width_min'] <= cv2.boundingRect(c)[2] <= cfg['bar_width_max'] and cv2.boundingRect(c)[3] >= bar_min_h for c in cnts)
        mid = gray[h//2-5:h//2+5, :]
        cols = np.mean(cv2.GaussianBlur(mid, (5,5), 0), axis=0)
        bri = np.mean(cols)
        bright_idxs = np.where(cols > bri + cfg['brightness_offset'])[0]
        return bar_box and len(bright_idxs) > 10

    def detection_and_click_loop(self, region: tuple[int, int, int, int]) -> None:
        """Continuously detect the progress bar and target region and click accordingly."""
        x, y, w, h = region
        cfg = self.config
        bar_height_min = int(h * cfg['bar_height_ratio'])
        margin = int(w * 0.1)
        gray_prev: np.ndarray | None = None
        with mss.mss() as sct:
            while self.running:
                self.global_pause_event.wait()
                img = np.array(sct.grab({"top": y, "left": x, "width": w, "height": h}))
                gray = cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)
                diff: np.ndarray | None = None
                if gray_prev is not None:
                    diff = cv2.absdiff(gray, gray_prev)
                _, thresh = cv2.threshold(gray, cfg['threshold'], 255, cv2.THRESH_BINARY_INV)
                contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                bar_box: tuple[int, int] | None = None
                for cnt in contours:
                    x0, y0, w0, h0 = cv2.boundingRect(cnt)
                    if cfg['bar_width_min'] <= w0 <= cfg['bar_width_max'] and h0 >= bar_height_min:
                        if diff is not None:
                            roi_diff = diff[y0:y0+h0, x0:x0+w0]
                            motion_level = np.mean(roi_diff)
                            motion_threshold = 10
                            if motion_level < motion_threshold:
                                continue
                        bar_box = (x + x0, x + x0 + w0)
                        break
                gray_prev = gray.copy()
                target_x: float | None = None
                center_band = gray[h//2-5:h//2+5, :]
                blurred = cv2.GaussianBlur(center_band, (5,5), 0)
                mean_col = np.mean(blurred, axis=0)
                mean_val = np.mean(mean_col)
                bright_indices = np.where(mean_col > mean_val + cfg['brightness_offset'])[0]
                if len(bright_indices) > 10:
                    left = bright_indices[0]
                    right = bright_indices[-1]
                    center = (left + right) / 2
                    offset = random.uniform(-1, 1)
                    target_x = x + center + offset
                can_click = False
                if bar_box and target_x is not None:
                    bar_left, bar_right = bar_box
                    left_marker = x + margin
                    right_marker = x + w - margin
                    tolerance = 30
                    if (abs(bar_left - left_marker) <= tolerance or abs(bar_right - right_marker) <= tolerance) and self.edge_armed:
                        self.edge_touched = True
                        self.edge_armed = False
                    edge_condition = (self.edge_touched if cfg.get('require_edge_touch', True) else True)
                    if edge_condition and (bar_left - cfg['click_tolerance']) <= target_x <= (bar_right + cfg['click_tolerance']):
                        pyautogui.click(target_x, y + h / 2)
                        self.edge_touched = False
                        self.edge_armed = True
                        can_click = True
                        click_coord = f"({int(target_x)}, {int(y + h / 2)})"
                        self.last_click_coords = click_coord
                        if not hasattr(self, 'click_history'):
                            self.click_history = []
                        self.click_history.insert(0, click_coord)
                        if len(self.click_history) > 15:
                            self.click_history.pop()
                        if self.debug_status_var.get() and hasattr(self, 'history_listbox'):
                            self.history_listbox.delete(0, tk.END)
                            for item in self.click_history:
                                self.history_listbox.insert(tk.END, item)
                        time.sleep(cfg['cooldown'])
                if self.debug_status_var.get() and hasattr(self, 'status_labels'):
                    def update_status(name: str, val: bool | str) -> None:
                        color = "green" if val else "red"
                        self.status_labels[name].config(text=str(val), fg=color)
                    update_status("bar_box", bar_box is not None)
                    update_status("target_x", target_x is not None)
                    update_status("edge_touched", self.edge_touched)
                    self.status_labels["can_click"].config(
                        text=self.last_click_coords,
                        fg="green" if can_click else "black")
                    self.status_labels["require_edge"].config(text=str(cfg.get("require_edge_touch", True)))
                if cfg['debug_view']:
                    debug = cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)
                    if target_x is not None:
                        cv2.line(debug, (int(target_x - x), 0), (int(target_x - x), h), (0, 255, 0), 1)
                    if bar_box:
                        cv2.rectangle(debug, (bar_box[0] - x, 0), (bar_box[1] - x, h), (0, 0, 255), 1)
                    cv2.line(debug, (margin, 0), (margin, h), (0, 255, 255), 1)
                    cv2.line(debug, (w - margin, 0), (w - margin, h), (0, 255, 255), 1)
                    self.debug_img = debug
                time.sleep(0.005)

    # ------------------------------------------------------------------
    #  Debug window rendering

    def display_debug_window(self) -> None:
        """Continuously update the debug window with the latest frame."""
        while self.running:
            if self.debug_img is not None:
                cv2.imshow("Debug View", self.debug_img)
                if cv2.waitKey(1) == 27:
                    self.running = False
                    break
            time.sleep(0.01)
        cv2.destroyAllWindows()

    # ------------------------------------------------------------------
    #  Auto sell functionality

    def perform_auto_sell(self) -> None:
        """Periodically execute a sequence of keypresses to perform an action."""
        while self.running:
            interval = self.config.get("auto_sell_interval", 600)
            point = self.config.get("auto_sell_point", (500, 500))
            time.sleep(max(0, interval - 2))
            self.partial_pause_event.clear()
            time.sleep(2.0)
            region = self.get_region()
            while self.running:
                if not self.is_event_active(region):
                    stable_start = time.time()
                    while time.time() - stable_start < 0.2:
                        if self.is_event_active(region) or not self.running:
                            break
                        time.sleep(0.05)
                    else:
                        break
                time.sleep(0.1)
            self.global_pause_event.clear()
            time.sleep(2.0)
            try:
                self.keyboard.press(Key.shift)
                self.keyboard.release(Key.shift)
                time.sleep(0.5)
                self.keyboard.press('g')
                self.keyboard.release('g')
                time.sleep(0.5)
                self.keyboard.press(Key.f8)
                self.keyboard.release(Key.f8)
                time.sleep(2.0)
                self.keyboard.press('g')
                self.keyboard.release('g')
                time.sleep(0.5)
                self.keyboard.press(Key.shift)
                self.keyboard.release(Key.shift)
                time.sleep(0.5)
            except Exception as exc:  # noqa: BLE001
                print(f"[AutoSell] Error: {exc}")
            self.global_pause_event.set()
            self.partial_pause_event.set()


# ---------------------------------------------------------------------------
#  Script entry point
#
#  The following block implements a simple command‑line prompt for a
#  licence key and launches the Tk user interface if verification
#  succeeds.  It has been retained from the original script with minor
#  edits for readability.

if __name__ == "__main__":
    import tkinter.simpledialog as simpledialog

    root = tk.Tk()
    root.withdraw()
    license_key = ""
    try:
        with open("license.txt", "r", encoding='utf-8') as f:
            license_key = f.read().strip()
    except Exception:
        pass
    valid, msg = decrypt_and_verify_license(license_key)
    if not valid:
        machine_id = get_machine_id()
        try:
            pyperclip.copy(machine_id)
        except Exception:
            pass
        messagebox.showinfo("Machine ID", f"Your ID has been copied to clipboard:\n{machine_id}")
        key_input = simpledialog.askstring("License Key", "Enter your encrypted license key:")
        if not key_input:
            messagebox.showerror("Missing", "License key not provided.")
            raise SystemExit
        valid, msg = decrypt_and_verify_license(key_input)
        if not valid:
            messagebox.showerror("Invalid", f"Key invalid: {msg}")
            raise SystemExit
        else:
            with open("license.txt", "w", encoding='utf-8') as f:
                f.write(key_input)
    root.deiconify()
    app = MacroAutomationTool(root)
    root.mainloop()