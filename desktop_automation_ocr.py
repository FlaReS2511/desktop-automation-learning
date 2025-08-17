from __future__ import annotations

import tkinter as tk
from tkinter import messagebox
import threading
import time
import json
import os
import pyautogui
import keyboard
import easyocr
import numpy as np
from PIL import ImageGrab, Image
import requests
import cv2
from rapidfuzz import fuzz
import io
import subprocess
import customtkinter as ctk


class AutomationApp:
    """A GUI tool for demonstrating OCR‑driven automation."""

    def __init__(self, root: tk.Tk) -> None:
        """Initialise the automation application and build the user interface.

        Args:
            root: the top level ``Tk`` instance provided by the caller.
        """
        self.root = root

        # Configure a dark themed appearance
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Attributes storing user defined regions and recovery point
        self.ocr_region: tuple[int, int, int, int] | None = None
        self.item_region: tuple[int, int, int, int] | None = None
        self.recovery_point: tuple[int, int] | None = None

        # Set up the OCR reader (GPU acceleration is used when available)
        self.reader = easyocr.Reader(['en'], gpu=True)

        # Flag used to gracefully stop background threads
        self.stop_flag: bool = False

        # Build the tabbed layout
        self.tabview = ctk.CTkTabview(root, corner_radius=15)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=20)

        # Name the tabs without decorative icons to keep the UI professional
        self.tab_config = self.tabview.add("Configuration")
        self.tab_ocr = self.tabview.add("OCR")
        self.tab_control = self.tabview.add("Control")

        # Construct each tab
        self._setup_config_tab()
        self._setup_ocr_tab()
        self._setup_control_tab()

        # Load any persisted configuration
        self._load_configuration()

    # ------------------------------------------------------------------
    # UI construction methods

    def _setup_config_tab(self) -> None:
        """Create controls for adjusting timing and key configuration."""

        def add_label_entry(parent: ctk.CTkBaseClass, text: str, row: int) -> ctk.CTkEntry:
            """Utility to create a label and entry on a specified row."""
            ctk.CTkLabel(parent, text=text).grid(row=row, column=0, sticky="w", pady=5)
            entry = ctk.CTkEntry(parent, width=180)
            entry.grid(row=row, column=1, pady=5)
            return entry

        # Timing and key parameters
        self.entry_hold = add_label_entry(self.tab_config, "Mouse hold duration (s):", 0)
        self.entry_start_delay = add_label_entry(self.tab_config, "Startup delay (s):", 1)
        self.entry_between_delay = add_label_entry(self.tab_config, "Delay between clicks (s):", 2)
        self.entry_click_delay = add_label_entry(self.tab_config, "Delay after a click (s):", 3)
        self.entry_shake_delay = add_label_entry(self.tab_config, "Shake delay after zero/full (s):", 4)
        self.entry_key_hold_time = add_label_entry(self.tab_config, "Key hold duration (s):", 5)
        self.entry_key1 = add_label_entry(self.tab_config, "Primary key (before shaking):", 6)
        self.entry_key2 = add_label_entry(self.tab_config, "Secondary key (after shaking):", 7)
        self.entry_stop_key = add_label_entry(self.tab_config, "Stop hotkey:", 8)
        self.entry_max_value = add_label_entry(self.tab_config, "Maximum counter value:", 9)

        # Checkbox controlling whether to repeat the workflow
        self.loop_var = ctk.IntVar()
        ctk.CTkCheckBox(
            self.tab_config,
            text="Repeat workflow",
            variable=self.loop_var
        ).grid(row=10, column=0, columnspan=2, pady=10)

        # Save configuration button
        ctk.CTkButton(
            self.tab_config,
            text="Save configuration",
            fg_color="orange",
            command=self.save_configuration
        ).grid(row=11, column=0, columnspan=2, pady=10)

    def _setup_ocr_tab(self) -> None:
        """Build controls used to define screen regions for OCR operations."""
        self.region_label = ctk.CTkLabel(self.tab_ocr, text="OCR region: Not set")
        self.region_label.pack(pady=5)

        self.recovery_point_label = ctk.CTkLabel(self.tab_ocr, text="Recovery point: Not set")
        self.recovery_point_label.pack(pady=5)

        ctk.CTkButton(
            self.tab_ocr,
            text="Select recovery point",
            command=self.set_recovery_point
        ).pack(pady=5)

        ctk.CTkButton(
            self.tab_ocr,
            text="Select OCR region",
            command=self.set_ocr_region
        ).pack(pady=5)

        self.item_region_label = ctk.CTkLabel(self.tab_ocr, text="Item OCR region: Not set")
        self.item_region_label.pack(pady=5)

        ctk.CTkButton(
            self.tab_ocr,
            text="Select item OCR region",
            command=self.set_item_region
        ).pack(pady=5)

    def _setup_control_tab(self) -> None:
        """Create controls to start and stop the automation and display progress."""
        ctk.CTkButton(
            self.tab_control,
            text="Start",
            fg_color="green",
            hover_color="darkgreen",
            command=self.start_automation
        ).pack(pady=10)

        self.status_label = ctk.CTkLabel(
            self.tab_control,
            text="Idle",
            font=("Arial", 14)
        )
        self.status_label.pack(pady=5)

        ctk.CTkButton(
            self.tab_control,
            text="Stop",
            fg_color="red",
            hover_color="darkred",
            command=self.stop_automation
        ).pack(pady=10)

        self.progress = ctk.CTkProgressBar(self.tab_control, width=300)
        self.progress.pack(pady=20)
        self.progress.set(0)

    # ------------------------------------------------------------------
    # Configuration persistence

    def _collect_configuration(self) -> dict | None:
        """Gather and validate configuration values from the UI.

        Returns:
            A dictionary containing validated configuration values or ``None``
            if any required values are missing or invalid.  An error dialog
            will be shown to the user when invalid input is detected.
        """
        try:
            config: dict = {
                "hold_time": float(self.entry_hold.get()),
                "start_delay": float(self.entry_start_delay.get()),
                "between_delay": float(self.entry_between_delay.get()),
                "click_delay": float(self.entry_click_delay.get()),
                "shake_delay": float(self.entry_shake_delay.get()),
                "key_hold_time": float(self.entry_key_hold_time.get()),
                "key1": self.entry_key1.get().strip(),
                "key2": self.entry_key2.get().strip(),
                "stop_key": self.entry_stop_key.get().strip(),
                "max_value": int(self.entry_max_value.get().strip()),
                "ocr_region": self.ocr_region,
                "item_region": self.item_region,
                "recovery_point": self.recovery_point,
            }

            # Validate presence of mandatory fields
            if not all([config["key1"], config["key2"], config["stop_key"]]):
                raise ValueError
            if config["ocr_region"] is None or config["item_region"] is None:
                raise ValueError
            return config
        except Exception:
            messagebox.showerror("Configuration error", "Please enter valid numbers and select all required regions.")
            return None

    def _load_configuration(self) -> None:
        """Load configuration from a JSON file located alongside this script."""
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                # Populate entries
                self.entry_hold.delete(0, "end")
                self.entry_hold.insert(0, str(cfg.get("hold_time", "")))
                self.entry_start_delay.delete(0, "end")
                self.entry_start_delay.insert(0, str(cfg.get("start_delay", "")))
                self.entry_between_delay.delete(0, "end")
                self.entry_between_delay.insert(0, str(cfg.get("between_delay", "")))
                self.entry_click_delay.delete(0, "end")
                self.entry_click_delay.insert(0, str(cfg.get("click_delay", "")))
                self.entry_shake_delay.delete(0, "end")
                self.entry_shake_delay.insert(0, str(cfg.get("shake_delay", "")))
                self.entry_key_hold_time.delete(0, "end")
                self.entry_key_hold_time.insert(0, str(cfg.get("key_hold_time", "")))
                self.entry_key1.delete(0, "end")
                self.entry_key1.insert(0, cfg.get("key1", ""))
                self.entry_key2.delete(0, "end")
                self.entry_key2.insert(0, cfg.get("key2", ""))
                self.entry_stop_key.delete(0, "end")
                self.entry_stop_key.insert(0, cfg.get("stop_key", ""))
                self.entry_max_value.delete(0, "end")
                self.entry_max_value.insert(0, str(cfg.get("max_value", "")))

                # Restore regions and recovery point
                self.ocr_region = tuple(cfg.get("ocr_region", ())) or None
                self.item_region = tuple(cfg.get("item_region", ())) or None
                self.recovery_point = tuple(cfg.get("recovery_point", ())) or None

                if self.ocr_region:
                    self.region_label.configure(text=f"OCR region: {self.ocr_region}")
                if self.item_region:
                    self.item_region_label.configure(text=f"Item OCR region: {self.item_region}")
                if self.recovery_point:
                    self.recovery_point_label.configure(text=f"Recovery point: {self.recovery_point}")
            except Exception as e:
                print(f"Failed to load configuration: {e}")

    def save_configuration(self) -> None:
        """Persist the current configuration to ``config.json`` in the script directory."""
        config = self._collect_configuration()
        if config:
            try:
                config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
                with open(config_path, "w", encoding="utf-8") as f:
                    json.dump(config, f, indent=4)
                messagebox.showinfo("Saved", f"Configuration saved to {config_path}")
            except Exception as e:
                messagebox.showerror("File error", str(e))

    # ------------------------------------------------------------------
    # Region selection methods

    def set_ocr_region(self) -> None:
        """Prompt the user to select the rectangular screen region used for OCR."""
        messagebox.showinfo(
            "Select region",
            "Move the mouse to the TOP‑LEFT corner and press ENTER.\n"
            "Then move the mouse to the BOTTOM‑RIGHT corner and press ENTER."
        )
        keyboard.wait('enter')
        x1, y1 = pyautogui.position()
        keyboard.wait('enter')
        x2, y2 = pyautogui.position()
        left = min(x1, x2)
        top = min(y1, y2)
        width = abs(x2 - x1)
        height = abs(y2 - y1)
        self.ocr_region = (left, top, width, height)
        self.region_label.configure(text=f"OCR region: {self.ocr_region}")

    def set_item_region(self) -> None:
        """Prompt the user to select the region for item OCR scanning."""
        messagebox.showinfo(
            "Select item region",
            "Move the mouse to the TOP‑LEFT corner and press ENTER.\n"
            "Then move the mouse to the BOTTOM‑RIGHT corner and press ENTER."
        )
        keyboard.wait('enter')
        x1, y1 = pyautogui.position()
        keyboard.wait('enter')
        x2, y2 = pyautogui.position()
        left = min(x1, x2)
        top = min(y1, y2)
        width = abs(x2 - x1)
        height = abs(y2 - y1)
        self.item_region = (left, top, width, height)
        self.item_region_label.configure(text=f"Item OCR region: {self.item_region}")

    def set_recovery_point(self) -> None:
        """Prompt the user to select a single recovery point on screen."""
        messagebox.showinfo("Select recovery point", "Move the mouse to the recovery location and press ENTER.")
        keyboard.wait('enter')
        x, y = pyautogui.position()
        self.recovery_point = (x, y)
        self.recovery_point_label.configure(text=f"Recovery point: {self.recovery_point}")

    # ------------------------------------------------------------------
    # Automation control

    def start_automation(self) -> None:
        """Begin the automation workflow in a background thread."""
        config = self._collect_configuration()
        if not config:
            return

        # Reset stop flag and register a hotkey to abort
        self.stop_flag = False
        keyboard.add_hotkey(config["stop_key"], lambda: self.stop_automation())

        def update_status(text: str, progress_value: float) -> None:
            self.status_label.configure(text=text)
            self.progress.set(progress_value)

        def workflow_loop() -> None:
            recovery_fail_count = 0
            first_time_run = True
            while not self.stop_flag:
                # Stage 1: click and hold to fill the gauge
                should_restart, first_time_run, recovery_fail_count = self.perform_click_and_hold_phase(
                    config, first_time_run, recovery_fail_count, update_status
                )
                if should_restart:
                    continue
                # Stage 2: shaking sequence until OCR bar changes
                self.perform_shake_phase(config, update_status)
                # Stage 3: item detection via OCR
                self.perform_item_detection_phase(update_status)
                # Break if not repeating
                if not self.loop_var.get():
                    break
            keyboard.remove_hotkey(config["stop_key"])

        threading.Thread(target=workflow_loop, daemon=True).start()
        messagebox.showinfo("Running", f"Workflow will start after {config['start_delay']} seconds")

    def stop_automation(self) -> None:
        """Stop the automation workflow at the next safe opportunity."""
        self.stop_flag = True

    # ------------------------------------------------------------------
    # Workflow phases

    def perform_click_and_hold_phase(
        self,
        config: dict,
        first_time_run: bool,
        recovery_fail_count: int,
        update_status
    ) -> tuple[bool, bool, int]:
        """Stage 1: repeatedly click and hold until the gauge is full or a recovery is required.

        Args:
            config: configuration values collected from the UI
            first_time_run: flag indicating whether this is the first iteration
            recovery_fail_count: count of consecutive recovery attempts
            update_status: callback to update the UI

        Returns:
            A tuple ``(should_restart, next_first_time_run, next_recovery_fail_count)``.
            ``should_restart`` signals whether the workflow loop should restart,
            ``next_first_time_run`` carries the updated first time state and
            ``next_recovery_fail_count`` carries the updated recovery count.
        """
        target_full = f"{config['max_value']}/{config['max_value']}"
        target_zero = f"0/{config['max_value']}"
        should_restart = False
        # On the very first run wait for the startup delay
        if first_time_run:
            update_status(f"Waiting {config['start_delay']} seconds before starting...", 0.0)
            time.sleep(config['start_delay'])
            first_time_run = False
        update_status("Phase 1: Filling gauge", 0.0)
        zero_stable_start: float | None = None
        while not self.stop_flag:
            # If gauge is full prior to clicking we can exit stage 1
            if self._contains_ocr_text(target_full, config['ocr_region']):
                break
            # Detect if zero persists for 3 seconds
            if zero_stable_start and time.time() - zero_stable_start >= 3.0:
                confirm = all(self._contains_ocr_text(target_zero, config['ocr_region']) for _ in range(3))
                if confirm:
                    # Perform auto recovery and send a notification
                    if self.stop_flag:
                        return False, first_time_run, recovery_fail_count
                    self.perform_auto_recovery()
                    self.send_webhook_alert()
                    recovery_fail_count += 1
                    # If recovery fails twice, try pressing the secondary key briefly
                    if recovery_fail_count > 1:
                        keyboard.press(config['key2'])
                        time.sleep(config['key_hold_time'] / 2)
                        keyboard.release(config['key2'])
                        recovery_fail_count = 0
                    should_restart = True
                    first_time_run = True
                    break
                zero_stable_start = None
            # Track when zero first appears
            if self._contains_ocr_text(target_zero, config['ocr_region']):
                if zero_stable_start is None:
                    zero_stable_start = time.time()
            else:
                zero_stable_start = None
            # Perform click and hold sequence
            pyautogui.mouseDown()
            time.sleep(config['hold_time'])
            pyautogui.mouseUp()
            time.sleep(config['between_delay'])
            # If gauge becomes full after the click we can exit stage 1
            if self._contains_ocr_text(target_full, config['ocr_region']):
                recovery_fail_count = 0
                break
        return should_restart, first_time_run, recovery_fail_count

    def perform_shake_phase(self, config: dict, update_status) -> None:
        """Stage 2: shake sequence until the OCR bar changes."""
        target_full = f"{config['max_value']}/{config['max_value']}"
        target_zero = f"0/{config['max_value']}"
        update_status("Phase 2: Shaking sequence", 0.33)
        if self.stop_flag:
            return
        self._execute_shake_sequence_until_bar_changes(config, target_full, target_zero)
        update_status("Phase 2: Completed", 0.66)

    def perform_item_detection_phase(self, update_status) -> None:
        """Stage 3: spawn a thread to scan for special items using OCR."""
        threading.Thread(target=self.detect_special_items, daemon=True).start()
        update_status("Phase 3: Scanning items", 1.0)

    # ------------------------------------------------------------------
    # Helper methods

    def _contains_ocr_text(self, target: str, region: tuple[int, int, int, int]) -> bool:
        """Check whether the OCR result from the given region contains the target string."""
        left, top, width, height = region
        right = left + width
        bottom = top + height
        img = ImageGrab.grab(bbox=(left, top, right, bottom))
        result = self.reader.readtext(np.array(img), detail=0)
        for text in result:
            cleaned = text.strip().replace(" ", "").replace("O", "0")
            if cleaned == target:
                return True
        return False

    def _is_ocr_bar_static(self, region: tuple[int, int, int, int], check_time: float = 4.0, interval: float = 0.5) -> bool:
        """Determine if the OCR bar remains unchanged for a specified duration."""
        last_text: str | None = None
        stable_start = time.time()
        while time.time() - stable_start < check_time:
            if self.stop_flag:
                return False
            left, top, width, height = region
            img = ImageGrab.grab(bbox=(left, top, left + width, top + height))
            result = self.reader.readtext(np.array(img), detail=0)
            cleaned = "".join(t.strip().replace(" ", "").replace("O", "0") for t in result)
            if last_text is None:
                last_text = cleaned
            elif cleaned != last_text:
                return False
            time.sleep(interval)
        return True

    def _wait_for_zero_or_full_during_hold(self, config: dict, target_full: str, target_zero: str) -> None:
        """Wait while holding the mouse button until zero appears three times, then press the secondary key."""
        zero_count = 0
        while not self.stop_flag:
            if self._contains_ocr_text(target_zero, config["ocr_region"]):
                zero_count += 1
                if zero_count >= 3:
                    pyautogui.mouseUp()
                    break
            else:
                zero_count = 0
            time.sleep(0.3)
        time.sleep(config["shake_delay"])
        keyboard.press(config["key2"])
        time.sleep(config["key_hold_time"] + 0.001)
        keyboard.release(config["key2"])

    def _execute_shake_sequence_until_bar_changes(
        self,
        config: dict,
        target_full: str,
        target_zero: str,
    ) -> None:
        """Execute the shake sequence until the OCR bar changes."""
        first_attempt = True
        while not self.stop_flag:
            hold_time = config["key_hold_time"] if first_attempt else config["key_hold_time"] / 2
            # press primary key
            keyboard.press(config["key1"])
            time.sleep(hold_time)
            keyboard.release(config["key1"])
            # simple click
            pyautogui.click()
            time.sleep(config["click_delay"])
            # click and hold
            pyautogui.mouseDown()
            time.sleep(0.1)
            # if bar is static, reset and retry
            if self._is_ocr_bar_static(config["ocr_region"], check_time=3.0):
                pyautogui.mouseUp()
                first_attempt = False
                continue
            else:
                break
        self._wait_for_zero_or_full_during_hold(config, target_full, target_zero)

    def detect_special_items(self) -> None:
        """Scan the item region for special keywords and halt automation if found."""
        time.sleep(0.2)
        region = self.item_region
        if not region:
            return
        left, top, width, height = region
        right = left + width
        bottom = top + height
        keywords = ["mythic", "exotic", "ex@tic", "aetherite"]
        found_keyword: str | None = None
        for _ in range(1):
            img = ImageGrab.grab(bbox=(left, top, right, bottom))
            img_np = np.array(img.convert("RGB"))
            results = self.reader.readtext(img_np, detail=1)
            for _, text, conf in results:
                if conf > 0.5:
                    text_clean = text.lower().strip()
                    for kw in keywords:
                        score = fuzz.partial_ratio(text_clean, kw)
                        if score >= 75:
                            found_keyword = text
                            break
                if found_keyword:
                    break
            if found_keyword:
                break
            time.sleep(0.2)
        if found_keyword:
            self.stop_automation()
            # send a user mention via webhook (ID could be replaced)
            self.send_discord_message(f"Special item detected: `{found_keyword.upper()}`")

    def perform_auto_recovery(self) -> None:
        """Attempt to recover from a stuck state by returning to a known point and clicking."""
        keyboard.press_and_release('g')
        time.sleep(0.5)
        if self.recovery_point:
            x, y = self.recovery_point
            pyautogui.moveTo(x, y, duration=0.5)
            if self.stop_flag:
                return
            time.sleep(0.5)
            # External helper to click; path may need adjustment in different environments
            subprocess.run(["C:\\Users\\ADMINA1\\OneDrive\\3_IT\\prospecting\\nudge_click.exe"])
        time.sleep(1.5)
        keyboard.press_and_release('g')

    # ------------------------------------------------------------------
    # Discord notification helpers

    def send_webhook_alert(self) -> None:
        """Send a generic alert to a Discord webhook indicating a recovery was attempted."""
        url = "" # Discord webhook URL
        try:
            content = "Workflow encountered a stuck state and performed a recovery."
            requests.post(url, json={"content": content})
        except Exception as e:
            print(f"Failed to send webhook alert: {e}")

    def send_discord_message(self, content: str) -> None:
        """Send a custom message to a Discord webhook."""
        url = "" # Discord webhook URL
        try:
            response = requests.post(url, json={"content": content})
            if response.status_code not in (200, 204):
                print(f"Discord webhook returned {response.status_code}: {response.text}")
        except Exception as e:
            print(f"Error sending Discord message: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AutomationApp(root)
    root.mainloop()