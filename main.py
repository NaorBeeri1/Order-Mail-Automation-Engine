"""
Order Mail Monitor & Win32 Automation Engine
--------------------------------------------
A multi-threaded desktop application that monitors incoming emails via IMAP,
extracts order barcodes from PDF attachments in parallel, and dispatches
recorded UI macro sequences to legacy desktop ERP/billing software via Win32 API.
"""

import ctypes
import email
import email.header
import imaplib
import io
import json
import logging
import os
import platform
import re
import subprocess
import threading
import time
import urllib.parse
import webbrowser
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Set, Tuple

import fitz  # PyMuPDF
import tkinter as tk
from dotenv import load_dotenv
from PIL import Image
from pyzbar.pyzbar import decode

# Load environment variables from .env file
load_dotenv()

# Configure structured logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("MailAutomationEngine")

# --- CONFIGURATION ---
MONITOR_EMAIL = os.getenv("MONITOR_EMAIL", "")
MONITOR_PASSWORD = os.getenv("MONITOR_PASSWORD", "")

BATCH_EMAIL = os.getenv("BATCH_EMAIL", "")
BATCH_APP_PASSWORD = os.getenv("BATCH_APP_PASSWORD", "")

USER_PROFILE = os.path.expanduser("~")
BASE_DATA_DIR = os.getenv(
    "APP_DATA_DIR", os.path.join(USER_PROFILE, "OrderMailAutomation", "Data")
)
MACRO_CONFIG_DIR = os.path.join(BASE_DATA_DIR, "MacroProfiles")
SAVE_FILE = os.path.join(BASE_DATA_DIR, "saved_targets.json")

os.makedirs(BASE_DATA_DIR, exist_ok=True)
os.makedirs(MACRO_CONFIG_DIR, exist_ok=True)

# --- WIN32 API CONSTANTS ---
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002
MOUSEEVENTF_LEFTDOWN = 0x02
MOUSEEVENTF_LEFTUP = 0x04
VK_F10 = 0x79

EXTENDED_VKS = {0x5B, 0x5C, 0x25, 0x26, 0x27, 0x28, 0x2E, 0x24, 0x23, 0x21, 0x22}

VK_MAP: Dict[str, int] = {
    "shift": 0x10,
    "ctrl": 0x11,
    "control": 0x11,
    "alt": 0x12,
    "win": 0x5B,
    "windows": 0x5B,
    "lwin": 0x5B,
    "rwin": 0x5C,
    "enter": 0x0D,
    "return": 0x0D,
    "tab": 0x09,
    "space": 0x20,
    "esc": 0x1B,
    "escape": 0x1B,
    "backspace": 0x08,
    "up": 0x26,
    "down": 0x28,
    "left": 0x25,
    "right": 0x27,
    "delete": 0x2E,
    "home": 0x24,
    "end": 0x23,
    "pageup": 0x21,
    "pagedown": 0x22,
}

BRANCH_EXPORT_MAP: Dict[str, str] = {
    "branch_main": "export_main.txt",
    "branch_north": "export_north.txt",
    "branch_south": "export_south.txt",
}


# =====================================================================
# WIN32 NATIVE INPUT & WINDOW CONTROLLER
# =====================================================================


class Win32Controller:
    """Encapsulates low-level Windows keyboard, mouse, and window management."""

    @staticmethod
    def resolve_vk(key_name: str) -> int:
        key_clean = key_name.strip().lower()
        if key_clean in VK_MAP:
            return VK_MAP[key_clean]
        if len(key_clean) == 1:
            return ord(key_clean.upper())
        if key_clean.startswith("vk_") and key_clean[3:].isdigit():
            return int(key_clean[3:])
        return 0

    @staticmethod
    def press_key_down(vk_code: int) -> None:
        if platform.system() != "Windows":
            return
        flags = KEYEVENTF_EXTENDEDKEY if vk_code in EXTENDED_VKS else 0
        ctypes.windll.user32.keybd_event(vk_code, 0, flags, 0)

    @staticmethod
    def release_key_up(vk_code: int) -> None:
        if platform.system() != "Windows":
            return
        flags = KEYEVENTF_KEYUP
        if vk_code in EXTENDED_VKS:
            flags |= KEYEVENTF_EXTENDEDKEY
        ctypes.windll.user32.keybd_event(vk_code, 0, flags, 0)

    @classmethod
    def tap_key(cls, vk_code: int) -> None:
        cls.press_key_down(vk_code)
        time.sleep(0.05)
        cls.release_key_up(vk_code)
        time.sleep(0.05)

    @staticmethod
    def click_at(x: int, y: int, clicks: int = 1) -> None:
        if platform.system() != "Windows":
            return
        user32 = ctypes.windll.user32
        user32.SetCursorPos(x, y)
        time.sleep(0.08)
        for _ in range(clicks):
            user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.05)
            user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            time.sleep(0.05)

    @staticmethod
    def is_key_pressed(vk_code: int) -> bool:
        if platform.system() != "Windows":
            return False
        return (ctypes.windll.user32.GetAsyncKeyState(vk_code) & 0x8000) != 0

    @staticmethod
    def focus_window_by_keyword(keyword: str) -> bool:
        if platform.system() != "Windows" or not keyword or not keyword.strip():
            return False

        user32 = ctypes.windll.user32
        keyword_clean = keyword.strip().lower()
        found_hwnd: List[Optional[int]] = [None]

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def enum_windows_callback(hwnd, _extra):
            if user32.IsWindowVisible(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buff = ctypes.create_unicode_buffer(length + 1)
                    user32.GetWindowTextW(hwnd, buff, length + 1)
                    if keyword_clean in buff.value.lower():
                        found_hwnd[0] = hwnd
                        return False
            return True

        user32.EnumWindows(enum_windows_callback, 0)
        if found_hwnd[0]:
            hwnd = found_hwnd[0]
            user32.ShowWindow(hwnd, 9)  # SW_RESTORE
            user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
            user32.SetForegroundWindow(hwnd)
            return True
        return False


# =====================================================================
# PDF & BARCODE EXTRACTION SERVICE
# =====================================================================


class DocumentParser:
    """Extracts barcodes and order metadata from binary PDF payloads."""

    @staticmethod
    def decode_mime_filename(encoded_filename: Optional[str]) -> str:
        if not encoded_filename:
            return ""
        decoded_parts = email.header.decode_header(encoded_filename)
        parts = []
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                parts.append(part.decode(encoding or "utf-8", errors="ignore"))
            else:
                parts.append(str(part))
        return "".join(parts)

    @staticmethod
    def extract_barcodes_from_pdf(filename: str, pdf_bytes: bytes) -> Tuple[str, List[str]]:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        order_number = "".join(filter(str.isdigit, filename)) or "Unknown"

        barcodes: List[str] = []
        full_text = ""

        for page in doc:
            full_text += page.get_text()
            pix = page.get_pixmap(dpi=150)
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            for obj in decode(image):
                code = obj.data.decode("utf-8")
                if code not in barcodes:
                    barcodes.append(code)

        # Fallback regex extraction if raster barcode scan yields no hits
        if not barcodes:
            matches = re.findall(r"\b(1000\d{6}|\d{4})\b", full_text)
            for match in matches:
                if match != order_number and match not in barcodes:
                    barcodes.append(match)

        return order_number, barcodes


# =====================================================================
# DATA MODELS
# =====================================================================


@dataclass
class TrackedOrderRow:
    frame: tk.Frame
    entry_widget: tk.Entry
    note_widget: tk.Entry
    date_widget: tk.Entry
    status_var: tk.StringVar
    status_label: tk.Label
    update_status_fn: callable
    checkbox_widget: tk.Checkbutton
    checkbox_var: tk.IntVar
    target_value: str = ""
    processed_uids: Set[str] = field(default_factory=set)
    is_active: bool = False
    is_alerted: bool = False
    msg_id_hex: str = ""


# =====================================================================
# MAIN APPLICATION UI & ORCHESTRATOR
# =====================================================================


class OrderMailAutomationApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Order Mail Monitor & ERP Macro Automation")
        self.root.geometry("1400x920")
        self.root.minsize(1100, 600)
        self.root.configure(bg="#0f172a")

        self.lock = threading.Lock()
        self.is_running = False
        self.selecting_mode = False
        self.sort_ascending = True
        self.completed_count = 0

        self.rows: List[TrackedOrderRow] = []
        self.completed_cards: List[Tuple[tk.Frame, TrackedOrderRow]] = []
        self.kept_history: List[Tuple[str, str]] = []

        self.profiles: Dict[str, dict] = {
            "branch_main": {"Steps": [], "Loop": False, "AppKeyword": "Billing"},
            "branch_north": {"Steps": [], "Loop": False, "AppKeyword": "NorthERP"},
            "branch_south": {"Steps": [], "Loop": False, "AppKeyword": "SouthERP"},
        }

        self._bind_clipboard_shortcuts()
        self._load_macro_profiles()
        self._build_ui()
        self._load_persisted_state()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # -----------------------------------------------------------------
    # UI CONSTRUCTION
    # -----------------------------------------------------------------

    def _build_ui(self) -> None:
        main_container = tk.Frame(self.root, bg="#0f172a")
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # Header
        header = tk.Frame(main_container, bg="#0f172a")
        header.pack(fill="x", pady=(0, 10))

        left_header = tk.Frame(header, bg="#0f172a")
        left_header.pack(side=tk.LEFT)

        tk.Label(
            left_header,
            text="Order Subject Monitor & ERP Automation",
            font=("Segoe UI", 16, "bold"),
            fg="#f8fafc",
            bg="#0f172a",
        ).pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(
            left_header,
            text="Filter:",
            font=("Segoe UI", 9, "bold"),
            fg="#94a3b8",
            bg="#0f172a",
        ).pack(side=tk.LEFT, padx=(0, 6))

        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", self._filter_rows)
        tk.Entry(
            left_header,
            textvariable=self.filter_var,
            font=("Segoe UI", 11),
            bg="#1e293b",
            fg="#f8fafc",
            insertbackground="white",
            relief="flat",
            width=18,
        ).pack(side=tk.LEFT, ipady=4, padx=4)

        self.sort_btn = tk.Button(
            left_header,
            text="Sort: Old -> New",
            command=self._toggle_sorting,
            bg="#334155",
            fg="#38bdf8",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=4,
        )
        self.sort_btn.pack(side=tk.LEFT, padx=(8, 0))

        tk.Button(
            header,
            text="View Kept History",
            command=self._open_history_modal,
            bg="#0ea5e9",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
        ).pack(side=tk.RIGHT)

        # Split Workspace
        split = tk.Frame(main_container, bg="#0f172a")
        split.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        split.columnconfigure(0, weight=1)
        split.columnconfigure(1, weight=1)
        split.rowconfigure(0, weight=1)

        # Left Pane (Active Tracking)
        left_pane = tk.Frame(split, bg="#0f172a")
        left_pane.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left_pane.rowconfigure(1, weight=1)
        left_pane.columnconfigure(0, weight=1)

        self.active_label = tk.Label(
            left_pane,
            text="Active Tracking (0):",
            font=("Segoe UI", 10, "bold"),
            fg="#38bdf8",
            bg="#0f172a",
        )
        self.active_label.grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.canvas_active, self.rows_frame = self._create_scrollable_panel(left_pane)

        # Right Pane (Completed Matches)
        right_pane = tk.Frame(split, bg="#0f172a")
        right_pane.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right_pane.rowconfigure(1, weight=1)
        right_pane.columnconfigure(0, weight=1)

        right_header = tk.Frame(right_pane, bg="#0f172a")
        right_header.grid(row=0, column=0, sticky="ew", pady=(0, 4))

        self.completed_label = tk.Label(
            right_header,
            text="Matched Emails (0):",
            font=("Segoe UI", 10, "bold"),
            fg="#f43f5e",
            bg="#0f172a",
        )
        self.completed_label.pack(side=tk.LEFT)

        tk.Button(
            right_header,
            text="Clear All Completed",
            command=self._clear_all_completed,
            bg="#dc2626",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=2,
        ).pack(side=tk.RIGHT)

        self.canvas_completed, self.completed_frame = self._create_scrollable_panel(right_pane)

        # Footer Controls
        footer = tk.Frame(main_container, bg="#0f172a")
        footer.pack(fill="x", pady=(5, 0))

        action_row = tk.Frame(footer, bg="#0f172a")
        action_row.pack(fill="x", pady=(0, 5))
        action_row.columnconfigure(0, weight=1)
        action_row.columnconfigure(1, weight=1)

        self.start_btn = tk.Button(
            action_row,
            text="Start / Update Live Monitoring",
            command=self.start_monitoring,
            bg="#10b981",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2",
        )
        self.start_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5), ipady=10)

        batch_container = tk.Frame(action_row, bg="#0f172a")
        batch_container.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        self.batch_btn = tk.Button(
            batch_container,
            text="Process & Dispatch Batch",
            command=self._handle_batch_button,
            bg="#2563eb",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2",
        )
        self.batch_btn.pack(side=tk.TOP, fill="x", ipady=10)

        self.select_pending_btn = tk.Button(
            batch_container,
            text="Select All (Pending Only)",
            command=self._select_all_pending,
            bg="#d97706",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
        )

        macro_row = tk.Frame(footer, bg="#0f172a")
        macro_row.pack(fill="x", pady=(4, 0))

        tk.Button(
            macro_row,
            text="Configure Macro Profiles",
            command=self._open_macro_builder,
            bg="#475569",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
        ).pack(side=tk.LEFT, ipady=4, padx=(0, 10))

        self.status_label = tk.Label(
            footer,
            text="Status: Idle & Ready",
            fg="#38bdf8",
            bg="#0f172a",
            font=("Segoe UI", 10, "bold"),
        )
        self.status_label.pack(anchor="w", pady=(4, 0))

    def _create_scrollable_panel(self, parent: tk.Frame) -> Tuple[tk.Canvas, tk.Frame]:
        card = tk.Frame(
            parent, bg="#1e293b", highlightbackground="#334155", highlightthickness=1
        )
        card.grid(row=1, column=0, sticky="nsew")
        card.rowconfigure(0, weight=1)
        card.columnconfigure(0, weight=1)

        canvas = tk.Canvas(card, bg="#1e293b", highlightthickness=0)
        scrollbar = tk.Scrollbar(card, orient="vertical", command=canvas.yview)
        inner_frame = tk.Frame(canvas, bg="#1e293b")

        inner_frame.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        win_id = canvas.create_window((0, 0), window=inner_frame, anchor="nw")
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfig(win_id, width=e.width),
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        scrollbar.grid(row=0, column=1, sticky="ns", pady=8, padx=(0, 4))

        for widget in (canvas, inner_frame):
            widget.bind(
                "<MouseWheel>",
                lambda e, c=canvas: c.yview_scroll(int(-1 * (e.delta / 120)), "units"),
            )

        return canvas, inner_frame

    # -----------------------------------------------------------------
    # ROW & STATE MANAGEMENT
    # -----------------------------------------------------------------

    def add_entry_row(
        self,
        initial_val: str = "",
        initial_note: str = "",
        initial_date: str = "",
        initial_status: str = "Pending",
        auto_save: bool = True,
        focus_new: bool = False,
    ) -> Optional[TrackedOrderRow]:
        today_str = datetime.now().strftime("%d/%m/%Y")
        if not initial_date or ("/" not in initial_date and "-" not in initial_date):
            initial_date = today_str

        if focus_new:
            with self.lock:
                for item in self.rows:
                    if not item.is_active and not item.entry_widget.get().strip():
                        item.entry_widget.focus_set()
                        self.root.update_idletasks()
                        self.canvas_active.yview_moveto(1.0)
                        return item

        row_frame = tk.Frame(
            self.rows_frame,
            bg="#0f172a",
            highlightbackground="#334155",
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        row_frame.pack(fill="x", expand=True, pady=6, padx=6)

        status_var = tk.StringVar(value=initial_status)
        status_lbl = tk.Label(
            row_frame, font=("Segoe UI", 9, "bold"), width=13, relief="flat", padx=4, pady=4
        )

        def update_status_display(val: str) -> None:
            if val == "Sent Status":
                status_lbl.config(text="Dispatched", bg="#065f46", fg="#34d399")
            else:
                status_lbl.config(text="Pending", bg="#78350f", fg="#fbbf24")

        update_status_display(initial_status)

        date_box = tk.Frame(
            row_frame, bg="#1e293b", highlightbackground="#64748b", highlightthickness=1, padx=2, pady=2
        )
        date_entry = tk.Entry(
            date_box,
            font=("Segoe UI", 11, "bold"),
            bg="#1e293b",
            fg="#38bdf8",
            insertbackground="white",
            relief="flat",
            width=12,
            justify="center",
        )
        date_entry.pack(side=tk.LEFT, fill="x", expand=True, ipady=6, padx=2)
        date_entry.insert(0, initial_date)

        note_box = tk.Frame(
            row_frame, bg="#1e293b", highlightbackground="#64748b", highlightthickness=1, padx=2, pady=2
        )
        note_entry = tk.Entry(
            note_box,
            font=("Segoe UI", 12),
            bg="#1e293b",
            fg="#f8fafc",
            insertbackground="white",
            relief="flat",
            width=14,
        )
        note_entry.pack(side=tk.LEFT, fill="x", expand=True, ipady=6, padx=4)

        placeholder = "Note..."
        if initial_note:
            note_entry.insert(0, initial_note)
        else:
            note_entry.insert(0, placeholder)
            note_entry.config(fg="#64748b")

        note_entry.bind(
            "<FocusIn>",
            lambda _e: (
                note_entry.delete(0, tk.END),
                note_entry.config(fg="#f8fafc"),
            )
            if note_entry.get() == placeholder
            else None,
        )
        note_entry.bind(
            "<FocusOut>",
            lambda _e: (
                (
                    note_entry.insert(0, placeholder),
                    note_entry.config(fg="#64748b"),
                )
                if not note_entry.get().strip()
                else None,
                self._save_state(),
            ),
        )

        entry = tk.Entry(
            row_frame,
            font=("Segoe UI", 13, "bold"),
            bg="#1e293b",
            fg="#f8fafc",
            insertbackground="white",
            relief="flat",
        )
        if initial_val:
            entry.insert(0, initial_val)
            entry.config(state="disabled", disabledbackground="#1e293b", disabledforeground="#94a3b8")

        chk_var = tk.IntVar(value=0)
        chk_box = tk.Checkbutton(
            row_frame,
            variable=chk_var,
            bg="#1e293b",
            activebackground="#1e293b",
            selectcolor="#0f172a",
            fg="#ffffff",
            highlightbackground="#38bdf8",
            highlightthickness=1,
            cursor="hand2",
        )

        row_obj = TrackedOrderRow(
            frame=row_frame,
            entry_widget=entry,
            note_widget=note_entry,
            date_widget=date_entry,
            status_var=status_var,
            status_label=status_lbl,
            update_status_fn=update_status_display,
            checkbox_widget=chk_box,
            checkbox_var=chk_var,
            target_value=initial_val,
        )

        # Pack right-to-left controls
        tk.Button(
            row_frame,
            text="-",
            command=lambda: self._remove_row(row_obj),
            bg="#dc2626",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2",
            width=3,
        ).pack(side=tk.RIGHT, padx=(4, 0), ipady=5)

        tk.Button(
            row_frame,
            text="+",
            command=lambda: self._commit_and_add_row(row_obj),
            bg="#2563eb",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2",
            width=3,
        ).pack(side=tk.RIGHT, padx=4, ipady=5)

        status_lbl.pack(side=tk.RIGHT, padx=6)
        date_box.pack(side=tk.RIGHT, padx=6)
        note_box.pack(side=tk.RIGHT, padx=6)

        if self.selecting_mode:
            chk_box.pack(side=tk.LEFT, padx=(2, 10))
        entry.pack(side=tk.LEFT, fill="x", expand=True, ipady=8, padx=(4, 8))

        # Bindings
        for widget in (entry, note_entry, date_entry):
            widget.bind("<Return>", lambda _e: self._commit_and_add_row(row_obj))
        note_entry.bind("<KeyRelease>", lambda _e: self._save_state())
        date_entry.bind("<KeyRelease>", lambda _e: [self._save_state(), self._reorder_rows()])

        with self.lock:
            self.rows.append(row_obj)

        if auto_save:
            self._save_state()
        self._reorder_rows()

        if focus_new:
            self.root.update_idletasks()
            self.canvas_active.yview_moveto(1.0)
            entry.focus_set()

        return row_obj

    def _commit_and_add_row(self, row_obj: TrackedOrderRow) -> str:
        self._save_state()
        if not row_obj.is_active:
            val = row_obj.entry_widget.get().strip()
            if val:
                row_obj.target_value = val
                row_obj.is_active = True
                row_obj.is_alerted = False
                row_obj.entry_widget.config(
                    state="disabled", disabledbackground="#1e293b", disabledforeground="#94a3b8"
                )

        self.start_monitoring()
        self.add_entry_row(auto_save=True, focus_new=True)
        return "break"

    def _remove_row(self, row_obj: TrackedOrderRow) -> None:
        with self.lock:
            row_obj.is_active = False
            if row_obj in self.rows:
                self.rows.remove(row_obj)

        try:
            row_obj.frame.destroy()
        except tk.TclError:
            pass

        self._save_state()
        with self.lock:
            active_count = sum(1 for r in self.rows if r.is_active)
            total_count = len(self.rows)

        if total_count == 0:
            self.add_entry_row(auto_save=True, focus_new=True)
            self.is_running = False
            self.status_label.config(text="Status: All tracking complete. Ready.")
        else:
            self.status_label.config(text=f"Status: Monitoring {active_count} target(s)...")
        self._update_counters()

    # -----------------------------------------------------------------
    # LIVE IMAP POLLING ENGINE
    # -----------------------------------------------------------------

    def start_monitoring(self) -> None:
        if not MONITOR_EMAIL or not MONITOR_PASSWORD:
            self.status_label.config(
                text="Status: Missing MONITOR_EMAIL / MONITOR_PASSWORD in .env"
            )
            return

        total_active = 0
        with self.lock:
            for item in self.rows:
                if item.is_active:
                    total_active += 1
                    continue
                val = item.entry_widget.get().strip()
                if val:
                    item.target_value = val
                    item.is_active = True
                    item.is_alerted = False
                    item.entry_widget.config(
                        state="disabled", disabledbackground="#1e293b", disabledforeground="#94a3b8"
                    )
                    total_active += 1

        if total_active == 0:
            return

        self._save_state()
        self._update_counters()
        self.status_label.config(text=f"Status: Monitoring {total_active} target(s) live...")

        if not self.is_running:
            self.is_running = True
            threading.Thread(target=self._imap_monitor_worker, daemon=True).start()

    def _imap_monitor_worker(self) -> None:
        mail: Optional[imaplib.IMAP4_SSL] = None
        while self.is_running:
            try:
                if mail is None:
                    mail = imaplib.IMAP4_SSL("imap.gmail.com")
                    mail.login(MONITOR_EMAIL, MONITOR_PASSWORD)
                    mail.select("INBOX", readonly=True)
                else:
                    mail.noop()

                with self.lock:
                    active_targets = [r for r in self.rows if r.is_active and not r.is_alerted]

                for item in active_targets:
                    target = item.target_value.strip()
                    if not target:
                        continue

                    status, response = mail.uid("SEARCH", None, f'(SUBJECT "{target}")')
                    if status != "OK" or not response[0]:
                        continue

                    for uid_bytes in response[0].split():
                        uid_str = uid_bytes.decode()
                        if uid_str in item.processed_uids:
                            continue

                        msg_id_hex = self._fetch_gmail_msgid_hex(mail, uid_str)

                        with self.lock:
                            item.processed_uids.add(uid_str)
                            item.is_alerted = True
                            item.msg_id_hex = msg_id_hex

                        self._save_state()
                        summary = self._format_row_summary(item)
                        self.root.after(
                            0,
                            lambda r=item, s=summary, m=msg_id_hex: self._render_completed_card(r, s, m),
                        )
                        break

                now_str = time.strftime("%H:%M:%S")
                self.root.after(
                    0,
                    lambda t=now_str, c=len(active_targets): self.status_label.config(
                        text=f"Live checking {c} target(s)... Last poll: {t}"
                    ),
                )
            except Exception as exc:
                logger.warning("IMAP polling error (reconnecting): %s", exc)
                self.root.after(
                    0, lambda e=exc: self.status_label.config(text=f"Connection retry: {e}")
                )
                mail = None

            time.sleep(3)

    @staticmethod
    def _fetch_gmail_msgid_hex(mail: imaplib.IMAP4_SSL, uid_str: str) -> str:
        try:
            status, data = mail.uid("FETCH", uid_str, "(X-GM-MSGID)")
            if status == "OK" and data and data[0]:
                match = re.search(r"X-GM-MSGID\s+(\d+)", data[0].decode("utf-8", errors="ignore"))
                if match:
                    return format(int(match.group(1)), "x")
        except Exception as exc:
            logger.debug("Failed to parse X-GM-MSGID: %s", exc)
        return ""

    # -----------------------------------------------------------------
    # CONCURRENT PDF FETCHING & MACRO DISPATCH
    # -----------------------------------------------------------------

    def _handle_batch_button(self) -> None:
        if not self.selecting_mode:
            self.selecting_mode = True
            with self.lock:
                for item in self.rows:
                    item.checkbox_widget.pack(side=tk.LEFT, padx=(2, 10), before=item.entry_widget)
            self.batch_btn.config(text="CONFIRM & DISPATCH", bg="#f59e0b")
            self.select_pending_btn.pack(side=tk.TOP, fill="x", pady=(4, 0), ipady=4)
            self.status_label.config(text="Status: Select rows to dispatch, then click CONFIRM.")
            return

        selected_orders: List[str] = []
        selected_rows: List[TrackedOrderRow] = []

        with self.lock:
            for item in self.rows:
                if item.checkbox_var.get() == 1:
                    val = item.target_value.strip() if item.is_active else item.entry_widget.get().strip()
                    if val:
                        selected_orders.append(val)
                        selected_rows.append(item)

        if not selected_orders:
            self.status_label.config(text="Status: No orders selected.")
            return

        self.selecting_mode = False
        self.select_pending_btn.pack_forget()
        with self.lock:
            for item in self.rows:
                item.checkbox_widget.pack_forget()
                item.checkbox_var.set(0)

        self.batch_btn.config(text="Process & Dispatch Batch", bg="#2563eb", state=tk.DISABLED)
        self.status_label.config(
            text=f"Status: Fetching PDFs concurrently for {len(selected_orders)} order(s)..."
        )
        threading.Thread(
            target=self._run_batch_workflow,
            args=(selected_orders, selected_rows),
            daemon=True,
        ).start()

    def _run_batch_workflow(self, orders: List[str], rows_to_update: List[TrackedOrderRow]) -> None:
        batch_folder = os.path.join(BASE_DATA_DIR, f"Batch_{orders[0]}")
        os.makedirs(batch_folder, exist_ok=True)

        def fetch_pdf(order_num: str) -> Optional[Tuple[str, bytes]]:
            try:
                mail = imaplib.IMAP4_SSL("imap.gmail.com")
                mail.login(BATCH_EMAIL, BATCH_APP_PASSWORD)
                mail.select('"[Gmail]/All Mail"', readonly=True)

                status, response = mail.uid("SEARCH", "X-GM-RAW", f'"{order_num}"')
                if status == "OK" and response[0]:
                    latest_uid = response[0].split()[-1]
                    _, fetch_data = mail.uid("FETCH", latest_uid, "(RFC822)")
                    if fetch_data and fetch_data[0]:
                        msg = email.message_from_bytes(fetch_data[0][1])
                        for part in msg.walk():
                            raw_name = part.get_filename()
                            if raw_name or "attachment" in str(part.get("Content-Disposition")):
                                clean_name = DocumentParser.decode_mime_filename(raw_name)
                                if clean_name.lower().endswith(".pdf") and order_num in clean_name:
                                    payload = part.get_payload(decode=True)
                                    if payload:
                                        mail.logout()
                                        return (clean_name, payload)
                mail.logout()
            except Exception as exc:
                logger.error("Failed fetching PDF for %s: %s", order_num, exc)
            return None

        with ThreadPoolExecutor(max_workers=5) as pool:
            fetched_pdfs = [res for res in pool.map(fetch_pdf, orders) if res is not None]

        if not fetched_pdfs:
            self.root.after(
                0,
                lambda: [
                    self.batch_btn.config(state=tk.NORMAL, bg="#2563eb"),
                    self.status_label.config(text="Status: No matching PDF attachments found."),
                ],
            )
            return

        # Group extracted barcodes by branch prefix (1 -> main, 2 -> north, 3 -> south)
        prefix_to_branch = {"1": "branch_main", "2": "branch_north", "3": "branch_south"}
        grouped_exports: Dict[str, str] = {b: "" for b in BRANCH_EXPORT_MAP}

        for filename, payload in fetched_pdfs:
            order_num, barcodes = DocumentParser.extract_barcodes_from_pdf(filename, payload)
            block = (
                f"============================\n"
                f"order number: {order_num}\n"
                f"----------------------------\n"
                + "".join(f"item number:  {code}\n" for code in barcodes)
                + f"============================\n\n"
            )
            branch_key = prefix_to_branch.get(order_num[0], "branch_main")
            grouped_exports[branch_key] += block

        active_branches: List[str] = []
        for branch_key, content in grouped_exports.items():
            if content.strip():
                out_path = os.path.join(batch_folder, BRANCH_EXPORT_MAP[branch_key])
                with open(out_path, "w", encoding="utf-8") as f:
                    f.write(content)
                if self.profiles.get(branch_key, {}).get("Steps"):
                    active_branches.append(branch_key)

        with self.lock:
            for r in rows_to_update:
                r.status_var.set("Sent Status")
                r.update_status_fn("Sent Status")
        self._save_state()

        if active_branches:
            self.root.after(0, self.root.iconify)
            time.sleep(0.8)
            for branch in active_branches:
                self._execute_macro_profile(branch, batch_folder)
                time.sleep(0.4)
            self.root.after(
                0,
                lambda: [self.root.deiconify(), self.root.lift(), self.root.focus_force()],
            )

        self.root.after(
            0,
            lambda: [
                self.batch_btn.config(state=tk.NORMAL, bg="#2563eb"),
                self.status_label.config(text="Status: Batch dispatch complete."),
            ],
        )

    def _execute_macro_profile(self, profile_name: str, batch_folder: str) -> None:
        profile = self.profiles.get(profile_name, {})
        steps = profile.get("Steps", [])
        export_path = os.path.join(batch_folder, BRANCH_EXPORT_MAP.get(profile_name, ""))

        if not steps or not os.path.exists(export_path):
            return

        running = True
        is_looping = bool(profile.get("Loop", False))
        app_keyword = profile.get("AppKeyword", profile_name)

        while running:
            for step in steps:
                if Win32Controller.is_key_pressed(VK_F10):
                    running = False
                    break

                stype = step.get("Type")
                delay = float(step.get("Delay", 1.0))

                if stype == "Mouse":
                    Win32Controller.click_at(
                        int(step["X"]), int(step["Y"]), int(step.get("Clicks", 1))
                    )
                elif stype == "Keyboard":
                    vk = Win32Controller.resolve_vk(step.get("Key", ""))
                    if vk:
                        Win32Controller.tap_key(vk)
                elif stype == "Combo":
                    keys = [
                        Win32Controller.resolve_vk(k)
                        for k in step.get("Combo", "").split("+")
                        if k.strip()
                    ]
                    valid_vks = [vk for vk in keys if vk != 0]
                    for vk in valid_vks:
                        Win32Controller.press_key_down(vk)
                        time.sleep(0.04)
                    time.sleep(0.08)
                    for vk in reversed(valid_vks):
                        Win32Controller.release_key_up(vk)
                        time.sleep(0.04)
                elif stype == "FocusApp":
                    Win32Controller.focus_window_by_keyword(app_keyword)
                    time.sleep(0.25)
                elif stype == "ReadFile":
                    self._inject_items_from_file(export_path, delay)

                if stype != "ReadFile":
                    time.sleep(max(0.05, delay))

            if not running or not is_looping:
                break

    def _inject_items_from_file(self, file_path: str, interval_delay: float) -> None:
        with open(file_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        items = [
            match.group(1)
            for line in lines
            if (match := re.search(r"item\s*number\s*:\s*(\d+)", line, re.IGNORECASE))
        ]

        for idx, item_code in enumerate(items):
            if Win32Controller.is_key_pressed(VK_F10):
                break

            self.root.clipboard_clear()
            self.root.clipboard_append(item_code)
            self.root.update()
            time.sleep(0.08)

            # Ctrl+V then Enter
            Win32Controller.press_key_down(0x11)
            Win32Controller.tap_key(0x56)
            Win32Controller.release_key_up(0x11)
            time.sleep(0.08)
            Win32Controller.tap_key(0x0D)

            if idx < len(items) - 1 and interval_delay > 0:
                time.sleep(interval_delay)

    # -----------------------------------------------------------------
    # COMPLETED CARDS & HISTORY MODAL
    # -----------------------------------------------------------------

    def _render_completed_card(
        self, row_obj: TrackedOrderRow, summary: str, msg_id_hex: str = "", reload_mode: bool = False
    ) -> None:
        if not reload_mode:
            if self.root.state() == "iconic":
                self.root.deiconify()
            self.root.lift()

        with self.lock:
            self.completed_count += 1
        self._update_counters()

        card = tk.Frame(
            self.completed_frame,
            bg="#1e293b",
            highlightbackground="#f43f5e",
            highlightthickness=2,
            padx=12,
            pady=12,
        )
        card.pack(fill="x", pady=6, padx=6)

        with self.lock:
            self.completed_cards.append((card, row_obj))

        tk.Label(
            card,
            text=f"MATCH FOUND:\n{summary}",
            font=("Segoe UI", 11, "bold"),
            fg="#f8fafc",
            bg="#1e293b",
            justify="left",
        ).pack(anchor="w", pady=(0, 10))

        btn_row = tk.Frame(card, bg="#1e293b")
        btn_row.pack(fill="x")

        token = summary.split()[0] if summary else ""
        encoded = urllib.parse.quote(token)
        mail_url = (
            f"https://mail.google.com/mail/?authuser={MONITOR_EMAIL}#search/{encoded}/{msg_id_hex}"
            if msg_id_hex
            else f"https://mail.google.com/mail/?authuser={MONITOR_EMAIL}#search/{encoded}"
        )

        tk.Button(
            btn_row,
            text="Open in Gmail",
            command=lambda: webbrowser.open(mail_url),
            bg="#3b82f6",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=2, fill="x", expand=True, ipady=5)

        def on_keep() -> None:
            card.destroy()
            with self.lock:
                self.completed_count = max(0, self.completed_count - 1)
                self.completed_cards = [c for c in self.completed_cards if c[0] != card]
                row_obj.is_alerted = False
                self.kept_history.append((summary, msg_id_hex))
            self._update_counters()
            self._save_state()

        tk.Button(
            btn_row,
            text="Keep Monitoring",
            command=on_keep,
            bg="#f59e0b",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=2, fill="x", expand=True, ipady=5)

        def on_done() -> None:
            card.destroy()
            with self.lock:
                self.completed_count = max(0, self.completed_count - 1)
                self.completed_cards = [c for c in self.completed_cards if c[0] != card]
            self._update_counters()
            self._remove_row(row_obj)

        tk.Button(
            btn_row,
            text="Done",
            command=on_done,
            bg="#ef4444",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=2, fill="x", expand=True, ipady=5)

    def _clear_all_completed(self) -> None:
        with self.lock:
            for card, row_obj in list(self.completed_cards):
                card.destroy()
                row_obj.is_active = False
                if row_obj in self.rows:
                    self.rows.remove(row_obj)
                row_obj.frame.destroy()
            self.completed_cards.clear()
            self.completed_count = 0

        self._save_state()
        self._update_counters()

    def _open_history_modal(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("Kept Monitoring History")
        win.geometry("650x500")
        win.configure(bg="#0f172a")
        win.grab_set()

        tk.Label(
            win,
            text="Archived Kept Matches",
            font=("Segoe UI", 14, "bold"),
            fg="#38bdf8",
            bg="#0f172a",
        ).pack(anchor="w", padx=15, pady=12)

        with self.lock:
            items = list(self.kept_history)

        if not items:
            tk.Label(
                win, text="No archived items yet.", fg="#94a3b8", bg="#0f172a", font=("Segoe UI", 11)
            ).pack(anchor="w", padx=15)
            return

        for text, _mid in items:
            box = tk.Frame(win, bg="#1e293b", padx=10, pady=8)
            box.pack(fill="x", padx=15, pady=4)
            tk.Label(box, text=text, fg="#f8fafc", bg="#1e293b", font=("Segoe UI", 10, "bold")).pack(
                anchor="w"
            )

    def _open_macro_builder(self) -> None:
        builder = tk.Toplevel(self.root)
        builder.title("ERP Macro Profile Builder")
        builder.geometry("540x420")
        builder.grab_set()

        tk.Label(builder, text="Profile:", font=("Segoe UI", 10, "bold")).pack(pady=(15, 5))
        prof_var = tk.StringVar(value="branch_main")
        tk.OptionMenu(builder, prof_var, *self.profiles.keys()).pack()

        tk.Label(builder, text="Target Window Title Keyword:", font=("Segoe UI", 9)).pack(pady=(12, 2))
        kw_entry = tk.Entry(builder, width=30)
        kw_entry.insert(0, self.profiles["branch_main"]["AppKeyword"])
        kw_entry.pack()

        def save_kw() -> None:
            cur = prof_var.get()
            self.profiles[cur]["AppKeyword"] = kw_entry.get().strip()
            path = os.path.join(MACRO_CONFIG_DIR, f"{cur}.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.profiles[cur], f, indent=2)

        tk.Button(builder, text="Save Profile Settings", command=save_kw, bg="#0ea5e9", fg="white").pack(
            pady=15
        )

    # -----------------------------------------------------------------
    # PERSISTENCE & HELPERS
    # -----------------------------------------------------------------

    def _format_row_summary(self, item: TrackedOrderRow) -> str:
        note = item.note_widget.get().strip()
        if note == "Note...":
            note = ""
        date_val = item.date_widget.get().strip()
        parts = [item.target_value.strip()]
        if note:
            parts.append(note)
        if date_val:
            parts.append(f"[{date_val}]")
        return "  •  ".join(parts)

    def _update_counters(self) -> None:
        with self.lock:
            self.active_label.config(text=f"Active Tracking ({len(self.rows)}):")
            self.completed_label.config(text=f"Matched Emails ({self.completed_count}):")

    def _toggle_sorting(self) -> None:
        self.sort_ascending = not self.sort_ascending
        self.sort_btn.config(
            text="Sort: Old -> New" if self.sort_ascending else "Sort: New -> Old"
        )
        self._reorder_rows()

    def _reorder_rows(self) -> None:
        with self.lock:
            def parse_dt(r: TrackedOrderRow) -> datetime:
                try:
                    return datetime.strptime(r.date_widget.get().strip(), "%d/%m/%Y")
                except ValueError:
                    return datetime.min

            self.rows.sort(key=parse_dt, reverse=not self.sort_ascending)
            for r in self.rows:
                r.frame.pack_forget()
            for r in self.rows:
                r.frame.pack(fill="x", expand=True, pady=6, padx=6)

    def _filter_rows(self, *_args) -> None:
        query = self.filter_var.get().strip().lower()
        with self.lock:
            for r in self.rows:
                num = r.entry_widget.get().strip().lower()
                note = r.note_widget.get().strip().lower()
                if not query or query in num or query in note:
                    r.frame.pack(fill="x", expand=True, pady=6, padx=6)
                else:
                    r.frame.pack_forget()

    def _select_all_pending(self) -> None:
        with self.lock:
            completed_targets = {r.target_value.strip() for _, r in self.completed_cards}
            for r in self.rows:
                val = r.target_value.strip() if r.is_active else r.entry_widget.get().strip()
                r.checkbox_var.set(1 if val and val not in completed_targets else 0)

    def _load_macro_profiles(self) -> None:
        for prof_name in self.profiles:
            path = os.path.join(MACRO_CONFIG_DIR, f"{prof_name}.json")
            if os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        self.profiles[prof_name].update(json.load(f))
                except Exception as exc:
                    logger.warning("Failed loading profile %s: %s", prof_name, exc)

    def _save_state(self) -> None:
        with self.lock:
            targets_payload = []
            for r in self.rows:
                num = r.entry_widget.get().strip()
                note = r.note_widget.get().strip()
                if note == "Note...":
                    note = ""
                date_val = r.date_widget.get().strip()
                if num or note or date_val:
                    targets_payload.append(
                        {
                            "number": num,
                            "note": note,
                            "date": date_val,
                            "status": r.status_var.get(),
                            "is_active": r.is_active,
                            "is_alerted": r.is_alerted,
                            "processed_uids": list(r.processed_uids),
                            "msg_id_hex": r.msg_id_hex,
                        }
                    )
            history_payload = [{"text": t, "msg_id_hex": m} for t, m in self.kept_history]

        try:
            with open(SAVE_FILE, "w", encoding="utf-8") as f:
                json.dump({"targets": targets_payload, "history": history_payload}, f, indent=2)
        except Exception as exc:
            logger.error("Failed persisting state: %s", exc)

    def _load_persisted_state(self) -> None:
        if not os.path.exists(SAVE_FILE):
            self.add_entry_row(auto_save=False)
            return

        try:
            with open(SAVE_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)

            self.kept_history = [
                (h.get("text", ""), h.get("msg_id_hex", "")) for h in raw.get("history", [])
            ]
            for entry in raw.get("targets", []):
                row_obj = self.add_entry_row(
                    initial_val=entry.get("number", ""),
                    initial_note=entry.get("note", ""),
                    initial_date=entry.get("date", ""),
                    initial_status=entry.get("status", "Pending"),
                    auto_save=False,
                )
                if row_obj:
                    row_obj.is_active = entry.get("is_active", False)
                    row_obj.is_alerted = entry.get("is_alerted", False)
                    row_obj.processed_uids = set(entry.get("processed_uids", []))
                    row_obj.msg_id_hex = entry.get("msg_id_hex", "")

            self._reorder_rows()
            self._update_counters()

            for r in self.rows:
                if r.is_alerted:
                    self._render_completed_card(
                        r, self._format_row_summary(r), r.msg_id_hex, reload_mode=True
                    )
        except Exception as exc:
            logger.error("Failed loading saved state: %s", exc)
            self.add_entry_row(auto_save=False)

    def _bind_clipboard_shortcuts(self) -> None:
        def handle_copy(e):
            if isinstance(e.widget, tk.Entry):
                self.root.clipboard_clear()
                self.root.clipboard_append(e.widget.get())
            return "break"

        self.root.bind_all("<Control-c>", handle_copy)

    def _on_close(self) -> None:
        self._save_state()
        self.is_running = False
        self.root.destroy()


if __name__ == "__main__":
    app_root = tk.Tk()
    OrderMailAutomationApp(app_root)
    app_root.mainloop()