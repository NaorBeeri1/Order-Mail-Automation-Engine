# Order Mail Monitor & Win32 ERP Automation Engine

A multi-threaded desktop automation application engineered in Python to eliminate manual order tracking and legacy ERP data entry.

## Architecture & Key Capabilities

1. **Real-Time IMAP Subject Polling (`imaplib`)**
   - Monitors incoming email queues via `UID SEARCH` and extracts Gmail `X-GM-MSGID` hex identifiers for direct one-click browser deep-linking.
   - Thread-safe state management (`threading.Lock`) with atomic JSON persistence across shared multi-user workspaces.

2. **Concurrent PDF Rasterization & Barcode Extraction (`ThreadPoolExecutor` + `PyMuPDF` + `pyzbar`)**
   - Fetches matching order PDFs concurrently across worker threads.
   - Renders PDF pages in memory (`150 DPI`) and decodes embedded 1D/2D barcodes, falling back to regex pattern matching when raster scans are unavailable.

3. **Native Win32 API Macro Dispatcher (`ctypes.windll.user32`)**
   - Focuses target desktop ERP windows dynamically via `EnumWindows` and `SetForegroundWindow`.
   - Executes configurable mouse, keyboard, and clipboard injection sequences with hardware-level `F10` emergency kill-switch support.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
