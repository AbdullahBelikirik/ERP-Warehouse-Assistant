# 📦 Shop & Warehouse Assistant (ERP-Integrated Mobile Automation)

A lightweight, mobile-first inventory assistant and invoice preparation system integrated with legacy **Akınsoft Wolvox ERP (Firebird SQL)**.

This application bridges on-field warehouse stock operations (barcode scanning, price lookup, invoice collection) and desk-level ERP invoice imports by generating standardized, ready-to-import OpenXML (`.xlsx`) datasets on the fly — and, optionally, by driving the ERP's own screens to import and print the finished invoice.

---

## 🎯 The Problem (Field Bottlenecks)

In traditional wholesale & retail warehouse operations:

1. **Barcode vs. Non-Barcode Chaos:** Products without barcodes require manual paperwork or split Excel sheets, causing invoice fragmentation and double entries during ERP batch imports.
2. **Slow Legacy ERP Lookups:** Staff frequently walk back to desktop terminals just to check recent customer-specific sale prices, balances or stock codes.
3. **Data Inconsistency:** Typing typos in customer titles or non-normalized Turkish characters (`i` vs. `İ`) break database search queries.
4. **Unit vs. Box Price Mistakes:** The same product is sold per piece and per box, so a wrong historical price silently produces a wrong invoice.
5. **Heavy Dependencies:** Running heavy spreadsheet engines (pandas/openpyxl) on constrained terminal servers slows down low-latency barcode operations.

---

## 💡 The Solution & Architecture

A localized web application built with **Flask**, **Vanilla JavaScript**, and **Tailwind CSS**, communicating with a **read-only** Firebird SQL connection and utilizing a custom zero-dependency **OpenXML Zip/XML engine**.

```
┌─────────────────────────────────────────────────────────────┐
│                 Mobile Device / Barcode PWA                 │
│  (Html5-Qrcode Camera / Torch Control / Web Audio Haptics)  │
└──────────────────────────────┬──────────────────────────────┘
                               │ HTTP JSON APIs (PIN-gated token)
                               ▼
┌─────────────────────────────────────────────────────────────┐
│                   Flask Application Server                  │
│   - Turkish Char Normalization (to_turkish_upper)           │
│   - 13-digit EAN Barcode Resolution (main + extra barcodes) │
│   - Smart Price Resolution + 6x Deviation Safety Check      │
│   - Customer Balance & Statement (CARI / CARIHR)            │
└──────────────┬──────────────────────────────┬───────────────┘
               │                              │
      Read-Only SQL (FDB)            Direct XML Generation
               │                              │
               ▼                              ▼
┌─────────────────────────────┐   ┌───────────────────────────┐
│  Akınsoft Wolvox Database   │   │ Auto-Merged Excel Files   │
│      (Firebird 2.5/3.0)     │   │  STOKKODU|MIKTARI|FIYATI  │
└─────────────────────────────┘   └─────────────┬─────────────┘
                                                │ (optional, Windows only)
                                                ▼
                                  ┌───────────────────────────┐
                                  │ Wolvox UI Automation      │
                                  │ import → VAT → save → print│
                                  └───────────────────────────┘
```

---

## ✨ Key Features

- **🔐 PIN-Gated Devices:** Each phone authorizes once with a staff PIN and receives a device token; every API call is token-protected.
- **📷 Hardware-Accelerated Camera Barcode Scanner:**
  * Real-time 8–14 digit validation preventing partial scans.
  * Flashlight / Torch toggle via MediaTrack constraints.
  * Web Audio API beep + Vibration (Haptic feedback) on successful reads.
  * Also accepts hardware (USB/Bluetooth) scanners through a global key listener.
- **🔍 Unified Fast Search & Turkish Normalization:**
  * Custom `to_turkish_upper` mapping solving standard Python `.upper()` `i` -> `I` mismatch against Turkish collation (`İ`).
  * Instant search across `STOK` (Stock Names/Codes) and `CARI` (Customer Trade Titles & Full Names).
- **💰 Smart Price Lookup:**
  * Combines the latest purchase (`AF`) and last three sales (`SFT`) from `STOKHR`.
  * **6x deviation guard:** if prices differ by 6x or more (piece vs. box), the user must pick the correct unit price before the item is added.
- **⚡ Native OpenXML Spreadsheet Engine:**
  * Uses Python's native `zipfile` and `xml.etree.ElementTree` to read/write 3-column `.xlsx` files (`STOKKODU`, `MIKTARI`, `FIYATI`).
  * Zero memory bloat, writes to `Desktop/fatura_listeleri/`.
- **🛒 Live Invoice Cart & Counter:**
  * Real-time item tallying and totals, inline quantity edit, one-tap **price edit mode**, undo-last, instant row deletion.
  * The quantity box remembers its value across consecutive scans.
- **👤 Customer Balance & Statement:** Net balance, debit/credit totals, tax ID and the last 50 transactions, straight from `CARIHR`.
- **🖨️ One-Tap ERP Import & Print (Windows only):** Automates the Wolvox desktop UI to import the Excel list, apply the VAT step, save and print the invoice, then archives the used file. Can also print a customer transaction report.
- **🏷️ Duplicate Barcode Guard:** Prevents registering barcodes already assigned in primary (`STOK.BARKODU`) or secondary (`STOK_BARKOD.BARKODU`) tables. New barcodes are queued in `eklenecekler.txt` for the ERP operator.

---

## 🚀 Tech Stack

- **Backend:** Python 3.10+, Flask, `fdb` (Firebird Driver), `python-dotenv`, `pyperclip`
- **Frontend:** Vanilla JavaScript (ES6+), Tailwind CSS, Html5-Qrcode
- **Database:** Firebird SQL (Akınsoft Wolvox ERP Backend)
- **Data Interchange:** OpenXML (`.xlsx`), REST JSON

---

## 🛠️ Installation & Setup

### 1. Clone the Repository

```
git clone https://github.com/AbdullahBelikirik/ERP-Warehouse-Assistant.git
cd ERP-Warehouse-Assistant
```

### 2. Set Up Virtual Environment

```
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux/macOS:
source venv/bin/activate

pip install -r requirements.txt
```

### 3. Configure Environment Variables

Copy `.env.example` to `.env` and fill in your values. The app refuses to start if a required secret is missing.

```
cp .env.example .env
```

| Variable | Required | Description |
| --- | --- | --- |
| `APP_PIN` | ✅ | Staff PIN entered once per device |
| `AUTH_TOKEN_SECRET` | ✅ | Long random string used as device token (`python -c "import secrets; print(secrets.token_hex(32))"`) |
| `DB_PATH` | ✅ | Full path of the Wolvox `.fdb` file |
| `DB_PASSWORD` | ✅ | Firebird user password |
| `DB_HOST` / `DB_PORT` / `DB_USER` / `DB_CHARSET` | – | Defaults: `127.0.0.1` / `3050` / `SYSDBA` / `WIN1254` |
| `APP_BASE_DIR` | – | Where lists, logs and `eklenecekler.txt` live (default: Desktop) |
| `PORT` | – | Web server port (default `5000`) |

### 4. Run the Server

```
python app.py
```

Access the interface via your local network (e.g., `http://192.168.1.X:5000`) or through an encrypted tunnel like `ngrok`.

---

## 🔌 API Overview

All endpoints except `/` and `/api/auth/verify-pin` require the `X-Shop-Token` header. Responses use a `status` field (`success`, `error`, `not_found`, ...) and a human-readable `message`.

| Endpoint | Purpose |
| --- | --- |
| `POST /api/auth/verify-pin` | Exchange the staff PIN for a device token |
| `POST /api/price/query` | Barcode / name lookup with price history |
| `GET /api/price/detail/<blkodu>` | Price detail for a chosen product |
| `POST /api/products/search` | Product search (name or stock code) |
| `POST /api/barcodes/resolve` | Barcode → stock code |
| `POST /api/barcodes/add` | Queue a new barcode in `eklenecekler.txt` |
| `POST /api/customers/search` | Customer search |
| `POST /api/customers/balance` | Balance + last 50 transactions |
| `POST /api/invoices/items` | List items of a customer's draft invoice |
| `POST /api/invoices/items/add` | Add an item (handles the 6x price deviation flow) |
| `POST /api/invoices/items/update` | Change quantity/price or delete an item |
| `POST /api/invoices/print` | Import the list into Wolvox, print it and archive the file (Windows) |
| `POST /api/customers/report/print` | Print the customer transaction report (Windows) |

---

## 🔒 Security & ERP Data Integrity Note

- **Read-Only Database Access:** The app executes only `SELECT` queries on the Firebird database, so core ERP transactional records are never written by SQL. For defense in depth, create a dedicated read-only Firebird user and use it as `DB_USER`.
- **File-Based Bridge:** Finalized invoice lines are compiled into standardized `STOKKODU` / `MIKTARI` / `FIYATI` Excel sheets, allowing ERP operators to review and import batches natively through the standard Wolvox Import Wizard.
- **Secrets stay out of the code:** PIN, device token secret and DB password come only from `.env` (which is git-ignored). Never commit a real `.env`.
- **UI automation caveat:** `/api/invoices/print` and `/api/customers/report/print` simulate mouse clicks at fixed screen coordinates (`COORD_*` constants in `app.py`) on the shop PC. They depend on screen resolution, window layout and a single running ERP window; recalibrate the coordinates for another machine and avoid triggering two prints at once.
- If you expose the app through a public tunnel, keep the PIN private and consider rate limiting.

---

## 📄 License

MIT License. Created for production warehouse automation & retail inventory workflow optimization.
