# 📦 Shop & Warehouse Assistant (ERP-Integrated Mobile Automation)

A lightweight, mobile-first inventory assistant and invoice preparation system integrated with legacy **Akınsoft Wolvox ERP (Firebird SQL)**. 

This application bridges on-field warehouse stock operations (barcode scanning, price lookup, invoice collection) and desk-level ERP invoice imports by generating standardized, ready-to-import OpenXML (`.xlsx`) datasets directly on the fly.

---

## 🎯 The Problem (Field Bottlenecks)

In traditional wholesale & retail warehouse operations:
1. **Barcode vs. Non-Barcode Chaos:** Products without barcodes require manual paperwork or split Excel sheets, causing invoice fragmentation and double entries during ERP batch imports.
2. **Slow Legacy ERP Lookups:** Staff frequently walk back to desktop terminals just to check recent customer-specific sale prices or verify stock codes.
3. **Data Inconsistency:** Typing typos in customer titles or non-normalized Turkish characters (`i` vs. `İ`) break database search queries.
4. **Heavy Dependencies:** Running heavy spreadsheet engines (pandas/openpyxl) on constrained terminal servers slows down low-latency barcode operations.

---

## 💡 The Solution & Architecture

A localized web application built with **Flask**, **Vanilla JavaScript**, and **Tailwind CSS**, communicating with a read-only **Firebird SQL** database and utilizing a custom zero-dependency **OpenXML Zip/XML streaming engine**.

```
  ┌─────────────────────────────────────────────────────────────┐
  │                 Mobile Device / Barcode PWA                 │
  │  (Html5-Qrcode Camera / Torch Control / Web Audio Haptics)  │
  └──────────────────────────────┬──────────────────────────────┘
                                 │ HTTP JSON APIs
                                 ▼
  ┌─────────────────────────────────────────────────────────────┐
  │                   Flask Application Server                  │
  │   - Turkish Char Normalization (to_turkish_upper)           │
  │   - Intelligent 13-digit EAN Barcode Verification           │
  │   - Customer Unification Engine (ADI_SOYADI + UNVAN)        │
  └──────────────┬──────────────────────────────┬───────────────┘
                 │                              │
        Read-Only SQL (FDB)            Direct XML Generation
                 │                              │
                 ▼                              ▼
  ┌─────────────────────────────┐   ┌───────────────────────────┐
  │  Akınsoft Wolvox Database   │   │ Auto-Merged Excel Files   │
  │      (Firebird 2.5/3.0)     │   │     (OpenXML / .xlsx)     │
  └─────────────────────────────┘   └───────────────────────────┘
```

---

## ✨ Key Features

* **📷 Hardware-Accelerated Camera Barcode Scanner:**
  * Real-time 13-digit EAN validation preventing partial scans.
  * Flashlight / Torch toggle via MediaTrack constraints.
  * Web Audio API beep synthesizer + Vibration (Haptic feedback) on successful reads.
* **🔍 Unified Fast Search & Turkish Normalization:**
  * Custom `to_turkish_upper` mapping solving standard Python `.upper()` `i` -> `I` mismatch against Turkish collation (`İ`).
  * Instant search across `STOK` (Stock Names/Codes) and `CARI` (Customer Trade Titles & Full Names).
* **⚡ Native OpenXML Spreadsheet Engine:**
  * Uses Python's native `zipfile` and `xml.etree.ElementTree` to read/write `.xlsx` files directly.
  * Zero memory bloat, high-speed atomic writes to `Desktop/fatura_listeleri/`.
* **🛒 Live Invoice Cart & Counter:**
  * Real-time item tallying, inline quantity increments/decrements, and instant row deletion without opening Excel files manually.
* **🏷️ Duplicate Barcode Guard:**
  * Prevents registering barcodes already assigned in primary (`STOK.BARKODU`) or secondary (`STOK_BARKOD.BARKODU`) tables.

---

## 🚀 Tech Stack

* **Backend:** Python 3.10+, Flask, `fdb` (Firebird Driver)
* **Frontend:** Vanilla JavaScript (ES6+), Tailwind CSS, Html5-Qrcode
* **Database:** Firebird SQL (Akınsoft Wolvox ERP Backend)
* **Data Interchange:** OpenXML (`.xlsx`), REST JSON

---

## 🛠️ Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/your-username/warehouse-erp-assistant.git
cd warehouse-erp-assistant
```

### 2. Set Up Virtual Environment
```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux/macOS:
source venv/bin/activate

pip install -r requirements.txt
```

### 3. Configure Environment Variables
Copy `.env.example` to `.env` and fill in your connection details:
```bash
cp .env.example .env
```

```env
PORT=5000
DB_HOST=127.0.0.1
DB_PORT=3050
DB_PATH=C:\AKINSOFT\Wolvox8\Database_FB\100\2026\wolvox.fdb
DB_USER=SYSDBA
DB_PASSWORD=your_password
DB_CHARSET=WIN1254
```

### 4. Run the Server
```bash
python app.py
```

Access the interface via your local network (e.g., `http://192.168.1.X:5000` or through an encrypted tunnel like `ngrok`).

---

## 🔒 Security & ERP Data Integrity Note

This system adheres to strict ERP data integrity rules:
* **Read-Only Database Access:** It executes only `SELECT` queries on the Firebird database, ensuring core ERP transactional records remain untouched.
* **File-Based Bridge:** Finalized invoice lines are compiled into standardized `STOKKODU` / `MIKTARI` Excel sheets on the host desktop, allowing ERP operators to review and import batches natively via standard Wolvox Import Wizard.

---

## 📄 License
MIT License. Created for production warehouse automation & retail inventory workflow optimization.
