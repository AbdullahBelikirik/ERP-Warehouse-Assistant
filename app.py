import hmac
import os
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, jsonify
import fdb
from dotenv import load_dotenv
import zipfile
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils
import pyperclip
import ctypes
import time
import re
import threading

app = Flask(__name__)

load_dotenv()


def require_env(name):
    value = os.getenv(name)
    if not value:
        raise SystemExit(f"Missing required setting {name}. Copy .env.example to .env and fill it in.")
    return value


# --- INVOICE SAFETY LIMITS ---
MAX_ITEM_QUANTITY = 100000   # Highest quantity for one line (rejects a barcode typed into the quantity box)
UNPRICED_THRESHOLD = 0.011   # Prices at or below this count as "not set" (never sent to Akinsoft)

# --- SECURITY & PIN CONFIGURATION (never hard-code these) ---
STAFF_PIN = require_env("APP_PIN")
AUTH_TOKEN_SECRET = require_env("AUTH_TOKEN_SECRET")

# --- PATHS ---
BASE_DIR = os.getenv("APP_BASE_DIR", os.path.join(os.path.expanduser("~"), "Desktop"))
INVOICE_OUTPUT_DIR = os.path.join(BASE_DIR, "fatura_listeleri")
PENDING_BARCODES_FILE = os.path.join(BASE_DIR, "eklenecekler.txt")

if not os.path.exists(INVOICE_OUTPUT_DIR):
    os.makedirs(INVOICE_OUTPUT_DIR)
    print(f"Invoice output folder created: {INVOICE_OUTPUT_DIR}")

# --- DATABASE (Firebird / Akinsoft Wolvox) ---
DB_HOST = os.getenv("DB_HOST", "127.0.0.1")
DB_PORT = int(os.getenv("DB_PORT", 3050))
DB_PATH = require_env("DB_PATH")
DB_USER = os.getenv("DB_USER", "SYSDBA")
DB_PASSWORD = require_env("DB_PASSWORD")
DB_CHARSET = os.getenv("DB_CHARSET", "WIN1254")


def to_turkish_upper(text):
    if not text:
        return ""
    char_map = {'i': 'İ', 'ı': 'I', 'ş': 'Ş', 'ğ': 'Ğ', 'ü': 'Ü', 'ö': 'Ö', 'ç': 'Ç'}
    for k, v in char_map.items():
        text = text.replace(k, v)
    return text.upper()

# Non-blocking: a fresh DB connection for every request
def get_db_connection():
    try:
        return fdb.connect(
            host=DB_HOST,
            port=DB_PORT,
            database=DB_PATH,
            user=DB_USER,
            password=DB_PASSWORD,
            charset=DB_CHARSET
        )
    except Exception as e:
        append_log(f"[DB BAGLANTI HATA] {str(e)}")
        print("Database connection error:", str(e))
        return None

def append_log(message):
    try:
        today = datetime.now().strftime("%Y-%m-%d")
        current_time = datetime.now().strftime("%H:%M:%S")
        log_file_path = os.path.join(INVOICE_OUTPUT_DIR, f"gunluk_log_{today}.txt")
        log_line = f"[{current_time}] {message}\n"
        with open(log_file_path, "a", encoding="utf-8") as f:
            f.write(log_line)
    except Exception as e:
        print("Log write error:", str(e))

def auth_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        token = request.headers.get('X-Shop-Token', '')
        if not hmac.compare_digest(token.encode(), AUTH_TOKEN_SECRET.encode()):
            return jsonify({"status": "unauthorized", "message": "Geçersiz veya eksik yetki tokenı! Lütfen PIN giriniz."}), 401
        return f(*args, **kwargs)
    return decorated_function

# Flows that drive Akinsoft with the mouse/keyboard (invoice, report, collection/payment) must
# never run at the same time: the clicks would mix. Returns 409 while one is running.
AUTOMATION_LOCK = threading.Lock()

def automation_lock_required(f):
    @wraps(f)
    def locked_function(*args, **kwargs):
        if not AUTOMATION_LOCK.acquire(blocking=False):
            return jsonify({
                "status": "error",
                "message": "Akınsoft şu an başka bir işlemle meşgul (fatura, rapor veya tahsilat/tediye çalışıyor). İşlem bitince tekrar dene."
            }), 409
        try:
            return f(*args, **kwargs)
        finally:
            AUTOMATION_LOCK.release()
    return locked_function

# --- 3-COLUMN EXCEL ENGINE (STOKKODU | MIKTARI | FIYATI) ---
def read_invoice_excel(xlsx_path):
    existing_records = []
    if not os.path.exists(xlsx_path):
        return existing_records
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for elem in tree.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                    shared_strings.append(elem.text or '')

            sheet_xml = z.read('xl/worksheets/sheet1.xml')
            tree = ET.fromstring(sheet_xml)
            rows = tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
            
            for r in rows[1:]:
                cells = r.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
                if len(cells) >= 2:
                    code_elem = cells[0].find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    code_val = ""
                    if code_elem is not None and code_elem.text:
                        if cells[0].get('t') == 's':
                            idx = int(code_elem.text)
                            if idx < len(shared_strings):
                                code_val = shared_strings[idx]
                        else:
                            code_val = code_elem.text

                    qty_elem = cells[1].find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    qty_val = 1
                    if qty_elem is not None and qty_elem.text:
                        try:
                            qty_val = int(float(qty_elem.text))
                        except:
                            qty_val = 1

                    price_value = 0.01
                    if len(cells) >= 3:
                        price_elem = cells[2].find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        if price_elem is not None and price_elem.text:
                            try:
                                price_value = float(price_elem.text)
                            except:
                                price_value = 0.01

                    if code_val:
                        existing_records.append([str(code_val).strip(), qty_val, price_value])
    except Exception as e:
        print("Excel okuma hatası:", str(e))
    return existing_records

def write_invoice_excel(xlsx_path, records):
    strings = ["STOKKODU", "MIKTARI", "FIYATI"]
    for record in records:
        code = str(record[0]).strip()
        if code not in strings:
            strings.append(code)

    total_string_count = 3 + len(records)

    sst_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    sst_xml += f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{total_string_count}" uniqueCount="{len(strings)}">'
    for s in strings:
        safe_string = saxutils.escape(str(s))
        sst_xml += f'<si><t>{safe_string}</t></si>'
    sst_xml += '</sst>'

    sheet_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    sheet_xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    sheet_xml += '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>'

    row_num = 2
    for record in records:
        code = str(record[0]).strip()
        qty = record[1]
        price = record[2] if len(record) >= 3 else 0.01
        if float(price) <= 0:
            price = 0.01

        code_idx = strings.index(code)
        sheet_xml += f'<row r="{row_num}">'
        sheet_xml += f'<c r="A{row_num}" t="s"><v>{code_idx}</v></c>'
        sheet_xml += f'<c r="B{row_num}"><v>{qty}</v></c>'
        sheet_xml += f'<c r="C{row_num}"><v>{float(price):.2f}</v></c>'
        sheet_xml += '</row>'
        row_num += 1

    sheet_xml += '</sheetData></worksheet>'

    content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    wb_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>'
    wb_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'

    with zipfile.ZipFile(xlsx_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('xl/workbook.xml', wb_xml)
        z.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        z.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        z.writestr('xl/sharedStrings.xml', sst_xml)

@app.route('/')
def index():
    return render_template('index.html')

@app.after_request
def add_cors_and_tunnel_headers(response):
    response.headers['ngrok-skip-browser-warning'] = 'true'
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = '*'
    return response

@app.route('/manifest.json')
def manifest():
    return jsonify({
        "name": "Dükkan Asistanı",
        "short_name": "Asistan",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#f3f4f6",
        "theme_color": "#2563eb",
        "icons": []
    })

@app.route('/api/auth/verify-pin', methods=['POST'])
def verify_pin():
    payload = request.json or {}
    entered_pin = str(payload.get('pin', '')).strip()

    if hmac.compare_digest(entered_pin.encode(), STAFF_PIN.encode()):
        append_log("[GÜVENLİK] Başarılı PIN doğrulaması yapıldı.")
        return jsonify({
            "status": "success",
            "token": AUTH_TOKEN_SECRET,
            "message": "Giriş başarılı."
        })
    else:
        append_log("[GÜVENLİK UYARISI] Hatalı PIN denemesi.")
        return jsonify({
            "status": "error",
            "message": "Hatalı PIN Kodu! Lütfen tekrar deneyin."
        }), 401

# --- SMART PRICE RESOLUTION & 6x DEVIATION SAFETY CHECK ---
def find_price_by_blkodu(conn, blkodu):
    cursor = conn.cursor()
    
    sale_query = """
        SELECT FIRST 3 H.KPB_FIYATI, H.EVRAK_NO, H.TARIHI
        FROM STOKHR H
        WHERE H.BLSTKODU = ?
        AND (UPPER(H.EVRAK_NO) LIKE 'SFT%')
        ORDER BY H.TARIHI DESC
    """
    cursor.execute(sale_query, (blkodu,))
    sale_records = cursor.fetchall()

    purchase_query = """
        SELECT FIRST 1 H.KPB_FIYATI, H.EVRAK_NO, H.TARIHI
        FROM STOKHR H
        WHERE H.BLSTKODU = ?
        AND (UPPER(H.EVRAK_NO) LIKE 'AF%')
        ORDER BY H.TARIHI DESC
    """
    cursor.execute(purchase_query, (blkodu,))
    purchase_record = cursor.fetchone()

    price_pool = []
    detail_parts = []

    if purchase_record and purchase_record[0] is not None:
        purchase_price = float(purchase_record[0])
        purchase_date = str(purchase_record[2])[:10] if purchase_record[2] else ""
        if purchase_price > 0:
            price_pool.append((purchase_price, purchase_date, "AF"))
            detail_parts.append(f"Son Alış (AF): {purchase_price:.2f} TL ({purchase_date})")

    sale_price_texts = []
    for s in sale_records:
        if s and s[0] is not None:
            sale_price = float(s[0])
            sale_date = str(s[2])[:10] if s[2] else ""
            if sale_price > 0:
                price_pool.append((sale_price, sale_date, "SFT"))
                sale_price_texts.append(f"{sale_price:.2f} TL ({sale_date})")

    if sale_price_texts:
        detail_parts.append(f"Son Satışlar (SFT): {' | '.join(sale_price_texts)}")

    if not price_pool:
        return 0.0, "", "Kayıt bulunamadı", None

    detail_text = " • ".join(detail_parts)
    price_values = [item[0] for item in price_pool]
    min_price = min(price_values)
    max_price = max(price_values)

    deviation_obj = None
    if min_price > 0 and (max_price / min_price) >= 6.0:
        deviation_obj = {
            "detected": True,
            "min_price": min_price,
            "max_price": max_price,
            "ratio": round(max_price / min_price, 1),
            "options": [
                {
                    "label": f"Düşük Fiyat (Muhtemel Adet): {min_price:.2f} TL",
                    "price": min_price,
                    "description": "Geçmiş hareketlerdeki en düşük birim fiyat"
                },
                {
                    "label": f"Yüksek Fiyat (Muhtemel Koli): {max_price:.2f} TL",
                    "price": max_price,
                    "description": "Geçmiş hareketlerdeki en yüksek birim fiyat"
                }
            ]
        }

    highest_price = max_price
    matching_prices = [item for item in price_pool if item[0] == highest_price]
    latest_date = matching_prices[0][1] if matching_prices else ""

    return highest_price, latest_date, detail_text, deviation_obj

def format_single_price_response(conn, blkodu, product_name, stock_code="", barcode=""):
    price, date_str, detail, deviation = find_price_by_blkodu(conn, blkodu)
    price_list = []
    if price > 0:
        price_list.append({"price": price, "date": date_str, "detail": detail})

    append_log(f"[SORGULA: TEK] BLKODU:{blkodu} | {stock_code} -> {product_name} | Fiyat: {price:.2f} TL")
    return jsonify({
        "status": "single_result",
        "blkodu": blkodu,
        "stock_code": str(stock_code).strip(),
        "barcode": barcode,
        "product_name": product_name,
        "price_deviation": deviation,
        "records": price_list
    })

@app.route('/api/price/query', methods=['POST'])
@auth_required
def query_price():
    payload = request.json or {}
    search_text = str(payload.get('barcode', '')).strip()
    
    if not search_text:
        return jsonify({"status": "error", "message": "Arama metni boş olamaz!"}), 400

    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanına bağlanılamıyor!"}), 500

    try:
        cursor = conn.cursor()
        
        barcode_query = """
            SELECT S.BLKODU, TRIM(S.STOKKODU), S.STOK_ADI, S.BARKODU
            FROM STOK S
            WHERE S.AKTIF = 1 AND S.BARKODU = ?
        """
        cursor.execute(barcode_query, (search_text,))
        barcode_results = cursor.fetchall()
        
        if not barcode_results:
            extra_barcode_query = """
                SELECT S.BLKODU, TRIM(S.STOKKODU), S.STOK_ADI, SB.BARKODU
                FROM STOK_BARKOD SB 
                INNER JOIN STOK S ON SB.BLSTKODU = S.BLKODU 
                WHERE S.AKTIF = 1 AND SB.BARKODU = ?
            """
            cursor.execute(extra_barcode_query, (search_text,))
            barcode_results = cursor.fetchall()

        if barcode_results:
            blkodu = barcode_results[0][0]
            stock_code = str(barcode_results[0][1]).strip() if barcode_results[0][1] else ""
            product_name = barcode_results[0][2] or "İsimsiz Ürün"
            barcode_value = barcode_results[0][3] or ""
            return format_single_price_response(conn, blkodu, product_name, stock_code, barcode_value)

        like_pattern = f"%{to_turkish_upper(search_text)}%"
        general_search_query = """
            SELECT FIRST 80 S.BLKODU, TRIM(S.STOKKODU), S.STOK_ADI, S.BARKODU
            FROM STOK S
            WHERE S.AKTIF = 1
            AND (UPPER(S.STOK_ADI) LIKE ? OR UPPER(S.STOKKODU) LIKE ?)
            ORDER BY S.STOK_ADI
        """
        cursor.execute(general_search_query, (like_pattern, like_pattern))
        matches = cursor.fetchall()

        if not matches:
            append_log(f"[SORGULA: BULUNAMADI] {search_text}")
            return jsonify({"status": "not_found", "message": f"'{search_text}' için aktif ürün bulunamadı."})

        if len(matches) == 1:
            blkodu = matches[0][0]
            stock_code = str(matches[0][1]).strip() if matches[0][1] else ""
            product_name = matches[0][2] or "İsimsiz Ürün"
            barcode_value = matches[0][3] or ""
            return format_single_price_response(conn, blkodu, product_name, stock_code, barcode_value)

        product_list = []
        for row in matches:
            product_list.append({
                "blkodu": row[0],
                "stock_code": str(row[1]).strip() if row[1] else "",
                "product_name": row[2] if row[2] else "İsimsiz Ürün",
                "barcode": row[3] if row[3] else ""
            })

        append_log(f"[SORGULA: ÇOKLU] '{search_text}' -> {len(product_list)} adet sonuç")
        return jsonify({"status": "multiple_results", "products": product_list})

    except Exception as e:
        append_log(f"[HATA] Sorgulama Hatası ({search_text}): {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        try: conn.close()
        except Exception: pass

@app.route('/api/price/detail/<int:blkodu>', methods=['GET'])
@auth_required
def get_price_detail(blkodu):
    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanı bağlantısı yok!"}), 500
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT TRIM(STOKKODU), STOK_ADI, BARKODU FROM STOK WHERE BLKODU = ?", (blkodu,))
        match = cursor.fetchone()
        stock_code = str(match[0]).strip() if match and match[0] else ""
        product_name = match[1] if match and match[1] else "Ürün"
        barcode = match[2] if match and match[2] else ""
        return format_single_price_response(conn, blkodu, product_name, stock_code, barcode)
    except Exception as e:
        append_log(f"[HATA] Fiyat Detay Hatası (BLKODU:{blkodu}): {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        try: conn.close()
        except Exception: pass

# --- CUSTOMER SEARCH (length-tolerant, non-blocking) ---
@app.route('/api/customers/search', methods=['POST'])
@auth_required
def search_customers():
    payload = request.json or {}
    search_text = str(payload.get('text', '')).strip()
    if not search_text:
        return jsonify({"status": "empty", "customers": []})

    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanı bağlantısı yok!"}), 500

    try:
        cursor = conn.cursor()
        search_upper = to_turkish_upper(search_text)
        
        # Guard against Firebird VARCHAR(30) CARIKODU overflow
        code_param = search_upper[:30]
        like_title = f"%{search_upper}%"

        sql_query = """
            SELECT FIRST 30 TRIM(C.CARIKODU), C.TICARI_UNVANI, C.ADI_SOYADI
            FROM CARI C
            WHERE C.AKTIF = 1
            AND (UPPER(TRIM(C.CARIKODU)) = ? 
                OR UPPER(COALESCE(C.TICARI_UNVANI, '')) LIKE ? 
                OR UPPER(COALESCE(C.ADI_SOYADI, '')) LIKE ?)
            ORDER BY C.TICARI_UNVANI
        """
        cursor.execute(sql_query, (code_param, like_title, like_title))
        matches = cursor.fetchall()

        customer_list = []
        for row in matches:
            customer_code = str(row[0]).strip() if row[0] else ""
            trade_title = str(row[1]).strip() if row[1] else ""
            full_name = str(row[2]).strip() if row[2] else ""

            if trade_title and full_name and trade_title != full_name:
                display_title = f"{trade_title} ({full_name})"
                excel_title = trade_title
            elif trade_title:
                display_title = trade_title
                excel_title = trade_title
            elif full_name:
                display_title = full_name
                excel_title = full_name
            else:
                display_title = "İsimsiz Cari"
                excel_title = customer_code

            customer_list.append({"customer_code": customer_code, "title": excel_title, "display_name": display_title})

        return jsonify({"status": "success", "customers": customer_list})
    except Exception as e:
        append_log(f"[CARI ARA HATA] {str(e)}")
        print("Cari Ara Hata:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        try: conn.close()
        except Exception: pass

# --- CUSTOMER BALANCE & STATEMENT (length-tolerant, non-blocking) ---
@app.route('/api/customers/balance', methods=['POST'])
@auth_required
def get_customer_balance():
    payload = request.json or {}
    search_query = str(payload.get('query', '')).strip()

    if not search_query:
        return jsonify({"status": "empty", "message": "Cari adı veya kodu boş."})

    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanı bağlantısı yok!"}), 500

    try:
        cursor = conn.cursor()
        
        # Guard against Firebird VARCHAR(30) CARIKODU overflow
        search_upper = to_turkish_upper(search_query)
        code_param = search_upper[:30]
        like_title = f"%{search_upper}%"

        # 1. Find the customer BLKODU, title and national/tax ID
        customer_query = """
            SELECT FIRST 1 C.BLKODU, TRIM(C.CARIKODU), C.TICARI_UNVANI, C.ADI_SOYADI, C.TC_KIMLIK_NO, C.VERGI_NO
            FROM CARI C
            WHERE C.AKTIF = 1
            AND (UPPER(TRIM(C.CARIKODU)) = ? 
                 OR UPPER(C.TICARI_UNVANI) LIKE ? 
                 OR UPPER(C.ADI_SOYADI) LIKE ?)
            ORDER BY C.TICARI_UNVANI
        """
        cursor.execute(customer_query, (code_param, like_title, like_title))
        customer_row = cursor.fetchone()

        if not customer_row:
            return jsonify({"status": "not_found", "message": "Cari kaydı bulunamadı."})

        customer_blkodu = customer_row[0]
        customer_code = str(customer_row[1]).strip() if customer_row[1] else ""
        title = customer_row[2] or customer_row[3] or customer_code
        national_id = str(customer_row[4]).strip() if customer_row[4] else ""
        tax_number = str(customer_row[5]).strip() if customer_row[5] else ""
        tax_id = national_id or tax_number or ""

        # 2. Net balance from CARIHR (KPB_BTUT - KPB_ATUT)
        balance_query = """
            SELECT 
                COALESCE(SUM(H.KPB_BTUT), 0) AS TOPLAM_BORC,
                COALESCE(SUM(H.KPB_ATUT), 0) AS TOPLAM_ALACAK
            FROM CARIHR H
            WHERE H.BLCRKODU = ? 
            AND (H.SILINDI = 0 OR H.SILINDI IS NULL)
        """
        cursor.execute(balance_query, (customer_blkodu,))
        balance_row = cursor.fetchone()

        total_debit = float(balance_row[0] or 0.0)
        total_credit = float(balance_row[1] or 0.0)
        net_balance = total_debit - total_credit

        # 3. Last 50 transactions from CARIHR (statement)
        statement_query = """
            SELECT FIRST 50
                H.TARIHI,
                H.VADESI,
                COALESCE(H.KPB_BTUT, 0),
                COALESCE(H.KPB_ATUT, 0),
                H.ACIKLAMA,
                H.EVRAK_NO
            FROM CARIHR H
            WHERE H.BLCRKODU = ?
            AND (H.SILINDI = 0 OR H.SILINDI IS NULL)
            ORDER BY H.TARIHI DESC, H.BLKODU DESC
        """
        cursor.execute(statement_query, (customer_blkodu,))
        raw_transactions = cursor.fetchall()

        transactions = []
        for r in raw_transactions:
            transaction_date = str(r[0])[:10] if r[0] else ""
            due_date = str(r[1])[:10] if r[1] else transaction_date
            debit_amount = float(r[2] or 0.0)
            credit_amount = float(r[3] or 0.0)
            description = r[4] or (r[5] if r[5] else "Cari Hareket")

            if debit_amount > 0:
                transactions.append({
                    "type": "debit",
                    "amount": round(debit_amount, 2),
                    "date": transaction_date,
                    "due_date": due_date,
                    "description": description
                })
            elif credit_amount > 0:
                transactions.append({
                    "type": "credit",
                    "amount": round(credit_amount, 2),
                    "date": transaction_date,
                    "due_date": due_date,
                    "description": description
                })

        append_log(f"[CARI BAKİYE] {title} -> Net: {net_balance:.2f}, Hareket: {len(transactions)}")

        return jsonify({
            "status": "success",
            "title": title,
            "tax_id": tax_id,
            "total_debit": round(total_debit, 2),
            "total_credit": round(total_credit, 2),
            "net_balance": round(net_balance, 2),
            "transactions": transactions
        })

    except Exception as e:
        append_log(f"[CARI BAKİYE HATA] {str(e)}")
        print("Cari Bakiye Hata:", str(e))
        return jsonify({"status": "error", "message": f"Sunucu Hatası: {str(e)}"}), 500
    finally:
        try: conn.close()
        except Exception: pass

@app.route('/api/barcodes/resolve', methods=['POST'])
@auth_required
def resolve_barcode():
    payload = request.json or {}
    barcode = str(payload.get('barcode', '')).strip()
    if not barcode:
        return jsonify({"status": "error", "message": "Barkod boş olamaz!"}), 400

    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanı bağlantısı yok!"}), 500

    try:
        cursor = conn.cursor()
        cursor.execute("SELECT S.BLKODU, TRIM(S.STOKKODU), S.STOK_ADI FROM STOK S WHERE S.AKTIF = 1 AND S.BARKODU = ?", (barcode,))
        match = cursor.fetchone()

        if not match or not match[1]:
            cursor.execute("""
                SELECT S.BLKODU, TRIM(S.STOKKODU), S.STOK_ADI 
                FROM STOK_BARKOD SB 
                INNER JOIN STOK S ON SB.BLSTKODU = S.BLKODU 
                WHERE S.AKTIF = 1 AND SB.BARKODU = ?
            """, (barcode,))
            match = cursor.fetchone()

        if match and match[1]:
            blkodu = match[0]
            _, _, _, deviation = find_price_by_blkodu(conn, blkodu)
            return jsonify({
                "status": "success",
                "blkodu": blkodu,
                "stock_code": str(match[1]).strip(),
                "product_name": str(match[2]).strip() if match[2] else str(match[1]).strip(),
                "price_deviation": deviation
            })
        else:
            return jsonify({"status": "not_found", "message": f"'{barcode}' barkodu Akınsoft sisteminde kayıtlı değil!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        try: conn.close()
        except Exception: pass

@app.route('/api/products/search', methods=['POST'])
@auth_required
def search_products():
    payload = request.json or {}
    search_text = str(payload.get('text', '')).strip()
    if not search_text:
        return jsonify({"status": "empty", "products": []})

    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanı bağlantısı yok!"}), 500

    try:
        cursor = conn.cursor()
        like_pattern = f"%{to_turkish_upper(search_text)}%"

        sql_query = """
            SELECT FIRST 80 S.BLKODU, TRIM(S.STOKKODU), S.STOK_ADI, S.BARKODU
            FROM STOK S
            WHERE S.AKTIF = 1
            AND (UPPER(S.STOK_ADI) LIKE ? OR UPPER(S.STOKKODU) LIKE ?)
            ORDER BY S.STOK_ADI
        """
        cursor.execute(sql_query, (like_pattern, like_pattern))
        matches = cursor.fetchall()

        product_list = []
        for row in matches:
            product_list.append({
                "blkodu": row[0],
                "stock_code": str(row[1]).strip() if row[1] else "",
                "product_name": row[2] if row[2] else "İsimsiz Ürün",
                "barcode": row[3] if row[3] else ""
            })

        return jsonify({"status": "success", "products": product_list})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        try: conn.close()
        except Exception: pass

@app.route('/api/barcodes/add', methods=['POST'])
@auth_required
def add_pending_barcode():
    payload = request.json or {}
    barcode = str(payload.get('barcode', '')).strip()
    note_text = str(payload.get('note', '')).strip()
    
    if len(barcode) != 13 or not barcode.isdigit():
        append_log(f"[HATA] Barkod Ekle Hatalı Karakter: {barcode}")
        return jsonify({"status": "error", "message": "Barkod tam 13 haneli sayı olmalıdır!"}), 400

    if not note_text:
        return jsonify({"status": "error", "message": "Ürün bilgisi boş bırakılamaz!"}), 400

    conn = get_db_connection()
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT STOK_ADI FROM STOK WHERE BARKODU = ?", (barcode,))
            match = cursor.fetchone()
            if match and match[0]:
                return jsonify({"status": "error", "message": f"Bu barkod zaten '{match[0]}' ürününe kayıtlı!"}), 400
                
            cursor.execute("""
                SELECT S.STOK_ADI 
                FROM STOK_BARKOD SB 
                INNER JOIN STOK S ON SB.BLSTKODU = S.BLKODU 
                WHERE SB.BARKODU = ?
            """, (barcode,))
            extra_match = cursor.fetchone()
            if extra_match and extra_match[0]:
                return jsonify({"status": "error", "message": f"Bu barkod zaten '{extra_match[0]}' ürününe ek barkod olarak kayıtlı!"}), 400
        except Exception:
            pass
        finally:
            try: conn.close()
            except Exception: pass

    try:
        timestamp = datetime.now().strftime("%d.%m.%Y %H:%M")
        entry_line = f"Barkod: {barcode} | Ürün/Not: {note_text} | Tarih: {timestamp}\n"
        with open(PENDING_BARCODES_FILE, "a", encoding="utf-8") as file_handle:
            file_handle.write(entry_line)
            
        append_log(f"[BARKOD EKLENDİ] {barcode} -> {note_text}")
        return jsonify({"status": "success"})
    except Exception as e:
        append_log(f"[HATA] Barkod Kayıt Hatası: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/invoices/items/add', methods=['POST'])
@auth_required
def add_invoice_item():
    conn = get_db_connection()
    if conn is None: 
        return jsonify({'status': 'error', 'message': 'Veritabanı bağlantısı yok!'}), 500

    try:
        data = request.json or {}
        customer_name = str(data.get('customer_name', '')).strip()
        input_value = str(data.get('barcode', '')).strip()
        quantity = data.get('quantity', 1)
        user_selected_price = data.get('selected_price', None)

        try: quantity = int(quantity)
        except (ValueError, TypeError): quantity = 1

        # A barcode typed into the quantity box (e.g. a 13-digit number) must never reach the list
        if quantity > MAX_ITEM_QUANTITY:
            append_log(f"[HATA] Aşırı miktar reddedildi: {quantity}")
            return jsonify({'status': 'error', 'message': f'Miktar çok yüksek ({quantity}). Barkod miktar kutusuna girmiş olabilir; en çok {MAX_ITEM_QUANTITY} adet girilebilir.'}), 400

        if not customer_name or not input_value:
            append_log("[HATA] Faturaya Ekleme Boş Alan")
            return jsonify({'status': 'error', 'message': 'Müşteri adı ve ürün/kod alanı boş olamaz!'}), 400

        safe_file_name = "".join([c for c in customer_name if c.isalnum() or c in (' ', '_', '-')]).strip()
        if not safe_file_name: safe_file_name = "isimsiz_liste"

        excel_stock_code = input_value
        display_name = input_value
        calculated_price = 0.01
        price_missing = False
        deviation_info = None
        blkodu = None  # Internal id of the matched product; the UI uses it to refresh the price info box

        cursor = conn.cursor()

        if len(input_value) == 13 and input_value.isdigit():
            try:
                cursor.execute("SELECT S.BLKODU, TRIM(S.STOKKODU), S.STOK_ADI FROM STOK S WHERE S.AKTIF = 1 AND S.BARKODU = ?", (input_value,))
                match = cursor.fetchone()
                
                if not match or not match[1]:
                    cursor.execute("""
                        SELECT S.BLKODU, TRIM(S.STOKKODU), S.STOK_ADI 
                        FROM STOK_BARKOD SB 
                        INNER JOIN STOK S ON SB.BLSTKODU = S.BLKODU 
                        WHERE S.AKTIF = 1 AND SB.BARKODU = ?
                    """, (input_value,))
                    match = cursor.fetchone()

                if match and match[1]:
                    blkodu = match[0]
                    excel_stock_code = str(match[1]).strip()
                    display_name = match[2] if match[2] else excel_stock_code
                    resolved_price, _, _, deviation_info = find_price_by_blkodu(conn, blkodu)
                    if resolved_price > 0:
                        calculated_price = resolved_price
                    else:
                        calculated_price = 0.01
                        price_missing = True
                else:
                    append_log(f"[HATA] Tanımsız Barkod: {input_value} ({safe_file_name})")
                    return jsonify({
                        'status': 'error',
                        'message': f"'{input_value}' barkodu Akınsoft sisteminde kayıtlı değil!"
                    }), 400
            except Exception:
                pass
        else:
            excel_stock_code = input_value
            display_name = input_value
            try:
                cursor.execute("SELECT S.BLKODU, S.STOK_ADI FROM STOK S WHERE S.AKTIF = 1 AND (TRIM(S.STOKKODU) = ? OR S.STOKKODU = ?)", (input_value, input_value))
                code_match = cursor.fetchone()
                if code_match:
                    blkodu = code_match[0]
                    display_name = code_match[1] if code_match[1] else excel_stock_code
                    resolved_price, _, _, deviation_info = find_price_by_blkodu(conn, blkodu)
                    if resolved_price > 0:
                        calculated_price = resolved_price
                    else:
                        calculated_price = 0.01
                        price_missing = True
                else:
                    price_missing = True
            except: 
                price_missing = True

        if user_selected_price is not None:
            calculated_price = float(user_selected_price)
            price_missing = False
        elif deviation_info and deviation_info.get('detected') and user_selected_price is None:
            return jsonify({
                'status': 'price_selection_required',
                'product_name': display_name,
                'stock_code': excel_stock_code,
                'blkodu': blkodu,
                'quantity': quantity,
                'price_deviation': deviation_info
            })

        if calculated_price <= 0:
            calculated_price = 0.01
            price_missing = True

        xlsx_path = os.path.join(INVOICE_OUTPUT_DIR, f"{safe_file_name}.xlsx")
        existing_records = read_invoice_excel(xlsx_path)
        code_found = False

        for row in existing_records:
            if str(row[0]).strip() == excel_stock_code:
                row[1] += quantity
                if calculated_price > 0.01:
                    row[2] = calculated_price
                code_found = True
                break

        if not code_found:
            existing_records.append([excel_stock_code, quantity, calculated_price])

        write_invoice_excel(xlsx_path, existing_records)
        append_log(f"[FATURA: {safe_file_name}] Kod: {excel_stock_code} ({display_name}) x {quantity} Adet | Fiyat: {calculated_price:.2f} TL")

        return jsonify({
            'status': 'success', 
            'product_name': display_name,
            'stock_code': excel_stock_code,
            'blkodu': blkodu,
            'price': calculated_price,
            'price_missing': price_missing,
            'message': f"[{excel_stock_code}] {display_name} listeye eklendi."
        })
    except Exception as e:
        append_log(f"[HATA] Fatura Ekleme Hatası: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500
    finally:
        try: conn.close()
        except Exception: pass

@app.route('/api/invoices/items', methods=['POST'])
@auth_required
def get_invoice_items():
    try:
        data = request.json or {}
        customer_name = str(data.get('customer_name', '')).strip()
        if not customer_name:
            return jsonify({
                'status': 'success', 
                'items': [], 
                'total_items': 0, 
                'total_quantity': 0,
                'total_amount': 0.0,
                'unpriced_item_count': 0
            })

        safe_file_name = "".join([c for c in customer_name if c.isalnum() or c in (' ', '_', '-')]).strip()
        xlsx_path = os.path.join(INVOICE_OUTPUT_DIR, f"{safe_file_name}.xlsx")

        if not os.path.exists(xlsx_path):
            return jsonify({
                'status': 'success', 
                'items': [], 
                'total_items': 0, 
                'total_quantity': 0,
                'total_amount': 0.0,
                'unpriced_item_count': 0
            })

        raw_records = read_invoice_excel(xlsx_path)
        if not raw_records:
            return jsonify({
                'status': 'success', 
                'items': [], 
                'total_items': 0, 
                'total_quantity': 0,
                'total_amount': 0.0,
                'unpriced_item_count': 0
            })

        stock_codes = [str(row[0]).strip() for row in raw_records]
        name_map = {}

        conn = get_db_connection()
        if conn and stock_codes:
            try:
                cursor = conn.cursor()
                placeholders = ",".join(["?"] * len(stock_codes))
                cursor.execute(f"SELECT TRIM(STOKKODU), STOK_ADI FROM STOK WHERE TRIM(STOKKODU) IN ({placeholders})", tuple(stock_codes))
                for r in cursor.fetchall():
                    if r[0]:
                        name_map[str(r[0]).strip()] = r[1] if r[1] else str(r[0]).strip()
            except Exception as e:
                print("İsim eşleştirme hatası:", str(e))
            finally:
                try: conn.close()
                except Exception: pass

        items = []
        total_quantity = 0
        total_amount = 0.0
        unpriced_count = 0

        for row in raw_records:
            code = str(row[0]).strip()
            quantity = row[1]
            price = float(row[2]) if len(row) >= 3 else 0.01
            
            total_quantity += quantity
            total_amount += (quantity * price)
            if price <= 0.01:
                unpriced_count += 1

            item_name = name_map.get(code, code)
            items.append({"code": code, "name": item_name, "quantity": quantity, "price": price})

        items.reverse()
        return jsonify({
            'status': 'success', 
            'items': items, 
            'total_items': len(items), 
            'total_quantity': total_quantity,
            'total_amount': round(total_amount, 2),
            'unpriced_item_count': unpriced_count
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/invoices/items/update', methods=['POST'])
@auth_required
def update_invoice_item():
    try:
        data = request.json or {}
        customer_name = str(data.get('customer_name', '')).strip()
        code = str(data.get('code', '')).strip()
        new_quantity = data.get('new_quantity', 0)
        new_price = data.get('new_price', None)
        action = data.get('action', 'update')

        if not customer_name or not code:
            return jsonify({'status': 'error', 'message': 'Müşteri adı ve ürün kodu zorunludur!'}), 400

        try:
            if action != 'update_price' and int(new_quantity) > MAX_ITEM_QUANTITY:
                append_log(f"[HATA] Aşırı miktar güncellemesi reddedildi: {new_quantity}")
                return jsonify({'status': 'error', 'message': f'Miktar çok yüksek ({new_quantity}). Barkod miktar kutusuna girmiş olabilir; en çok {MAX_ITEM_QUANTITY} adet girilebilir.'}), 400
        except (ValueError, TypeError):
            pass

        safe_file_name = "".join([c for c in customer_name if c.isalnum() or c in (' ', '_', '-')]).strip()
        xlsx_path = os.path.join(INVOICE_OUTPUT_DIR, f"{safe_file_name}.xlsx")

        if not os.path.exists(xlsx_path):
            return jsonify({'status': 'error', 'message': 'Müşteriye ait fatura listesi bulunamadı!'}), 404

        existing_records = read_invoice_excel(xlsx_path)
        updated_list = []

        for row in existing_records:
            if str(row[0]).strip() == code:
                if action == 'delete' or (action != 'update_price' and int(new_quantity) <= 0):
                    continue
                elif action == 'update_price':
                    f = float(new_price) if new_price is not None else float(row[2])
                    if f <= 0:
                        f = 0.01
                    m = int(row[1]) if int(new_quantity) <= 0 else int(new_quantity)
                    updated_list.append([code, m, f])
                else:
                    price = float(row[2]) if len(row) >= 3 else 0.01
                    updated_list.append([code, int(new_quantity), price])
            else:
                updated_list.append(row)

        write_invoice_excel(xlsx_path, updated_list)
        append_log(f"[FATURA GÜNCELLEME: {safe_file_name}] Kod: {code} -> İşlem: {action}, Miktar: {new_quantity}, Fiyat: {new_price}")
        return jsonify({'status': 'success', 'message': 'Liste başarıyla güncellendi.'})
    except Exception as e:
        append_log(f"[HATA] Kalem Güncelleme Hatası: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ==============================================================================
# AKINSOFT NATIVE WINDOWS UI AUTOMATION & SCREEN COORDINATES
# ==============================================================================
# The UI automation only works on Windows; the rest of the app runs anywhere.
IS_WINDOWS = os.name == "nt"
user32 = ctypes.windll.user32 if IS_WINDOWS else None
SCREEN_WIDTH = user32.GetSystemMetrics(0) if IS_WINDOWS else 1920
SCREEN_HEIGHT = user32.GetSystemMetrics(1) if IS_WINDOWS else 1080

MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_ABSOLUTE = 0x8000
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
KEYEVENTF_KEYUP = 0x0002

VK_RETURN = 0x0D
VK_CONTROL = 0x11
VK_F5 = 0x74
VK_V = 0x56

# Tested screen coordinates (specific to the shop PC resolution/layout).
# Use koordinat_olc.py to measure a coordinate on the shop PC.
COORD_ERP_TASKBAR = (467, 750)          # Akinsoft button on the taskbar
COORD_SHORTCUTS_TAB = (127, 65)         # "Shortcuts" tab
COORD_OPEN_INVOICE = (919, 528)         # "Domestic Sales Invoice" on the main screen
COORD_ACTIONS_MENU = (1219, 112)        # "Actions" menu inside the invoice
COORD_IMPORT_FROM_EXCEL = (1282, 457)   # "Import from Excel"
COORD_FILE_PICKER = (966, 228)          # File picker trigger
COORD_IMPORT_OPEN = (545, 555)          # Excel import "Open"
COORD_IMPORT_CONFIRM = (692, 438)       # Import confirmation 1
COORD_IMPORT_DONE = (753, 445)          # Import completed confirmation
COORD_VAT_ZERO_BUTTON = (767, 190)      # VAT 0% button
COORD_VAT_CONFIRM = (700, 439)          # VAT confirmation dialog
COORD_SAVE = (794, 688)                 # Save invoice
COORD_PRINT_MENU = (1158, 110)          # Print menu
COORD_PRINT_TEMPLATE = (1190, 158)      # Template selection
COORD_PRINT_CONFIRM = (734, 570)        # Print confirmation dialog
COORD_CLOSE_INVOICE_TAB = (390, 56)     # Close invoice tab

# After printing a multi-page invoice, Akinsoft shows a page-counter notice (2, 3, ...)
# and the invoice tab will not close until it is confirmed. The button is clicked for
# single-page invoices too. Set to None to skip this step.
COORD_PAGE_COUNTER_OK = (634, 408)
PAGE_COUNTER_WAIT_SECONDS = 1.0  # Extra wait for the notice to appear after the print confirmation

# Akinsoft's "Import from Excel" takes longer the more line items the list has.
# wait = BASE + item_count * PER_ITEM, clamped to [MIN, MAX].
# If the item count cannot be read, the old fixed FALLBACK wait is used.
IMPORT_WAIT_BASE = 4.0
IMPORT_WAIT_PER_ITEM = 0.07
IMPORT_WAIT_MIN = 5.0
IMPORT_WAIT_MAX = 30.0
IMPORT_WAIT_FALLBACK = 11.0


def import_wait_seconds(excel_path):
    """Return how long (seconds) to wait for the Excel import, based on the item count."""
    try:
        item_count = len(read_invoice_excel(excel_path))
    except Exception:
        item_count = 0
    if item_count <= 0:
        return IMPORT_WAIT_FALLBACK
    wait = IMPORT_WAIT_BASE + item_count * IMPORT_WAIT_PER_ITEM
    return max(IMPORT_WAIT_MIN, min(IMPORT_WAIT_MAX, wait))


def mouse_click(x, y):
    nx = int(x * 65535 / SCREEN_WIDTH)
    ny = int(y * 65535 / SCREEN_HEIGHT)
    user32.mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, nx, ny, 0, 0)
    time.sleep(0.1)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.08)
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def double_click(x, y):
    mouse_click(x, y)
    time.sleep(0.2)
    mouse_click(x, y)


def press_key(vk_code):
    user32.keybd_event(vk_code, 0, 0, 0)
    time.sleep(0.08)
    user32.keybd_event(vk_code, 0, KEYEVENTF_KEYUP, 0)


def press_ctrl_v():
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    user32.keybd_event(VK_V, 0, 0, 0)
    time.sleep(0.08)
    user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)


def copy_and_paste(text):
    pyperclip.copy(text)
    time.sleep(0.15)
    press_ctrl_v()
    time.sleep(0.2)


def is_erp_window_foreground():
    """Check whether the foreground window is Akinsoft/Wolvox."""
    active_hwnd = user32.GetForegroundWindow()
    if not active_hwnd:
        return False

    length = user32.GetWindowTextLengthW(active_hwnd)
    if length > 0:
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(active_hwnd, buffer, length + 1)
        window_title = buffer.value.upper()
        if "AKINSOFT" in window_title or "WOLVOX" in window_title:
            return True
    return False


def bring_erp_to_front():
    """Leave Akinsoft alone if it is in front; otherwise click it on the taskbar."""
    if is_erp_window_foreground():
        return  # Already in front, do not touch anything

    # Otherwise (behind or minimized) click it on the taskbar
    mouse_click(*COORD_ERP_TASKBAR)
    time.sleep(0.8)


def run_invoice_import_and_print_flow(customer_name_or_code, excel_path, print_invoice=True):
    """Tested 14-step Akinsoft invoice import and print flow.

    With print_invoice=False the invoice is finished at step 13 (Save): the print
    steps are skipped and the invoice tab is closed ("save only" mode).
    """

    # 0. Bring Akinsoft to the front
    bring_erp_to_front()

    # 0.5 Open the shortcuts tab
    mouse_click(*COORD_SHORTCUTS_TAB)
    time.sleep(0.8)

    # 1. Open the invoice from the main menu
    mouse_click(*COORD_OPEN_INVOICE)
    time.sleep(0.8)

    # 2. Customer lookup (F5)
    press_key(VK_F5)
    time.sleep(0.7)

    # 3. Paste the customer and select it
    copy_and_paste(customer_name_or_code)
    press_key(VK_RETURN)
    time.sleep(0.6)
    press_key(VK_RETURN)
    time.sleep(0.5)

    # 4. Actions menu
    mouse_click(*COORD_ACTIONS_MENU)
    time.sleep(0.5)

    # 5. Import from Excel
    mouse_click(*COORD_IMPORT_FROM_EXCEL)
    time.sleep(0.6)

    # 6. Trigger the file picker
    mouse_click(*COORD_FILE_PICKER)
    time.sleep(0.5)

    # 7. Paste the file path and press Enter
    copy_and_paste(excel_path)
    press_key(VK_RETURN)
    time.sleep(0.7)

    # 8. Start the import (Open)
    mouse_click(*COORD_IMPORT_OPEN)
    time.sleep(0.8)

    # 9. Import confirmation 1 (wait scales with the number of items)
    mouse_click(*COORD_IMPORT_CONFIRM)
    time.sleep(import_wait_seconds(excel_path))

    # 10. Import completed confirmation
    mouse_click(*COORD_IMPORT_DONE)
    time.sleep(0.6)

    # 11. VAT 0% button
    mouse_click(*COORD_VAT_ZERO_BUTTON)
    time.sleep(0.5)

    # 12. VAT confirmation dialog
    mouse_click(*COORD_VAT_CONFIRM)
    time.sleep(0.6)

    # 13. Save
    mouse_click(*COORD_SAVE)
    time.sleep(1.0)

    if print_invoice:
        # 14. Print steps
        mouse_click(*COORD_PRINT_MENU)
        time.sleep(0.5)
        mouse_click(*COORD_PRINT_TEMPLATE)
        time.sleep(0.6)
        mouse_click(*COORD_PRINT_CONFIRM)
        time.sleep(1.0)

        # 14.5 Dismiss the page-counter notice (the tab will not close until it is confirmed)
        if COORD_PAGE_COUNTER_OK is not None:
            time.sleep(PAGE_COUNTER_WAIT_SECONDS)
            mouse_click(*COORD_PAGE_COUNTER_OK)
            time.sleep(0.6)
    else:
        # Save-only mode: no printing. Give the save a little extra time
        # before the tab is closed.
        time.sleep(1.5)

    # 15. Close the domestic sales invoice tab
    mouse_click(*COORD_CLOSE_INVOICE_TAB)
    time.sleep(0.5)

    return True


# Coordinates
COORD_CUSTOMER_REPORT = (1161, 212)
COORD_REPORT_PRINT = (1200, 111)
COORD_REPORT_CLOSE_TAB = (393, 59)
COORD_REPORT_PRINT_CONFIRM = (1203, 139)


def run_customer_report_print_flow(customer_name_or_code):
    """Customer transaction report print flow recorded with a mouse recorder."""
    # 0. Bring Akinsoft to the front if needed
    bring_erp_to_front()
    time.sleep(0.5)

    # 1. Click the shortcuts tab (safety)
    mouse_click(*COORD_SHORTCUTS_TAB)
    time.sleep(0.6)

    # 2. Double-click "Customer Transaction Report"
    mouse_click(*COORD_CUSTOMER_REPORT)
    time.sleep(0.15)
    mouse_click(*COORD_CUSTOMER_REPORT)
    time.sleep(2.5)

    # 3. Press F5 to open the customer lookup
    press_key(VK_F5)
    time.sleep(0.8)

    # 4. Paste the customer from the clipboard and select it (Enter x3)
    copy_and_paste(customer_name_or_code)
    time.sleep(0.3)
    press_key(VK_RETURN)
    time.sleep(2.0)
    press_key(VK_RETURN)
    time.sleep(2.0)
    press_key(VK_RETURN)
    time.sleep(2.0)

    # 5. Click the print button
    mouse_click(*COORD_REPORT_PRINT)
    time.sleep(2.0)
    mouse_click(*COORD_REPORT_PRINT_CONFIRM)
    time.sleep(2.0)

    # 6. Close the transaction report tab
    mouse_click(*COORD_REPORT_CLOSE_TAB)
    time.sleep(0.5)

    return True


@app.route('/api/customers/report/print', methods=['POST'])
@auth_required
@automation_lock_required
def print_customer_report():
    try:
        data = request.get_json() or {}
        customer_name = str(data.get('customer_name', '')).strip()
        customer_code = str(data.get('customer_code', '')).strip()

        if not customer_name and not customer_code:
            return jsonify({'status': 'error', 'message': 'Yazdırılacak cari bilgisi bulunamadı.'}), 400

        target = customer_code if customer_code else customer_name
        run_customer_report_print_flow(target)

        append_log(f"[AKINSOFT RAPOR] Cari hareket raporu yazdırıldı: {target}")
        return jsonify({'status': 'success', 'message': f'"{target}" hareket raporu yazdırıldı.'})
    except Exception as e:
        append_log(f"[HATA] Cari Rapor Yazdırma Hatası: {str(e)}")
        return jsonify({'status': 'error', 'message': f'Rapor yazdırma hatası: {str(e)}'}), 500


# ==============================================================================
# POST-INVOICE CHECK: verify against Akinsoft's customer transactions (CARIHR)
# Read-only: it only runs SELECT queries and never writes to the database.
# How it works: before the invoice, the latest CARIHR record id is stored; after the
# invoice, the transactions newer than that id are summed and compared with the list total.
# ==============================================================================
VERIFY_ATTEMPTS = 8           # How many times to look for the new transaction
VERIFY_WAIT_SECONDS = 1.0     # Wait between attempts


def format_try(amount):
    """1250.5 -> '1.250,50' (Turkish number format)"""
    s = f"{float(amount):,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def expected_invoice_total(excel_path):
    """Total (quantity x price) and item count of the invoice list."""
    items = read_invoice_excel(excel_path)
    total = 0.0
    for record in items:
        price = float(record[2]) if len(record) >= 3 else 0.01
        total += record[1] * price
    return round(total, 2), len(items)


def find_customer_candidates(conn, customer_name, customer_code):
    """Customer record(s) the invoice is expected to land on (exact match first, then similar)."""
    cursor = conn.cursor()
    name_upper = to_turkish_upper(customer_name)

    if customer_code:
        cursor.execute("""
            SELECT FIRST 10 C.BLKODU, TRIM(C.CARIKODU), C.TICARI_UNVANI, C.ADI_SOYADI
            FROM CARI C
            WHERE C.AKTIF = 1
            AND UPPER(TRIM(C.CARIKODU)) = ?
        """, (to_turkish_upper(customer_code)[:30],))
        rows = cursor.fetchall()
    else:
        cursor.execute("""
            SELECT FIRST 10 C.BLKODU, TRIM(C.CARIKODU), C.TICARI_UNVANI, C.ADI_SOYADI
            FROM CARI C
            WHERE C.AKTIF = 1
            AND (UPPER(TRIM(C.CARIKODU)) = ?
                 OR UPPER(TRIM(COALESCE(C.TICARI_UNVANI, ''))) = ?
                 OR UPPER(TRIM(COALESCE(C.ADI_SOYADI, ''))) = ?)
        """, (name_upper[:30], name_upper, name_upper))
        rows = cursor.fetchall()

        if not rows:
            cursor.execute("""
                SELECT FIRST 5 C.BLKODU, TRIM(C.CARIKODU), C.TICARI_UNVANI, C.ADI_SOYADI
                FROM CARI C
                WHERE C.AKTIF = 1
                AND (UPPER(COALESCE(C.TICARI_UNVANI, '')) LIKE ?
                     OR UPPER(COALESCE(C.ADI_SOYADI, '')) LIKE ?)
            """, (f"%{name_upper}%", f"%{name_upper}%"))
            rows = cursor.fetchall()

    candidates = []
    for r in rows:
        code = str(r[1]).strip() if r[1] else ""
        title = str(r[2]).strip() if r[2] else (str(r[3]).strip() if r[3] else code)
        candidates.append({"blkodu": r[0], "code": code, "title": title})
    return candidates


def new_transactions_by_customer(conn, previous_max_id):
    """Sum the (non-deleted) customer transactions added after previous_max_id, per customer.
    Returns: {customer_blkodu: (total_debit, total_credit, count)}"""
    cursor = conn.cursor()
    cursor.execute("""
        SELECT H.BLCRKODU, COALESCE(SUM(H.KPB_BTUT), 0), COALESCE(SUM(H.KPB_ATUT), 0), COUNT(*)
        FROM CARIHR H
        WHERE H.BLKODU > ?
        AND (H.SILINDI = 0 OR H.SILINDI IS NULL)
        GROUP BY H.BLCRKODU
    """, (previous_max_id,))
    groups = {}
    for r in cursor.fetchall():
        groups[r[0]] = (float(r[1] or 0.0), float(r[2] or 0.0), int(r[3] or 0))
    return groups


def get_customer_summary(conn, blkodu):
    """A customer's code, title and current net balance (debit - credit)."""
    cursor = conn.cursor()
    cursor.execute("SELECT TRIM(CARIKODU), TICARI_UNVANI, ADI_SOYADI FROM CARI WHERE BLKODU = ?", (blkodu,))
    r = cursor.fetchone()
    code = str(r[0]).strip() if r and r[0] else ""
    title = str(r[1]).strip() if r and r[1] else (str(r[2]).strip() if r and r[2] else code)

    cursor.execute("""
        SELECT COALESCE(SUM(H.KPB_BTUT), 0), COALESCE(SUM(H.KPB_ATUT), 0)
        FROM CARIHR H
        WHERE H.BLCRKODU = ?
        AND (H.SILINDI = 0 OR H.SILINDI IS NULL)
    """, (blkodu,))
    b = cursor.fetchone()
    balance = float(b[0] or 0.0) - float(b[1] or 0.0)
    return {"code": code, "title": title, "balance": balance}


def get_latest_transaction_id(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT COALESCE(MAX(BLKODU), 0) FROM CARIHR")
    return int(cursor.fetchone()[0] or 0)


def prepare_invoice_check(customer_name, customer_code, excel_path):
    """Called BEFORE the invoice is created. Returns None on failure (the check is
    skipped; the invoice flow is never blocked)."""
    try:
        expected, item_count = expected_invoice_total(excel_path)
        conn = get_db_connection()
        if conn is None:
            return None
        try:
            candidates = find_customer_candidates(conn, customer_name, customer_code)
            previous_max_id = get_latest_transaction_id(conn)
        finally:
            try: conn.close()
            except Exception: pass
        return {
            "expected": expected,
            "item_count": item_count,
            "candidates": candidates,
            "previous_max_id": previous_max_id,
            "customer_name": customer_name,
        }
    except Exception as e:
        append_log(f"[FATURA KONTROL HATA] Hazırlık yapılamadı: {str(e)}")
        return None


def evaluate_invoice_check(groups, candidate_ids, expected, tolerance):
    """Compare the new customer transactions with the expected invoice total."""
    equal = lambda a, b: abs(a - b) <= tolerance

    for blkodu in candidate_ids:
        if blkodu in groups and equal(groups[blkodu][0], expected):
            return {"status": "verified", "customer_blkodu": blkodu, "posted": groups[blkodu][0]}

    for blkodu, (debit, _credit, _count) in groups.items():
        if blkodu not in candidate_ids and equal(debit, expected):
            return {"status": "wrong_customer", "customer_blkodu": blkodu, "posted": debit}

    candidate_debit = sum(groups[b][0] for b in candidate_ids if b in groups)
    if candidate_debit > 0:
        return {"status": "amount_mismatch", "customer_blkodu": None, "posted": candidate_debit}

    return {"status": "not_found", "customer_blkodu": None, "posted": 0.0}


def balance_text(balance):
    if balance > 0.01:
        return f"{format_try(balance)} TL (Borçlu)"
    if balance < -0.01:
        return f"{format_try(abs(balance))} TL (Alacaklı)"
    return "0,00 TL (Hesap kapalı)"


def poll_new_transactions(previous_max_id, evaluate, final_statuses, log_prefix):
    """Read the new CARIHR transactions up to VERIFY_ATTEMPTS times until `evaluate`
    returns one of `final_statuses`. Returns (result, groups)."""
    result, groups = None, {}
    for attempt in range(VERIFY_ATTEMPTS):
        conn = get_db_connection()
        if conn is not None:
            try:
                groups = new_transactions_by_customer(conn, previous_max_id)
            except Exception as e:
                append_log(f"[{log_prefix} HATA] Hareket okunamadı: {str(e)}")
            finally:
                try: conn.close()
                except Exception: pass
        result = evaluate(groups)
        if result["status"] in final_statuses:
            break
        if attempt < VERIFY_ATTEMPTS - 1:
            time.sleep(VERIFY_WAIT_SECONDS)
    return result, groups


def load_customer_summary(blkodu):
    if blkodu is None:
        return None
    conn = get_db_connection()
    if conn is None:
        return None
    try:
        return get_customer_summary(conn, blkodu)
    except Exception:
        return None
    finally:
        try: conn.close()
        except Exception: pass


def describe_groups(groups):
    return ", ".join(
        f"cari {code}: borç {v[0]:.2f} alacak {v[1]:.2f} adet {v[2]}" for code, v in groups.items()
    ) or "yeni hareket yok"


def finish_invoice_check(check):
    """Called AFTER the invoice is created; reads the customer transactions and returns
    the result. Never raises."""
    if check is None:
        return {
            "status": "skipped",
            "message": "Cari hareket kontrolü yapılamadı (veritabanına ulaşılamadı). Faturayı Akınsoft'ta elle kontrol et."
        }
    try:
        expected = check["expected"]
        tolerance = max(0.05, 0.01 * check["item_count"])
        candidate_ids = [c["blkodu"] for c in check["candidates"]]

        result, groups = poll_new_transactions(
            check["previous_max_id"],
            lambda g: evaluate_invoice_check(g, candidate_ids, expected, tolerance),
            ("verified", "wrong_customer"),
            "FATURA KONTROL",
        )

        status = result["status"]
        posted = result["posted"]
        customer = load_customer_summary(result["customer_blkodu"])

        many_candidates = len(check["candidates"]) > 1
        if status == "verified":
            message = (f"✓ Cari hareketi doğrulandı: {customer['title'] if customer else 'cari'} carisine "
                       f"{format_try(posted)} TL borç işlendi (liste toplamı {format_try(expected)} TL).")
            if customer:
                message += f" Güncel bakiye: {balance_text(customer['balance'])}."
            if many_candidates:
                message += " Not: benzer birden fazla cari var, yukarıdaki cariyi kontrol et."
        elif status == "wrong_customer":
            target = f"{customer['title']} ({customer['code']})" if customer else "başka bir cari"
            message = (f"YANLIŞ CARİ OLABİLİR! Fatura '{check['customer_name']}' carisine değil, "
                       f"'{target}' carisine {format_try(posted)} TL olarak işlenmiş görünüyor. "
                       f"Akınsoft'ta faturayı HEMEN kontrol et.")
        elif status == "amount_mismatch":
            message = (f"TUTAR FARKI! Cariye {format_try(posted)} TL işlendi ama listenin toplamı "
                       f"{format_try(expected)} TL. Akınsoft'taki faturayı kontrol et.")
        else:
            message = (f"FATURA GÖRÜNMÜYOR! Cari hareketlerinde {format_try(expected)} TL'lik yeni bir fatura "
                       f"borcu bulunamadı. Fatura oluşmamış olabilir; Akınsoft'ı kontrol et.")

        append_log(f"[FATURA KONTROL] {status} | {check['customer_name']} | beklenen: {expected:.2f} | "
                   f"islenen: {posted:.2f} | {describe_groups(groups)}")

        return {
            "status": status,
            "message": message,
            "expected": expected,
            "posted": round(posted, 2),
            "customer_code": customer["code"] if customer else "",
            "customer_title": customer["title"] if customer else "",
            "balance": round(customer["balance"], 2) if customer else None,
        }
    except Exception as e:
        append_log(f"[FATURA KONTROL HATA] {str(e)}")
        return {
            "status": "skipped",
            "message": "Cari hareket kontrolü sırasında hata oluştu. Faturayı Akınsoft'ta elle kontrol et."
        }


@app.route('/api/invoices/print', methods=['POST'])
@auth_required
@automation_lock_required
def print_invoice_via_erp():
    try:
        data = request.get_json() or {}
        customer_name = str(data.get('customer_name', '')).strip()
        customer_code = str(data.get('customer_code', '')).strip()
        # Printing is skipped only when save_only is explicitly true; otherwise the invoice is printed.
        save_only = data.get('save_only') is True

        if not customer_name:
            return jsonify({'status': 'error', 'message': 'Müşteri adı boş gönderilemez.'}), 400

        safe_file_name = "".join([c for c in customer_name if c.isalnum() or c in (' ', '_', '-')]).strip()
        file_path = os.path.join(INVOICE_OUTPUT_DIR, f"{safe_file_name}.xlsx")

        if not os.path.exists(file_path):
            alt_file_path = os.path.join(INVOICE_OUTPUT_DIR, f"{safe_file_name}.xls")
            if os.path.exists(alt_file_path):
                file_path = alt_file_path
            else:
                return jsonify({'status': 'error', 'message': f'Fatura Excel dosyası bulunamadı: {safe_file_name}.xlsx'}), 404

        # Never send an invoice that still has unpriced (0.01 TL) items to Akinsoft
        unpriced = [str(r[0]) for r in read_invoice_excel(file_path) if float(r[2]) <= UNPRICED_THRESHOLD]
        if unpriced:
            shown = ", ".join(unpriced[:8]) + (f" ... (+{len(unpriced) - 8})" if len(unpriced) > 8 else "")
            append_log(f"[AKINSOFT ENGEL] Fiyatsız kalem var, fatura gönderilmedi: {safe_file_name} -> {shown}")
            return jsonify({'status': 'error', 'message': f'Listede fiyatı belirlenmemiş (0,01 TL) {len(unpriced)} kalem var: {shown}. Önce "Fiyat Değiştir" ile fiyatlarını girin, sonra tekrar deneyin.'}), 400

        # Customer lookup parameter: customer code if available, otherwise name
        target_customer = customer_code if customer_code else customer_name

        # Before the invoice: expected total, candidate customers and the CARIHR snapshot
        check = prepare_invoice_check(customer_name, customer_code, file_path)

        # Run the full, tested automation flow (print steps are skipped in save-only mode)
        run_invoice_import_and_print_flow(target_customer, file_path, print_invoice=not save_only)

        # After the invoice: verify against the customer transactions (never raises)
        verification = finish_invoice_check(check)

        # Move the processed file to the archive folder (prevents duplicates)
        archive_dir = os.path.join(INVOICE_OUTPUT_DIR, "arsiv")
        os.makedirs(archive_dir, exist_ok=True)

        timestamp = time.strftime("%Y%m%d_%H%M%S")
        extension = os.path.splitext(file_path)[1]
        archive_target = os.path.join(archive_dir, f"{safe_file_name}_{timestamp}{extension}")

        archived = False
        try:
            os.replace(file_path, archive_target)
            archived = True
        except Exception as e:
            print(f"[UYARI] Dosya arşive taşınamadı: {e}")

        # If the check found a problem, say where the list file is (so it can be restored by hand)
        if verification.get('status') != 'verified' and archived:
            verification['message'] += f" (Liste dosyası: arsiv\\{os.path.basename(archive_target)})"

        if save_only:
            append_log(f"[AKINSOFT KAYDET] Fatura Akınsoft'a aktarıldı ve kaydedildi (YAZDIRILMADI), arşivlendi: {safe_file_name}")
            message = f'"{customer_name}" faturası Akınsoft\'a aktarıldı ve kaydedildi (yazdırılmadı), arşive taşındı.'
        else:
            append_log(f"[AKINSOFT YAZDIR] Fatura başarıyla Akınsoft'a aktarıldı, yazdırıldı ve arşivlendi: {safe_file_name}")
            message = f'"{customer_name}" faturası Akınsoft\'a aktarıldı, yazdırıldı ve arşive taşındı.'
        return jsonify({'status': 'success', 'message': message, 'verification': verification})

    except Exception as e:
        append_log(f"[HATA] Akınsoft Yazdırma Hatası: {str(e)}")
        return jsonify({'status': 'error', 'message': f'Otomasyon sırasında hata oluştu: {str(e)}'}), 500


# ==============================================================================
# CUSTOMER COLLECTION / PAYMENT (Cari Tahsilat / Cari Tediye — cash or document)
# Fills in and saves Akinsoft's collection/payment screen through UI automation.
# It never writes to the database. Read-only queries are used to: match the customer
# exactly (nothing starts if it is ambiguous), read the balance before the transaction,
# and verify the new CARIHR record afterwards.
# ==============================================================================
MAX_TRANSACTION_AMOUNT = 1000000   # Highest amount (TL) accepted at once; guards against extra zeros
DUPLICATE_WINDOW_SECONDS = 120     # Warn if the same customer + amount comes again within this window

# --- Coordinates: measure them on the shop PC (koordinat_olc.py helps). Automation does NOT start while one is None. ---
# Taken from the recorder files (cash collection + document payment).
COORD_COLLECTION_MENU = (1171, 247)         # "Cari Tahsilat" in the shortcuts list (double-clicked)
COORD_PAYMENT_MENU = (1194, 328)            # "Cari Tediye" in the shortcuts list (double-clicked)
COORD_CUSTOMER_LIST_FIRST_ROW = (595, 181)  # First row of the customer finder after filtering (double-click selects it)
COORD_MAIN_SCREEN_TAB = (127, 65)           # Workaround for an Akinsoft bug: the "main screen" tab (= Shortcuts tab)
COORD_COLLECTION_TAB = (288, 58)            # The "Cari Tahsilat" tab at the top
COORD_PAYMENT_TAB = None                    # The "Cari Tediye" tab; None = same place as the collection tab
COORD_METHOD_DROPDOWN = (709, 331)          # The dropdown that shows "Nakit" (used only to pick "Evrak")
COORD_METHOD_DOCUMENT = (708, 366)          # "Evrak" in the opened dropdown
COORD_TRANSACTION_AMOUNT = (523, 439)       # Credit / debit amount box
COORD_TRANSACTION_DESCRIPTION = (771, 518)  # Description box
COORD_TRANSACTION_SAVE = (840, 686)         # Green confirm (Save) button
COORD_TRANSACTION_CLOSE = (1020, 184)       # "Close" button of the screen

TRANSACTION_TYPES = ('collection', 'payment')   # collection = tahsilat, payment = tediye
PAYMENT_METHODS = ('cash', 'document')          # cash = nakit, document = evrak

_recent_transactions = {}
_recent_transactions_lock = threading.Lock()


def type_label(tx_type):
    return "tahsilat" if tx_type == 'collection' else "tediye"


def method_label(method):
    return "NAKİT" if method == 'cash' else "EVRAK"


def missing_transaction_coordinates(tx_type, method, has_description):
    """Names of the coordinates this transaction needs that are not defined yet (None)."""
    required = {
        "COORD_SHORTCUTS_TAB": COORD_SHORTCUTS_TAB,
        "COORD_CUSTOMER_LIST_FIRST_ROW": COORD_CUSTOMER_LIST_FIRST_ROW,
        "COORD_MAIN_SCREEN_TAB": COORD_MAIN_SCREEN_TAB,
        "COORD_TRANSACTION_AMOUNT": COORD_TRANSACTION_AMOUNT,
        "COORD_TRANSACTION_SAVE": COORD_TRANSACTION_SAVE,
        "COORD_TRANSACTION_CLOSE": COORD_TRANSACTION_CLOSE,
    }
    if tx_type == 'collection':
        required["COORD_COLLECTION_MENU"] = COORD_COLLECTION_MENU
        required["COORD_COLLECTION_TAB"] = COORD_COLLECTION_TAB
    else:
        required["COORD_PAYMENT_MENU"] = COORD_PAYMENT_MENU
        if COORD_PAYMENT_TAB is None:
            required["COORD_COLLECTION_TAB (or COORD_PAYMENT_TAB)"] = COORD_COLLECTION_TAB
    if method == 'document':
        required["COORD_METHOD_DROPDOWN"] = COORD_METHOD_DROPDOWN
        required["COORD_METHOD_DOCUMENT"] = COORD_METHOD_DOCUMENT
    if has_description:
        required["COORD_TRANSACTION_DESCRIPTION"] = COORD_TRANSACTION_DESCRIPTION
    return [name for name, value in required.items() if value is None]


def parse_transaction_request(data):
    """Validate the request. Returns (params, error_message). The amount must be a positive WHOLE number."""
    tx_type = str(data.get('type', '')).strip().lower()
    method = str(data.get('method', 'cash')).strip().lower()
    if tx_type not in TRANSACTION_TYPES:
        return None, "İşlem türü tahsilat veya tediye olmalı."
    if method not in PAYMENT_METHODS:
        return None, "Ödeme şekli nakit veya evrak olmalı."

    raw_amount = str(data.get('amount', '')).strip()
    if not re.fullmatch(r'\d{1,9}', raw_amount):
        return None, "Tutar sadece rakamlardan oluşan tam sayı olmalı (örn: 10000)."
    amount = int(raw_amount)
    if amount <= 0:
        return None, "Tutar sıfırdan büyük olmalı."
    if amount > MAX_TRANSACTION_AMOUNT:
        return None, f"Tutar çok yüksek ({format_try(amount)} TL). Güvenlik sınırı {format_try(MAX_TRANSACTION_AMOUNT)} TL."

    description = " ".join(str(data.get('description', '')).split())[:200]

    return {
        "type": tx_type, "method": method, "amount": amount, "description": description,
        "customer_code": str(data.get('customer_code', '')).strip(),
        "customer_name": str(data.get('customer_name', '')).strip(),
    }, None


def resolve_customer_exactly(conn, customer_code, customer_name):
    """Find the customer by an EXACT match only (no guessing by similar names).
    Returns: ('single', {blkodu, code, title, full_name}) | ('none', None) | ('ambiguous', None)"""
    cursor = conn.cursor()
    if customer_code:
        cursor.execute("""
            SELECT FIRST 5 C.BLKODU, TRIM(C.CARIKODU), C.TICARI_UNVANI, C.ADI_SOYADI
            FROM CARI C
            WHERE C.AKTIF = 1 AND UPPER(TRIM(C.CARIKODU)) = ?
        """, (to_turkish_upper(customer_code)[:30],))
    elif customer_name:
        name_upper = to_turkish_upper(customer_name)
        cursor.execute("""
            SELECT FIRST 5 C.BLKODU, TRIM(C.CARIKODU), C.TICARI_UNVANI, C.ADI_SOYADI
            FROM CARI C
            WHERE C.AKTIF = 1
            AND (UPPER(TRIM(C.CARIKODU)) = ?
                 OR UPPER(TRIM(COALESCE(C.TICARI_UNVANI, ''))) = ?
                 OR UPPER(TRIM(COALESCE(C.ADI_SOYADI, ''))) = ?)
        """, (name_upper[:30], name_upper, name_upper))
    else:
        return 'none', None
    rows = cursor.fetchall()
    if not rows:
        return 'none', None
    if len(rows) > 1:
        return 'ambiguous', None
    r = rows[0]
    code = str(r[1]).strip() if r[1] else ""
    title = str(r[2]).strip() if r[2] else ""
    full_name = str(r[3]).strip() if r[3] else ""
    return 'single', {"blkodu": r[0], "code": code, "title": title or full_name or code, "full_name": full_name}


def choose_erp_filter_text(conn, customer):
    """Pick the text to paste into Akinsoft's customer search: with this text ONLY this
    customer must show up. Returns None if no candidate gives a single match."""
    candidates = []
    for text in (customer["title"], customer["code"], customer["full_name"]):
        if text and text not in candidates:
            candidates.append(text)
    cursor = conn.cursor()
    for text in candidates:
        pattern = to_turkish_upper(text)
        cursor.execute("""
            SELECT COUNT(*) FROM CARI C
            WHERE C.AKTIF = 1
            AND (UPPER(TRIM(C.CARIKODU)) LIKE ?
                 OR UPPER(COALESCE(C.TICARI_UNVANI, '')) LIKE ?
                 OR UPPER(COALESCE(C.ADI_SOYADI, '')) LIKE ?)
        """, (f"%{pattern}%", f"%{pattern}%", f"%{pattern}%"))
        if int(cursor.fetchone()[0] or 0) == 1:
            return text
    return None


def expected_balance_after(balance, tx_type, amount):
    """A collection is posted as CREDIT (balance goes down), a payment as DEBIT (balance goes up)."""
    return balance - amount if tx_type == 'collection' else balance + amount


def prepare_transaction(p):
    """BEFORE the transaction: resolve the customer exactly, pick the Akinsoft search text,
    read the balance and the latest CARIHR id. Returns (context, error_message).
    Nothing starts if the database is unreachable or the customer is ambiguous."""
    conn = get_db_connection()
    if conn is None:
        return None, "Veritabanına ulaşılamadı. Para hareketi doğrulanamayacağı için işlem başlatılmadı."
    try:
        match, customer = resolve_customer_exactly(conn, p["customer_code"], p["customer_name"])
        if match == 'none':
            return None, ("Bu cari birebir bulunamadı. Cari adını yazıp büyüteç (Cari Ara) listesinden "
                          "seçerek tekrar dene.")
        if match == 'ambiguous':
            return None, ("Bu isim birden fazla cariyle eşleşiyor. Büyüteçle doğru cariyi seç. Hâlâ olmuyorsa "
                          "bu işlemi Akınsoft'ta elle gir.")
        filter_text = choose_erp_filter_text(conn, customer)
        if filter_text is None:
            return None, (f"'{customer['title']}' adı Akınsoft cari aramasında başka carilerle de eşleşiyor; "
                          f"otomasyon yanlış cariyi seçebilir. Bu işlemi Akınsoft'ta elle gir.")
        summary = get_customer_summary(conn, customer["blkodu"])
        previous_max_id = get_latest_transaction_id(conn)
    except Exception as e:
        append_log(f"[CARI ISLEM HATA] Hazırlık yapılamadı: {str(e)}")
        return None, "Cari bilgisi okunurken hata oluştu. İşlem başlatılmadı."
    finally:
        try: conn.close()
        except Exception: pass

    return {
        "customer": customer, "filter_text": filter_text, "balance_before": summary["balance"],
        "expected_balance": expected_balance_after(summary["balance"], p["type"], p["amount"]),
        "previous_max_id": previous_max_id,
    }, None


def is_duplicate_transaction(customer_blkodu, p):
    key = (customer_blkodu, p["type"], p["amount"])
    now = time.time()
    with _recent_transactions_lock:
        for k in [k for k, t in _recent_transactions.items() if now - t > DUPLICATE_WINDOW_SECONDS]:
            del _recent_transactions[k]
        return key in _recent_transactions


def remember_transaction(customer_blkodu, p):
    with _recent_transactions_lock:
        _recent_transactions[(customer_blkodu, p["type"], p["amount"])] = time.time()


def run_customer_transaction_flow(filter_text, tx_type, method, amount, description):
    """Click chain that fills in and saves Akinsoft's collection/payment screen.
    Waits come from the user's recordings (with some margin)."""
    bring_erp_to_front()

    # 1. Shortcuts tab -> Cari Tahsilat / Cari Tediye (the menu item is double-clicked)
    mouse_click(*COORD_SHORTCUTS_TAB)
    time.sleep(0.8)
    menu = COORD_COLLECTION_MENU if tx_type == 'collection' else COORD_PAYMENT_MENU
    double_click(*menu)
    time.sleep(2.0)

    # 2. The customer finder opens with the cursor in its search box: paste + Enter filters it.
    #    Enter does not select a single match by itself, so the first row is double-clicked.
    #    (The server already checked that this text yields exactly one customer in Akinsoft.)
    copy_and_paste(filter_text)
    time.sleep(0.3)
    press_key(VK_RETURN)
    time.sleep(1.5)
    double_click(*COORD_CUSTOMER_LIST_FIRST_ROW)
    time.sleep(1.5)

    # 3. Akinsoft bug: the screen stays behind. Click the main screen tab, then the transaction tab.
    mouse_click(*COORD_MAIN_SCREEN_TAB)
    time.sleep(0.8)
    if tx_type == 'collection':
        mouse_click(*COORD_COLLECTION_TAB)
    else:
        mouse_click(*(COORD_PAYMENT_TAB or COORD_COLLECTION_TAB))
    time.sleep(2.0)

    # 4. For a document, change the transaction method (default is cash)
    if method == 'document':
        mouse_click(*COORD_METHOD_DROPDOWN)
        time.sleep(0.8)
        mouse_click(*COORD_METHOD_DOCUMENT)
        time.sleep(0.8)

    # 5. Amount (clicked for cash too; harmless if the cursor is already there)
    mouse_click(*COORD_TRANSACTION_AMOUNT)
    time.sleep(0.7)
    copy_and_paste(str(int(amount)))
    time.sleep(0.8)

    # 6. Description (only if there is one)
    if description:
        mouse_click(*COORD_TRANSACTION_DESCRIPTION)
        time.sleep(0.6)
        copy_and_paste(description)
        time.sleep(0.8)

    # 7. Save, then close the screen
    mouse_click(*COORD_TRANSACTION_SAVE)
    time.sleep(2.0)
    mouse_click(*COORD_TRANSACTION_CLOSE)
    time.sleep(1.0)
    return True


def evaluate_transaction(groups, target_blkodu, tx_type, amount, tolerance=0.01):
    """Compare the new customer transactions with what was expected.
    groups: {customer: (debit, credit, count)}. A collection is expected as CREDIT, a payment as DEBIT."""
    side = 1 if tx_type == 'collection' else 0
    other_side = 1 - side
    equal = lambda a, b: abs(a - b) <= tolerance

    if target_blkodu in groups and equal(groups[target_blkodu][side], amount):
        return {"status": "verified", "customer_blkodu": target_blkodu, "posted": groups[target_blkodu][side]}
    for blkodu, g in groups.items():
        if blkodu != target_blkodu and equal(g[side], amount):
            return {"status": "wrong_customer", "customer_blkodu": blkodu, "posted": g[side]}
    if target_blkodu in groups:
        if equal(groups[target_blkodu][other_side], amount):
            return {"status": "reversed", "customer_blkodu": target_blkodu, "posted": groups[target_blkodu][other_side]}
        if groups[target_blkodu][side] > 0:
            return {"status": "amount_mismatch", "customer_blkodu": target_blkodu, "posted": groups[target_blkodu][side]}
    return {"status": "not_found", "customer_blkodu": None, "posted": 0.0}


def verify_transaction(context, p):
    """AFTER the transaction: read CARIHR and return the result. Never raises."""
    label = type_label(p["type"])
    amount = float(p["amount"])
    target = context["customer"]["blkodu"]
    title = context["customer"]["title"]
    try:
        result, groups = poll_new_transactions(
            context["previous_max_id"],
            lambda g: evaluate_transaction(g, target, p["type"], amount),
            ("verified", "wrong_customer", "reversed"),
            "CARI ISLEM KONTROL",
        )
        status, posted = result["status"], result["posted"]
        customer = load_customer_summary(result["customer_blkodu"])

        if status == "verified":
            message = f"✓ {title} carisine {format_try(posted)} TL {label} işlendi ({method_label(p['method'])})."
            if customer:
                message += f" Bakiye: {balance_text(context['balance_before'])} → {balance_text(customer['balance'])}."
                if abs(customer["balance"] - context["expected_balance"]) > 0.01:
                    message += (f" Not: beklenen bakiye {balance_text(context['expected_balance'])} idi; "
                                f"arada başka bir hareket işlenmiş olabilir.")
        elif status == "wrong_customer":
            other = f"{customer['title']} ({customer['code']})" if customer else "başka bir cari"
            message = (f"YANLIŞ CARİ OLABİLİR! {format_try(posted)} TL {label} '{title}' carisine değil, "
                       f"'{other}' carisine işlenmiş görünüyor. Akınsoft'ta HEMEN kontrol et ve gerekirse sil.")
        elif status == "reversed":
            message = (f"TERS YÖNDE İŞLENMİŞ OLABİLİR! '{title}' carisinde {format_try(posted)} TL "
                       f"{'borç' if p['type'] == 'collection' else 'alacak'} hareketi görünüyor. Akınsoft'ta kontrol et.")
        elif status == "amount_mismatch":
            message = (f"TUTAR FARKI! '{title}' carisine {format_try(posted)} TL işlendi ama girilen "
                       f"{format_try(amount)} TL idi. Akınsoft'ta kontrol et.")
        else:
            message = (f"KAYIT GÖRÜNMÜYOR! '{title}' carisinde {format_try(amount)} TL'lik yeni {label} "
                       f"hareketi bulunamadı. İşlem kaydedilmemiş olabilir; Akınsoft'ı kontrol et.")

        append_log(f"[CARI ISLEM KONTROL] {status} | {label} {p['method']} | {title} | beklenen: {amount:.2f} | "
                   f"islenen: {posted:.2f} | {describe_groups(groups)}")

        return {
            "status": status, "message": message, "amount": amount, "posted": round(posted, 2),
            "balance_before": round(context["balance_before"], 2),
            "balance": round(customer["balance"], 2) if customer else None,
        }
    except Exception as e:
        append_log(f"[CARI ISLEM KONTROL HATA] {str(e)}")
        return {"status": "skipped", "message": "Kayıt kontrolü sırasında hata oluştu. İşlemi Akınsoft'ta elle kontrol et."}


@app.route('/api/customers/transactions/preview', methods=['POST'])
@auth_required
def preview_customer_transaction():
    """For the confirmation screen: resolves the customer exactly and returns the current and
    expected balance. Does NOT touch Akinsoft."""
    try:
        p, error = parse_transaction_request(request.get_json() or {})
        if error:
            return jsonify({'status': 'error', 'message': error}), 400
        missing = missing_transaction_coordinates(p["type"], p["method"], bool(p["description"]))
        if missing:
            return jsonify({'status': 'error', 'message': "Otomasyon koordinatları tanımlı değil (app.py): " + ", ".join(missing)}), 400
        context, error = prepare_transaction(p)
        if error:
            return jsonify({'status': 'error', 'message': error}), 400
        return jsonify({
            'status': 'success',
            'customer_code': context["customer"]["code"],
            'title': context["customer"]["title"],
            'balance': round(context["balance_before"], 2),
            'balance_text': balance_text(context["balance_before"]),
            'expected_balance': round(context["expected_balance"], 2),
            'expected_balance_text': balance_text(context["expected_balance"]),
            'amount_text': format_try(p["amount"]),
        })
    except Exception as e:
        append_log(f"[CARI ISLEM ONIZLE HATA] {str(e)}")
        return jsonify({'status': 'error', 'message': f'Önizleme hatası: {str(e)}'}), 500


@app.route('/api/customers/transactions', methods=['POST'])
@auth_required
@automation_lock_required
def create_customer_transaction():
    """Posts a customer collection / payment in Akinsoft and verifies it from CARIHR."""
    try:
        data = request.get_json() or {}
        p, error = parse_transaction_request(data)
        if error:
            return jsonify({'status': 'error', 'message': error}), 400
        missing = missing_transaction_coordinates(p["type"], p["method"], bool(p["description"]))
        if missing:
            return jsonify({'status': 'error', 'message': "Otomasyon koordinatları tanımlı değil (app.py): " + ", ".join(missing)}), 400

        context, error = prepare_transaction(p)
        if error:
            return jsonify({'status': 'error', 'message': error}), 400

        label = type_label(p["type"])
        customer = context["customer"]
        if data.get('confirm_duplicate') is not True and is_duplicate_transaction(customer["blkodu"], p):
            return jsonify({
                'status': 'duplicate',
                'message': (f"Son {DUPLICATE_WINDOW_SECONDS // 60} dakika içinde '{customer['title']}' için aynı tutarda "
                            f"({format_try(p['amount'])} TL) {label} zaten girildi. Yanlışlıkla ikinci kez mi basıldı?")
            }), 409

        append_log(f"[CARI ISLEM BASLADI] {label} {p['method']} | {customer['title']} ({customer['code']}) | "
                   f"{p['amount']} TL | filtre: {context['filter_text']} | açıklama: {p['description'] or '-'}")
        # Even if the flow breaks halfway the record may exist, so mark it before running
        remember_transaction(customer["blkodu"], p)

        flow_error = None
        try:
            run_customer_transaction_flow(context["filter_text"], p["type"], p["method"], p["amount"], p["description"])
        except Exception as e:
            flow_error = str(e)
            append_log(f"[CARI ISLEM AKIS HATA] {flow_error}")

        verification = verify_transaction(context, p)

        if flow_error:
            return jsonify({
                'status': 'error',
                'message': f"Otomasyon sırasında hata oluştu: {flow_error}\nKayıt kontrolü: {verification['message']}",
                'verification': verification,
            }), 500

        message = f"{customer['title']} için {format_try(p['amount'])} TL {label} Akınsoft'a işlendi."
        return jsonify({'status': 'success', 'message': message, 'verification': verification})
    except Exception as e:
        append_log(f"[CARI ISLEM HATA] {str(e)}")
        return jsonify({'status': 'error', 'message': f'Cari işlem hatası: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv("PORT", 5000)), threaded=True, debug=False)
