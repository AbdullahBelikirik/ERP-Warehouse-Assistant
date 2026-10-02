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

app = Flask(__name__)

load_dotenv()


def require_env(name):
    value = os.getenv(name)
    if not value:
        raise SystemExit(f"Missing required setting {name}. Copy .env.example to .env and fill it in.")
    return value


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

# Tested screen coordinates (specific to the shop PC resolution/layout)
COORD_ERP_TASKBAR = (467,750) # Akinsoft button on the taskbar
COORD_SHORTCUTS_TAB = (127, 65)  # "Shortcuts" tab
COORD_OPEN_INVOICE = (919, 528)          # "Domestic Sales Invoice" on the main screen
COORD_ACTIONS_MENU = (1219, 112)          # "Actions" menu inside the invoice
COORD_IMPORT_FROM_EXCEL = (1282, 457)          # "Import from Excel"
COORD_FILE_PICKER = (966, 228)        # File picker trigger
COORD_IMPORT_OPEN = (545, 555)          # Excel import "Open"
COORD_IMPORT_CONFIRM = (692, 438)        # Import confirmation 1
COORD_IMPORT_DONE = (753, 445)       # Import completed confirmation
COORD_VAT_ZERO_BUTTON = (767, 190)           # VAT 0% button
COORD_VAT_CONFIRM = (700, 439)            # VAT confirmation dialog
COORD_SAVE = (794, 688)              # Save invoice
COORD_PRINT_MENU = (1158, 110)        # Print menu
COORD_PRINT_TEMPLATE = (1190, 158)      # Template selection
COORD_PRINT_CONFIRM = (734, 570)         # Print confirmation dialog
COORD_CLOSE_INVOICE_TAB = (390,56) # Close invoice tab

def mouse_click(x, y):
    nx = int(x * 65535 / SCREEN_WIDTH)
    ny = int(y * 65535 / SCREEN_HEIGHT)
    user32.mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, nx, ny, 0, 0)
    time.sleep(0.1)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.08)
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

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
    
def run_invoice_import_and_print_flow(customer_name_or_code, excel_path):
    """Tested 14-step Akinsoft invoice import and print flow."""

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

    # 9. Import confirmation 1
    mouse_click(*COORD_IMPORT_CONFIRM)
    time.sleep(11)

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

    # 14. Print steps
    mouse_click(*COORD_PRINT_MENU)
    time.sleep(0.5)
    mouse_click(*COORD_PRINT_TEMPLATE)
    time.sleep(0.6)
    mouse_click(*COORD_PRINT_CONFIRM)
    time.sleep(3.5)

    # 15. Close the domestic sales invoice tab
    mouse_click(*COORD_CLOSE_INVOICE_TAB)
    time.sleep(0.5)

    return True


# Coordinates
COORD_CUSTOMER_REPORT = (1161, 212)
COORD_REPORT_PRINT = (1200, 111)
COORD_REPORT_CLOSE_TAB = (393, 59)
COORD_REPORT_PRINT_CONFIRM = (1203,139)

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



@app.route('/api/invoices/print', methods=['POST'])
@auth_required
def print_invoice_via_erp():
    try:
        data = request.get_json() or {}
        customer_name = str(data.get('customer_name', '')).strip()
        customer_code = str(data.get('customer_code', '')).strip()

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

        # Customer lookup parameter: customer code if available, otherwise name
        target_customer = customer_code if customer_code else customer_name

        # Run the full, tested automation flow
        run_invoice_import_and_print_flow(target_customer, file_path)

        # Move the printed file to the archive folder (prevents duplicates)
        archive_dir = os.path.join(INVOICE_OUTPUT_DIR, "arsiv")
        os.makedirs(archive_dir, exist_ok=True)

        timestamp = time.strftime("%Y%m%d_%H%M%S")
        extension = os.path.splitext(file_path)[1]
        archive_target = os.path.join(archive_dir, f"{safe_file_name}_{timestamp}{extension}")

        try:
            os.replace(file_path, archive_target)
        except Exception as e:
            print(f"[UYARI] Dosya arşive taşınamadı: {e}")

        append_log(f"[AKINSOFT YAZDIR] Fatura başarıyla Akınsoft'a aktarıldı, yazdırıldı ve arşivlendi: {safe_file_name}")
        return jsonify({
            'status': 'success',
            'message': f'"{customer_name}" faturası Akınsoft\'a aktarıldı, yazdırıldı ve arşive taşındı.'
        })

    except Exception as e:
        append_log(f"[HATA] Akınsoft Yazdırma Hatası: {str(e)}")
        return jsonify({'status': 'error', 'message': f'Otomasyon sırasında hata oluştu: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv("PORT", 5000)), threaded=True, debug=False)