import os
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from flask import Flask, render_template, request, jsonify
import fdb

app = Flask(__name__)

# --- CONFIGURATION & PATHS ---
BASE_DIR = os.getenv("APP_BASE_DIR", os.path.join(os.path.expanduser("~"), "Desktop"))
INVOICE_OUTPUT_DIR = os.path.join(BASE_DIR, "fatura_listeleri")
PENDING_BARCODES_FILE = os.path.join(BASE_DIR, "eklenecekler.txt")

if not os.path.exists(INVOICE_OUTPUT_DIR):
    os.makedirs(INVOICE_OUTPUT_DIR)

DB_HOST = os.getenv("DB_HOST", "127.0.0.1")
DB_PORT = int(os.getenv("DB_PORT", 3050))
DB_PATH = os.getenv("DB_PATH", r"C:\AKINSOFT\Wolvox8\Database_FB\100\2026\wolvox.fdb")
DB_USER = os.getenv("DB_USER", "SYSDBA")
DB_PASSWORD = os.getenv("DB_PASSWORD", "2604123")
DB_CHARSET = os.getenv("DB_CHARSET", "WIN1254")

global_db_connection = None


# --- TURKISH STRING NORMALIZER ---
def to_turkish_upper(text):
    if not text:
        return ""
    char_map = {
        'i': 'İ',
        'ı': 'I',
        'ş': 'Ş',
        'ğ': 'Ğ',
        'ü': 'Ü',
        'ö': 'Ö',
        'ç': 'Ç'
    }
    for lower_char, upper_char in char_map.items():
        text = text.replace(lower_char, upper_char)
    return text.upper()


# --- DATABASE CONNECTION ---
def get_db_connection():
    global global_db_connection
    try:
        if global_db_connection is None or global_db_connection.closed:
            global_db_connection = fdb.connect(
                host=DB_HOST,
                port=DB_PORT,
                database=DB_PATH,
                user=DB_USER,
                password=DB_PASSWORD,
                charset=DB_CHARSET
            )
    except Exception as error:
        print("Database connection error:", str(error))
        return None
    return global_db_connection


# --- LOGGING HELPER ---
def append_log(message):
    try:
        today = datetime.now().strftime("%Y-%m-%d")
        current_time = datetime.now().strftime("%H:%M:%S")
        log_file_path = os.path.join(INVOICE_OUTPUT_DIR, f"gunluk_log_{today}.txt")
        with open(log_file_path, "a", encoding="utf-8") as file:
            file.write(f"[{current_time}] {message}\n")
    except Exception as error:
        print("Log error:", str(error))


# --- LIGHTWEIGHT EXCEL (OPENXML) ENGINE ---
def read_invoice_excel(file_path):
    records = []
    if not os.path.exists(file_path):
        return records

    try:
        with zipfile.ZipFile(file_path, 'r') as archive:
            shared_strings = []
            if 'xl/sharedStrings.xml' in archive.namelist():
                tree = ET.fromstring(archive.read('xl/sharedStrings.xml'))
                for elem in tree.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                    shared_strings.append(elem.text or '')

            sheet_tree = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            rows = sheet_tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')

            for row in rows[1:]:
                cells = row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
                if len(cells) >= 2:
                    code_elem = cells[0].find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    code_val = ""
                    if code_elem is not None and code_elem.text:
                        if cells[0].get('t') == 's':
                            code_val = shared_strings[int(code_elem.text)]
                        else:
                            code_val = code_elem.text

                    qty_elem = cells[1].find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    qty_val = 1
                    if qty_elem is not None and qty_elem.text:
                        try:
                            qty_val = int(float(qty_elem.text))
                        except (ValueError, TypeError):
                            qty_val = 1

                    if code_val:
                        records.append([code_val, qty_val])
    except Exception as error:
        print("Excel read error:", str(error))
    return records


def write_invoice_excel(file_path, records):
    strings = ["STOKKODU", "MIKTARI"]
    for code, _ in records:
        if code not in strings:
            strings.append(code)

    sst_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    sst_xml += f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(strings)}" uniqueCount="{len(strings)}">'
    for string_entry in strings:
        sst_xml += f'<si><t>{string_entry}</t></si>'
    sst_xml += '</sst>'

    sheet_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    sheet_xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    sheet_xml += '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'

    row_index = 2
    for code, quantity in records:
        code_idx = strings.index(code)
        sheet_xml += f'<row r="{row_index}"><c r="A{row_index}" t="s"><v>{code_idx}</v></c><c r="B{row_index}"><v>{quantity}</v></c></row>'
        row_index += 1

    sheet_xml += '</sheetData></worksheet>'

    content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    wb_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>'
    wb_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'

    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('[Content_Types].xml', content_types)
        archive.writestr('_rels/.rels', rels)
        archive.writestr('xl/workbook.xml', wb_xml)
        archive.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        archive.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        archive.writestr('xl/sharedStrings.xml', sst_xml)


# --- ROUTES ---
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


def format_single_price_response(conn, blkodu, product_name, stock_code="", barcode=""):
    cursor = conn.cursor()
    price_query = f"""
        SELECT FIRST 3 H.KPB_FIYATI, H.EVRAK_NO, H.TARIHI
        FROM STOKHR H
        WHERE H.BLSTKODU = {blkodu}
        AND (H.EVRAK_NO STARTING WITH 'SFT' OR H.EVRAK_NO STARTING WITH 'sft')
        ORDER BY H.TARIHI DESC
    """
    cursor.execute(price_query)
    rows = cursor.fetchall()

    history = []
    if rows:
        for row in rows:
            price = row[0]
            history.append({
                "price": float(price) if price is not None else 0.0,
                "date": str(row[2])[:10]
            })

        if len(history) > 1 and all(item["price"] == history[0]["price"] for item in history):
            history = [history[0]]

    append_log(f"[PRICE QUERY: SINGLE] BLKODU:{blkodu} | {stock_code} -> {product_name}")
    return jsonify({
        "status": "single_match",
        "blkodu": blkodu,
        "stock_code": stock_code,
        "barcode": barcode,
        "product_name": product_name,
        "data": history
    })


@app.route('/api/price/query', methods=['POST'])
def query_price():
    payload = request.json or {}
    search_text = payload.get('search_query', '').strip()

    if not search_text:
        return jsonify({"status": "error", "message": "Arama metni boş olamaz!"}), 400

    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanına bağlanılamıyor!"}), 500

    try:
        conn.commit()
        cursor = conn.cursor()

        # Step 1: Exact Barcode Search
        barcode_query = f"""
            SELECT S.BLKODU, S.STOKKODU, S.STOK_ADI, S.BARKODU
            FROM STOK S
            WHERE S.AKTIF = 1 AND S.BARKODU = '{search_text}'
        """
        cursor.execute(barcode_query)
        match = cursor.fetchall()

        if not match:
            additional_barcode_query = f"""
                SELECT S.BLKODU, S.STOKKODU, S.STOK_ADI, SB.BARKODU
                FROM STOK_BARKOD SB
                INNER JOIN STOK S ON SB.BLSTKODU = S.BLKODU
                WHERE S.AKTIF = 1 AND SB.BARKODU = '{search_text}'
            """
            cursor.execute(additional_barcode_query)
            match = cursor.fetchall()

        if match:
            blkodu = match[0][0]
            stock_code = match[0][1] or ""
            product_name = match[0][2] or "İsimsiz Ürün"
            barcode = match[0][3] or ""
            return format_single_price_response(conn, blkodu, product_name, stock_code, barcode)

        # Step 2: Like Search on Stock Name and Stock Code
        clean_text = to_turkish_upper(search_text).replace("'", "''")
        general_query = f"""
            SELECT FIRST 80 S.BLKODU, S.STOKKODU, S.STOK_ADI, S.BARKODU
            FROM STOK S
            WHERE S.AKTIF = 1
            AND (UPPER(S.STOK_ADI) LIKE '%{clean_text}%' OR UPPER(S.STOKKODU) LIKE '%{clean_text}%')
            ORDER BY S.STOK_ADI
        """
        cursor.execute(general_query)
        results = cursor.fetchall()

        if not results:
            append_log(f"[PRICE QUERY: NOT FOUND] {search_text}")
            return jsonify({"status": "not_found", "message": f"'{search_text}' için aktif ürün bulunamadı."})

        if len(results) == 1:
            blkodu = results[0][0]
            stock_code = results[0][1] or ""
            product_name = results[0][2] or "İsimsiz Ürün"
            barcode = results[0][3] or ""
            return format_single_price_response(conn, blkodu, product_name, stock_code, barcode)

        products = []
        for row in results:
            products.append({
                "blkodu": row[0],
                "stock_code": row[1] if row[1] else "",
                "stock_name": row[2] if row[2] else "İsimsiz Ürün",
                "barcode": row[3] if row[3] else ""
            })

        append_log(f"[PRICE QUERY: MULTIPLE] '{search_text}' -> {len(products)} adet sonuç")
        return jsonify({
            "status": "multiple_match",
            "products": products
        })

    except Exception as error:
        try:
            conn.rollback()
        except:
            pass
        append_log(f"[ERROR] Price Query Failed ({search_text}): {str(error)}")
        return jsonify({"status": "error", "message": str(error)}), 500


@app.route('/api/price/detail/<int:blkodu>', methods=['GET'])
def get_price_detail(blkodu):
    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanı bağlantısı yok!"}), 500

    try:
        conn.commit()
        cursor = conn.cursor()
        cursor.execute(f"SELECT STOKKODU, STOK_ADI, BARKODU FROM STOK WHERE BLKODU = {blkodu}")
        row = cursor.fetchone()

        stock_code = row[0] if row and row[0] else ""
        product_name = row[1] if row and row[1] else "Ürün"
        barcode = row[2] if row and row[2] else ""

        return format_single_price_response(conn, blkodu, product_name, stock_code, barcode)
    except Exception as error:
        try:
            conn.rollback()
        except:
            pass
        append_log(f"[ERROR] Price Detail Failed (BLKODU:{blkodu}): {str(error)}")
        return jsonify({"status": "error", "message": str(error)}), 500


@app.route('/api/customers/search', methods=['POST'])
def search_customers():
    payload = request.json or {}
    search_text = payload.get('search_query', '').strip()

    if not search_text:
        return jsonify({"status": "empty", "customers": []})

    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanı bağlantısı yok!"}), 500

    try:
        conn.commit()
        cursor = conn.cursor()
        clean_text = to_turkish_upper(search_text).replace("'", "''")

        query = f"""
            SELECT FIRST 30 
                C.CARIKODU, 
                C.TICARI_UNVANI, 
                C.ADI_SOYADI
            FROM CARI C
            WHERE C.AKTIF = 1
            AND (
                UPPER(COALESCE(C.TICARI_UNVANI, '')) LIKE '%{clean_text}%' 
                OR UPPER(COALESCE(C.CARIKODU, '')) LIKE '%{clean_text}%'
                OR UPPER(COALESCE(C.ADI_SOYADI, '')) LIKE '%{clean_text}%'
            )
            ORDER BY C.TICARI_UNVANI
        """
        cursor.execute(query)
        results = cursor.fetchall()

        customers = []
        for row in results:
            customer_code = row[0] if row[0] else ""
            trade_title = str(row[1]).strip() if row[1] else ""
            full_name = str(row[2]).strip() if row[2] else ""

            if trade_title and full_name and trade_title != full_name:
                display_name = f"{trade_title} ({full_name})"
                excel_title = trade_title
            elif trade_title:
                display_name = trade_title
                excel_title = trade_title
            elif full_name:
                display_name = full_name
                excel_title = full_name
            else:
                display_name = "İsimsiz Cari"
                excel_title = customer_code

            customers.append({
                "customer_code": customer_code,
                "title": excel_title,
                "display_name": display_name
            })

        return jsonify({"status": "success", "customers": customers})
    except Exception as error:
        return jsonify({"status": "error", "message": str(error)}), 500


@app.route('/api/products/search', methods=['POST'])
def search_products():
    payload = request.json or {}
    search_text = payload.get('search_query', '').strip()

    if not search_text:
        return jsonify({"status": "empty", "products": []})

    conn = get_db_connection()
    if conn is None:
        return jsonify({"status": "error", "message": "Veritabanı bağlantısı yok!"}), 500

    try:
        conn.commit()
        cursor = conn.cursor()
        clean_text = to_turkish_upper(search_text).replace("'", "''")

        query = f"""
            SELECT FIRST 80 S.BLKODU, S.STOKKODU, S.STOK_ADI, S.BARKODU
            FROM STOK S
            WHERE S.AKTIF = 1
            AND (UPPER(S.STOK_ADI) LIKE '%{clean_text}%' OR UPPER(S.STOKKODU) LIKE '%{clean_text}%')
            ORDER BY S.STOK_ADI
        """
        cursor.execute(query)
        results = cursor.fetchall()

        products = []
        for row in results:
            products.append({
                "blkodu": row[0],
                "stock_code": row[1] if row[1] else "",
                "stock_name": row[2] if row[2] else "İsimsiz Ürün",
                "barcode": row[3] if row[3] else ""
            })

        return jsonify({"status": "success", "products": products})
    except Exception as error:
        return jsonify({"status": "error", "message": str(error)}), 500


@app.route('/api/barcodes/add', methods=['POST'])
def add_pending_barcode():
    payload = request.json or {}
    barcode = str(payload.get('barcode', '')).strip()
    product_note = payload.get('note', '').strip()

    if len(barcode) != 13:
        append_log(f"[ERROR] Add Barcode Invalid Length: {barcode}")
        return jsonify({"status": "error", "message": "Barkod tam 13 haneli olmalıdır!"}), 400

    if not product_note:
        return jsonify({"status": "error", "message": "Ürün bilgisi boş bırakılamaz!"}), 400

    conn = get_db_connection()
    if conn:
        try:
            conn.commit()
            cursor = conn.cursor()

            cursor.execute(f"SELECT STOK_ADI FROM STOK WHERE BARKODU = '{barcode}'")
            match = cursor.fetchone()
            if match and match[0]:
                return jsonify({"status": "error", "message": f"Bu barkod zaten '{match[0]}' ürününe kayıtlı!"}), 400

            cursor.execute(f"SELECT S.STOK_ADI FROM STOK_BARKOD SB INNER JOIN STOK S ON SB.BLSTKODU = S.BLKODU WHERE SB.BARKODU = '{barcode}'")
            additional_match = cursor.fetchone()
            if additional_match and additional_match[0]:
                return jsonify({"status": "error", "message": f"Bu barkod zaten '{additional_match[0]}' ürününe ek barkod olarak kayıtlı!"}), 400
        except:
            try:
                conn.rollback()
            except:
                pass

    try:
        timestamp = datetime.now().strftime("%d.%m.%Y %H:%M")
        log_line = f"Barkod: {barcode} | Ürün/Not: {product_note} | Tarih: {timestamp}\n"
        with open(PENDING_BARCODES_FILE, "a", encoding="utf-8") as file:
            file.write(log_line)

        append_log(f"[BARCODE REGISTERED] {barcode} -> {product_note}")
        return jsonify({"status": "success"})
    except Exception as error:
        append_log(f"[ERROR] Barcode Save Failed: {str(error)}")
        return jsonify({"status": "error", "message": str(error)})


@app.route('/api/invoices/items/add', methods=['POST'])
def add_invoice_item():
    try:
        payload = request.json or {}
        customer_name = str(payload.get('customer_name', '')).strip()
        input_value = str(payload.get('item_input', '')).strip()
        quantity = payload.get('quantity', 1)

        try:
            quantity = int(quantity)
        except (ValueError, TypeError):
            quantity = 1

        if not customer_name or not input_value:
            append_log("[ERROR] Add Invoice Item Missing Required Fields")
            return jsonify({'status': 'error', 'message': 'Müşteri adı ve ürün/kod alanı boş olamaz!'}), 400

        safe_filename = "".join([c for c in customer_name if c.isalnum() or c in (' ', '_', '-')]).strip()
        if not safe_filename:
            safe_filename = "isimsiz_liste"

        conn = get_db_connection()
        if conn is None:
            return jsonify({'status': 'error', 'message': 'Veritabanı bağlantısı yok!'}), 500

        target_stock_code = input_value
        display_product_name = input_value

        # Step 1: 13-digit EAN Barcode Verification
        if len(input_value) == 13 and input_value.isdigit():
            try:
                conn.commit()
                cursor = conn.cursor()

                cursor.execute(f"SELECT STOKKODU, STOK_ADI FROM STOK WHERE BARKODU = '{input_value}'")
                match = cursor.fetchone()

                if not match or not match[0]:
                    cursor.execute(f"SELECT S.STOKKODU, S.STOK_ADI FROM STOK_BARKOD SB INNER JOIN STOK S ON SB.BLSTKODU = S.BLKODU WHERE SB.BARKODU = '{input_value}'")
                    match = cursor.fetchone()

                if match and match[0]:
                    target_stock_code = str(match[0]).strip()
                    display_product_name = match[1] if match[1] else target_stock_code
                else:
                    append_log(f"[ERROR] Unregistered Barcode: {input_value} ({safe_filename})")
                    return jsonify({
                        'status': 'error',
                        'message': f"UYARI: '{input_value}' barkodu Akınsoft sisteminde kayıtlı değil!\n\nLütfen barkodu kontrol edin veya ürün adıyla arama yapın."
                    }), 400
            except:
                try:
                    conn.rollback()
                except:
                    pass
        else:
            target_stock_code = input_value
            display_product_name = input_value

        excel_path = os.path.join(INVOICE_OUTPUT_DIR, f"{safe_filename}.xlsx")
        existing_records = read_invoice_excel(excel_path)
        is_item_found = False

        for record in existing_records:
            if record[0] == target_stock_code:
                record[1] += quantity
                is_item_found = True
                break

        if not is_item_found:
            existing_records.append([target_stock_code, quantity])

        write_invoice_excel(excel_path, existing_records)
        append_log(f"[INVOICE: {safe_filename}] Code: {target_stock_code} ({display_product_name}) x {quantity} Qty")

        return jsonify({
            'status': 'success',
            'product_name': display_product_name,
            'message': f"[{target_stock_code}] {display_product_name} listeye eklendi."
        })

    except Exception as error:
        append_log(f"[ERROR] Add Invoice Item Failed: {str(error)}")
        return jsonify({'status': 'error', 'message': str(error)}), 500


@app.route('/api/invoices/items', methods=['POST'])
def get_invoice_items():
    try:
        payload = request.json or {}
        customer_name = str(payload.get('customer_name', '')).strip()
        if not customer_name:
            return jsonify({'status': 'success', 'items': [], 'total_items': 0, 'total_quantity': 0})

        safe_filename = "".join([c for c in customer_name if c.isalnum() or c in (' ', '_', '-')]).strip()
        excel_path = os.path.join(INVOICE_OUTPUT_DIR, f"{safe_filename}.xlsx")

        if not os.path.exists(excel_path):
            return jsonify({'status': 'success', 'items': [], 'total_items': 0, 'total_quantity': 0})

        raw_records = read_invoice_excel(excel_path)
        if not raw_records:
            return jsonify({'status': 'success', 'items': [], 'total_items': 0, 'total_quantity': 0})

        stock_codes = ["'" + str(row[0]).replace("'", "''") + "'" for row in raw_records]
        name_map = {}

        conn = get_db_connection()
        if conn and stock_codes:
            try:
                conn.commit()
                cursor = conn.cursor()
                codes_joined = ",".join(stock_codes)
                cursor.execute(f"SELECT STOKKODU, STOK_ADI FROM STOK WHERE STOKKODU IN ({codes_joined})")
                for match in cursor.fetchall():
                    if match[0]:
                        name_map[str(match[0]).strip()] = match[1] if match[1] else str(match[0]).strip()
            except Exception as error:
                print("Name mapping error:", str(error))

        items = []
        total_quantity = 0
        for code, quantity in raw_records:
            total_quantity += quantity
            resolved_name = name_map.get(code, code)
            items.append({
                "code": code,
                "name": resolved_name,
                "quantity": quantity
            })

        items.reverse()

        return jsonify({
            'status': 'success',
            'items': items,
            'total_items': len(items),
            'total_quantity': total_quantity
        })
    except Exception as error:
        return jsonify({'status': 'error', 'message': str(error)}), 500


@app.route('/api/invoices/items/update', methods=['POST'])
def update_invoice_item():
    try:
        payload = request.json or {}
        customer_name = str(payload.get('customer_name', '')).strip()
        item_code = str(payload.get('code', '')).strip()
        new_quantity = payload.get('new_quantity', 0)
        action = payload.get('action', 'update')

        if not customer_name or not item_code:
            return jsonify({'status': 'error', 'message': 'Müşteri adı ve ürün kodu zorunludur!'}), 400

        safe_filename = "".join([c for c in customer_name if c.isalnum() or c in (' ', '_', '-')]).strip()
        excel_path = os.path.join(INVOICE_OUTPUT_DIR, f"{safe_filename}.xlsx")

        if not os.path.exists(excel_path):
            return jsonify({'status': 'error', 'message': 'Müşteriye ait fatura listesi bulunamadı!'}), 404

        existing_records = read_invoice_excel(excel_path)
        updated_records = []

        for code, quantity in existing_records:
            if code == item_code:
                if action == 'delete' or int(new_quantity) <= 0:
                    continue
                else:
                    updated_records.append([code, int(new_quantity)])
            else:
                updated_records.append([code, quantity])

        write_invoice_excel(excel_path, updated_records)
        append_log(f"[INVOICE UPDATE: {safe_filename}] Code: {item_code} -> Action: {action}, Qty: {new_quantity}")

        return jsonify({'status': 'success', 'message': 'Liste başarıyla güncellendi.'})
    except Exception as error:
        append_log(f"[ERROR] Invoice Update Failed: {str(error)}")
        return jsonify({'status': 'error', 'message': str(error)}), 500


if __name__ == '__main__':
    port = int(os.getenv("PORT", 5000))
    app.run(host='0.0.0.0', port=port, threaded=True, debug=False)