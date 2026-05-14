import os
import io
import json
import datetime
import pandas as pd
import pytz
import gspread
from google.oauth2 import service_account
from google.cloud import bigquery
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.oauth2.credentials import Credentials


# ==========================================
# 1. CẤU HÌNH BIẾN (BẠN ĐIỀN THÔNG TIN VÀO ĐÂY)
# ==========================================
SERVICE_ACCOUNT_FILE = os.environ.get('SERVICE_ACCOUNT_FILE', 'service_account.json')
SERVICE_ACCOUNT_JSON = os.environ.get('SERVICE_ACCOUNT_JSON')
TOKEN_FILE = os.environ.get('USER_TOKEN_FILE', 'token.json')
TOKEN_JSON = os.environ.get('TOKEN_JSON')
BQ_PROJECT_ID = os.environ.get('BQ_PROJECT_ID', 'pacc-analytics-prod')
BQ_DATASET_ID = os.environ.get('BQ_DATASET_ID', 'pacc')
BQ_TABLE_NAME = os.environ.get('BQ_TABLE_NAME', 'int_grn_po')

TEMPLATE_ID = os.environ.get('TEMPLATE_ID', '1qqXARPKGghXpdwyHaa5B8-QrTp1b-WCWWRtrezD6yBw')
ROOT_FOLDER_ID = os.environ.get('ROOT_FOLDER_ID', '1rijuHPPJ5KoC4wgrwF9F1TeySuhLBA_z')

WAREHOUSE_FOLDERS = {
    'HCM-BG7-NKKN': '10xuPahVis7D9kozAWNShTgM7fr3-p4UW',
    'HCM-BG8-QT'  : '1q1TkUEeHf1HsFfGFXJFFlsFi0YkU7vAR',
    'HCM-BG10-XT' : '1oKHQGwroCavqtAPKH4Y_diMdbiKUsGnt',
    'HCM-CK-PACC' : '1i82KQNNnst7wN9z_bPeDPjc2ubGpa3SE',
    'HCM-LAB1'    : '14wfaeL85ix4p2azI9_QNC4dVbGjBs1Jt',
    'HCM-HO1'     : '1MAQn-lMcykDVKNXwXvYmDtXOukjJhmBS'
}

SHEET_URL = os.environ.get('SHEET_URL', 'https://docs.google.com/spreadsheets/d/1HWEEgcMOzDjOg-zm4LbSEnGHz00kC0d20XVW5zqDeFY/edit?gid=1573501749#gid=1573501749')
LOG_SHEET_NAME = os.environ.get('LOG_SHEET_NAME', 'PO_pdf_log')
MAX_PER_RUN = int(os.environ.get('MAX_PER_RUN', '100'))
TIMEZONE = pytz.timezone(os.environ.get('TIMEZONE', 'Asia/Ho_Chi_Minh'))

# ==========================================
# 2. KHỞI TẠO CÁC SERVICE GOOGLE
# ==========================================
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/spreadsheets"
]

# 1. BigQuery: Vẫn xài Service Account (như cũ)
if SERVICE_ACCOUNT_JSON:
    sa_info = json.loads(SERVICE_ACCOUNT_JSON)
    sa_creds = service_account.Credentials.from_service_account_info(sa_info)
else:
    sa_creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE)
bq_client = bigquery.Client(credentials=sa_creds, project=BQ_PROJECT_ID)

# 2. Drive, Docs, Sheets: Xài Token cá nhân của bạn
# Nó sẽ tự động làm mới token mỗi khi chạy mà không cần đăng nhập
if TOKEN_JSON:
    token_info = json.loads(TOKEN_JSON)
    user_creds = Credentials.from_authorized_user_info(token_info, SCOPES)
else:
    user_creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

drive_service = build('drive', 'v3', credentials=user_creds)
docs_service = build('docs', 'v1', credentials=user_creds)
gc = gspread.authorize(user_creds)

print("✅ Đã kết nối thành công với danh nghĩa email cá nhân!")

# ==========================================
# 3. CÁC HÀM HELPER
# ==========================================
def map_warehouse(warehouse_id, delivery_date_obj):
    date_str = delivery_date_obj.strftime('%Y-%m-%d') if pd.notnull(delivery_date_obj) else ''
    wh_map = {
        'HCM-BG7-NKKN': {'a': 'Sân thượng - 61 Nam Kỳ Khởi Nghĩa, p. Bến Thành, TPHCM', 'm': 'Trường Nguyễn', 'p': '0938329634', 't': '2:00pm - 4:00pm'},
        'HCM-BG8-QT'  : {'a': 'Sân thượng - 01 Quang Trung, p. Gò Vấp, TPHCM', 'm': 'Danh Phạm', 'p': '0937873937', 't': '2:00pm - 4:00pm'},
        'HCM-BG10-XT' : {'a': '39 Xuân Thủy, p. An Khánh, TPHCM', 'm': 'Trường Đào', 'p': '0375779817', 't': '2:00pm - 4:00pm'},
        'HCM-CK-PACC' : {'a': '39 Xuân Thủy, p. An Khánh, TPHCM', 'm': 'Danh Phạm', 'p': '0937873937', 't': '7:00am - 9:00am'},
        'HCM-LAB1'    : {'a': '27 Đường 19, Khu Đô Thị Lakeview city, p. Bình Trưng, TPHCM', 'm': 'Linh Đinh', 'p': '0933133311', 't': '9:00am - 12:00am'},
        'HCM-HO1'     : {'a': '39 Xuân Thủy, p. An Khánh, TPHCM', 'm': 'Diễm Nguyễn', 'p': '0396880989', 't': '2:00pm - 4:00pm'}
    }
    info = wh_map.get(str(warehouse_id).strip(), {'a':'', 'm':'', 'p':'', 't':''})
    delivery_datetime = f"{info['t']} {date_str}".strip()
    return info['a'], info['m'], info['p'], delivery_datetime

def get_existing_group_keys(worksheet):
    try:
        records = worksheet.get_all_records()
        return [str(row['group_key']).strip() for row in records if 'group_key' in row]
    except Exception:
        return []

def replace_text_in_doc(document_id, replace_dict):
    requests = []
    for key, value in replace_dict.items():
        requests.append({
            'replaceAllText': {
                'containsText': {'text': key, 'matchCase': True},
                'replaceText': str(value) if value else ''
            }
        })
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

# Hàm tạo PDF từ Template (Tương đương createSinglePO_ trong GAS)
def create_pdf_from_template(po_data, items_df):
    folder_id = WAREHOUSE_FOLDERS.get(po_data['warehouse_id'], ROOT_FOLDER_ID)
    file_name = f"{po_data['comments'] or 'PO'}_{po_data['partner_tax'] or ''}"

    # 1. Copy Template Docs
    body = {'name': file_name, 'parents': [folder_id]}
    copied_file = drive_service.files().copy(fileId=TEMPLATE_ID, body=body, supportsAllDrives=True).execute()
    doc_id = copied_file.get('id')

    # XỬ LÝ BẢNG BẰNG CÁCH NỐI CHUỖI (NEWLINE HACK)
    # Lấy cột tương ứng, ép kiểu về String, xử lý dữ liệu trống, rồi nối các dòng lại bằng ký tự xuống dòng (\n)
    col_item_id = '\n'.join(items_df['item_id'].astype(str).replace('nan', ''))
    col_item_name = '\n'.join(items_df['item_name'].astype(str).replace('nan', ''))
    col_unit_id = '\n'.join(items_df['unit_id'].astype(str).replace('nan', ''))
    
    # Đối với quantity, format cho đẹp (bỏ đuôi .0 nếu có)
    items_df['quantity_order'] = items_df['quantity_order'].fillna(0)
    col_qty = '\n'.join(items_df['quantity_order'].apply(lambda x: f"{x:g}"))

    # 2. Replace Text Variables (Gộp cả thông tin PO và Bảng)
    replace_dict = {
        '{{warehouse_id}}': po_data['warehouse_id'],
        '{{delivery_date}}': po_data['delivery_datetime'],
        '{{partner_tax}}': po_data['partner_tax'],
        '{{partner_name}}': po_data['partner_name'],
        '{{po_tran_no}}': po_data['po_tran_no'],
        '{{po_tran_date}}': po_data['po_tran_date'],
        '{{comments}}': po_data['comments'],
        '{{outlet_address}}': po_data['outlet_address'],
        '{{mod}}': po_data['recipient'],
        '{{mod_phone}}': po_data['recipient_phone'],
        # Đưa các cột dữ liệu vào bảng
        '{{col_item_id}}': col_item_id,
        '{{col_item_name}}': col_item_name,
        '{{col_unit_id}}': col_unit_id,
        '{{col_qty}}': col_qty
    }
    
    # Hàm replace text (Đã định nghĩa ở trên)
    replace_text_in_doc(doc_id, replace_dict)

    # 3. Export to PDF
    request = drive_service.files().export_media(fileId=doc_id, mimeType='application/pdf')
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()

    # 4. Upload PDF back to Drive folder
    fh.seek(0)
    file_metadata = {'name': f"{file_name}.pdf", 'parents': [folder_id]}
    media = MediaIoBaseUpload(fh, mimetype='application/pdf', resumable=True)
    pdf_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()

    # 5. Xóa file Docs tạm (Giải phóng dung lượng cho Bot)
    try:
        drive_service.files().delete(fileId=doc_id, supportsAllDrives=True).execute()
    except:
        pass

    return pdf_file.get('webViewLink')

# ==========================================
# 4. LOGIC CHÍNH
# ==========================================
def main():
    print("🚀 Đang khởi chạy tiến trình...")

    # A. Kết nối Google Sheet
    sheet = gc.open_by_url(SHEET_URL)
    try:
        worksheet = sheet.worksheet(LOG_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sheet.add_worksheet(title=LOG_SHEET_NAME, rows="100", cols="20")
        headers = ['group_key', 'po_pr_key', 'po_tran_no', 'po_tran_date', 'delivery_date', 'warehouse_id', 
                   'po_type', 'comments', 'partner_id', 'partner_tax', 'partner_name', 'po_pdf_status', 
                   'pdf_url', 'created_at', 'zalo_group_id', 'frequency', 'sent_channel', 'sent_status']
        worksheet.append_row(headers)

    existing_keys = get_existing_group_keys(worksheet)
    print(f"📊 Tìm thấy {len(existing_keys)} PO đã được tạo trước đó.")

    # B. Query BigQuery
    sql = f"""
        SELECT po_pr_key, warehouse_id, po_tran_no, po_tran_date, delivery_date, po_type, 
               comments, partner_id, partner_tax, partner_name, item_id, item_name, unit_id, 
               quantity_order, zalo_id, frequency, po_review_status
        FROM `{BQ_PROJECT_ID}.{BQ_DATASET_ID}.{BQ_TABLE_NAME}`
        WHERE date(po_tran_date) >= DATE_SUB(CURRENT_DATE(), INTERVAL 3 DAY)
          AND ((LOWER(frequency) = 'weekly' AND LOWER(po_review_status) = 'checked') 
          OR (LOWER(frequency) <> 'weekly' OR frequency IS NULL))
    """
    df = bq_client.query(sql).to_dataframe()
    if df.empty:
        print("✅ Không có data mới từ BigQuery.")
        return

    # C. Group data (Tạo group_key)
    df['warehouse_id'] = df['warehouse_id'].fillna('')
    df['partner_id'] = df['partner_id'].fillna('')
    df['group_key'] = df['po_tran_no'].astype(str) + '||' + df['warehouse_id'].astype(str) + '||' + df['partner_id'].astype(str)

    # Lọc bỏ các PO đã tạo
    df_new = df[~df['group_key'].isin(existing_keys)]
    groups = df_new.groupby('group_key')
    
    if len(groups) == 0:
        print("✅ Tất cả PO đã được up to date.")
        return

    print(f"📦 Có {len(groups)} PO mới cần tạo PDF.")
    
    log_buffer = []
    count = 0

    for group_key, group_df in groups:
        if count >= MAX_PER_RUN:
            break
            
        try:
            print(f"⏳ Đang xử lý: {group_key}")
            first_row = group_df.iloc[0]

            # Parse ngày giao hàng & map warehouse
            delivery_date = pd.to_datetime(first_row['delivery_date'])
            addr, mod, phone, deliv_dt = map_warehouse(first_row['warehouse_id'], delivery_date)

            po_data = {
                'group_key': group_key,
                'po_pr_key': first_row['po_pr_key'],
                'warehouse_id': first_row['warehouse_id'],
                'po_tran_no': first_row['po_tran_no'],
                'po_tran_date': str(first_row['po_tran_date']),
                'delivery_date': str(first_row['delivery_date']),
                'po_type': first_row['po_type'],
                'comments': first_row['comments'],
                'partner_id': first_row['partner_id'],
                'partner_tax': first_row['partner_tax'],
                'partner_name': first_row['partner_name'],
                'zalo_group_id': first_row['zalo_id'],
                'frequency': first_row['frequency'],
                'recipient': mod,
                'recipient_phone': phone,
                'outlet_address': addr,
                'delivery_datetime': deliv_dt
            }

            items_df = group_df[['item_id', 'item_name', 'unit_id', 'quantity_order']]
            
            # TẠO PDF (Gọi hàm)
            pdf_url = create_pdf_from_template(po_data, items_df)

            # CHUẨN BỊ DATA ĐỂ LOG VÀO SHEET
            created_at = datetime.datetime.now(TIMEZONE).strftime('%d/%m/%Y %H:%M:%S')
            
            log_row = [
                po_data['group_key'], str(po_data['po_pr_key']), po_data['po_tran_no'], 
                po_data['po_tran_date'], po_data['delivery_date'], po_data['warehouse_id'], 
                po_data['po_type'], po_data['comments'], po_data['partner_id'], 
                str(po_data['partner_tax']), po_data['partner_name'], 
                'done',  # po_pdf_status
                pdf_url, # pdf_url
                created_at, 
                str(po_data['zalo_group_id']), 
                po_data['frequency'], 
                'zalo',  # sent_channel
                ''       # sent_status (Để trống chuẩn bị cho Zalo sau này)
            ]
            log_buffer.append(log_row)
            count += 1
            print(f"   => Thành công: {pdf_url}")

        except Exception as e:
            print(f"❌ Lỗi ở {group_key}: {e}")

    # D. Lưu log vào Google Sheet
    if log_buffer:
        print(f"📝 Đang lưu {len(log_buffer)} dòng vào Google Sheet...")
        worksheet.append_rows(log_buffer, value_input_option='USER_ENTERED')
        print("✅ Đã lưu xong!")

    # E. Ghi metrics của lần chạy này
    metrics = {'pdf_created': count}
    with open('run_metrics.json', 'w', encoding='utf-8') as f:
        json.dump(metrics, f, ensure_ascii=False)
    print(f"📊 Ghi metrics: {count} PDF tạo")

if __name__ == '__main__':
    main()