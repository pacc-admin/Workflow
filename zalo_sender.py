import sys
sys.stdout.reconfigure(encoding='utf-8')

import argparse
import json
from zlapi import ZaloAPI
from zlapi.models import Message, ThreadType
import time
import os
import gspread
import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import datetime
import pytz

# ==========================================
# 1. CẤU HÌNH BIẾN
# ==========================================
TOKEN_FILE = os.environ.get('USER_TOKEN_FILE', 'token.json')
TOKEN_JSON = os.environ.get('TOKEN_JSON')
SHEET_URL = os.environ.get('SHEET_URL', 'https://docs.google.com/spreadsheets/d/1HWEEgcMOzDjOg-zm4LbSEnGHz00kC0d20XVW5zqDeFY/edit?gid=1573501749#gid=1573501749')
LOG_SHEET_NAME = os.environ.get('LOG_SHEET_NAME', 'PO_pdf_log')
TIMEZONE = pytz.timezone(os.environ.get('TIMEZONE', 'Asia/Ho_Chi_Minh'))
ALLOWED_FREQUENCIES = ['DAILY', 'WEEKLY', 'MONTHLY']

# ========== ZALO CREDENTIALS ==========
PHONE = os.environ.get('ZALO_PHONE', "0396880989")
IMEI = os.environ.get('ZALO_IMEI', "acb23337-932b-4088-aa3f-e26b0fac0cba-90daa551604269dbcdcf237b5cc700f3")
COOKIE = os.environ.get('ZALO_COOKIE', "__zi=3000.SSZzejyD7D4ecRUgqn5LsYBOiwdR7HgNFvQziPX129rnshouqmu9a77Vkl3K3m_OSzowyDu22DKn.1; __zi-legacy=3000.SSZzejyD7D4ecRUgqn5LsYBOiwdR7HgNFvQziPX129rnshouqmu9a77Vkl3K3m_OSzowyDu22DKn.1; zoaw_sek=Mwg6.304597146.3.6_P2zTaXX6j3wQDvsI64QzaXX6lP8OHhs620ax0XX6i; zoaw_type=0; _gid=GA1.2.384305111.1779073780; _zlang=vn; _gat=1; _ga_RYD7END4JE=GS2.2.s1779087842$o4$g1$t1779087843$j59$l0$h0; _ga_YS1V643LGV=GS2.1.s1779087842$o3$g0$t1779087843$j59$l0$h0; _ga=GA1.2.2098202610.1741932988; _ga_3EM8ZPYYN3=GS2.2.s1779087850$o5$g0$t1779087850$j60$l0$h0; zpsid=MRPl.196304666.25.M-fanEtwiB3tfGExuVgmi9Y9me7Kp8E0sSo6XjT0SJarLxITx089keVwiB0; zpw_sek=OzSy.196304666.a0.M52Ao8J6sP2UcMgRfSRPaF_alFkc_FlpvQkmySwqjz_jfOFYjl6S_zItdVRp--Io-VnZKGqUT9Ifno4fZgpPa0; app.event.zalo.me=2794251787792731214")

# ========== TEST MODE (Gửi vào Group ID cố định để test) ==========
USE_TEST_MODE = os.environ.get('USE_TEST_MODE', 'True').strip().lower() in ('1', 'true', 'yes')
TEST_GROUP_ID = os.environ.get('TEST_GROUP_ID', "3410445310157946128")

# ========== PRODUCTION MODE (Gửi vào từng Group ID từ Sheet) ==========
# Uncomment khối code dưới để chuyển sang production mode
# ──────────────────────────────────────────────────────────────────────────────
USE_TEST_MODE = False
TEST_GROUP_ID = None
# ──────────────────────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ==========================================
# 2. KHỞI TẠO SERVICE
# ==========================================
if TOKEN_JSON:
    token_info = json.loads(TOKEN_JSON)
    user_creds = Credentials.from_authorized_user_info(token_info, SCOPES)
else:
    user_creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

gc = gspread.authorize(user_creds)

def parse_args():
    parser = argparse.ArgumentParser(description='Gửi tin nhắn Zalo theo frequency')
    parser.add_argument(
        '--frequency',
        type=str,
        default='DAILY',
        help='Frequency filter: DAILY, WEEKLY, or MONTHLY. Default is DAILY.'
    )
    args = parser.parse_args()
    frequency = args.frequency.strip().upper()
    if frequency not in ALLOWED_FREQUENCIES:
        raise SystemExit(f"Giá trị frequency không hợp lệ: {args.frequency}. Chọn trong {ALLOWED_FREQUENCIES}")
    return frequency

# Parse Cookie thành Dictionary
def parse_cookie_string(cookie_str):
    cookie_dict = {}
    for item in cookie_str.split(';'):
        if '=' in item:
            k, v = item.strip().split('=', 1)
            cookie_dict[k] = v
    return cookie_dict

# ==========================================
# 3. HÀM TẠO MESSAGE
# ==========================================
def create_message(partner_name, warehouse_id, delivery_date, pdf_url):
    """
    Tạo nội dung tin nhắn Zalo
    """
    message_content = f"""PACC (BA GÁC) gửi {partner_name} đơn hàng:
- Chi nhánh: {warehouse_id}
- Ngày giao hàng: {delivery_date}
- Bấm vào link để xem chi tiết: {pdf_url}

Vui lòng hồi âm và xác nhận đơn hàng."""
    return message_content

# ==========================================
# 4. HÀM GỬI TIN NHẮN ZALO
# ==========================================
def send_zalo_message(client, group_id, message_content):
    """
    Gửi tin nhắn tới Group Zalo
    Return: True nếu thành công, False nếu thất bại
    """
    try:
        msg = Message(text=message_content)
        client.send(msg, thread_id=group_id, thread_type=ThreadType.GROUP)
        return True
    except Exception as e:
        print(f"   ❌ Lỗi gửi tin nhắn: {e}")
        return False

# ==========================================
# 5. HÀM CẬP NHẬT STATUS TRÊN SHEET
# ==========================================
def update_sent_status(worksheet, group_key, status):
    """
    Cập nhật sent_status cho dòng có group_key
    """
    try:
        all_records = worksheet.get_all_records()
        for i, record in enumerate(all_records):
            if str(record.get('group_key', '')).strip() == str(group_key).strip():
                # i+2 vì dòng đầu là header (dòng 1), và list index từ 0
                cell_row = i + 2
                # Column 'sent_status' là cột thứ 18 (R)
                worksheet.update_cell(cell_row, 18, 'done')
                print(f"   ✅ Cập nhật sent_status = 'done' cho row {cell_row}")
                return True
        print(f"   ⚠️  Không tìm thấy group_key: {group_key}")
        return False
    except Exception as e:
        print(f"   ❌ Lỗi cập nhật sheet: {e}")
        return False

# ==========================================
# 6. LOGIC CHÍNH
# ==========================================
def main(frequency='DAILY'):
    print(f"🚀 Bắt đầu job gửi tin nhắn Zalo (frequency={frequency})...\n")

    # A. Kết nối Google Sheet
    try:
        sheet = gc.open_by_url(SHEET_URL)
        worksheet = sheet.worksheet(LOG_SHEET_NAME)
        print("✅ Kết nối Google Sheet thành công!")
    except Exception as e:
        print(f"❌ Lỗi kết nối Sheet: {e}")
        return

    # B. Lấy dữ liệu từ sheet
    try:
        records = worksheet.get_all_records()
        print(f"📊 Tổng số dòng trong sheet: {len(records)}")
    except Exception as e:
        print(f"❌ Lỗi lấy dữ liệu từ sheet: {e}")
        return

    # C. Lọc dữ liệu: sent_status trống AND zalo_group_id không null AND frequency đúng
    requested_frequency = str(frequency).strip().upper()
    pending_records = []
    for record in records:
        sent_status = str(record.get('sent_status', '')).strip()
        zalo_group_id = str(record.get('zalo_group_id', '')).strip()
        frequency_value = str(record.get('frequency', '') or '').strip().upper()

        if requested_frequency == 'DAILY':
            valid_frequency = frequency_value in ('', 'DAILY')
        else:
            valid_frequency = frequency_value == requested_frequency

        if sent_status == '' and zalo_group_id and zalo_group_id != 'None' and valid_frequency:
            pending_records.append(record)

    if not pending_records:
        print("✅ Không có đơn hàng nào cần gửi Zalo. Job kết thúc.")
        return

    print(f"📦 Tìm thấy {len(pending_records)} đơn hàng cần gửi Zalo\n")

    # D. Khởi tạo Zalo Client
    print("🔄 Đang kết nối tới máy chủ Zalo...")
    try:
        cookie_dict = parse_cookie_string(COOKIE)
        client = ZaloAPI(PHONE, "password_bat_ky", imei=IMEI, cookies=cookie_dict)
        print("✅ Kết nối Zalo THÀNH CÔNG!\n")
    except Exception as e:
        print(f"❌ Lỗi kết nối Zalo: {e}")
        return

    # E. Gửi tin nhắn cho từng đơn hàng
    sent_count = 0
    failed_count = 0

    for idx, record in enumerate(pending_records, 1):
        try:
            group_key = record.get('group_key', 'N/A')
            partner_name = record.get('partner_name', 'Khách hàng')
            warehouse_id = record.get('warehouse_id', 'N/A')
            delivery_date = record.get('delivery_date', 'N/A')
            pdf_url = record.get('pdf_url', 'N/A')
            zalo_group_id = record.get('zalo_group_id', '')

            # Xác định Group ID (Test mode or Production mode)
            if USE_TEST_MODE:
                target_group_id = TEST_GROUP_ID
                print(f"[{idx}/{len(pending_records)}] 🧪 TEST MODE - Group Key: {group_key}")
                print(f"   📌 Sẽ gửi vào Group: {TEST_GROUP_ID} (thay vì {zalo_group_id})")
            else:
                target_group_id = zalo_group_id
                print(f"[{idx}/{len(pending_records)}] 📤 Group Key: {group_key}")

            # Tạo message
            message_content = create_message(partner_name, warehouse_id, delivery_date, pdf_url)
            print(f"   Nội dung: {message_content[:60]}...")

            # Gửi tin nhắn
            if send_zalo_message(client, target_group_id, message_content):
                print(f"   ✅ Gửi thành công!")
                
                # Cập nhật sent_status trên sheet
                if update_sent_status(worksheet, group_key, 'done'):
                    sent_count += 1
                else:
                    failed_count += 1
            else:
                failed_count += 1
                print(f"   ❌ Gửi thất bại!")

            # Ngủ 0.5s để tránh spam request
            time.sleep(1)

        except Exception as e:
            print(f"❌ Lỗi xử lý dòng {idx}: {e}")
            failed_count += 1

    # F. Tổng kết
    print("\n" + "=" * 50)
    print(f"✅ Gửi thành công: {sent_count} đơn hàng")
    print(f"❌ Gửi thất bại: {failed_count} đơn hàng")
    print("=" * 50)

    # G. Ghi metrics của lần chạy này
    try:
        if os.path.exists('run_metrics.json'):
            with open('run_metrics.json', 'r', encoding='utf-8') as f:
                metrics = json.load(f)
        else:
            metrics = {}
        metrics['zalo_sent'] = sent_count
        with open('run_metrics.json', 'w', encoding='utf-8') as f:
            json.dump(metrics, f, ensure_ascii=False)
        print(f"📊 Ghi metrics: {sent_count} Zalo gửi")
    except Exception as e:
        print(f"⚠️  Lỗi ghi metrics: {e}")

if __name__ == '__main__':
    frequency = parse_args()
    main(frequency)
