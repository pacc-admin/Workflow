import sys
sys.stdout.reconfigure(encoding='utf-8')

import os
import json
import gspread
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.oauth2 import service_account
from google.cloud import bigquery
from google.oauth2.credentials import Credentials
from datetime import datetime
import pytz

# ==========================================
# CẤU HÌNH
# ==========================================
SERVICE_ACCOUNT_FILE = os.environ.get('SERVICE_ACCOUNT_FILE', 'service_account.json')
SERVICE_ACCOUNT_JSON = os.environ.get('SERVICE_ACCOUNT_JSON')
TOKEN_FILE = os.environ.get('USER_TOKEN_FILE', 'token.json')
TOKEN_JSON = os.environ.get('TOKEN_JSON')

BQ_PROJECT_ID = os.environ.get('BQ_PROJECT_ID', 'pacc-analytics-prod')
BQ_DATASET_ID = os.environ.get('BQ_DATASET_ID', 'pacc')
BQ_TABLE_NAME = os.environ.get('BQ_TABLE_NAME', 'int_grn_po')

SHEET_URL = os.environ.get('SHEET_URL', 'https://docs.google.com/spreadsheets/d/1HWEEgcMOzDjOg-zm4LbSEnGHz00kC0d20XVW5zqDeFY/edit?gid=1573501749#gid=1573501749')
LOG_SHEET_NAME = os.environ.get('LOG_SHEET_NAME', 'PO_pdf_log')

EMAIL_RECIPIENTS = os.environ.get('EMAIL_RECIPIENTS', 'dataengineerpacc@gmail.com,pacc.qc@gmail.com,pacc.workplace@gmail.com').split(',')
EMAIL_RECIPIENTS = [e.strip() for e in EMAIL_RECIPIENTS if e.strip()]

TIMEZONE = pytz.timezone(os.environ.get('TIMEZONE', 'Asia/Ho_Chi_Minh'))
SMTP_USER = os.environ.get('SMTP_USER', 'dataengineerpacc@gmail.com')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD', 'wlvankiktxabffdu')

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ==========================================
# KHỞI TẠO SERVICE
# ==========================================
# BigQuery
if SERVICE_ACCOUNT_JSON:
    sa_info = json.loads(SERVICE_ACCOUNT_JSON)
    sa_creds = service_account.Credentials.from_service_account_info(sa_info)
else:
    sa_creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE)
bq_client = bigquery.Client(credentials=sa_creds, project=BQ_PROJECT_ID)

# Google Sheets
if TOKEN_JSON:
    token_info = json.loads(TOKEN_JSON)
    user_creds = Credentials.from_authorized_user_info(token_info, SCOPES)
else:
    user_creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
gc = gspread.authorize(user_creds)

# ==========================================
# HÀM TẠO SUMMARY
# ==========================================
def get_po_expected_count():
    """
    Query BigQuery để lấy số PO cần tạo theo warehouse
    """
    sql = f"""
    SELECT 
        warehouse_id, 
        COUNT(DISTINCT partner_name) as total_partners,
        COUNT(DISTINCT po_pr_key) as total_pos
    FROM `{BQ_PROJECT_ID}.{BQ_DATASET_ID}.{BQ_TABLE_NAME}` 
    WHERE LOWER(frequency) = 'daily'
      AND DATE(po_tran_date) = CURRENT_DATE('Asia/Ho_Chi_Minh')
    GROUP BY 1
    ORDER BY 3 DESC
    """
    try:
        df = bq_client.query(sql).to_dataframe()
        return df
    except Exception as e:
        print(f"❌ Lỗi query BigQuery: {e}")
        return None

def get_run_metrics():
    """
    Lấy metrics từ file JSON (tạo trong lần chạy hiện tại)
    """
    try:
        if os.path.exists('run_metrics.json'):
            with open('run_metrics.json', 'r', encoding='utf-8') as f:
                metrics = json.load(f)
            pdf_created = metrics.get('pdf_created', 0)
            zalo_sent = metrics.get('zalo_sent', 0)
            return pdf_created, zalo_sent
        else:
            print("⚠️  Không tìm thấy file run_metrics.json (có thể PDF/Zalo job không chạy)")
            return 0, 0
    except Exception as e:
        print(f"❌ Lỗi đọc metrics: {e}")
        return 0, 0

def send_email_summary(po_expected_df, total_created, total_sent):
    """
    Gửi email summary
    """
    if not SMTP_PASSWORD:
        print("⚠️  SMTP_PASSWORD không được set, bỏ qua gửi email")
        return False
    
    try:
        # Tạo nội dung email
        timezone = pytz.timezone('Asia/Ho_Chi_Minh')
        today = datetime.now(timezone).strftime('%Y-%m-%d %H:%M:%S')
        
        html_content = f"""
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{ font-family: Arial, sans-serif; }}
                table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #4CAF50; color: white; }}
                .summary {{ font-size: 16px; margin: 20px 0; }}
                .stat {{ font-weight: bold; color: #2196F3; }}
            </style>
        </head>
        <body>
            <h2>📊 Workflow Summary Report</h2>
            <p><strong>Thời gian:</strong> {today}</p>
            
            <div class="summary">
                <p>📦 <span class="stat">Tổng PO cần tạo hôm nay:</span> {po_expected_df['total_pos'].sum() if po_expected_df is not None else 'N/A'} PO</p>
                <p>✅ <span class="stat">Tổng PO đã tạo:</span> {total_created} PO</p>
                <p>📤 <span class="stat">Tổng PO đã gửi Zalo:</span> {total_sent} PO</p>
            </div>
            
            <h3>Chi tiết theo Chi Nhánh:</h3>
        """
        
        if po_expected_df is not None and len(po_expected_df) > 0:
            html_content += "<table>"
            html_content += "<tr><th>Chi Nhánh</th><th>Số Partner</th><th>Số PO</th></tr>"
            for idx, row in po_expected_df.iterrows():
                html_content += f"<tr><td>{row['warehouse_id']}</td><td>{row['total_partners']}</td><td>{row['total_pos']}</td></tr>"
            html_content += "</table>"
        else:
            html_content += "<p>Không có dữ liệu PO cho hôm nay.</p>"
        
        html_content += """
            <hr>
            <p><em>Email tự động từ Workflow Job</em></p>
        </body>
        </html>
        """
        
        # Gửi email
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"[PACC Workflow] Daily Summary - {today[:10]}"
        msg['From'] = SMTP_USER
        msg['To'] = ', '.join(EMAIL_RECIPIENTS)
        
        msg.attach(MIMEText(html_content, 'html', 'utf-8'))
        
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, EMAIL_RECIPIENTS, msg.as_string())
        
        print(f"✅ Email gửi thành công tới: {', '.join(EMAIL_RECIPIENTS)}")
        return True
        
    except Exception as e:
        print(f"❌ Lỗi gửi email: {e}")
        return False

# ==========================================
# LOGIC CHÍNH
# ==========================================
def main():
    print("📊 Bắt đầu summarize log...\n")
    
    # 1. Lấy số PO cần tạo từ BigQuery
    print("🔍 Query số PO cần tạo từ BigQuery...")
    po_expected_df = get_po_expected_count()
    if po_expected_df is not None:
        print(f"✅ Tìm thấy {po_expected_df['total_pos'].sum()} PO cần tạo hôm nay")
    
    # 2. Lấy metrics từ lần chạy workflow này
    print("📊 Đọc metrics từ lần chạy này...")
    total_created, total_sent = get_run_metrics()
    print(f"✅ Lần chạy này: {total_created} PDF tạo, {total_sent} Zalo gửi")
    
    # 3. Gửi email summary
    print("📧 Gửi email summary...")
    send_email_summary(po_expected_df, total_created, total_sent)
    
    # 4. Xóa file metrics sau khi xong
    try:
        if os.path.exists('run_metrics.json'):
            os.remove('run_metrics.json')
            print("🧹 Đã xóa file metrics")
    except Exception as e:
        print(f"⚠️  Lỗi xóa metrics: {e}")
    
    print("✅ Summarize log hoàn tất!\n")

if __name__ == '__main__':
    main()
