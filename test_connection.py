#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script kiểm tra kết nối với Google Sheet
"""

import os
import sys
import json
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding for Windows console
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Load environment variables
load_dotenv()

def test_connection():
    """Kiểm tra kết nối với Google Sheet"""
    print("=" * 60)
    print("KIỂM TRA KẾT NỐI GOOGLE SHEET")
    print("=" * 60)
    
    # 1. Kiểm tra biến môi trường
    print("\n[1] Kiểm tra biến môi trường...")
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    json_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    json_content = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT")
    
    if not sheet_id:
        print("❌ GOOGLE_SHEET_ID không được tìm thấy trong .env")
        return False
    print(f"✅ GOOGLE_SHEET_ID: {sheet_id}")
    
    if not json_path and not json_content:
        print("❌ GOOGLE_SERVICE_ACCOUNT_JSON hoặc GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT không được tìm thấy")
        return False
    
    if json_path:
        print(f"✅ GOOGLE_SERVICE_ACCOUNT_JSON: {json_path}")
        if not os.path.exists(json_path):
            print(f"❌ File {json_path} không tồn tại")
            return False
        print(f"✅ File {json_path} tồn tại")
    
    if json_content:
        print("✅ GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT được tìm thấy")
    
    # 2. Kiểm tra Service Account JSON
    print("\n[2] Kiểm tra Service Account JSON...")
    try:
        if json_content:
            creds_info = json.loads(json_content)
            print("✅ Parse JSON từ GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT thành công")
        else:
            with open(json_path, 'r', encoding='utf-8') as f:
                creds_info = json.load(f)
            print("✅ Đọc file JSON thành công")
        
        # Kiểm tra các trường cần thiết
        required_fields = ['type', 'project_id', 'private_key_id', 'private_key', 'client_email']
        missing_fields = [field for field in required_fields if field not in creds_info]
        if missing_fields:
            print(f"❌ Thiếu các trường: {', '.join(missing_fields)}")
            return False
        print(f"✅ Service Account Email: {creds_info.get('client_email')}")
    except json.JSONDecodeError as e:
        print(f"❌ Lỗi parse JSON: {e}")
        return False
    except Exception as e:
        print(f"❌ Lỗi đọc Service Account: {e}")
        return False
    
    # 3. Khởi tạo Google Sheets client
    print("\n[3] Khởi tạo Google Sheets client...")
    try:
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(creds_info, scopes=scope)
        client = gspread.authorize(creds)
        print("✅ Khởi tạo client thành công")
    except Exception as e:
        print(f"❌ Lỗi khởi tạo client: {e}")
        return False
    
    # 4. Mở Google Sheet
    print("\n[4] Mở Google Sheet...")
    try:
        sheet = client.open_by_key(sheet_id)
        print(f"✅ Mở sheet thành công: {sheet.title}")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ Không tìm thấy sheet với ID: {sheet_id}")
        print("   → Kiểm tra lại GOOGLE_SHEET_ID hoặc quyền truy cập")
        return False
    except gspread.exceptions.APIError as e:
        print(f"❌ Lỗi API: {e}")
        print("   → Kiểm tra lại quyền truy cập của Service Account")
        return False
    except Exception as e:
        print(f"❌ Lỗi mở sheet: {e}")
        return False
    
    # 5. Kiểm tra các worksheet (tab)
    print("\n[5] Kiểm tra các worksheet (tab)...")
    try:
        worksheets = sheet.worksheets()
        print(f"✅ Tìm thấy {len(worksheets)} worksheet(s):")
        for ws in worksheets:
            print(f"   - {ws.title} (ID: {ws.id}, {ws.row_count} rows)")
        
        # Kiểm tra tab "codes"
        try:
            codes_ws = sheet.worksheet("codes")
            print(f"\n✅ Tab 'codes' tồn tại ({codes_ws.row_count} rows)")
            
            # Kiểm tra headers
            headers = codes_ws.row_values(1)
            print(f"   Headers: {', '.join(headers) if headers else '(trống)'}")
            
            # Đếm số codes
            all_records = codes_ws.get_all_records()
            print(f"   Số codes: {len(all_records)}")
            
            # Hiển thị 3 codes đầu tiên (nếu có)
            if all_records:
                print("   Một số codes mẫu:")
                for i, record in enumerate(all_records[:3], 1):
                    code = record.get('code', 'N/A')
                    status = record.get('status', 'N/A')
                    email = record.get('email', 'N/A')
                    print(f"     {i}. Code: {code}, Status: {status}, Email: {email}")
        except gspread.exceptions.WorksheetNotFound:
            print("❌ Tab 'codes' không tồn tại")
            print("   → Vui lòng tạo tab 'codes' với header 'code'")
        
        # Kiểm tra tab "activations"
        try:
            activations_ws = sheet.worksheet("activations")
            print(f"\n✅ Tab 'activations' tồn tại ({activations_ws.row_count} rows)")
            
            # Kiểm tra headers
            headers = activations_ws.row_values(1)
            print(f"   Headers: {', '.join(headers) if headers else '(trống)'}")
            
            # Đếm số activations
            all_records = activations_ws.get_all_records()
            print(f"   Số activations: {len(all_records)}")
        except gspread.exceptions.WorksheetNotFound:
            print("❌ Tab 'activations' không tồn tại")
            print("   → Vui lòng tạo tab 'activations' với headers: timestamp, code, email, team_id")
    
    except Exception as e:
        print(f"❌ Lỗi kiểm tra worksheets: {e}")
        return False
    
    print("\n" + "=" * 60)
    print("✅ KẾT NỐI THÀNH CÔNG!")
    print("=" * 60)
    return True

if __name__ == "__main__":
    try:
        success = test_connection()
        exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⚠️  Đã hủy bởi người dùng")
        exit(1)
    except Exception as e:
        print(f"\n\n❌ Lỗi không mong đợi: {e}")
        import traceback
        traceback.print_exc()
        exit(1)
