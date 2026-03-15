import subprocess
import json
import os
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

load_dotenv()

app = Flask(__name__)

# Google Sheets setup
def get_google_sheets_client():
    """Khởi tạo client Google Sheets"""
    try:
        # Ưu tiên dùng JSON content từ env (cho Railway/deploy)
        json_content = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT")
        if json_content:
            creds_info = json.loads(json_content)
        else:
            # Hoặc dùng file path
            json_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
            if not json_path:
                return None
            with open(json_path, 'r', encoding='utf-8') as f:
                creds_info = json.load(f)
        
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(creds_info, scopes=scope)
        return gspread.authorize(creds)
    except Exception as e:
        print(f"Lỗi khởi tạo Google Sheets client: {e}")
        return None

def get_sheet():
    """Lấy Google Sheet theo ID"""
    client = get_google_sheets_client()
    if not client:
        return None
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    if not sheet_id:
        return None
    try:
        return client.open_by_key(sheet_id)
    except Exception as e:
        print(f"Lỗi mở Google Sheet: {e}")
        return None

def find_code_in_sheet(code: str):
    """Tìm code trong tab 'codes' và trả về row index và data"""
    sheet = get_sheet()
    if not sheet:
        return None, None
    
    try:
        worksheet = sheet.worksheet("codes")
        all_records = worksheet.get_all_records()
        
        for idx, record in enumerate(all_records, start=2):  # start=2 vì row 1 là header
            # Thử cả "code" và "CODE" để tương thích với cả hai format
            code_value = record.get("code") or record.get("CODE") or ""
            if code_value.strip().upper() == code.strip().upper():
                return idx, record
        return None, None
    except Exception as e:
        print(f"Lỗi tìm code trong sheet: {e}")
        return None, None

def update_code_row(row_idx: int, email: str, team_id: str, status: str = "activated", error: str = ""):
    """Cập nhật row trong tab 'codes' với thông tin kích hoạt"""
    sheet = get_sheet()
    if not sheet:
        return False
    
    try:
        worksheet = sheet.worksheet("codes")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Lấy headers
        headers = worksheet.row_values(1)
        
        # Tạo dict để cập nhật
        updates = {}
        if "activated_at" in headers:
            updates["activated_at"] = now
        if "email" in headers:
            updates["email"] = email
        if "team_id" in headers:
            updates["team_id"] = team_id
        if "status" in headers:
            updates["status"] = status
        if "error" in headers:
            updates["error"] = error
        
        # Tính expires_at nếu cần
        ttl_months = int(os.getenv("CODE_TTL_MONTHS", "3"))
        if "expires_at" in headers and not updates.get("expires_at"):
            expires = datetime.now() + timedelta(days=ttl_months * 30)
            updates["expires_at"] = expires.strftime("%Y-%m-%d %H:%M:%S")
        
        # Cập nhật từng cell
        for col_name, value in updates.items():
            col_idx = headers.index(col_name) + 1
            worksheet.update_cell(row_idx, col_idx, value)
        
        return True
    except Exception as e:
        print(f"Lỗi cập nhật code row: {e}")
        return False

def log_activation(code: str, email: str, team_id: str):
    """Ghi log vào tab 'activations'"""
    sheet = get_sheet()
    if not sheet:
        return False
    
    try:
        worksheet = sheet.worksheet("activations")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Thêm row mới
        worksheet.append_row([now, code, email, team_id])
        return True
    except Exception as e:
        print(f"Lỗi ghi log activation: {e}")
        return False


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/add-member")
def add_member():
    data = request.get_json(silent=True) if request.is_json else None
    data = data if isinstance(data, dict) else {}
    
    email = (request.form.get("email") or data.get("email") or "").strip()
    code = (request.form.get("code") or data.get("code") or "").strip()
    
    if not email:
        return jsonify({"success": False, "error": "Vui lòng nhập email."}), 400
    
    if not code:
        return jsonify({"success": False, "error": "Vui lòng nhập code kích hoạt."}), 400
    
    # Kiểm tra code trong Google Sheets
    row_idx, code_record = find_code_in_sheet(code)
    if not code_record:
        return jsonify({"success": False, "error": "Code không hợp lệ hoặc không tồn tại."}), 400
    
    # Kiểm tra code đã được kích hoạt chưa
    if code_record.get("status", "").lower() == "activated":
        activated_email = code_record.get("email", "")
        if activated_email:
            return jsonify({"success": False, "error": f"Code đã được sử dụng bởi {activated_email}."}), 400
    
    # Kiểm tra code hết hạn
    expires_at = code_record.get("expires_at", "")
    if expires_at:
        try:
            expires_date = datetime.strptime(expires_at, "%Y-%m-%d %H:%M:%S")
            if datetime.now() > expires_date:
                return jsonify({"success": False, "error": "Code đã hết hạn."}), 400
        except:
            pass
    
    # Tạo lệnh PowerShell với email được thay thế
    command = [
        "powershell.exe",
        "-Command",
        f'$body = @{{email="{email}"}} | ConvertTo-Json; Invoke-RestMethod -Uri "https://trandinhat.tokyo/api/public/add-member" -Method Post -Headers @{{"Content-Type"="application/json"}} -Body $body'
    ]
    
    team_id = "Unknown"
    try:
        # Chạy lệnh PowerShell
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0:
            # Thử parse JSON response nếu có
            try:
                response_data = json.loads(result.stdout) if result.stdout.strip() else {}
                team_id = response_data.get("team", "Unknown")
                
                # Ghi log vào Google Sheets
                log_activation(code, email, team_id)
                update_code_row(row_idx, email, team_id, "activated", "")
                
                return jsonify({"success": True, "data": response_data, "message": "Thêm thành viên thành công!"})
            except Exception as parse_err:
                # Vẫn cố gắng ghi log nếu có thể
                log_activation(code, email, team_id)
                update_code_row(row_idx, email, team_id, "activated", "")
                return jsonify({"success": True, "message": "Thêm thành viên thành công!", "output": result.stdout})
        else:
            error_msg = result.stderr or result.stdout or "Lỗi không xác định"
            # Ghi lỗi vào Google Sheets
            update_code_row(row_idx, email, "", "error", error_msg)
            return jsonify({"success": False, "error": error_msg}), 500
            
    except subprocess.TimeoutExpired:
        update_code_row(row_idx, email, "", "error", "Timeout")
        return jsonify({"success": False, "error": "Lệnh chạy quá lâu (timeout)."}), 500
    except Exception as e:
        update_code_row(row_idx, email, "", "error", str(e))
        return jsonify({"success": False, "error": str(e)}), 500


if __name__ == "__main__":
    import os
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=True)
