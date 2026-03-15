import os
import json
from datetime import datetime, timezone
from calendar import monthrange
from email.message import EmailMessage
import smtplib
import time

import requests
from requests.exceptions import RequestException, Timeout, ConnectionError, SSLError
import ssl
import socket
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request
import gspread
from google.oauth2.service_account import Credentials


load_dotenv()


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def bool_env(name: str, default: bool = False) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "y", "on"}


def get_sheets_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    sa_json = (os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT", "") or "").strip()
    if sa_json:
        try:
            info = json.loads(sa_json)
        except Exception as e:
            raise RuntimeError(f"Invalid GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT: {e}") from e
        creds = Credentials.from_service_account_info(info, scopes=scopes)
        return gspread.authorize(creds)

    sa_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not sa_path:
        raise RuntimeError("Missing env GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT")

    creds = Credentials.from_service_account_file(sa_path, scopes=scopes)
    return gspread.authorize(creds)


def open_worksheets():
    sheet_id = os.getenv("GOOGLE_SHEET_ID", "").strip()
    if not sheet_id:
        raise RuntimeError("Missing env GOOGLE_SHEET_ID")

    codes_tab = os.getenv("SHEET_CODES_TAB", "codes").strip() or "codes"
    activations_tab = os.getenv("SHEET_ACTIVATIONS_TAB", "activations").strip() or "activations"

    client = get_sheets_client()
    sh = client.open_by_key(sheet_id)
    return sh.worksheet(codes_tab), sh.worksheet(activations_tab)


def find_code_row(ws_codes, code: str):
    """
    Expected header columns (case-insensitive):
      code

    Optional columns (will be written back on activation):
      activated_at, email, team_id, status, error
    """
    code_lower = code.strip().lower()

    # Prefer header-based reading if header exists (case-insensitive).
    headers_raw = [h.strip() for h in ws_codes.row_values(1)]
    headers_lower = [h.lower() for h in headers_raw]
    if "code" in headers_lower:
        code_col = headers_lower.index("code") + 1  # 1-indexed
        col_values = ws_codes.col_values(code_col)
        for row_idx, value in enumerate(col_values, start=1):
            if row_idx == 1:
                continue  # header
            if str(value).strip().lower() == code_lower:
                row_values = ws_codes.row_values(row_idx)
                row: dict[str, str] = {}
                for i, header in enumerate(headers_raw):
                    key = header.strip().lower()
                    if not key:
                        continue
                    row[key] = str(row_values[i]).strip() if i < len(row_values) else ""
                if "code" not in row:
                    row["code"] = str(value).strip()
                return row_idx, row
        return None, None

    # Fallback: no header row (or missing 'code'). Treat column A as codes.
    col_a = ws_codes.col_values(1)
    for row_idx, value in enumerate(col_a, start=1):
        if str(value).strip().lower() == code_lower:
            return row_idx, {"code": value}
    return None, None


def normalize_auth(auth: str):
    base = os.getenv("MANAGETEAM_BASE_URL", "https://trandinhat.tokyo/api").rstrip("/")
    auth = auth.strip()
    if auth.lower().startswith("base64:"):
        return base, auth
    if ":" in auth:
        return base, auth
    return base, f"base64:{auth}"


def _request_with_cloudflare_retry(method: str, url: str, timeout: int = 30, retries: int = 3, backoff: float = 1.5, json=None):
    """
    Thực hiện HTTP request với retry khi:
    - Bị Cloudflare trả về trang "Just a moment..." (HTTP 403 + HTML challenge)
    - Gặp lỗi kết nối (ConnectionError, SSLError, Timeout)
    - Gặp lỗi khi đọc response body (SSL errors, socket errors)
    - Gặp lỗi ở mức thấp khi đọc status line hoặc response headers

    Điều này giúp giảm các lỗi lặt vặt do Cloudflare thỉnh thoảng bật
    challenge ngẫu nhiên hoặc các lỗi mạng tạm thời.
    
    Sử dụng tuple timeout (connect_timeout, read_timeout) để phát hiện
    lỗi kết nối nhanh hơn và tránh hang quá lâu.
    
    Args:
        json: Optional dict to send as JSON in the request body
    """
    last_exception = None
    last_resp = None
    
    # Sử dụng tuple timeout: (connect_timeout, read_timeout)
    # Connect timeout ngắn hơn để phát hiện lỗi kết nối nhanh
    # Read timeout dài hơn để đợi response từ server
    if isinstance(timeout, (int, float)):
        connect_timeout = min(10, timeout * 0.3)  # 30% của timeout hoặc tối đa 10s
        read_timeout = timeout
        timeout_tuple = (connect_timeout, read_timeout)
    else:
        timeout_tuple = timeout
    
    # Sử dụng session để kiểm soát connection tốt hơn
    session = requests.Session()
    # Tắt connection pooling để tránh reuse connection bị lỗi
    session.mount('http://', requests.adapters.HTTPAdapter(pool_connections=1, pool_maxsize=1, max_retries=0))
    session.mount('https://', requests.adapters.HTTPAdapter(pool_connections=1, pool_maxsize=1, max_retries=0))
    
    try:
        for attempt in range(retries):
            try:
                # Sử dụng stream=True để kiểm soát tốt hơn việc đọc response
                # Sử dụng tuple timeout để phát hiện lỗi kết nối nhanh hơn
                resp = session.request(method=method, url=url, timeout=timeout_tuple, stream=True, json=json)
                last_resp = resp
                
                # Đọc response body có thể gây ra lỗi SSL/socket
                # nếu connection bị đóng trong lúc đọc
                try:
                    text = resp.text or ""
                except (SSLError, ConnectionError, socket.error, ssl.SSLError, OSError, Exception) as read_err:
                    # Nếu lỗi khi đọc response body, coi như request failed và retry
                    last_exception = read_err
                    # Đóng response để cleanup
                    try:
                        resp.close()
                    except:
                        pass
                    if attempt < retries - 1:
                        time.sleep(backoff)
                        backoff *= 2
                        continue
                    raise
                
                # Nếu không phải 403, hoặc nội dung không giống trang challenge,
                # trả về luôn (để logic cũ xử lý).
                if resp.status_code != 403 or "Just a moment" not in text:
                    return resp
                
                # Nếu là 403 kiểu Cloudflare challenge và vẫn còn lượt retry,
                # đợi một chút rồi thử lại.
                if attempt < retries - 1:
                    time.sleep(backoff)
                    backoff *= 2
                    continue
                    
            except (ConnectionError, SSLError, Timeout, RequestException, socket.error, ssl.SSLError, OSError, Exception) as e:
                # Bỏ qua SystemExit và KeyboardInterrupt để không chặn shutdown
                if isinstance(e, (SystemExit, KeyboardInterrupt)):
                    raise
                last_exception = e
                # Nếu vẫn còn lượt retry, đợi một chút rồi thử lại
                if attempt < retries - 1:
                    time.sleep(backoff)
                    backoff *= 2
                    continue
                # Nếu hết lượt retry, raise exception cuối cùng
                raise
    finally:
        # Đóng session để cleanup connections
        try:
            session.close()
        except:
            pass
    
    # Nếu có response cuối cùng (dù là 403), trả về nó
    if last_resp is not None:
        return last_resp
    
    # Nếu không có response nào thành công, raise exception cuối cùng
    if last_exception is not None:
        raise last_exception
    
    # Fallback (không nên xảy ra)
    raise RuntimeError(f"Request failed after {retries} attempts")


def call_list_api(team_id: str, auth: str):
    base, path_auth = normalize_auth(auth)
    url = f"{base}/{path_auth}/{team_id}/list"
    # Timeout 20s với 2 retries = worst case ~60s per request
    resp = _request_with_cloudflare_retry("GET", url, timeout=20, retries=2)
    try:
        data = resp.json()
    except Exception:
        try:
            error_text = resp.text
        except (SSLError, ConnectionError, socket.error, ssl.SSLError, OSError):
            error_text = f"Failed to read response body (status={resp.status_code})"
        data = {"success": False, "error": error_text}

    if resp.status_code >= 400 or not data.get("success", False):
        code = data.get("code") or f"HTTP_{resp.status_code}"
        err = data.get("error") or "List failed"
        raise RuntimeError(f"{code}: {err}")

    return data


def call_teams_api(auth: str):
    base, path_auth = normalize_auth(auth)
    url = f"{base}/{path_auth}/teams"
    # Timeout 20s với 2 retries = worst case ~60s per request
    resp = _request_with_cloudflare_retry("GET", url, timeout=20, retries=2)
    try:
        data = resp.json()
    except Exception:
        try:
            error_text = resp.text
        except (SSLError, ConnectionError, socket.error, ssl.SSLError, OSError):
            error_text = f"Failed to read response body (status={resp.status_code})"
        data = {"success": False, "error": error_text}

    if resp.status_code >= 400 or not data.get("success", False):
        code = data.get("code") or f"HTTP_{resp.status_code}"
        err = data.get("error") or "Teams failed"
        raise RuntimeError(f"{code}: {err}")

    return data


def call_invite_api(team_id: str, auth: str, member_email: str):
    # Sử dụng endpoint mới: https://trandinhat.tokyo/api/public/add-member
    base = os.getenv("MANAGETEAM_BASE_URL", "https://trandinhat.tokyo/api").rstrip("/")
    url = f"{base}/public/add-member"
    # Timeout 20s với 2 retries = worst case ~60s per request
    resp = _request_with_cloudflare_retry("POST", url, timeout=20, retries=2, json={"email": member_email})
    try:
        data = resp.json()
    except Exception:
        try:
            error_text = resp.text
        except (SSLError, ConnectionError, socket.error, ssl.SSLError, OSError):
            error_text = f"Failed to read response body (status={resp.status_code})"
        data = {"success": False, "error": error_text}

    if resp.status_code >= 400 or not data.get("success", False):
        code = data.get("code") or f"HTTP_{resp.status_code}"
        err = data.get("error") or "Invite failed"
        raise RuntimeError(f"{code}: {err}")

    return data


def pick_team_with_capacity(auth: str, max_size: int = 5):
    teams_payload = call_teams_api(auth=auth)
    teams_data = (teams_payload.get("data") or {}).get("teams") or []
    team_ids = []
    for t in teams_data:
        tid = str(t.get("id") or t.get("_id") or t.get("teamId") or "").strip()
        if tid:
            team_ids.append(tid)

    if not team_ids:
        raise RuntimeError("Không lấy được danh sách teamId từ endpoint /teams.")

    last_err = None
    for team_id in team_ids:
        try:
            payload = call_list_api(team_id=team_id, auth=auth)
            data = payload.get("data") or {}
            members = data.get("members") or []
            # Một số API trả về pending dưới key `pendingInvites`,
            # một số khác là `pending`. Ưu tiên pendingInvites nhưng
            # luôn fallback sang pending để đảm bảo không vượt quá
            # giới hạn 5 người (tính cả pending).
            pending = data.get("pendingInvites")
            if pending is None:
                pending = data.get("pending") or []
            else:
                # nếu cả hai cùng tồn tại, cộng gộp.
                extra_pending = data.get("pending") or []
                if extra_pending:
                    pending = list(pending) + list(extra_pending)
            total = len(members) + len(pending)
            if total < max_size:
                return team_id, {"members": len(members), "pendingInvites": len(pending), "total": total}
        except Exception as e:
            last_err = str(e)
            continue
    if last_err:
        raise RuntimeError(f"Không tìm được team phù hợp. Lỗi cuối: {last_err}")
    raise RuntimeError("Không tìm được team phù hợp (tất cả team đã đủ 5 người tính cả pending).")


def invite_with_failover(auth: str, member_email: str, max_size: int):
    teams_payload = call_teams_api(auth=auth)
    teams_data = (teams_payload.get("data") or {}).get("teams") or []
    team_ids: list[str] = []
    # Lưu lại meta để có thể in ra tên team cùng với id
    team_meta: dict[str, dict] = {}
    for t in teams_data:
        tid = str(t.get("id") or t.get("_id") or t.get("teamId") or "").strip()
        if not tid:
            continue
        name = str(
            t.get("name")
            or t.get("teamName")
            or t.get("label")
            or t.get("title")
            or ""
        ).strip()
        team_ids.append(tid)
        team_meta[tid] = {"name": name}

    if not team_ids:
        raise RuntimeError("Không lấy được danh sách teamId từ endpoint /teams.")

    last_err = None
    tried: list[str] = []
    for team_id in team_ids:
        tried.append(team_id)
        try:
            cap = assert_team_has_capacity(team_id=team_id, auth=auth, max_size=max_size)
            invited_payload = call_invite_api(team_id=team_id, auth=auth, member_email=member_email)
            team_name = (team_meta.get(team_id) or {}).get("name") or ""
            # danh sách "name(id)" để debug / hiển thị
            pretty_tried: list[str] = []
            for tid in tried:
                meta = team_meta.get(tid) or {}
                tname = (meta.get("name") or "").strip()
                if tname:
                    pretty_tried.append(f"{tname}({tid})")
                else:
                    pretty_tried.append(tid)
            return {
                "team_id": team_id,
                "team_name": team_name,
                "capacity": cap,
                "invited": invited_payload,
                "tried_ids": tried,
                "tried": pretty_tried,
            }
        except Exception as e:
            last_err = str(e)
            continue

    msg = f"Không mời được vào bất kỳ team nào. Lỗi cuối: {last_err}"
    if tried:
        # thêm tên team cho dễ debug: "TeamName(id)"
        pretty_tried_err: list[str] = []
        for tid in tried[:8]:
            meta = team_meta.get(tid) or {}
            tname = (meta.get("name") or "").strip()
            if tname:
                pretty_tried_err.append(f"{tname}({tid})")
            else:
                pretty_tried_err.append(tid)
        msg += f" | tried={','.join(pretty_tried_err)}{'...' if len(tried) > 8 else ''}"
    raise RuntimeError(msg)

def ensure_code_sheet_columns(ws_codes, required_headers: list[str]):
    headers = ws_codes.row_values(1)
    headers_norm = [h.strip() for h in headers]

    # If A1 is empty, we can safely seed the primary header.
    if (len(headers_norm) == 0 or not headers_norm[0]) and "code" in required_headers:
        ws_codes.update_cell(1, 1, "code")
        headers = ws_codes.row_values(1)
        headers_norm = [h.strip() for h in headers]

    # Case-insensitive header matching to avoid duplicates like CODE vs code.
    headers_lower = [h.lower() for h in headers_norm]
    col_map: dict[str, int] = {}
    missing: list[str] = []
    for h in required_headers:
        hl = h.lower()
        if hl in headers_lower:
            col_map[h] = headers_lower.index(hl) + 1
        else:
            missing.append(h)
    if not missing:
        return col_map

    new_headers = headers_norm[:]
    for h in missing:
        new_headers.append(h)
    ws_codes.update("1:1", [new_headers])
    new_lower = [h.lower() for h in new_headers]
    return {h: new_lower.index(h.lower()) + 1 for h in required_headers}


def parse_iso_dt(value: str) -> datetime | None:
    s = (value or "").strip()
    if not s:
        return None
    try:
        dt = datetime.fromisoformat(s)
    except Exception:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def add_months(dt: datetime, months: int) -> datetime:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    y = dt.year + (dt.month - 1 + months) // 12
    m = (dt.month - 1 + months) % 12 + 1
    dmax = monthrange(y, m)[1]
    d = min(dt.day, dmax)
    return dt.replace(year=y, month=m, day=d)


def get_code_ttl_months() -> int:
    raw = (os.getenv("CODE_TTL_MONTHS", "") or "").strip()
    if not raw:
        return 3
    try:
        v = int(raw)
        return v if v > 0 else 3
    except Exception:
        return 3


def get_max_team_size() -> int:
    raw = (os.getenv("MAX_TEAM_SIZE", "") or "").strip()
    if not raw:
        return 5
    try:
        v = int(raw)
        return v if v > 0 else 5
    except Exception:
        return 5


def assert_team_has_capacity(team_id: str, auth: str, max_size: int) -> dict:
    payload = call_list_api(team_id=team_id, auth=auth)
    data = payload.get("data") or {}
    members = data.get("members") or []
    # Tương tự như pick_team_with_capacity: pending có thể nằm ở
    # `pendingInvites` hoặc `pending`. Đảm bảo tính đủ cả pending.
    pending = data.get("pendingInvites")
    if pending is None:
        pending = data.get("pending") or []
    else:
        extra_pending = data.get("pending") or []
        if extra_pending:
            pending = list(pending) + list(extra_pending)
    total = len(members) + len(pending)
    if total >= max_size:
        raise RuntimeError(
            f"Team đã đủ {max_size} (members={len(members)}, pending={len(pending)}). Vui lòng thử lại."
        )
    return {"members": len(members), "pendingInvites": len(pending), "total": total}


def maybe_send_smtp_email(to_email: str, activation_code: str):
    if not bool_env("ENABLE_SMTP_EMAIL", False):
        return

    host = os.getenv("SMTP_HOST", "").strip()
    port = int(os.getenv("SMTP_PORT", "587").strip() or "587")
    username = os.getenv("SMTP_USERNAME", "").strip()
    password = os.getenv("SMTP_PASSWORD", "").strip()
    from_addr = os.getenv("SMTP_FROM", username).strip() or username

    if not (host and username and password and from_addr):
        raise RuntimeError("SMTP is enabled but missing SMTP_HOST/SMTP_USERNAME/SMTP_PASSWORD/SMTP_FROM")

    msg = EmailMessage()
    msg["Subject"] = "Lời mời tham gia team"
    msg["From"] = from_addr
    msg["To"] = to_email
    msg.set_content(
        "\n".join(
            [
                "Bạn đã được gửi lời mời tham gia team.",
                f"Mã kích hoạt: {activation_code}",
                "",
                "Nếu bạn không yêu cầu, hãy bỏ qua email này.",
            ]
        )
    )

    with smtplib.SMTP(host, port) as smtp:
        smtp.starttls()
        smtp.login(username, password)
        smtp.send_message(msg)


app = Flask(__name__)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/activate")
def activate():
    payload = request.get_json(silent=True) if request.is_json else None
    payload = payload if isinstance(payload, dict) else {}

    code = (request.form.get("code") or payload.get("code") or "").strip()
    email = (request.form.get("email") or payload.get("email") or "").strip()

    if not code or not email:
        return jsonify({"success": False, "error": "Thiếu code hoặc email."}), 400

    ws_codes, ws_acts = open_worksheets()
    row_idx, row = find_code_row(ws_codes, code)
    if not row:
        return jsonify({"success": False, "error": "Code không tồn tại."}), 404

    ts = utc_now_iso()

    auth = (os.getenv("MANAGETEAM_AUTH", "") or "").strip()
    if not auth:
        return jsonify({"success": False, "error": "Thiếu env MANAGETEAM_AUTH."}), 500

    cols = ensure_code_sheet_columns(
        ws_codes,
        ["code", "activated_at", "expires_at", "email", "team_id", "team_name", "status", "error"],
    )

    existing_email = str(row.get("email", "")).strip()
    existing_activated_raw = str(row.get("activated_at", "")).strip()
    existing_expires_raw = str(row.get("expires_at", "")).strip()

    # Enforce: 1 code ↔ 1 email, TTL months from first activation.
    ttl_months = get_code_ttl_months()
    now_dt = parse_iso_dt(ts) or datetime.now(timezone.utc)

    if existing_email:
        if existing_email.lower() != email.lower():
            return jsonify({"success": False, "error": "Code đã được gán cho email khác."}), 409

    activated_dt = parse_iso_dt(existing_activated_raw)
    expires_dt = parse_iso_dt(existing_expires_raw)
    if activated_dt and not expires_dt:
        expires_dt = add_months(activated_dt, ttl_months)

    if expires_dt and now_dt > expires_dt:
        ws_codes.update_cell(row_idx, cols["status"], "expired")
        ws_codes.update_cell(row_idx, cols["error"], f"expired_at={expires_dt.isoformat(timespec='seconds')}")
        return jsonify({"success": False, "error": "Code đã hết hạn."}), 410

    try:
        max_size = get_max_team_size()
        invite_info = invite_with_failover(auth=auth, member_email=email, max_size=max_size)
        team_id = invite_info["team_id"]
        team_name = invite_info.get("team_name", "") or ""
        cap = invite_info["capacity"]
        tried = invite_info.get("tried_ids") or []

        ws_acts.append_row([ts, code, email, team_id], value_input_option="RAW")
        # Bind activation info on first activation.
        if not activated_dt:
            activated_dt = now_dt
            expires_dt = add_months(activated_dt, ttl_months)
            ws_codes.update_cell(row_idx, cols["activated_at"], activated_dt.isoformat(timespec="seconds"))
            ws_codes.update_cell(row_idx, cols["expires_at"], expires_dt.isoformat(timespec="seconds"))
        else:
            # If expires_at wasn't present, backfill.
            if expires_dt and not existing_expires_raw:
                ws_codes.update_cell(row_idx, cols["expires_at"], expires_dt.isoformat(timespec="seconds"))

        ws_codes.update_cell(row_idx, cols["email"], email)
        ws_codes.update_cell(row_idx, cols["team_id"], team_id)
        if "team_name" in cols:
            ws_codes.update_cell(row_idx, cols["team_name"], team_name)
        ws_codes.update_cell(
            row_idx,
            cols["status"],
            f"invited to {team_name or team_id} (total={cap['total']}, tried={len(tried)})",
        )
        ws_codes.update_cell(row_idx, cols["error"], "")
    except Exception as e:
        # If first activation attempt, bind email/time (so code can't be stolen later)
        if not activated_dt:
            activated_dt = now_dt
            expires_dt = add_months(activated_dt, ttl_months)
            ws_codes.update_cell(row_idx, cols["activated_at"], activated_dt.isoformat(timespec="seconds"))
            ws_codes.update_cell(row_idx, cols["expires_at"], expires_dt.isoformat(timespec="seconds"))
        ws_codes.update_cell(row_idx, cols["email"], email)
        ws_codes.update_cell(row_idx, cols["status"], "failed")
        ws_codes.update_cell(row_idx, cols["error"], str(e))
        return jsonify({"success": False, "error": str(e)}), 500

    maybe_send_smtp_email(to_email=email, activation_code=code)

    # Bổ sung thông tin team để client / log có thể xem được tên + id
    return jsonify(
        {
            "success": True,
            "invited": invite_info["invited"],
            "activatedAt": ts,
            "teamId": team_id,
            "teamName": team_name,
            "teamTried": invite_info.get("tried"),
        }
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=bool_env("FLASK_DEBUG", True))
