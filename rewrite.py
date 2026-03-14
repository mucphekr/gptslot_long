import re

file_path = "app.py"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Update normalize_auth
content = content.replace(
    'base = os.getenv("MANAGETEAM_BASE_URL", "https://kendev.id.vn/api").rstrip("/")',
    'base = os.getenv("MANAGETEAM_BASE_URL", "https://trandinhat.tokyo/api").rstrip("/")'
)

# 2. Update _request_with_cloudflare_retry signature
content = content.replace(
    'def _request_with_cloudflare_retry(method: str, url: str, timeout: int = 30, retries: int = 3, backoff: float = 1.5):',
    'def _request_with_cloudflare_retry(method: str, url: str, timeout: int = 30, retries: int = 3, backoff: float = 1.5, **kwargs):'
)

# 3. Update session.request call
content = content.replace(
    'resp = session.request(method=method, url=url, timeout=timeout_tuple, stream=True)',
    'resp = session.request(method=method, url=url, timeout=timeout_tuple, stream=True, **kwargs)'
)

# 4. Replace call_list_api, call_teams_api, call_invite_api, pick_team_with_capacity, invite_with_failover
# We'll regex replace from "def call_list_api" up to "def ensure_code_sheet_columns"

new_invite_with_failover = """
def invite_with_failover(auth: str, member_email: str, max_size: int):
    base = os.getenv("MANAGETEAM_BASE_URL", "https://trandinhat.tokyo/api").rstrip("/")
    url = f"{base}/public/add-member"
    
    resp = _request_with_cloudflare_retry("POST", url, timeout=30, retries=2, json={"email": member_email})
    try:
        data = resp.json()
    except Exception:
        try:
            error_text = resp.text
        except Exception:
            error_text = f"Failed to read response body (status={resp.status_code})"
        raise RuntimeError(f"Lỗi phản hồi từ server (HTTP {resp.status_code}): {error_text}")

    if resp.status_code >= 400 or not data.get("success"):
        err = data.get("error") or "Lỗi thêm thành viên"
        raise RuntimeError(f"{err}")

    team_name = data.get("team") or "Unknown Team"
    members_count = data.get("members") or "?/5"
    
    return {
        "team_id": team_name,
        "team_name": team_name,
        "capacity": {"total": members_count},
        "invited": data,
        "tried_ids": [team_name],
        "tried": [team_name]
    }

"""

# Regex replacing from def call_list_api up to def ensure_code_sheet_columns
content = re.sub(
    r'def call_list_api\(.*?def ensure_code_sheet_columns',
    new_invite_with_failover + 'def ensure_code_sheet_columns',
    content,
    flags=re.DOTALL
)

# 5. We also need to remove assert_team_has_capacity
content = re.sub(
    r'def assert_team_has_capacity\(.*?\n\n\n',
    '\n\n',
    content,
    flags=re.DOTALL
)

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Updated app.py successfully!")
