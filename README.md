# 📘 Manager-Team API Documentation / Tài liệu API (EN / VI)

[![Docs](https://img.shields.io/badge/docs-API-blue)](https://github.com/KenDzai1122/ManageTeam)
[![Version](https://img.shields.io/badge/version-1.0.0-brightgreen)]()
[![License](https://img.shields.io/badge/license-MIT-lightgrey)]()

Short, bilingual reference for the Manager-Team external API — English (EN) and Tiếng Việt (VI).

---

## Mini web kích hoạt (Python + Google Sheet)

VI: Repo này có kèm một mini web Python (Flask) để:
- User nhập **code kích hoạt** + **email**
- Server tra code trong **Google Sheet** (chỉ cần cột `code`)
- Tự kiểm tra team còn slot trống (tổng **members + pendingInvites < 5**) rồi gọi API **Invite Member**
- Ghi log **thời gian kích hoạt + email + code** vào Google Sheet

### 1) Chuẩn bị Google Sheet

Tạo 2 worksheet (tab) trong cùng 1 Google Sheet:

- Tab `codes` (dòng đầu là header):
  - Bắt buộc: `code` (VD: ABC123)
  - (App sẽ tự tạo/ghi thêm nếu thiếu): `activated_at`, `expires_at`, `email`, `team_id`, `status`, `error`

- Tab `activations` (dòng đầu là header):
  - `activated_at`
  - `code`
  - `email`
  - `team_id`

Tạo Google Service Account, tải file JSON và **share Google Sheet** cho email của service account (quyền Editor).

### 2) Cấu hình env

- Copy `.env.example` thành `.env`
- Điền:
  - `MANAGETEAM_AUTH` (credential gọi API, không lưu trên Google Sheet)
  - `CODE_TTL_MONTHS` (mặc định 3: code chỉ hợp lệ trong 3 tháng từ lần kích hoạt đầu)
  - `GOOGLE_SERVICE_ACCOUNT_JSON` (đường dẫn file JSON)
  - (deploy Railway khuyến nghị) `GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT` (dán toàn bộ nội dung JSON service account)
  - `GOOGLE_SHEET_ID` (ID của sheet trên URL)
  - (tuỳ chọn) SMTP nếu muốn gửi email xác nhận

### 3) Chạy

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python .\app.py
```

Mở `http://localhost:5000` và thử nhập code + email.

---

## Deploy lên Railway

- **Không commit secrets**: `.env`, `service-account*.json`, password/base64 auth… phải đưa vào **Railway Variables**.
- **Start command**: Railway sẽ dùng `Procfile` trong repo:
  - `gunicorn -w 2 -b 0.0.0.0:${PORT} app:app`
- **Google Service Account**:
  - Trên Railway, tạo biến `GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT` và **paste nguyên JSON** (1 dòng hoặc nhiều dòng đều được).
  - Không cần upload file `service-account.json` lên server.

---

## Table of contents / Mục lục
- [Quick links / Liên kết nhanh](#quick-links--liên-kết-nhanh)
- [Authentication / Xác thực](#authentication--xác-thực)
- [Base URL / Địa chỉ gốc](#base-url--địa-chỉ-gốc)
- [Endpoints / Các endpoint](#endpoints--các-endpoint)
  - [List Teams / Liệt kê team](#list-teams--liệt-kê-team)
  - [List Members / Liệt kê thành viên](#list-members--liệt-kê-thành-viên)
  - [Invite Member / Mời thành viên](#invite-member--mời-thành-viên)
  - [Kick Member / Hủy lời mời hoặc loại thành viên](#kick-member--hủy-lời-mời-hoặc-loại-thành-viên)
  - [Sync Team / Đồng bộ team](#sync-team--đồng-bộ-team)
- [Error codes / Mã lỗi](#error-codes--mã-lỗi)
- [Examples / Ví dụ](#examples--ví-dụ)
- [Notes / Ghi chú](#notes--ghi-chú)
- [Contributing / Đóng góp](#contributing--đóng-góp)
- [Changelog / Lịch sử](#changelog--lịch-sử)

---

## Quick links / Liên kết nhanh
- Repository: https://github.com/KenDzai1122/ManageTeam  
- Base docs file: `README.md` (this file)  
- Original doc: `API-Endpoint-EN.md` (kept for reference)

---

## Authentication / Xác thực

EN: The API uses credentials in the URL path. Two supported formats:
VI: API sử dụng credential trong đường dẫn URL. Hỗ trợ 2 định dạng:

- Plain text:
  - /api/{email}:{password}/...
- Base64 (recommended if password has special chars):
  - /api/base64:{base64_encoded_credentials}/...

How to encode:
```javascript
const email = 'hoangviet1991997@gmail.com';
const password = 'VietNhung9700';
const base64Auth = Buffer.from(`${email}:${password}`).toString('base64');
// Use: /api/base64:{base64Auth}/...
```

Notes:
- Use Base64 when password contains characters like `#`, `+`, `=`, `&`.  
- Không lưu credential công khai; lưu trữ an toàn.

---

## Base URL / Địa chỉ gốc

- Production: `https://kendev.id.vn/api`  
- Local: `http://localhost:3000/api`

---

## Endpoints / Các endpoint

> Each endpoint shows a short EN description followed by VI translation.

### List Teams / Liệt kê team
GET `/api/{auth}/teams`  
EN: Get teams owned by authenticated user.  
VI: Lấy danh sách team thuộc sở hữu người dùng.

Response (abridged):
```json
{
  "success": true,
  "data": { "teams": [ /* ... */ ], "total": 1 }
}
```

---

### List Members / Liệt kê thành viên
GET `/api/{auth}/{teamId}/list`  
EN: Returns members and pending invites.  
VI: Trả về thành viên và lời mời đang chờ.

Response includes `members`, `pendingInvites`, and `stats`.

---

### Invite Member / Mời thành viên
POST `/api/{auth}/{teamId}/invite/{memberEmail}`  
EN: Send invite to new member.  
VI: Gửi lời mời cho email.

Success response contains invited email and timestamp.

---

### Kick Member / Cancel Invite / Loại thành viên hoặc hủy lời mời
DELETE `/api/{auth}/{teamId}/kick/{memberEmail}`  
EN: Remove member or cancel invite (owner protected).  
VI: Xoá thành viên hoặc huỷ lời mời (chủ sở hữu được bảo vệ).

Response type indicates whether `member` or `invite` was affected.

---

### Sync Team / Đồng bộ team
POST `/api/{auth}/{teamId}/sync`  
EN: Sync data from ChatGPT API and trigger auto-kick if enabled.  
VI: Đồng bộ dữ liệu và thực thi auto-kick nếu bật.

Response includes `members`, `pendingInvites`, `autoKick` details and `lastSyncAt`.

---

## Error codes / Mã lỗi

| Code | HTTP | EN | VI |
|------|------|----|----|
| INVALID_CREDENTIALS | 401 | Invalid email/password | Email hoặc mật khẩu không đúng |
| TEAM_NOT_FOUND | 404 | Team not found / not owned | Không tìm thấy team hoặc không thuộc sở hữu |
| MEMBER_NOT_FOUND | 404 | Member/invite not found | Không tìm thấy thành viên/lời mời |
| INVITE_FAILED | 500 | Invite sending failed | Gửi lời mời thất bại |
| KICK_FAILED | 500 | Remove member failed | Xoá thành viên thất bại |
| SYNC_FAILED | 500 | Sync failed | Đồng bộ thất bại |
| TOKEN_EXPIRED | 401 | ChatGPT token expired | Token ChatGPT đã hết hạn |

Error format:
```json
{ "success": false, "error": "Error message", "code": "ERROR_CODE" }
```

---

## Examples / Ví dụ

cURL
```bash
# List members
curl "https://kendev.id.vn/api/user@email.com:password/TEAM_ID/list"

# Invite
curl -X POST "https://kendev.id.vn/api/user@email.com:password/TEAM_ID/invite/new@email.com"
```

JavaScript (fetch)
```javascript
const BASE = 'https://kendev.id.vn/api';
const AUTH = 'user@example.com:password';
const TEAM = 'your-team-id';

await fetch(`${BASE}/${AUTH}/${TEAM}/list`);
```

Python (requests)
```python
import requests
resp = requests.get(f'{BASE}/{AUTH}/{TEAM}/list')
```

---

## Notes / Ghi chú
- Always use HTTPS in production. / Luôn dùng HTTPS cho production.  
- Rate limiting may apply — avoid excessive calls. / Có thể có giới hạn tần suất.  
- Owner and whitelisted members are protected from auto-kick. / Chủ sở hữu và whitelist được bảo vệ.  
- Store credentials securely; do not embed in client-side code. / Lưu credential an toàn.

---

## Contributing / Đóng góp
- Suggest fixes via Pull Requests.  
- For translation improvements, send PR updating the VN phrasing.

---

## Changelog / Lịch sử
- v1.0 — 2026-01-25 — Bilingual refactor & formatting.

---

If you'd like, I can:
- convert this into the repository `README.md` automatically (generate a PowerShell script you can run), or
- create a `README.md` on a new branch and open a PR (you'll need to push the branch or grant collaborator access).

Tell me which action you want next and I'll prepare the exact commands/script.  
Nếu bạn muốn, tôi có thể: tạo script PowerShell để ghi file README.md, hoặc đẩy lên nhánh mới và tạo PR — bạn chọn bước tiếp theo.
