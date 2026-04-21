#!/usr/bin/env python3
"""Web app: paste Google Sheet link + drop image → auto extract + fill sheet."""

import os
import json
import base64
import re
from http.server import HTTPServer, BaseHTTPRequestHandler

import gspread
from google.oauth2.service_account import Credentials
from google import genai
from google.genai import types

# Load .env
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.exists(_env_path):
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

API_KEY = os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    raise SystemExit("Lỗi: Không tìm thấy GEMINI_API_KEY trong .env")

CREDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
PORT = int(os.environ.get("PORT", 8765))

PROMPT = """
Look at this image carefully. Find and extract these 4 fields:
1. PO No. (Purchase Order Number) - alphanumeric code like BJQ4003969
2. GR Date (Goods Receipt Date) - a date like 06/11/2024
3. WorkScope / Scope - short code like REP, RCT, OHC, RBTH, INS, etc.
4. Vendor Name - company/supplier name

Return ONLY a JSON object (no extra text, no markdown):
{"po_no": "...", "gr_date": "...", "scope": "...", "vendor_name": "..."}

If a field is not found, use null. Date must be in DD/MM/YYYY format.
"""

HTML = """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<title>Auto Fill Google Sheet từ ảnh</title>
<style>
  * { box-sizing: border-box; }
  body {
    font-family: 'Segoe UI', Arial, sans-serif;
    max-width: 680px; margin: 36px auto; padding: 20px;
    background: #f0f2f5; color: #222;
  }
  h2 { color: #1a73e8; margin-bottom: 4px; }
  .subtitle { color: #888; font-size: 13px; margin-bottom: 24px; }

  .card {
    background: white; border-radius: 12px; padding: 20px 22px;
    margin-bottom: 16px; box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }
  .card-title { font-weight: 600; color: #444; font-size: 13px; margin-bottom: 10px; text-transform: uppercase; letter-spacing: .5px; }

  #sheet-url {
    width: 100%; padding: 10px 12px; border: 1.5px solid #ddd;
    border-radius: 8px; font-size: 14px; transition: border-color .2s;
    outline: none;
  }
  #sheet-url:focus { border-color: #1a73e8; }
  #sheet-url.valid { border-color: #34a853; }
  #sheet-url.invalid { border-color: #ea4335; }
  #url-hint { font-size: 12px; margin-top: 6px; color: #888; min-height: 16px; }
  #url-hint.ok { color: #34a853; }
  #url-hint.err { color: #ea4335; }

  #drop-zone {
    border: 2.5px dashed #1a73e8; border-radius: 12px; padding: 44px 20px;
    text-align: center; cursor: pointer; background: #f8fbff;
    transition: all .2s; position: relative;
  }
  #drop-zone.over { background: #e8f0fe; border-color: #1557b0; }
  #drop-zone.has-image { border-style: solid; border-color: #34a853; background: #f6fef8; padding: 14px; }
  #drop-zone p { color: #666; margin: 0; font-size: 15px; }
  #preview { max-width: 100%; max-height: 220px; border-radius: 8px; display: none; margin: 0 auto; }
  #drop-hint { font-size: 12px; color: #aaa; margin-top: 6px; }
  input[type=file] { display: none; }

  /* Processing overlay */
  #processing {
    display: none; text-align: center; padding: 20px 0; color: #1a73e8;
  }
  .spinner {
    width: 36px; height: 36px; border: 3px solid #e8f0fe;
    border-top-color: #1a73e8; border-radius: 50%;
    animation: spin .8s linear infinite; margin: 0 auto 10px;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* Result card */
  #result-card { display: none; }
  .result-header {
    display: flex; align-items: center; gap: 8px;
    font-weight: 600; font-size: 15px; margin-bottom: 14px;
  }
  .badge-success {
    background: #e6f4ea; color: #137333; padding: 3px 10px;
    border-radius: 20px; font-size: 12px; font-weight: 600;
  }
  .badge-error {
    background: #fce8e6; color: #c5221f; padding: 3px 10px;
    border-radius: 20px; font-size: 12px; font-weight: 600;
  }
  .fields { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 14px; }
  .field-box {
    background: #f8f9fa; border-radius: 8px; padding: 10px 12px;
    border-left: 3px solid #1a73e8;
  }
  .field-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: .4px; margin-bottom: 3px; }
  .field-value { font-size: 14px; font-weight: 600; color: #222; word-break: break-all; }
  .field-value.missing { color: #bbb; font-weight: 400; font-style: italic; }

  .success-msg {
    background: #e6f4ea; border-radius: 8px; padding: 12px 14px;
    color: #137333; font-size: 14px; display: flex; align-items: center; gap: 8px;
  }
  .error-msg {
    background: #fce8e6; border-radius: 8px; padding: 12px 14px;
    color: #c5221f; font-size: 14px;
  }

  .btn-retry {
    background: #1a73e8; color: white; border: none; padding: 9px 20px;
    border-radius: 8px; font-size: 14px; cursor: pointer; margin-top: 12px;
  }
  .btn-retry:hover { background: #1557b0; }

  /* Setup guide */
  #setup-guide {
    background: #fff8e1; border-radius: 10px; padding: 16px 18px;
    margin-bottom: 16px; border-left: 4px solid #f9a825; display: none;
  }
  #setup-guide h4 { margin: 0 0 8px; color: #e65100; }
  #setup-guide ol { margin: 0; padding-left: 18px; font-size: 13px; color: #555; line-height: 1.8; }
  #setup-guide code { background: #eee; padding: 1px 5px; border-radius: 3px; font-size: 12px; }
</style>
</head>
<body>

<h2>📊 Auto Fill Google Sheet</h2>
<p class="subtitle">Kéo ảnh PO vào → tự động trích xuất và điền thông tin vào Google Sheet</p>

<div id="setup-guide">
  <h4>⚙️ Cần cài đặt Service Account lần đầu</h4>
  <ol>
    <li>Vào <a href="https://console.cloud.google.com/" target="_blank">Google Cloud Console</a> → Tạo project mới (hoặc dùng project cũ)</li>
    <li>Vào <strong>APIs & Services → Enable APIs</strong> → Bật <strong>Google Sheets API</strong> và <strong>Google Drive API</strong></li>
    <li>Vào <strong>APIs & Services → Credentials</strong> → <strong>Create Credentials → Service Account</strong></li>
    <li>Tạo xong → vào Service Account → tab <strong>Keys → Add Key → JSON</strong> → tải file về</li>
    <li>Đổi tên file thành <code>credentials.json</code> và đặt vào thư mục <code>st_en/</code></li>
    <li>Mở file <code>credentials.json</code>, copy email <code>client_email</code></li>
    <li>Mở Google Sheet → <strong>Share</strong> → paste email service account → <strong>Editor</strong></li>
  </ol>
</div>

<div class="card">
  <div class="card-title">🔗 Link Google Sheet</div>
  <input type="text" id="sheet-url" placeholder="Paste link Google Sheet vào đây..."
         value="https://docs.google.com/spreadsheets/d/1zZXbWk8fybdshMVkixN7oHU4Darj70Wb/edit?usp=sharing" />
  <div id="url-hint">Paste link Google Sheet của bạn</div>
</div>

<div class="card">
  <div class="card-title">🖼️ Ảnh PO</div>
  <div id="drop-zone">
    <p>Kéo thả ảnh vào đây<br><small style="color:#aaa">hoặc click để chọn file</small></p>
    <input type="file" id="file-input" accept="image/*">
    <img id="preview" alt="Preview">
    <div id="drop-hint" style="display:none">Click để đổi ảnh khác</div>
  </div>
</div>

<div id="processing" class="card">
  <div class="spinner"></div>
  <div id="processing-text">Đang đọc ảnh và điền vào Google Sheet...</div>
</div>

<div id="result-card" class="card">
  <div class="result-header">
    <span>Kết quả</span>
    <span id="result-badge"></span>
  </div>
  <div class="fields" id="fields-grid" style="display:none">
    <div class="field-box">
      <div class="field-label">PO No.</div>
      <div class="field-value" id="r-po"></div>
    </div>
    <div class="field-box">
      <div class="field-label">GR Date</div>
      <div class="field-value" id="r-date"></div>
    </div>
    <div class="field-box">
      <div class="field-label">Scope</div>
      <div class="field-value" id="r-scope"></div>
    </div>
    <div class="field-box" style="border-left-color:#34a853">
      <div class="field-label">Vendor Name</div>
      <div class="field-value" id="r-vendor"></div>
    </div>
  </div>
  <div id="result-msg"></div>
  <button class="btn-retry" onclick="resetImage()">📷 Xử lý ảnh khác</button>
</div>

<script>
let currentImageB64 = null;
let currentMediaType = null;
let isProcessing = false;

// URL validation
const urlInput = document.getElementById('sheet-url');
const urlHint = document.getElementById('url-hint');

urlInput.addEventListener('input', validateUrl);

function validateUrl() {
  const val = urlInput.value.trim();
  if (!val) {
    urlInput.className = '';
    urlHint.className = '';
    urlHint.textContent = 'Paste link Google Sheet của bạn';
    return false;
  }
  const match = val.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) {
    urlInput.className = 'valid';
    urlHint.className = 'ok';
    urlHint.textContent = '✓ Link hợp lệ (Sheet ID: ' + match[1].substring(0, 12) + '...)';
    return true;
  } else {
    urlInput.className = 'invalid';
    urlHint.className = 'err';
    urlHint.textContent = '✗ Link không đúng định dạng Google Sheets';
    return false;
  }
}

// Image drop zone
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const preview = document.getElementById('preview');

dropZone.addEventListener('click', () => { if (!isProcessing) fileInput.click(); });
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('over');
  if (e.dataTransfer.files[0] && !isProcessing) handleImage(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => { if (e.target.files[0]) handleImage(e.target.files[0]); });

function handleImage(file) {
  const reader = new FileReader();
  reader.onload = e => {
    const dataUrl = e.target.result;
    currentMediaType = file.type || 'image/png';
    currentImageB64 = dataUrl.split(',')[1];

    // Show preview
    preview.src = dataUrl;
    preview.style.display = 'block';
    dropZone.querySelector('p').style.display = 'none';
    dropZone.classList.add('has-image');
    document.getElementById('drop-hint').style.display = 'block';

    // Auto process
    processNow();
  };
  reader.readAsDataURL(file);
}

async function processNow() {
  if (!validateUrl()) {
    showError('Vui lòng nhập link Google Sheet hợp lệ trước khi tải ảnh.');
    return;
  }
  if (!currentImageB64) return;

  isProcessing = true;
  document.getElementById('result-card').style.display = 'none';
  document.getElementById('processing').style.display = 'block';

  try {
    const res = await fetch('/process', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        image: currentImageB64,
        media_type: currentMediaType,
        sheet_url: urlInput.value.trim()
      })
    });
    const data = await res.json();
    document.getElementById('processing').style.display = 'none';

    if (data.error) {
      showError(data.error);
    } else {
      showSuccess(data);
    }
  } catch (err) {
    document.getElementById('processing').style.display = 'none';
    showError('Lỗi kết nối: ' + err.message);
  }
  isProcessing = false;
}

function showSuccess(data) {
  document.getElementById('result-badge').innerHTML = '<span class="badge-success">✓ Thành công</span>';
  document.getElementById('fields-grid').style.display = 'grid';

  const setField = (id, val) => {
    const el = document.getElementById(id);
    if (val) { el.textContent = val; el.className = 'field-value'; }
    else { el.textContent = 'Không tìm thấy'; el.className = 'field-value missing'; }
  };
  setField('r-po', data.po_no);
  setField('r-date', data.gr_date);
  setField('r-scope', data.scope);
  setField('r-vendor', data.vendor_name);

  document.getElementById('result-msg').innerHTML =
    '<div class="success-msg">✅ Đã điền thành công vào dòng <strong>' + data.row + '</strong> trong Google Sheet!</div>';
  document.getElementById('result-card').style.display = 'block';
}

function showError(msg) {
  document.getElementById('result-badge').innerHTML = '<span class="badge-error">✗ Lỗi</span>';
  document.getElementById('fields-grid').style.display = 'none';
  document.getElementById('result-msg').innerHTML = '<div class="error-msg">❌ ' + msg + '</div>';
  document.getElementById('result-card').style.display = 'block';
}

function resetImage() {
  currentImageB64 = null;
  currentMediaType = null;
  isProcessing = false;
  preview.src = '';
  preview.style.display = 'none';
  dropZone.querySelector('p').style.display = 'block';
  dropZone.classList.remove('has-image');
  document.getElementById('drop-hint').style.display = 'none';
  document.getElementById('result-card').style.display = 'none';
  document.getElementById('processing').style.display = 'none';
  fileInput.value = '';
}

// Init
validateUrl();
</script>
</body>
</html>"""


def check_credentials():
    return os.path.exists(CREDS_FILE) or bool(os.environ.get("GOOGLE_CREDENTIALS_JSON"))


def extract_from_image(image_b64, media_type):
    client = genai.Client(api_key=API_KEY)
    image_part = types.Part.from_bytes(data=base64.b64decode(image_b64), mime_type=media_type)
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[PROMPT, image_part],
    )

    # For thinking models, collect only non-thought text parts
    text = None
    try:
        if response.candidates:
            parts_text = []
            for part in response.candidates[0].content.parts:
                if getattr(part, "thought", False):
                    continue
                if getattr(part, "text", None):
                    parts_text.append(part.text)
            if parts_text:
                text = "".join(parts_text)
    except Exception:
        pass

    # Fallback to response.text
    if not text:
        try:
            text = response.text
        except Exception:
            pass

    print(f"[DEBUG] raw text from Gemini: {repr(text)}")

    if not text or not text.strip():
        raise ValueError("Gemini trả về phản hồi rỗng hoặc bị chặn.")

    text = text.strip()
    # Strip markdown code fences if present
    text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\s*```\s*$', '', text)
    text = text.strip()

    print(f"[DEBUG] text after strip: {repr(text)}")

    match = re.search(r'\{.*\}', text, re.DOTALL)
    print(f"[DEBUG] regex match: {repr(match.group()) if match else None}")
    if match:
        return json.loads(match.group())
    raise ValueError(f"Không thể đọc dữ liệu từ ảnh: {text[:200]}")


def col_letter(n):
    """Convert 1-based column index to letter (1→A, 4→D, ...)."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def get_gspread_client():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if creds_json:
        creds_json = creds_json.strip()
        # Support both raw JSON and base64-encoded JSON
        if creds_json.startswith("{"):
            info = json.loads(creds_json)
        else:
            # Fix base64 padding if needed
            padding = 4 - len(creds_json) % 4
            if padding != 4:
                creds_json += "=" * padding
            info = json.loads(base64.b64decode(creds_json).decode())
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def fill_gsheet(data, sheet_url):
    gc = get_gspread_client()
    sh = gc.open_by_url(sheet_url)
    ws = sh.sheet1

    # Find first empty row in column D (PO_NO)
    col_d = ws.col_values(4)
    next_row = 2
    for i, val in enumerate(col_d):
        if i == 0:
            continue
        if not str(val).strip():
            next_row = i + 1
            break
        next_row = i + 2

    # Get row_num from col B
    row_num_cell = ws.cell(next_row, 2).value
    row_num = int(row_num_cell) if row_num_cell else (next_row - 1)

    # Write all 4 cells in one batch_update call
    updates = [
        {"range": f"{col_letter(4)}{next_row}", "values": [[data.get("po_no") or ""]]},
        {"range": f"{col_letter(6)}{next_row}", "values": [[data.get("gr_date") or ""]]},
        {"range": f"{col_letter(8)}{next_row}", "values": [[data.get("scope") or ""]]},
        {"range": f"{col_letter(10)}{next_row}", "values": [[data.get("vendor_name") or ""]]},
    ]
    ws.batch_update(updates, value_input_option="USER_ENTERED")

    return row_num


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"[{self.path}]", format % args)

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/debug":
            self.send_json({
                "GEMINI_API_KEY": bool(os.environ.get("GEMINI_API_KEY")),
                "GOOGLE_CREDENTIALS_JSON": bool(os.environ.get("GOOGLE_CREDENTIALS_JSON")),
                "CREDS_FILE_EXISTS": os.path.exists(CREDS_FILE),
                "check_credentials": check_credentials(),
            })
            return
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(HTML.encode())

    def do_POST(self):
        try:
            self._handle_post()
        except Exception as e:
            import traceback
            print(f"[do_POST CRASH] {traceback.format_exc()}")
            try:
                self.send_json({"error": f"Server error: {e}"}, 500)
            except Exception:
                pass

    def _handle_post(self):
        if self.path.rstrip("/") != "/process":
            self.send_json({"error": "Not found"}, 404)
            return

        length = int(self.headers.get("Content-Length", 0))
        raw_body = self.rfile.read(length)

        try:
            body = json.loads(raw_body)
        except json.JSONDecodeError:
            self.send_json({"error": "Body JSON không hợp lệ."}, 400)
            return

        required_fields = ["image", "media_type", "sheet_url"]
        missing_fields = [field for field in required_fields if not body.get(field)]
        if missing_fields:
            self.send_json(
                {"error": f"Thiếu trường bắt buộc: {', '.join(missing_fields)}."},
                400,
            )
            return

        if not check_credentials():
            self.send_json(
                {"error": "Chưa có file credentials.json. Xem hướng dẫn cài đặt Service Account ở góc trên trang."},
                500,
            )
            return

        try:
            print("[process] Đang đọc ảnh...")
            extracted = extract_from_image(body["image"], body["media_type"])
            print(f"[process] Đọc được: {extracted}")

            print("[process] Đang ghi vào Google Sheet...")
            row_num = fill_gsheet(extracted, body["sheet_url"])
            print(f"[process] Ghi thành công dòng {row_num}")

            self.send_json({**extracted, "row": row_num}, 200)

        except Exception as e:
            import traceback
            print(f"[process] ERROR: {traceback.format_exc()}")
            self.send_json({"error": str(e)}, 500)


if __name__ == "__main__":
    print(f"🚀 Mở trình duyệt tại: http://localhost:{PORT}")
    if not check_credentials():
        print("⚠️  Chưa có credentials.json — xem hướng dẫn trên trang web")
    print("Nhấn Ctrl+C để dừng.\n")
    HTTPServer(("", PORT), Handler).serve_forever()
