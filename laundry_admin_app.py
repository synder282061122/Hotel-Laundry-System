"""
╔══════════════════════════════════════════════════════════════╗
║   Hotel Laundry Admin System  v3.0  — Full Edition           ║
║   ✅ OCR  ✅ Audit Logs  ✅ Duplicate Check                   ║
║   ✅ Editable Table  ✅ Monthly Dashboard                     ║
╚══════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import date, datetime
import os, base64, json, io
import anthropic
from PIL import Image
import fitz  # PyMuPDF

# ══════════════════════════════════════════════
#  PAGE CONFIG
# ══════════════════════════════════════════════
st.set_page_config(page_title="Laundry Admin System", page_icon="🧺", layout="wide")

# ══════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════
ADMIN_PASSWORD   = "admin1234"
SPREADSHEET_NAME = "HotelLaundryDB"
WS_LAUNDRY       = "Laundry"
WS_AUDITLOG      = "AuditLog"

DEPARTMENTS = [
    "HK (Housekeeping)", "FB (Food & Beverage)", "Spa",
    "Kitchen", "Engineering", "Front Office", "HR", "อื่นๆ",
]
LINEN_ITEMS = [
    "ผ้าปูที่นอน Single", "ผ้าปูที่นอน Double", "ผ้าปูที่นอน King",
    "ปลอกหมอน", "ผ้านวม", "ผ้าขนหนูใหญ่", "ผ้าขนหนูกลาง",
    "ผ้าขนหนูเล็ก", "ผ้าเช็ดหน้า", "ผ้ากันเปื้อน", "ผ้าเช็ดโต๊ะ",
    "ผ้าเช็ดแก้ว", "ชุด Uniform", "อื่นๆ",
]

# Main data columns (Audit timestamp อยู่ท้ายสุดเสมอ)
LAUNDRY_COLS = [
    "วันที่บันทึก", "วันที่ในบิล", "เลขที่บิล", "แผนก",
    "รายการผ้า", "จำนวน", "ราคาต่อหน่วย (บาท)", "ยอดรวม (บาท)",
    "ชื่อไฟล์อ้างอิง", "หมายเหตุ", "แก้ไขล่าสุด",          # ← Audit col
]
AUDITLOG_COLS = [
    "Timestamp", "Action", "เลขที่บิล", "แผนก", "รายละเอียด", "ผู้ดำเนินการ",
]

THAI_MONTHS = {
    1:"มกราคม",2:"กุมภาพันธ์",3:"มีนาคม",4:"เมษายน",
    5:"พฤษภาคม",6:"มิถุนายน",7:"กรกฎาคม",8:"สิงหาคม",
    9:"กันยายน",10:"ตุลาคม",11:"พฤศจิกายน",12:"ธันวาคม",
}

# ══════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════
_defaults = {
    "logged_in": False, "gc": None,
    "ws_laundry": None, "ws_audit": None,
    "ocr_result": None, "uploaded_filename": "",
    "entry_mode": "ocr",
    "edit_row_idx": None,   # index ของแถวที่กำลัง edit
    "confirm_delete": None, # index รอยืนยันลบ
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ══════════════════════════════════════════════
#  GOOGLE SHEETS HELPERS
# ══════════════════════════════════════════════
def _get_or_create_ws(spreadsheet, name: str, rows: int, cols: int, header: list):
    try:
        ws = spreadsheet.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=name, rows=rows, cols=cols)
        ws.append_row(header)
    if not ws.row_values(1):
        ws.append_row(header)
    return ws


def connect_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds",
                 "https://www.googleapis.com/auth/drive"]
        if "gcp_service_account" in st.secrets:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(
                dict(st.secrets["gcp_service_account"]), scope)
        elif os.path.exists("credentials.json"):
            creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
        else:
            st.error("❌ ไม่พบ credentials!"); return None, None, None

        gc = gspread.authorize(creds)
        try:    sp = gc.open(SPREADSHEET_NAME)
        except gspread.SpreadsheetNotFound:
            sp = gc.create(SPREADSHEET_NAME)

        ws_l = _get_or_create_ws(sp, WS_LAUNDRY,  3000, len(LAUNDRY_COLS),  LAUNDRY_COLS)
        ws_a = _get_or_create_ws(sp, WS_AUDITLOG, 5000, len(AUDITLOG_COLS), AUDITLOG_COLS)
        return gc, ws_l, ws_a
    except Exception as e:
        st.error(f"❌ เชื่อมต่อ Google Sheets ไม่สำเร็จ: {e}")
        return None, None, None


def load_df(ws) -> pd.DataFrame:
    try:
        recs = ws.get_all_records()
        return pd.DataFrame(recs) if recs else pd.DataFrame(columns=LAUNDRY_COLS)
    except Exception as e:
        st.error(f"❌ โหลดข้อมูลไม่สำเร็จ: {e}"); return pd.DataFrame(columns=LAUNDRY_COLS)


def load_audit_df(ws) -> pd.DataFrame:
    try:
        recs = ws.get_all_records()
        return pd.DataFrame(recs) if recs else pd.DataFrame(columns=AUDITLOG_COLS)
    except Exception:
        return pd.DataFrame(columns=AUDITLOG_COLS)


def append_row_laundry(ws, row: list) -> bool:
    """เพิ่มแถว + เติม Audit timestamp ท้ายสุด"""
    # ตรวจให้มีครบ LAUNDRY_COLS columns
    while len(row) < len(LAUNDRY_COLS) - 1:
        row.append("")
    # คอลัมน์สุดท้าย "แก้ไขล่าสุด" — ตอน INSERT ใช้ค่า "สร้างใหม่"
    row_with_audit = list(row) + ["สร้างใหม่"]
    try:
        ws.append_row(row_with_audit, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"❌ บันทึกไม่สำเร็จ: {e}"); return False


def update_row_laundry(ws, sheet_row_idx: int, new_values: list):
    """อัปเดตแถว (sheet_row_idx คือ 1-based row ใน Sheets รวม header)"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # ตรวจให้ครบ
    while len(new_values) < len(LAUNDRY_COLS) - 1:
        new_values.append("")
    new_values_with_ts = list(new_values) + [timestamp]
    try:
        ws.update(f"A{sheet_row_idx}:{chr(64+len(LAUNDRY_COLS))}{sheet_row_idx}",
                  [new_values_with_ts])
        return True
    except Exception as e:
        st.error(f"❌ อัปเดตไม่สำเร็จ: {e}"); return False


def delete_row_laundry(ws, sheet_row_idx: int) -> bool:
    try:
        ws.delete_rows(sheet_row_idx); return True
    except Exception as e:
        st.error(f"❌ ลบไม่สำเร็จ: {e}"); return False


def write_audit(ws_audit, action: str, bill_no: str, dept: str, detail: str):
    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        action, bill_no, dept, detail, "Admin",
    ]
    try: ws_audit.append_row(row, value_input_option="USER_ENTERED")
    except Exception: pass   # ไม่ให้ crash หลักถ้า audit sheet มีปัญหา


def check_duplicate(ws, bill_no: str) -> bool:
    """True = มีเลขที่บิลซ้ำอยู่แล้ว"""
    try:
        cells = ws.findall(bill_no)
        # เช็กว่า cell อยู่คอลัมน์ "เลขที่บิล" (col 3)
        return any(c.col == 3 for c in cells)
    except Exception:
        return False


# ══════════════════════════════════════════════
#  CLAUDE VISION OCR
# ══════════════════════════════════════════════
OCR_SYSTEM_PROMPT = """
You are an OCR assistant for a hotel laundry billing system.
Extract information from the bill/receipt image and return ONLY valid JSON — no markdown, no code fences.

JSON schema:
{
  "bill_date":   "YYYY-MM-DD or empty string",
  "bill_no":     "invoice number or empty string",
  "department":  "department or empty string",
  "linen_items": [{"item":"","qty":null,"unit_price":null,"total":null}],
  "grand_total": null,
  "confidence":  "high|medium|low",
  "raw_text":    "full OCR text"
}
Rules: numeric fields = numbers not strings. Dates = YYYY-MM-DD. Return ONLY JSON.
"""

def get_anthropic_client():
    api_key = None
    try:
        if "anthropic_api_key" in st.secrets:
            api_key = st.secrets["anthropic_api_key"]
    except Exception: pass
    if not api_key: api_key = os.environ.get("ANTHROPIC_API_KEY")
    return anthropic.Anthropic(api_key=api_key) if api_key else None


def pdf_to_image_bytes(pdf_bytes: bytes) -> bytes:
    doc  = fitz.open(stream=pdf_bytes, filetype="pdf")
    pix  = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
    return pix.tobytes("png")


def run_claude_ocr(img_bytes: bytes, media_type: str) -> dict:
    client = get_anthropic_client()
    if not client: raise ValueError("ไม่พบ Anthropic API Key")
    b64 = base64.standard_b64encode(img_bytes).decode()
    resp = client.messages.create(
        model="claude-opus-4-5", max_tokens=1500,
        system=OCR_SYSTEM_PROMPT,
        messages=[{"role":"user","content":[
            {"type":"image","source":{"type":"base64","media_type":media_type,"data":b64}},
            {"type":"text","text":"Extract billing information from this laundry bill."}
        ]}],
    )
    raw = resp.content[0].text.strip()
    if raw.startswith("```"):
        raw = "\n".join(raw.split("\n")[1:]).rstrip("`").strip()
    return json.loads(raw)


def map_department(s: str) -> str:
    if not s: return DEPARTMENTS[0]
    d = s.upper()
    for k,v in {"HK":"HK (Housekeeping)","HOUSEKEEPING":"HK (Housekeeping)",
                "FB":"FB (Food & Beverage)","F&B":"FB (Food & Beverage)",
                "FOOD":"FB (Food & Beverage)","SPA":"Spa","KITCHEN":"Kitchen",
                "ENG":"Engineering","FRONT":"Front Office","FO":"Front Office",
                "HR":"HR"}.items():
        if k in d: return v
    return "อื่นๆ"


# ══════════════════════════════════════════════
#  LOGIN PAGE
# ══════════════════════════════════════════════
def login_page():
    c1, c2, c3 = st.columns([1,1.5,1])
    with c2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
        <div style='text-align:center;padding:2.5rem;background:#f8f9fa;
                    border-radius:14px;border:1px solid #dee2e6;box-shadow:0 2px 12px #0001'>
            <h1 style='color:#1a73e8;margin:0'>🧺 Laundry Admin</h1>
            <p style='color:#888;margin:.5rem 0 0'>Hotel Laundry Management System v3</p>
        </div>""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        with st.form("login"):
            pw = st.text_input("🔑 Admin Password", type="password")
            ok = st.form_submit_button("เข้าสู่ระบบ", use_container_width=True, type="primary")
        if ok:
            if pw == ADMIN_PASSWORD:
                st.session_state.logged_in = True
                with st.spinner("กำลังเชื่อมต่อ Google Sheets..."):
                    gc, ws_l, ws_a = connect_google_sheets()
                    st.session_state.gc         = gc
                    st.session_state.ws_laundry = ws_l
                    st.session_state.ws_audit   = ws_a
                st.rerun()
            else:
                st.error("❌ Password ไม่ถูกต้อง")


# ══════════════════════════════════════════════
#  SHARED: REVIEW / MANUAL FORM  (with Duplicate Check)
# ══════════════════════════════════════════════
def render_entry_form(ws_l, ws_a, ocr_data=None, filename=""):
    is_ocr = ocr_data is not None

    if is_ocr:
        conf  = ocr_data.get("confidence","low")
        badge = {"high":"🟢","medium":"🟡","low":"🔴"}.get(conf,"⚪")
        st.markdown(f"### 📝 Step 2 — ตรวจสอบ & แก้ไขข้อมูล  {badge} ความมั่นใจ AI: **{conf.upper()}**")
        st.info("💡 ตรวจสอบทุกช่องก่อนกดบันทึก")
    else:
        st.markdown("### 📝 กรอกข้อมูลบิล")

    # defaults
    if is_ocr:
        try:    def_date = datetime.strptime(ocr_data.get("bill_date",""),"%Y-%m-%d").date()
        except: def_date = date.today()
        def_bill  = ocr_data.get("bill_no","") or ""
        def_dept  = map_department(ocr_data.get("department",""))
        items     = ocr_data.get("linen_items",[{}]); fi = items[0] if items else {}
        _ri       = fi.get("item","อื่นๆ")
        def_item  = _ri if _ri in LINEN_ITEMS else "อื่นๆ"
        def_qty   = max(1, int(fi.get("qty") or 1))
        def_price = float(fi.get("unit_price") or 0.0)
    else:
        def_date  = date.today(); def_bill = ""; def_dept = DEPARTMENTS[0]
        def_item  = LINEN_ITEMS[0]; def_qty = 1; def_price = 0.0

    with st.form("entry_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            bill_date  = st.date_input("📅 วันที่ในบิล", value=def_date)
            bill_no    = st.text_input("🧾 เลขที่บิล", value=def_bill,
                                        placeholder="เช่น INV-2025-001")
            di = DEPARTMENTS.index(def_dept) if def_dept in DEPARTMENTS else 0
            department = st.selectbox("🏨 แผนก", DEPARTMENTS, index=di)
        with c2:
            ii = LINEN_ITEMS.index(def_item) if def_item in LINEN_ITEMS else len(LINEN_ITEMS)-1
            linen_item = st.selectbox("👕 รายการผ้า", LINEN_ITEMS, index=ii)
            qty        = st.number_input("📦 จำนวน (ชิ้น)", min_value=1, value=def_qty, step=1)
            unit_price = st.number_input("💰 ราคาต่อหน่วย (บาท)", min_value=0.0,
                                          value=float(def_price), step=0.50, format="%.2f")
        ref_file = st.text_input("📎 ชื่อไฟล์อ้างอิง", value=filename)
        note     = st.text_input("📝 หมายเหตุ")

        total = qty * unit_price
        st.metric("💵 ยอดรวม", f"฿{total:,.2f}")

        if is_ocr and ocr_data.get("raw_text"):
            with st.expander("📄 Raw Text จาก AI"):
                st.text(ocr_data["raw_text"])
        if is_ocr and len(ocr_data.get("linen_items",[])) > 1:
            with st.expander(f"📋 รายการอื่น ({len(ocr_data['linen_items'])-1} รายการ)"):
                for i, it in enumerate(ocr_data["linen_items"][1:], 2):
                    st.write(f"**{i}.** {it.get('item','-')} | qty:{it.get('qty','-')} | ฿{it.get('unit_price','-')}")

        st.markdown("---")
        cs, cc = st.columns([3,1])
        save_btn  = cs.form_submit_button("✅ ยืนยันและบันทึกลง Google Sheets",
                                           use_container_width=True, type="primary")
        clear_btn = cc.form_submit_button("🔄 ล้าง", use_container_width=True)

    if save_btn:
        if not bill_no.strip():
            st.error("❌ กรุณากรอกเลขที่บิล"); return

        # ── DUPLICATE CHECK ──
        if check_duplicate(ws_l, bill_no.strip()):
            st.warning(
                f"⚠️ **เลขที่บิล `{bill_no}` มีอยู่แล้วในระบบ!**  \n"
                "หากต้องการบันทึกซ้ำจริง กรุณาแก้ไขเลขที่บิลก่อน"
            )
            return

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            bill_date.strftime("%Y-%m-%d"),
            bill_no.strip(), department, linen_item,
            int(qty), float(unit_price), float(total),
            ref_file.strip(), note.strip(),
        ]
        with st.spinner("⏳ กำลังบันทึก..."):
            ok = append_row_laundry(ws_l, row)
        if ok:
            write_audit(ws_a, "INSERT", bill_no.strip(), department,
                        f"{linen_item} x{qty} = ฿{total:,.2f}")
            st.success(f"✅ บันทึกสำเร็จ! บิล `{bill_no}` | ฿{total:,.2f}")
            st.balloons()
            st.session_state.ocr_result = None
            st.session_state.uploaded_filename = ""

    if clear_btn:
        st.session_state.ocr_result = None
        st.session_state.uploaded_filename = ""
        st.rerun()


# ══════════════════════════════════════════════
#  PAGE 1: บันทึกรายการใหม่
# ══════════════════════════════════════════════
def page_new_entry():
    ws_l = st.session_state.ws_laundry
    ws_a = st.session_state.ws_audit
    st.title("📋 บันทึกรายการซักรีด")
    if not ws_l:
        st.warning("⚠️ ยังไม่ได้เชื่อมต่อ Google Sheets"); return

    c1, c2 = st.columns(2)
    with c1:
        if st.button("🤖 OCR อ่านบิลอัตโนมัติ", use_container_width=True,
                     type="primary" if st.session_state.entry_mode=="ocr" else "secondary"):
            st.session_state.entry_mode = "ocr"; st.session_state.ocr_result = None; st.rerun()
    with c2:
        if st.button("✏️ กรอกข้อมูลเอง (Manual)", use_container_width=True,
                     type="primary" if st.session_state.entry_mode=="manual" else "secondary"):
            st.session_state.entry_mode = "manual"; st.session_state.ocr_result = None; st.rerun()
    st.markdown("---")

    if st.session_state.entry_mode == "ocr":
        st.subheader("🤖 Step 1 — อัปโหลดบิล")
        has_api = get_anthropic_client() is not None
        if not has_api:
            st.error("❌ ไม่พบ Anthropic API Key — ใช้ Manual แทน หรือเพิ่ม `anthropic_api_key` ใน secrets")

        uploaded = st.file_uploader("เลือกบิล (PDF / JPG / PNG)",
                                     type=["pdf","jpg","jpeg","png"], key="uploader")
        if uploaded:
            st.session_state.uploaded_filename = uploaded.name
            fb = uploaded.read()
            if uploaded.type == "application/pdf":
                with st.spinner("แปลง PDF..."):
                    ib = pdf_to_image_bytes(fb); mt = "image/png"
            else:
                ib = fb; mt = uploaded.type

            ci, cb = st.columns([2,1])
            ci.image(Image.open(io.BytesIO(ib)), caption=uploaded.name, use_container_width=True)
            with cb:
                st.markdown("<br><br>", unsafe_allow_html=True)
                if st.button("🔍 วิเคราะห์ด้วย AI", use_container_width=True,
                             type="primary", disabled=not has_api):
                    with st.spinner("🤖 กำลังอ่านบิล..."):
                        try:
                            st.session_state.ocr_result = run_claude_ocr(ib, mt)
                            st.success("✅ อ่านสำเร็จ!")
                        except json.JSONDecodeError:
                            st.error("❌ AI ตอบรูปแบบผิด ลองใหม่หรือใช้ Manual")
                        except Exception as e:
                            st.error(f"❌ {e}")

        if st.session_state.ocr_result is not None:
            st.markdown("---")
            render_entry_form(ws_l, ws_a, st.session_state.ocr_result,
                              st.session_state.uploaded_filename)
    else:
        st.subheader("✏️ กรอกข้อมูลด้วยตนเอง")
        render_entry_form(ws_l, ws_a, ocr_data=None, filename="")


# ══════════════════════════════════════════════
#  PAGE 2: DATA MANAGEMENT (Editable Table + Delete)
# ══════════════════════════════════════════════
def page_manage():
    ws_l = st.session_state.ws_laundry
    ws_a = st.session_state.ws_audit
    st.title("🗂️ Data Management")
    if not ws_l:
        st.warning("⚠️ ยังไม่ได้เชื่อมต่อ Google Sheets"); return

    with st.spinner("โหลดข้อมูล..."):
        df = load_df(ws_l)

    if df.empty:
        st.info("📭 ยังไม่มีข้อมูล"); return

    for c in ["จำนวน","ราคาต่อหน่วย (บาท)","ยอดรวม (บาท)"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # ── Search bar ──
    search = st.text_input("🔍 ค้นหาเลขที่บิล / แผนก / รายการ", placeholder="พิมพ์เพื่อค้นหา...")
    disp = df.copy()
    if search:
        mask = (
            disp["เลขที่บิล"].astype(str).str.contains(search, case=False, na=False) |
            disp["แผนก"].astype(str).str.contains(search, case=False, na=False) |
            disp["รายการผ้า"].astype(str).str.contains(search, case=False, na=False)
        )
        disp = disp[mask]

    st.caption(f"แสดง {len(disp):,} จาก {len(df):,} รายการ")
    st.markdown("---")

    # ── Render rows with Edit / Delete actions ──
    edit_idx    = st.session_state.edit_row_idx
    confirm_del = st.session_state.confirm_delete

    for i, row in disp.iterrows():
        sheet_row = i + 2   # +2 เพราะ header อยู่ row 1, pandas 0-indexed

        with st.container():
            # Header row
            hc1, hc2, hc3, hc4, hc5, hc6 = st.columns([1,2,2,2,2,2])
            hc1.caption(f"#{i}")
            hc2.write(f"**{row.get('เลขที่บิล','')}**")
            hc3.write(f"🏨 {row.get('แผนก','')}")
            hc4.write(f"📅 {row.get('วันที่ในบิล','')}")
            hc5.write(f"💵 ฿{float(row.get('ยอดรวม (บาท)',0)):,.2f}")

            with hc6:
                ec, dc = st.columns(2)
                if ec.button("✏️", key=f"edit_{i}", help="แก้ไข"):
                    st.session_state.edit_row_idx = i if edit_idx != i else None
                    st.session_state.confirm_delete = None
                    st.rerun()
                if dc.button("🗑️", key=f"del_{i}", help="ลบ"):
                    st.session_state.confirm_delete = i if confirm_del != i else None
                    st.session_state.edit_row_idx = None
                    st.rerun()

            # ── CONFIRM DELETE ──
            if confirm_del == i:
                st.warning(f"⚠️ ยืนยันลบบิล **{row.get('เลขที่บิล','')}** ?")
                ya, na = st.columns(2)
                if ya.button("✅ ยืนยันลบ", key=f"conf_del_{i}", type="primary"):
                    if delete_row_laundry(ws_l, sheet_row):
                        write_audit(ws_a, "DELETE",
                                    str(row.get("เลขที่บิล","")),
                                    str(row.get("แผนก","")),
                                    f"ลบแถว row={sheet_row}")
                        st.success("✅ ลบสำเร็จ")
                        st.session_state.confirm_delete = None
                        st.rerun()
                if na.button("❌ ยกเลิก", key=f"cancel_del_{i}"):
                    st.session_state.confirm_delete = None
                    st.rerun()

            # ── INLINE EDIT FORM ──
            if edit_idx == i:
                with st.form(f"edit_form_{i}"):
                    st.markdown(f"##### ✏️ แก้ไขบิล `{row.get('เลขที่บิล','')}`")
                    ec1, ec2 = st.columns(2)
                    with ec1:
                        try: ed = datetime.strptime(str(row.get("วันที่ในบิล","")), "%Y-%m-%d").date()
                        except: ed = date.today()
                        e_date  = st.date_input("📅 วันที่ในบิล", value=ed, key=f"ed_{i}")
                        e_billno= st.text_input("🧾 เลขที่บิล",
                                                 value=str(row.get("เลขที่บิล","")), key=f"ebn_{i}")
                        _ed = str(row.get("แผนก",""))
                        di  = DEPARTMENTS.index(_ed) if _ed in DEPARTMENTS else 0
                        e_dept  = st.selectbox("🏨 แผนก", DEPARTMENTS, index=di, key=f"edp_{i}")
                    with ec2:
                        _ei = str(row.get("รายการผ้า",""))
                        ii  = LINEN_ITEMS.index(_ei) if _ei in LINEN_ITEMS else len(LINEN_ITEMS)-1
                        e_item  = st.selectbox("👕 รายการผ้า", LINEN_ITEMS, index=ii, key=f"eit_{i}")
                        e_qty   = st.number_input("📦 จำนวน", min_value=1,
                                                   value=max(1,int(row.get("จำนวน",1))),
                                                   key=f"eq_{i}")
                        e_price = st.number_input("💰 ราคา/ชิ้น", min_value=0.0,
                                                   value=float(row.get("ราคาต่อหน่วย (บาท)",0)),
                                                   step=0.5, format="%.2f", key=f"ep_{i}")
                    e_ref  = st.text_input("📎 ชื่อไฟล์", value=str(row.get("ชื่อไฟล์อ้างอิง","")), key=f"er_{i}")
                    e_note = st.text_input("📝 หมายเหตุ",  value=str(row.get("หมายเหตุ","")),          key=f"en_{i}")
                    e_total = e_qty * e_price
                    st.metric("💵 ยอดรวมใหม่", f"฿{e_total:,.2f}")

                    sb, cb2 = st.columns(2)
                    save_edit  = sb.form_submit_button("💾 บันทึกการแก้ไข", type="primary",
                                                        use_container_width=True)
                    cancel_edit= cb2.form_submit_button("❌ ยกเลิก", use_container_width=True)

                if save_edit:
                    # Duplicate check (ยกเว้นแถวตัวเอง)
                    if e_billno.strip() != str(row.get("เลขที่บิล","")).strip():
                        if check_duplicate(ws_l, e_billno.strip()):
                            st.warning(f"⚠️ เลขที่บิล `{e_billno}` ซ้ำกับรายการอื่น!")
                            st.stop()

                    new_vals = [
                        str(row.get("วันที่บันทึก","")),
                        e_date.strftime("%Y-%m-%d"),
                        e_billno.strip(), e_dept, e_item,
                        int(e_qty), float(e_price), float(e_total),
                        e_ref.strip(), e_note.strip(),
                    ]
                    with st.spinner("กำลังบันทึก..."):
                        ok = update_row_laundry(ws_l, sheet_row, new_vals)
                    if ok:
                        write_audit(ws_a, "UPDATE", e_billno.strip(), e_dept,
                                    f"{e_item} x{e_qty} = ฿{e_total:,.2f}")
                        st.success("✅ แก้ไขสำเร็จ!")
                        st.session_state.edit_row_idx = None
                        st.rerun()

                if cancel_edit:
                    st.session_state.edit_row_idx = None
                    st.rerun()

            st.divider()

    # ── Bulk export ──
    csv = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button("⬇️ Export ทั้งหมด CSV", data=csv,
                       file_name=f"laundry_all_{date.today()}.csv", mime="text/csv")


# ══════════════════════════════════════════════
#  PAGE 3: AUDIT LOG
# ══════════════════════════════════════════════
def page_audit():
    st.title("📜 Audit Log")
    ws_a = st.session_state.ws_audit
    if not ws_a:
        st.warning("⚠️ ยังไม่ได้เชื่อมต่อ"); return

    with st.spinner("โหลด Audit Log..."):
        adf = load_audit_df(ws_a)

    if adf.empty:
        st.info("📭 ยังไม่มี Log"); return

    # Summary chips
    total  = len(adf)
    inserts= len(adf[adf["Action"]=="INSERT"]) if "Action" in adf.columns else 0
    updates= len(adf[adf["Action"]=="UPDATE"]) if "Action" in adf.columns else 0
    deletes= len(adf[adf["Action"]=="DELETE"]) if "Action" in adf.columns else 0
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📋 Log ทั้งหมด", f"{total:,}")
    c2.metric("➕ INSERT",  f"{inserts:,}")
    c3.metric("✏️ UPDATE",  f"{updates:,}")
    c4.metric("🗑️ DELETE", f"{deletes:,}")
    st.markdown("---")

    # Filter
    act_filter = st.multiselect("กรองประเภท Action", ["INSERT","UPDATE","DELETE"],
                                 default=["INSERT","UPDATE","DELETE"])
    fdf = adf[adf["Action"].isin(act_filter)] if "Action" in adf.columns else adf

    # Color-code
    def _color_action(val):
        c = {"INSERT":"#d4edda","UPDATE":"#fff3cd","DELETE":"#f8d7da"}.get(val,"")
        return f"background-color:{c}" if c else ""

    styled = fdf.style.applymap(_color_action, subset=["Action"]) if "Action" in fdf.columns else fdf
    st.dataframe(styled, use_container_width=True, hide_index=True)

    csv = fdf.to_csv(index=False, encoding="utf-8-sig")
    st.download_button("⬇️ Export Audit CSV", data=csv,
                       file_name=f"audit_{date.today()}.csv", mime="text/csv")


# ══════════════════════════════════════════════
#  PAGE 4: MONTHLY DASHBOARD
# ══════════════════════════════════════════════
def page_dashboard():
    ws_l = st.session_state.ws_laundry
    st.title("📊 Summary Dashboard")
    if not ws_l:
        st.warning("⚠️ ยังไม่ได้เชื่อมต่อ"); return

    with st.spinner("โหลดข้อมูล..."):
        df = load_df(ws_l)

    if df.empty:
        st.info("📭 ยังไม่มีข้อมูล"); return

    # Prep
    for c in ["จำนวน","ราคาต่อหน่วย (บาท)","ยอดรวม (บาท)"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["_date"] = pd.to_datetime(df["วันที่ในบิล"], errors="coerce")
    df.dropna(subset=["_date"], inplace=True)
    df["_year"]  = df["_date"].dt.year
    df["_month"] = df["_date"].dt.month
    df["_ym"]    = df["_date"].dt.to_period("M").astype(str)

    # ── Year / Month selectors ──
    years  = sorted(df["_year"].unique(), reverse=True)
    months = list(range(1, 13))

    col_y, col_m, _ = st.columns([1,1,2])
    sel_year  = col_y.selectbox("📅 ปี", years, index=0)
    sel_month = col_m.selectbox("📅 เดือน",
                                 [f"{m} — {THAI_MONTHS[m]}" for m in months],
                                 index=datetime.today().month - 1)
    sel_m_num = int(sel_month.split(" — ")[0])

    mdf = df[(df["_year"]==sel_year) & (df["_month"]==sel_m_num)]

    st.markdown(f"### 📅 {THAI_MONTHS[sel_m_num]} {sel_year}")

    if mdf.empty:
        st.info(f"📭 ไม่มีข้อมูลในเดือน {THAI_MONTHS[sel_m_num]} {sel_year}")
    else:
        # ── KPI ──
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("💵 ยอดรวมเดือนนี้",  f"฿{mdf['ยอดรวม (บาท)'].sum():,.2f}")
        k2.metric("📄 จำนวนรายการ",      f"{len(mdf):,}")
        k3.metric("🧾 จำนวนบิล",         f"{mdf['เลขที่บิล'].nunique():,}")
        k4.metric("🏨 แผนกที่ใช้บริการ",  f"{mdf['แผนก'].nunique():,}")
        st.markdown("---")

        # ── Bar: ยอดตามแผนก ──
        dept_m = (mdf.groupby("แผนก")["ยอดรวม (บาท)"].sum()
                     .reset_index().sort_values("ยอดรวม (บาท)", ascending=False))
        st.markdown("#### 🏨 ยอดแยกตามแผนก")
        st.bar_chart(dept_m.set_index("แผนก"), use_container_width=True)

        # ── Table: แผนก ──
        dept_m["% ของเดือน"] = (dept_m["ยอดรวม (บาท)"] /
                                   dept_m["ยอดรวม (บาท)"].sum() * 100).round(1)
        dept_m["ยอดรวม (บาท)"] = dept_m["ยอดรวม (บาท)"].apply(lambda x: f"฿{x:,.2f}")
        dept_m["% ของเดือน"]    = dept_m["% ของเดือน"].apply(lambda x: f"{x}%")
        st.dataframe(dept_m, use_container_width=True, hide_index=True)

    st.markdown("---")
    # ── 12-month trend ──
    st.markdown("#### 📈 ยอดรวมย้อนหลัง 12 เดือน")
    trend = (df.groupby("_ym")["ยอดรวม (บาท)"].sum()
               .reset_index().sort_values("_ym").tail(12))
    if not trend.empty:
        st.bar_chart(trend.set_index("_ym"), use_container_width=True)

    # ── YTD table ──
    st.markdown(f"#### 📋 สรุปยอดแยกแผนก ทั้งปี {sel_year}")
    ytd = df[df["_year"]==sel_year]
    if not ytd.empty:
        ytd_pivot = (ytd.groupby(["_month","แผนก"])["ยอดรวม (บาท)"]
                       .sum().unstack(fill_value=0))
        ytd_pivot.index = [f"{THAI_MONTHS.get(m,m)}" for m in ytd_pivot.index]
        ytd_pivot["รวม"] = ytd_pivot.sum(axis=1)
        st.dataframe(ytd_pivot.style.format("฿{:,.0f}"), use_container_width=True)

        csv = ytd_pivot.reset_index().to_csv(index=False, encoding="utf-8-sig")
        st.download_button(f"⬇️ Export สรุปปี {sel_year}", data=csv,
                           file_name=f"summary_{sel_year}.csv", mime="text/csv")


# ══════════════════════════════════════════════
#  PAGE 5: SETTINGS
# ══════════════════════════════════════════════
def page_settings():
    st.title("⚙️ การตั้งค่าระบบ")
    st.markdown("---")

    st.markdown("### 🔗 Google Sheets")
    if st.session_state.ws_laundry:
        st.success(f"✅ เชื่อมต่อสำเร็จ → **{SPREADSHEET_NAME}**")
        st.write(f"- Sheet หลัก: `{WS_LAUNDRY}` ({len(LAUNDRY_COLS)} คอลัมน์)")
        st.write(f"- Sheet Log:  `{WS_AUDITLOG}` ({len(AUDITLOG_COLS)} คอลัมน์)")
    else:
        st.error("❌ ยังไม่ได้เชื่อมต่อ")
        if st.button("🔄 เชื่อมต่อใหม่", type="primary"):
            gc, ws_l, ws_a = connect_google_sheets()
            st.session_state.gc = gc
            st.session_state.ws_laundry = ws_l
            st.session_state.ws_audit   = ws_a
            st.rerun()

    st.markdown("### 🤖 Claude Vision OCR")
    if get_anthropic_client():
        st.success("✅ Anthropic API Key พร้อมใช้งาน")
    else:
        st.error("❌ ไม่พบ API Key")
        st.code('anthropic_api_key = "sk-ant-api03-..."', language="toml")

    st.markdown("---")
    st.markdown("### 📖 การตั้งค่า `.streamlit/secrets.toml`")
    st.code("""
anthropic_api_key = "sk-ant-api03-..."

[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\\n...\\n-----END RSA PRIVATE KEY-----\\n"
client_email = "sa@project.iam.gserviceaccount.com"
client_id = "..."
""", language="toml")

    st.markdown("---")
    st.markdown("### 🗄️ Database Schema")
    st.markdown("**Sheet: Laundry** (Main Data)")
    st.dataframe(pd.DataFrame({
        "คอลัมน์": LAUNDRY_COLS,
        "หมายเหตุ": [
            "Timestamp บันทึก","วันที่ในบิล","เลขที่บิล (Unique)",
            "แผนก","รายการผ้า","จำนวนชิ้น","ราคาต่อชิ้น",
            "ยอดรวม (auto)","ชื่อไฟล์ภาพ","หมายเหตุ","Audit Timestamp ✅"
        ]
    }), hide_index=True, use_container_width=True)

    st.markdown("**Sheet: AuditLog** (ประวัติการเปลี่ยนแปลง)")
    st.dataframe(pd.DataFrame({
        "คอลัมน์": AUDITLOG_COLS,
        "หมายเหตุ": ["เวลา","ประเภท (INSERT/UPDATE/DELETE)",
                    "เลขที่บิล","แผนก","รายละเอียด","ผู้ดำเนินการ"]
    }), hide_index=True, use_container_width=True)

    st.markdown("---")
    st.json({"App Version":"3.0","OCR Model":"claude-opus-4-5",
             "Duplicate Check":"เลขที่บิล","Audit Col":"แก้ไขล่าสุด (LAUNDRY_COLS[-1])"})


# ══════════════════════════════════════════════
#  MAIN SHELL
# ══════════════════════════════════════════════
def main_app():
    with st.sidebar:
        st.markdown("## 🧺 Laundry Admin")
        st.caption("v3.0 — Full Edition")
        st.markdown("---")
        menu = st.radio("เมนู", [
            "📋 บันทึกรายการใหม่",
            "🗂️ Data Management",
            "📊 Summary Dashboard",
            "📜 Audit Log",
            "⚙️ การตั้งค่า",
        ], label_visibility="collapsed")
        st.markdown("---")
        if st.button("🚪 Logout", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    {
        "📋 บันทึกรายการใหม่": page_new_entry,
        "🗂️ Data Management":  page_manage,
        "📊 Summary Dashboard":page_dashboard,
        "📜 Audit Log":         page_audit,
        "⚙️ การตั้งค่า":        page_settings,
    }[menu]()


# ══════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════
if not st.session_state.logged_in:
    login_page()
else:
    main_app()
