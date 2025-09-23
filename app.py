# -*- coding: utf-8 -*-
# =============================================================
# Streamlit App: PDF Template Overlay + CSV -> Batch PDF Export (PDF-only)
# Update:
#   ✅ แสดงสถานะชัดเจนเมื่อโหลดเทมเพลตจาก GitHub อัตโนมัติ (Body/Cover)
#   ✅ นำเข้า Preset (.json) จาก GitHub (raw) อัตโนมัติถ้าไม่อัปโหลด พร้อม URL ให้แก้ได้
#   ✅ CSV เดียว (No, Student ID, Name, Semester 1, Semester 2, Total, Rating, Grade, Year)
#   ✅ Cover ใช้ข้อมูล "แถว 0 เสมอ"
#   ✅ Preset รวมไฟล์เดียว (Body + Cover) — บันทึก data_row_index=0 เสมอ
#   ✅ พรีวิวสด — ใช้ use_container_width
#   ✅ CSV Auto-load จาก GitHub (เหมือน Body/Cover) + แสดงสถานะ
#   ✅ Preset (.json) Auto-load จาก GitHub แบบเดียวกับ Body/Cover/CSV + สถานะ
#   ✅ รองรับการจัดตำแหน่ง Align (left/center/right) ด้วยการคำนวณความกว้างข้อความ
#
# Install deps:
#   pip install streamlit pandas pillow pymupdf requests
# =============================================================

import io
import json
from typing import List, Optional

import streamlit as st
import pandas as pd
import requests

# PDF dependency
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

from PIL import Image  # used to render pixmap previews

# ------------------ Default URLs ------------------
DEFAULT_COVER_URL = "https://github.com/firstnattapon/Canva/blob/main/Cover.pdf"
DEFAULT_BODY_URL  = "https://github.com/firstnattapon/Canva/blob/main/Template.pdf"
# ✅ ใช้ลิงก์หน้าเว็บก็ได้ เดี๋ยวแปลงเป็น raw ให้อัตโนมัติ
DEFAULT_PRESET_URL = "https://github.com/firstnattapon/Canva/blob/main/layout_preset.json"
DEFAULT_CSV_URL = "https://github.com/firstnattapon/Canva/blob/main/Data.csv"

# ------------------ Canonical columns & defaults ------------------
CANONICAL_COLS = {
    "No": "no",
    "Student ID": "student_id",
    "StudentID": "student_id",
    "ID": "student_id",
    "Name - Surname": "name",
    "Name": "name",
    "Semester 1": "sem1",
    "Semester1": "sem1",
    "Sem 1": "sem1",
    "Sem1": "sem1",
    "Semester 2": "sem2",
    "Semester2": "sem2",
    "Sem 2": "sem2",
    "Sem2": "sem2",
    "Total (50)": "total",
    "Total": "total",
    "Rating": "rating",
    "Grade": "grade",
    "Year": "year",
}

# Body defaults
DEFAULT_FIELDS = [
    ("name", "Name", True, 140.0, 160.0, "helv", 14, "title", "left"),
    ("student_id", "Student ID", True, 140.0, 185.0, "helv", 12, "none", "left"),
    ("sem1", "Semester 1", True, 420.0, 160.0, "helv", 14, "none", "left"),
    ("sem2", "Semester 2", True, 520.0, 160.0, "helv", 14, "none", "left"),
    ("total", "Total", True, 640.0, 160.0, "helv", 16, "none", "left"),
    ("rating", "Rating", False, 420.0, 190.0, "helv", 12, "upper", "left"),
    ("grade", "Grade", False, 520.0, 190.0, "helv", 12, "upper", "left"),
    ("year", "Year", False, 640.0, 190.0, "helv", 12, "none", "left"),
]

# Cover defaults
DEFAULT_COVER_FIELDS = [
    ("name", "Name", True, 220.0, 260.0, "helv", 24, "title", "left"),
    ("student_id", "Student ID", True, 220.0, 292.0, "helv", 16, "none", "left"),
    ("year", "Year", False, 220.0, 324.0, "helv", 14, "none", "left"),
    ("sem1", "Semester 1", False, 420.0, 260.0, "helv", 16, "none", "left"),
    ("sem2", "Semester 2", False, 520.0, 260.0, "helv", 16, "none", "left"),
    ("total", "Total", True, 420.0, 292.0, "helv", 20, "none", "left"),
    ("rating", "Rating", False, 520.0, 292.0, "helv", 16, "upper", "left"),
    ("grade", "Grade", False, 620.0, 292.0, "helv", 16, "upper", "left"),
]

STD_FONTS = ["helv", "times", "cour"]  # Built-in fonts for PyMuPDF

# ------------------ Helpers ------------------

def to_raw_github(url: str) -> str:
    """Transform GitHub web URL -> raw.githubusercontent URL when needed."""
    if "github.com/" in url and "/blob/" in url:
        return url.replace("github.com/", "raw.githubusercontent.com/").replace("/blob/", "/")
    return url

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_default_pdf(url: str) -> Optional[bytes]:
    try:
        raw_url = to_raw_github(url)
        resp = requests.get(raw_url, timeout=10)
        resp.raise_for_status()
        return resp.content
    except Exception as e:
        st.warning(f"โหลดค่าเริ่มต้นจาก {url} ไม่ได้: {e}")
        return None

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_default_json(url: str) -> Optional[bytes]:
    """Fetch JSON bytes from GitHub (supports normal or raw URLs)."""
    try:
        raw_url = to_raw_github(url)
        resp = requests.get(raw_url, timeout=10)
        resp.raise_for_status()
        return resp.content
    except Exception as e:
        st.warning(f"โหลด Preset จาก {url} ไม่ได้: {e}")
        return None

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_default_csv(url: str) -> Optional[bytes]:
    """Fetch CSV bytes from GitHub (supports normal or raw URLs)."""
    try:
        raw_url = to_raw_github(url)
        resp = requests.get(raw_url, timeout=10)
        resp.raise_for_status()
        return resp.content
    except Exception as e:
        st.warning(f"โหลด CSV เริ่มต้นจาก {url} ไม่ได้: {e}")
        return None

def read_csv_bytes(b: bytes) -> pd.DataFrame:
    """Read CSV bytes into DataFrame with BOM fallback + header trim."""
    try:
        df = pd.read_csv(io.BytesIO(b))
    except UnicodeDecodeError:
        df = pd.read_csv(io.BytesIO(b), encoding="utf-8-sig")
    df = df.rename(columns=lambda c: " ".join(str(c).split()))
    return df

def try_read_table(uploaded_file) -> pd.DataFrame:
    """Read CSV/Excel into DataFrame and normalize header whitespace."""
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".csv"):
            try:
                df = pd.read_csv(uploaded_file)
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            df = pd.read_excel(uploaded_file)
        else:
            st.warning(f"ไม่รองรับไฟล์: {uploaded_file.name}")
            return pd.DataFrame()
        df = df.rename(columns=lambda c: " ".join(str(c).split()))
    except Exception as e:
        st.error(f"อ่านไฟล์ {uploaded_file.name} ไม่ได้: {e}")
        return pd.DataFrame()
    return df

def canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    new_cols = {}
    for c in df.columns:
        key = c
        if c in CANONICAL_COLS:
            key = CANONICAL_COLS[c]
        else:
            c2 = str(c).strip().lower().replace(" ", "").replace("-", "").replace("_", "")
            if c2 in ["studentid", "id"]:
                key = "student_id"
            elif c2 in ["name", "namesurname"]:
                key = "name"
            elif c2 in ["semester1", "sem1"]:
                key = "sem1"
            elif c2 in ["semester2", "sem2"]:
                key = "sem2"
            elif "total" in c2:
                key = "total"
            elif "rating" in c2:
                key = "rating"
            elif "grade" in c2:
                key = "grade"
            elif "year" in c2:
                key = "year"
            elif c2 == "no":
                key = "no"
        new_cols[c] = key
    out = df.rename(columns=new_cols)
    return out

def build_field_df(existing_cols: List[str], defaults) -> pd.DataFrame:
    rows = []
    existing = set(existing_cols)
    known = set()
    for k, label, active, x, y, font, size, transform, align in defaults:
        rows.append({
            "field_key": k,
            "label": label,
            "active": active if k in existing or k in ["name", "student_id", "total"] else False,
            "x": x, "y": y, "font": font, "size": size,
            "transform": transform, "align": align
        })
        known.add(k)
    for c in existing:
        if c not in known and c not in ["no"]:
            rows.append({
                "field_key": c,
                "label": c.title(),
                "active": False, "x": 100.0, "y": 100.0,
                "font": "helv", "size": 12,
                "transform": "none", "align": "left"
            })
    df = pd.DataFrame(rows)
    return df

def apply_transform(text, mode: str) -> str:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    s = str(text)
    if mode == "upper":
        return s.upper()
    if mode == "lower":
        return s.lower()
    if mode == "title":
        return s.title()
    return s

def get_record_display(rec: pd.Series, key_cols=("student_id", "name")) -> str:
    parts = []
    for k in key_cols:
        if k in rec and pd.notnull(rec[k]):
            parts.append(str(rec[k]))
    return " • ".join(parts) if parts else "(no id / name)"

def _aligned_xy(page, text: str, x: float, y: float, font: str, size: float, align: str):
    """Return (x_adj, y) after applying align using text width."""
    try:
        width = page.get_text_length(text, fontname=font if font in STD_FONTS else "helv", fontsize=size)
    except Exception:
        width = page.get_text_length(text, fontname="helv", fontsize=size)
    if align == "center":
        return x - width / 2.0, y
    if align == "right":
        return x - width, y
    return x, y

def render_preview_with_pymupdf(template_bytes: bytes, fields_df: pd.DataFrame,
                                record: pd.Series, scale: float = 2.0):
    if fitz is None:
        raise RuntimeError("PyMuPDF (fitz) is not available")
    td = fitz.open(stream=template_bytes, filetype="pdf")
    if td.page_count != 1:
        st.warning("เทมเพลตต้องเป็น PDF หน้าเดียว จะใช้หน้าแรกแทน")
    newdoc = fitz.open()
    newdoc.insert_pdf(td, from_page=0, to_page=0)
    p = newdoc[0]

    for _, row in fields_df.iterrows():
        if not row["active"]:
            continue
        key = row["field_key"]
        if key not in record or pd.isna(record[key]):
            continue
        text = apply_transform(record[key], row["transform"])
        x, y = float(row["x"]), float(row["y"])
        font = row.get("font", "helv")
        size = float(row.get("size", 12))
        align = row.get("align", "left")
        ax, ay = _aligned_xy(p, str(text), x, y, font, size, align)
        try:
            p.insert_text((ax, ay), str(text), fontname=font if font in STD_FONTS else "helv",
                          fontsize=size, color=(0, 0, 0))
        except Exception:
            p.insert_text((ax, ay), str(text), fontname="helv", fontsize=size, color=(0, 0, 0))

    mat = fitz.Matrix(scale, scale)
    pix = p.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    td.close(); newdoc.close()
    return img

# ------------------ Streamlit UI ------------------

st.set_page_config(page_title="PDF Layout Editor — CSV (Unified) → Batch PDF [PDF-only]", layout="wide")
st.title("🖨️ PDF Layout Editor")

if fitz is None:
    st.error("ต้องติดตั้ง PyMuPDF ก่อนใช้งาน: `pip install pymupdf`")
    st.stop()

colL, colR = st.columns([1.2, 1.0], gap="large")

with st.sidebar:
    st.header("📄 เทมเพลต — Body (PDF เท่านั้น)")
    tpl_pdf = st.file_uploader("Template PDF (Body)", type=["pdf"]) 

    st.header("🧾 เทมเพลต — ปก (PDF เท่านั้น)")
    cover_active = st.checkbox("Active ปก (หน้าแรกเสมอ; ไม่ต่อคน)", value=False)
    tpl_cover_pdf = st.file_uploader("Cover Template PDF", type=["pdf"]) 

    st.header("📥 ข้อมูล (CSV เดียว)")
    csv_main = st.file_uploader("CSV หลัก (ตามสคีมาใหม่)", type=["csv", "xlsx", "xls"]) 

# Auto-fetch defaults if not uploaded
default_body_bytes = None
default_cover_bytes = None
default_csv_bytes = None
default_preset_bytes = None  # ✅ NEW

body_source = "uploaded" if tpl_pdf is not None else "github"
cover_source = "uploaded" if tpl_cover_pdf is not None else "github"
csv_source = "uploaded" if csv_main is not None else "github"
preset_source = "github"  # จะอัปเดตตามการกระทำด้านล่าง

if tpl_pdf is None:
    default_body_bytes = fetch_default_pdf(DEFAULT_BODY_URL)
    if default_body_bytes is None:
        body_source = "missing"
if cover_active and tpl_cover_pdf is None:
    default_cover_bytes = fetch_default_pdf(DEFAULT_COVER_URL)
    if default_cover_bytes is None:
        cover_source = "missing"
if csv_main is None:
    default_csv_bytes = fetch_default_csv(DEFAULT_CSV_URL)
    if default_csv_bytes is None:
        csv_source = "missing"

# Visible status for template sources
st.subheader("🔔 สถานะเทมเพลต (Body/Cover)")
c1, c2 = st.columns(2)
with c1:
    if body_source == "uploaded":
        st.success("Body Template: ใช้ไฟล์ที่อัปโหลด")
    elif body_source == "github":
        st.info(f"Body Template: โหลดจาก GitHub อัตโนมัติ\n{to_raw_github(DEFAULT_BODY_URL)}")
    else:
        st.error("Body Template: ไม่พบทั้งไฟล์อัปโหลดและค่าเริ่มต้นจาก GitHub")
with c2:
    if not cover_active:
        st.warning("Cover Template: ปิดการใช้งานปก (จะไม่มีหน้าแรก)")
    else:
        if cover_source == "uploaded":
            st.success("Cover Template: ใช้ไฟล์ที่อัปโหลด")
        elif cover_source == "github":
            st.info(f"Cover Template: โหลดจาก GitHub อัตโนมัติ\n{to_raw_github(DEFAULT_COVER_URL)}")
        else:
            st.error("Cover Template: ไม่พบทั้งไฟล์อัปโหลดและค่าเริ่มต้นจาก GitHub")

# CSV status
st.subheader("🔔 สถานะข้อมูล (CSV)")
if csv_source == "uploaded":
    st.success("CSV: ใช้ไฟล์ที่อัปโหลด")
elif csv_source == "github":
    st.info(f"CSV: โหลดจาก GitHub อัตโนมัติ\n{to_raw_github(DEFAULT_CSV_URL)}")
else:
    st.error("CSV: ไม่พบทั้งไฟล์อัปโหลดและค่าเริ่มต้นจาก GitHub")

with colL:
    # Load data — single CSV (uploaded or default GitHub)
    if csv_main is not None:
        df = canonicalize_columns(try_read_table(csv_main))
    else:
        if default_csv_bytes is not None:
            df = canonicalize_columns(read_csv_bytes(default_csv_bytes))
        else:
            st.warning("อัปโหลด CSV ตามสคีมาใหม่ก่อน หรือระบบโหลดจาก GitHub ไม่สำเร็จ")
            st.stop()

    if df.empty:
        st.warning("CSV ว่างเปล่า")
        st.stop()

    # Ensure important columns exist
    for c in ["no", "student_id", "name", "sem1", "sem2", "total", "rating", "grade", "year"]:
        if c not in df.columns:
            df[c] = ""

    # Order columns nicely
    pref = ["no", "student_id", "name", "sem1", "sem2", "total", "rating", "grade", "year"]
    ordered = [c for c in pref if c in df.columns] + [c for c in df.columns if c not in pref]
    active_df = df[ordered]

    st.subheader("📚 ข้อมูล (preview)")
    st.dataframe(active_df.head(12), use_container_width=True)

with colR:
    st.subheader("🧩 Preset (.json)")

    # Init session states
    if "fields_df" not in st.session_state:
        st.session_state["fields_df"] = build_field_df(active_df.columns.tolist(), DEFAULT_FIELDS)
    if "cover_fields_df" not in st.session_state:
        st.session_state["cover_fields_df"] = build_field_df(active_df.columns.tolist(), DEFAULT_COVER_FIELDS)
    if "preset_loaded" not in st.session_state:
        st.session_state["preset_loaded"] = False
    if "preset_url_used" not in st.session_state:
        st.session_state["preset_url_used"] = ""

    with st.expander("Import / Export", expanded=False):
        col_i, col_e = st.columns(2)
        with col_i:
            preset_json = st.file_uploader("นำเข้า Preset (.json)", type=["json"], key="unified_preset_upload")
            preset_url = st.text_input("หรือระบุ URL (GitHub/Raw) สำหรับ Preset", value=DEFAULT_PRESET_URL)
            load_from_url = st.button("⬇️ โหลด Preset จาก URL")

            def _apply_unified_preset_bytes(preset_bytes: bytes, source_label: str):
                try:
                    raw = json.loads(preset_bytes.decode("utf-8"))
                    # Back-compat: list/fields => Body only
                    if isinstance(raw, list) or "fields" in raw:
                        fields_list = raw.get("fields", raw if isinstance(raw, list) else [])
                        new_df = pd.DataFrame(fields_list)
                        req = ["field_key", "label", "active", "x", "y", "font", "size", "transform", "align"]
                        missing = [c for c in req if c not in new_df.columns]
                        if missing:
                            st.error(f"Preset JSON ขาดคีย์: {missing}")
                            return
                        st.session_state["fields_df"] = new_df[req]
                        st.session_state["preset_loaded"] = True
                        st.session_state["preset_url_used"] = source_label
                        st.info("โหลดเฉพาะ Body (legacy) จาก Preset แล้ว")
                    else:
                        body = raw.get("body", {})
                        cover = raw.get("cover", {})
                        if "fields" in body:
                            st.session_state["fields_df"] = pd.DataFrame(body["fields"])
                        if "fields" in cover:
                            st.session_state["cover_fields_df"] = pd.DataFrame(cover["fields"])
                        # data_row_index ถูกบังคับเป็น 0 เสมอ
                        st.session_state["preset_loaded"] = True
                        st.session_state["preset_url_used"] = source_label
                        st.success("นำเข้า Preset (Body + Cover) สำเร็จ")
                except Exception as e:
                    st.error(f"อ่านไฟล์/URL Preset ไม่ได้: {e}")

            if preset_json is not None:
                _apply_unified_preset_bytes(preset_json.read(), "uploaded")
                preset_source = "uploaded"

            if load_from_url:
                b = fetch_default_json(preset_url)
                if b:
                    _apply_unified_preset_bytes(b, to_raw_github(preset_url))
                    preset_source = "github"

            # Auto-load from DEFAULT_PRESET_URL once if nothing uploaded yet
            if not st.session_state["preset_loaded"] and preset_json is None and not load_from_url:
                auto_b = fetch_default_json(DEFAULT_PRESET_URL)
                if auto_b:
                    _apply_unified_preset_bytes(auto_b, to_raw_github(DEFAULT_PRESET_URL))
                    st.info(f"Preset: โหลดจาก GitHub อัตโนมัติ\n{to_raw_github(DEFAULT_PRESET_URL)}")
                    preset_source = "github"

        with col_e:
            try:
                payload = {
                    "version": 10,
                    "body": {"fields": st.session_state["fields_df"].to_dict(orient="records")},
                    "cover": {
                        "fields": st.session_state["cover_fields_df"].to_dict(orient="records"),
                        "data_row_index": 0,  # always zero by design
                    },
                }
                buf = io.StringIO(); json.dump(payload, buf, ensure_ascii=False, indent=2)
                st.download_button("⬇️ Export Preset (.json)", data=buf.getvalue().encode("utf-8"),
                                   file_name="layout_preset_body_cover.json", mime="application/json")
            except Exception as e:
                st.error(f"Export JSON ผิดพลาด: {e}")

    # ✅ แสดงสถานะ Preset เหมือน Body/Cover/CSV
    st.subheader("🔔 สถานะ Preset (.json)")
    if preset_source == "uploaded":
        st.success("Preset: ใช้ไฟล์ที่อัปโหลด")
    elif preset_source == "github":
        used = st.session_state.get("preset_url_used", to_raw_github(DEFAULT_PRESET_URL))
        st.info(f"Preset: โหลดจาก GitHub อัตโนมัติ\n{used}")
    else:
        st.error("Preset: ไม่พบทั้งไฟล์อัปโหลดและค่าเริ่มต้นจาก GitHub")

    # Tabs for editing
    tab_body, tab_cover = st.tabs(["⚙️ Body Layout", "⚙️ Cover Layout"])

    with tab_body:
        edited_body = st.data_editor(
            st.session_state["fields_df"],
            use_container_width=True, hide_index=True,
            column_config={
                "field_key": st.column_config.TextColumn("field_key", disabled=True),
                "label": st.column_config.TextColumn("Label"),
                "active": st.column_config.CheckboxColumn("Active"),
                "x": st.column_config.NumberColumn("X", step=1, format="%.1f"),
                "y": st.column_config.NumberColumn("Y", step=1, format="%.1f"),
                "font": st.column_config.SelectboxColumn("Font", options=STD_FONTS),
                "size": st.column_config.NumberColumn("Size (pt)", step=1, format="%.0f"),
                "transform": st.column_config.SelectboxColumn("Case", options=["none", "upper", "lower", "title"]),
                "align": st.column_config.SelectboxColumn("Align", options=["left", "center", "right"]),
            },
            key="fields_editor_body",
        )
        st.session_state["fields_df"] = edited_body

    with tab_cover:
        edited_cover = st.data_editor(
            st.session_state["cover_fields_df"],
            use_container_width=True, hide_index=True,
            column_config={
                "field_key": st.column_config.TextColumn("field_key", disabled=True),
                "label": st.column_config.TextColumn("Label"),
                "active": st.column_config.CheckboxColumn("Active"),
                "x": st.column_config.NumberColumn("X", step=1, format="%.1f"),
                "y": st.column_config.NumberColumn("Y", step=1, format="%.1f"),
                "font": st.column_config.SelectboxColumn("Font", options=STD_FONTS),
                "size": st.column_config.NumberColumn("Size (pt)", step=1, format="%.0f"),
                "transform": st.column_config.SelectboxColumn("Case", options=["none", "upper", "lower", "title"]),
                "align": st.column_config.SelectboxColumn("Align", options=["left", "center", "right"]),
            },
            key="fields_editor_cover",
        )
        st.session_state["cover_fields_df"] = edited_cover

# -------- Preview --------
st.subheader("🔎 พรีวิว")
idx_options = list(range(len(active_df)))
if len(idx_options) == 0:
    st.stop()

rec_idx = st.number_input("แถวที่ต้องการพรีวิว (Body)", min_value=0, max_value=len(idx_options)-1, value=0, step=1)
record_body = active_df.iloc[int(rec_idx)]

# Cover record = row 0 ALWAYS
cov_idx = 0
record_cover = active_df.iloc[cov_idx]

page_type = st.radio("หน้าไหน", ["Body", "Cover"], index=0, horizontal=True)

def _draw_fields_on_page(page, fields_df: pd.DataFrame, record: pd.Series):
    for _, row in fields_df.iterrows():
        if not row["active"]:
            continue
        key = row["field_key"]
        if key not in record or pd.isna(record[key]):
            continue
        text = apply_transform(record[key], row["transform"])
        x, y = float(row["x"]), float(row["y"])
        font = row.get("font", "helv")
        size = float(row.get("size", 12))
        align = row.get("align", "left")
        ax, ay = _aligned_xy(page, str(text), x, y, font, size, align)
        try:
            page.insert_text((ax, ay), str(text), fontname=font if font in STD_FONTS else "helv",
                             fontsize=size, color=(0,0,0))
        except Exception:
            page.insert_text((ax, ay), str(text), fontname="helv", fontsize=size, color=(0,0,0))

try:
    if page_type == "Body":
        body_src = tpl_pdf.getvalue() if tpl_pdf is not None else default_body_bytes
        if body_src is not None:
            st.image(
                render_preview_with_pymupdf(body_src, st.session_state["fields_df"], record_body, 2.0),
                caption=f"Body — {get_record_display(record_body)}",
                use_container_width=True,
            )
            if body_source == "github" and tpl_pdf is None:
                st.caption(f"กำลังใช้ Body จาก GitHub: {to_raw_github(DEFAULT_BODY_URL)}")
            st.caption("Body: หน่วย X/Y = จุด (pt) — มุมซ้ายบนคือ (0,0)")
        else:
            st.info("อัปโหลด Template PDF ของ Body หรือให้ระบบโหลดค่าเริ่มต้นจาก GitHub (ตรวจสอบเครือข่าย)")
    else:  # Cover
        if cover_active:
            cover_src = tpl_cover_pdf.getvalue() if tpl_cover_pdf is not None else default_cover_bytes
            if cover_src is not None:
                st.image(
                    render_preview_with_pymupdf(cover_src, st.session_state["cover_fields_df"], record_cover, 2.0),
                    caption=f"Cover — ใช้ข้อมูลแถวที่ 0 (แถวแรก) — {get_record_display(record_cover)}",
                    use_container_width=True,
                )
                if cover_source == "github" and tpl_cover_pdf is None:
                    st.caption(f"กำลังใช้ Cover จาก GitHub: {to_raw_github(DEFAULT_COVER_URL)}")
                st.caption("Cover: หน่วย X/Y = จุด (pt) — มุมซ้ายบนคือ (0,0) • ใช้ข้อมูลจากแถวที่ 0 เสมอ")
            else:
                st.info("เปิด Active ปก แล้ว—อัปโหลด Cover Template PDF หรือให้ระบบโหลดค่าเริ่มต้นจาก GitHub (ตรวจสอบเครือข่าย)")
        else:
            st.info("ยังไม่ได้เปิด Active ปก (หน้าแรกครั้งเดียว)")
except Exception as e:
    st.error(f"พรีวิวผิดพลาด: {e}")

st.divider()
st.subheader("📦 ส่งออก PDF ทั้งชุด (ปกหน้าแรก 1 ครั้ง + Body 1 หน้า/คน)")

if st.button("🚀 Export PDF"):
    try:
        body_src = tpl_pdf.getvalue() if tpl_pdf is not None else default_body_bytes
        if body_src is None:
            st.error("ไม่มี Template PDF ของ Body (อัปโหลดหรือให้ระบบโหลดค่าเริ่มต้นจาก GitHub)")
        else:
            out = fitz.open()

            # Insert global cover once using row 0
            if cover_active:
                cover_src = tpl_cover_pdf.getvalue() if tpl_cover_pdf is not None else default_cover_bytes
                if cover_src is not None:
                    t_cover = fitz.open(stream=cover_src, filetype="pdf")
                    out.insert_pdf(t_cover, from_page=0, to_page=0)
                    page0 = out[-1]
                    _draw_fields_on_page(page0, st.session_state["cover_fields_df"], record_cover)
                    t_cover.close()
                else:
                    st.warning("ไม่พบ Cover Template — จะข้ามหน้า Cover")

            # Insert body pages per student
            for _, rec in active_df.iterrows():
                t_body = fitz.open(stream=body_src, filetype="pdf")
                out.insert_pdf(t_body, from_page=0, to_page=0)
                page = out[-1]
                _draw_fields_on_page(page, st.session_state["fields_df"], rec)
                t_body.close()

            pdf_bytes = out.tobytes(); out.close()
            total_pages = len(active_df) + (1 if (cover_active and (tpl_cover_pdf is not None or default_cover_bytes is not None)) else 0)
            st.success(f"เสร็จแล้ว: {total_pages} หน้า (ปก 1 + เนื้อหา {len(active_df)})")
            st.download_button("⬇️ ดาวน์โหลด PDF", data=pdf_bytes,
                               file_name="exported_batch_with_global_cover.pdf", mime="application/pdf")
    except Exception as e:
        st.error(f"ส่งออกไม่สำเร็จ: {e}")

st.markdown("---")
st.caption(
    "CSV: No, Student ID, Name, Semester 1, Semester 2, Total, Rating, Grade, Year • "
    "Preset รวม (data_row_index=0) • ใช้ PDF เท่านั้น • "
    "โหลดอัตโนมัติจาก GitHub ได้ทั้ง Template + Preset + CSV (รองรับลิงก์หน้าเว็บ GitHub และ raw)"
)
