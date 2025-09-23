# -*- coding: utf-8 -*-
# =============================================================
# Streamlit App: PDF Template Overlay + CSV -> Batch PDF Export (PDF-only)
# Layout v2:
#   ✅ Sidebar = สถานะ (Body/Cover/Preset/CSV) + คู่มือใช้งานแบบย่อ
#   ✅ Tab 1 = เทมเพลต + ส่งออก PDF ทั้งชุด
#   ✅ Tab 2 = 🔎 พรีวิว
#   ✅ Tab 3 = 📚 ข้อมูล (preview) + Preset (.json)
#   ✅ โหลดอัตโนมัติจาก GitHub: Body/Cover/Preset/CSV
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
DEFAULT_PRESET_URL = "https://raw.githubusercontent.com/firstnattapon/Canva/refs/heads/main/layout_preset.json"
DEFAULT_CSV_URL = "https://raw.githubusercontent.com/firstnattapon/Canva/refs/heads/main/Data.csv"

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
    try:
        resp = requests.get(url, timeout=10)
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
    # Normalize header whitespace
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
        try:
            p.insert_text((x, y), str(text), fontname=font if font in STD_FONTS else "helv",
                          fontsize=size, color=(0, 0, 0))
        except Exception:
            p.insert_text((x, y), str(text), fontname="helv", fontsize=size, color=(0, 0, 0))

    mat = fitz.Matrix(scale, scale)
    pix = p.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    td.close(); newdoc.close()
    return img


# ------------------ Streamlit UI ------------------

st.set_page_config(page_title="PDF Layout Editor — Batch PDF (Tabs Layout)", layout="wide")
st.title("🖨️ PDF Layout Editor — CSV เดียว → Batch PDF (PDF-only, ปกครั้งเดียว)")
st.caption("Sidebar = สถานะ + คู่มือย่อ • Tab1 = เทมเพลต + ส่งออก • Tab2 = พรีวิว • Tab3 = ข้อมูล + Preset • ถ้าไม่อัปโหลดจะโหลดค่าเริ่มต้นจาก GitHub อัตโนมัติ")

if fitz is None:
    st.error("ต้องติดตั้ง PyMuPDF ก่อนใช้งาน: `pip install pymupdf`")
    st.stop()

# ------------------ Tabs skeleton (widgets live here) ------------------
tab1, tab2, tab3 = st.tabs([
    "1) เทมเพลต + ส่งออก PDF ทั้งชุด",
    "2) 🔎 พรีวิว",
    "3) 📚 ข้อมูล (preview) + Preset (.json)",
])

with tab1:
    st.subheader("📄 เทมเพลต (PDF เท่านั้น)")
    c1, c2 = st.columns([1,1])
    with c1:
        tpl_pdf = st.file_uploader("Template PDF (Body)", type=["pdf"], key="tpl_pdf")
    with c2:
        cover_active = st.checkbox("Active ปก (หน้าแรกเสมอ; ไม่ต่อคน)", value=st.session_state.get("cover_active", False), key="cover_active")
        tpl_cover_pdf = st.file_uploader("Cover Template PDF", type=["pdf"], key="tpl_cover_pdf")

    st.markdown("---")
    st.subheader("📥 ข้อมูล (CSV เดียว)")
    csv_main = st.file_uploader("CSV หลัก (ตามสคีมาใหม่)", type=["csv", "xlsx", "xls"], key="csv_main")

with tab2:
    st.subheader("🔎 พรีวิวหน้า PDF")
    st.caption("เลือกแถว/หน้าเพื่อพรีวิว หลังระบบโหลดเทมเพลตและข้อมูลแล้ว")

with tab3:
    st.subheader("📚 ข้อมูล (preview) + Preset (.json)")
    st.caption("จัดการคอลัมน์และเลย์เอาต์, นำเข้า/ส่งออก Preset รวม (Body+Cover)")

# ------------------ Load defaults (after widgets exist) ------------------
# Decide sources
body_source = "uploaded" if st.session_state.get("tpl_pdf") is not None else "github"
cover_source = "uploaded" if st.session_state.get("tpl_cover_pdf") is not None else "github"
csv_source = "uploaded" if st.session_state.get("csv_main") is not None else "github"
preset_source = st.session_state.get("preset_source", None)  # will be set during preset load

# Fetch bytes
default_body_bytes = None
default_cover_bytes = None
default_csv_bytes = None

if st.session_state.get("tpl_pdf") is None:
    default_body_bytes = fetch_default_pdf(DEFAULT_BODY_URL)
    if default_body_bytes is None:
        body_source = "missing"
if st.session_state.get("cover_active", False) and st.session_state.get("tpl_cover_pdf") is None:
    default_cover_bytes = fetch_default_pdf(DEFAULT_COVER_URL)
    if default_cover_bytes is None:
        cover_source = "missing"
if st.session_state.get("csv_main") is None:
    default_csv_bytes = fetch_default_csv(DEFAULT_CSV_URL)
    if default_csv_bytes is None:
        csv_source = "missing"

# ------------------ Load DataFrame ------------------
# CSV -> df
if st.session_state.get("csv_main") is not None:
    df = canonicalize_columns(try_read_table(st.session_state["csv_main"]))
else:
    if default_csv_bytes is not None:
        df = canonicalize_columns(read_csv_bytes(default_csv_bytes))
    else:
        df = pd.DataFrame()

if df.empty:
    # Still render sidebar + guidance
    pass
else:
    for c in ["no", "student_id", "name", "sem1", "sem2", "total", "rating", "grade", "year"]:
        if c not in df.columns:
            df[c] = ""
    pref = ["no", "student_id", "name", "sem1", "sem2", "total", "rating", "grade", "year"]
    ordered = [c for c in pref if c in df.columns] + [c for c in df.columns if c not in pref]
    active_df = df[ordered]

# ------------------ Preset import/export (lives in Tab3), plus fields_df init ------------------
if "fields_df" not in st.session_state:
    st.session_state["fields_df"] = build_field_df(df.columns.tolist() if not df.empty else [], DEFAULT_FIELDS)
if "cover_fields_df" not in st.session_state:
    st.session_state["cover_fields_df"] = build_field_df(df.columns.tolist() if not df.empty else [], DEFAULT_COVER_FIELDS)
if "preset_loaded" not in st.session_state:
    st.session_state["preset_loaded"] = False

with tab3:
    # Data preview
    if df is None or df.empty:
        st.warning("ยังไม่มีข้อมูล CSV (อัปโหลดหรือให้ระบบโหลดจาก GitHub)")
    else:
        st.dataframe(active_df.head(12), use_container_width=True)

    st.markdown("---")
    st.subheader("🧩 Preset (.json) — รวม Body + Cover (cover ใช้แถว 0 เสมอ)")

    col_i, col_e = st.columns(2)
    with col_i:
        preset_json = st.file_uploader("นำเข้า Preset (.json)", type=["json"], key="unified_preset_upload")
        preset_url = st.text_input("หรือระบุ URL (raw GitHub) สำหรับ Preset", value=DEFAULT_PRESET_URL, key="preset_url")
        load_from_url = st.button("⬇️ โหลด Preset จาก URL", key="btn_load_preset")

        def _apply_unified_preset_bytes(preset_bytes: bytes, source_tag: str):
            try:
                raw = json.loads(preset_bytes.decode("utf-8"))
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
                    st.session_state["preset_source"] = source_tag
                    st.info("โหลดเฉพาะ Body (legacy) จาก Preset แล้ว")
                else:
                    body = raw.get("body", {})
                    cover = raw.get("cover", {})
                    if "fields" in body:
                        st.session_state["fields_df"] = pd.DataFrame(body["fields"])
                    if "fields" in cover:
                        st.session_state["cover_fields_df"] = pd.DataFrame(cover["fields"])
                    st.session_state["preset_loaded"] = True
                    st.session_state["preset_source"] = source_tag
                    st.success("นำเข้า Preset (Body + Cover) สำเร็จ")
            except Exception as e:
                st.error(f"อ่านไฟล์/URL Preset ไม่ได้: {e}")

        if preset_json is not None:
            _apply_unified_preset_bytes(preset_json.read(), source_tag="uploaded")

        if load_from_url:
            b = fetch_default_json(st.session_state.get("preset_url", DEFAULT_PRESET_URL))
            if b:
                _apply_unified_preset_bytes(b, source_tag="url")

        # Auto-load once if nothing loaded yet
        if not st.session_state["preset_loaded"] and preset_json is None and not load_from_url:
            auto_b = fetch_default_json(DEFAULT_PRESET_URL)
            if auto_b:
                _apply_unified_preset_bytes(auto_b, source_tag="github")
                st.info(f"Preset: โหลดจาก GitHub อัตโนมัติ\n{DEFAULT_PRESET_URL}")

    with col_e:
        try:
            payload = {
                "version": 10,
                "body": {"fields": st.session_state["fields_df"].to_dict(orient="records")},
                "cover": {
                    "fields": st.session_state["cover_fields_df"].to_dict(orient="records"),
                    "data_row_index": 0,
                },
            }
            buf = io.StringIO(); json.dump(payload, buf, ensure_ascii=False, indent=2)
            st.download_button("⬇️ Export Preset (.json)", data=buf.getvalue().encode("utf-8"),
                               file_name="layout_preset_body_cover.json", mime="application/json")
        except Exception as e:
            st.error(f"Export JSON ผิดพลาด: {e}")

    # Editors
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

# ------------------ Preview (Tab 2) ------------------
with tab2:
    if df is None or df.empty:
        st.info("อัปโหลด CSV หรือให้ระบบโหลดจาก GitHub ก่อน แล้วค่อยพรีวิว")
    else:
        idx_options = list(range(len(active_df)))
        rec_idx = st.number_input("แถวที่ต้องการพรีวิว (Body)", min_value=0, max_value=len(idx_options)-1, value=0, step=1)
        record_body = active_df.iloc[int(rec_idx)]
        record_cover = active_df.iloc[0]
        page_type = st.radio("หน้าไหน", ["Body", "Cover"], index=0, horizontal=True)

        try:
            if page_type == "Body":
                body_src = st.session_state["tpl_pdf"].getvalue() if st.session_state.get("tpl_pdf") is not None else fetch_default_pdf(DEFAULT_BODY_URL)
                if body_src is not None:
                    st.image(
                        render_preview_with_pymupdf(body_src, st.session_state["fields_df"], record_body, 2.0),
                        caption=f"Body — {get_record_display(record_body)}",
                        use_container_width=True,
                    )
                    st.caption("Body: หน่วย X/Y = จุด (pt) — มุมซ้ายบนคือ (0,0)")
                else:
                    st.info("อัปโหลด Template PDF ของ Body หรือให้ระบบโหลดค่าเริ่มต้นจาก GitHub (ตรวจสอบเครือข่าย)")
            else:
                if st.session_state.get("cover_active", False):
                    cover_src = st.session_state["tpl_cover_pdf"].getvalue() if st.session_state.get("tpl_cover_pdf") is not None else fetch_default_pdf(DEFAULT_COVER_URL)
                    if cover_src is not None:
                        st.image(
                            render_preview_with_pymupdf(cover_src, st.session_state["cover_fields_df"], record_cover, 2.0),
                            caption=f"Cover — ใช้ข้อมูลแถวที่ 0 (แถวแรก) — {get_record_display(record_cover)}",
                            use_container_width=True,
                        )
                        st.caption("Cover: หน่วย X/Y = จุด (pt) — มุมซ้ายบนคือ (0,0) • ใช้ข้อมูลจากแถวที่ 0 เสมอ")
                    else:
                        st.info("เปิด Active ปก แล้ว—อัปโหลด Cover Template PDF หรือให้ระบบโหลดจาก GitHub")
                else:
                    st.info("ยังไม่ได้เปิด Active ปก (หน้าแรกครั้งเดียว)")
        except Exception as e:
            st.error(f"พรีวิวผิดพลาด: {e}")

# ------------------ Export (Tab 1) ------------------
with tab1:
    st.markdown("---")
    st.subheader("📦 ส่งออก PDF ทั้งชุด (ปกหน้าแรก 1 ครั้ง + Body 1 หน้า/คน)")

    if st.button("🚀 Export PDF", key="btn_export"):
        try:
            # Prepare bytes
            body_src = st.session_state["tpl_pdf"].getvalue() if st.session_state.get("tpl_pdf") is not None else default_body_bytes
            if body_src is None:
                st.error("ไม่มี Template PDF ของ Body (อัปโหลดหรือให้ระบบโหลดค่าเริ่มต้นจาก GitHub)")
            else:
                out = fitz.open()

                # Insert cover once (row 0)
                if st.session_state.get("cover_active", False):
                    cover_src = st.session_state["tpl_cover_pdf"].getvalue() if st.session_state.get("tpl_cover_pdf") is not None else default_cover_bytes
                    if cover_src is not None:
                        t_cover = fitz.open(stream=cover_src, filetype="pdf")
                        out.insert_pdf(t_cover, from_page=0, to_page=0)
                        page0 = out[-1]
                        record_cover = active_df.iloc[0] if not df.empty else pd.Series()
                        for _, row in st.session_state["cover_fields_df"].iterrows():
                            if not row["active"]:
                                continue
                            key = row["field_key"]
                            if key not in record_cover or pd.isna(record_cover[key]):
                                continue
                            text = apply_transform(record_cover[key], row["transform"])
                            x, y = float(row["x"]), float(row["y"]) 
                            font = row.get("font", "helv"); size = float(row.get("size", 12))
                            try:
                                page0.insert_text((x, y), str(text), fontname=font if font in STD_FONTS else "helv",
                                                  fontsize=size, color=(0,0,0))
                            except Exception:
                                page0.insert_text((x, y), str(text), fontname="helv", fontsize=size, color=(0,0,0))
                        t_cover.close()
                    else:
                        st.warning("ไม่พบ Cover Template — จะข้ามหน้า Cover")

                # Insert body pages
                if df is None or df.empty:
                    st.error("ไม่มีข้อมูล CSV สำหรับการส่งออก")
                else:
                    for _, rec in active_df.iterrows():
                        t_body = fitz.open(stream=body_src, filetype="pdf")
                        out.insert_pdf(t_body, from_page=0, to_page=0)
                        page = out[-1]
                        for _, row in st.session_state["fields_df"].iterrows():
                            if not row["active"]:
                                continue
                            key = row["field_key"]
                            if key not in rec or pd.isna(rec[key]):
                                continue
                            text = apply_transform(rec[key], row["transform"])
                            x, y = float(row["x"]), float(row["y"]) 
                            font = row.get("font", "helv"); size = float(row.get("size", 12))
                            try:
                                page.insert_text((x, y), str(text), fontname=font if font in STD_FONTS else "helv",
                                                 fontsize=size, color=(0,0,0))
                            except Exception:
                                page.insert_text((x, y), str(text), fontname="helv", fontsize=size, color=(0,0,0))
                        t_body.close()

                    pdf_bytes = out.tobytes(); out.close()
                    total_pages = len(active_df) + (1 if (st.session_state.get("cover_active", False) and (st.session_state.get("tpl_cover_pdf") is not None or default_cover_bytes is not None)) else 0)
                    st.success(f"เสร็จแล้ว: {total_pages} หน้า (ปก 1 + เนื้อหา {len(active_df)})")
                    st.download_button("⬇️ ดาวน์โหลด PDF", data=pdf_bytes,
                                       file_name="exported_batch_with_global_cover.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"ส่งออกไม่สำเร็จ: {e}")

# ------------------ Sidebar: Status + Quick Guide ------------------
st.sidebar.header("📊 สถานะ (แหล่งที่ใช้)")
# Body
if body_source == "uploaded":
    st.sidebar.success("Body: ใช้ไฟล์ที่อัปโหลด")
elif body_source == "github":
    st.sidebar.info(f"Body: โหลดจาก GitHub\n{to_raw_github(DEFAULT_BODY_URL)}")
else:
    st.sidebar.error("Body: ไม่พบไฟล์")
# Cover
if not st.session_state.get("cover_active", False):
    st.sidebar.warning("Cover: ปิดการใช้งาน")
else:
    if cover_source == "uploaded":
        st.sidebar.success("Cover: ใช้ไฟล์ที่อัปโหลด")
    elif cover_source == "github":
        st.sidebar.info(f"Cover: โหลดจาก GitHub\n{to_raw_github(DEFAULT_COVER_URL)}")
    else:
        st.sidebar.error("Cover: ไม่พบไฟล์")
# Preset
if st.session_state.get("preset_loaded", False):
    tag = st.session_state.get("preset_source", "")
    label = "Preset: จาก " + ("อัปโหลด" if tag=="uploaded" else ("GitHub" if tag=="github" else ("URL" if tag=="url" else "ไม่ทราบแหล่ง")))
    st.sidebar.success(label)
else:
    st.sidebar.info("Preset: ยังไม่โหลด (จะลองโหลดจากค่าเริ่มต้น GitHub อัตโนมัติ)")
# CSV
if csv_source == "uploaded":
    st.sidebar.success("CSV: ใช้ไฟล์ที่อัปโหลด")
elif csv_source == "github":
    st.sidebar.info(f"CSV: โหลดจาก GitHub\n{to_raw_github(DEFAULT_CSV_URL)}")
else:
    st.sidebar.error("CSV: ไม่พบไฟล์")

st.sidebar.markdown("---")
st.sidebar.header("ℹ️ วิธีใช้ (ย่อ)")
st.sidebar.markdown(
    "1) ไปที่ **Tab 1** อัปโหลดเทมเพลต/CSV (หรือปล่อยให้โหลดอัตโนมัติ) แล้วกด **Export**\\n"
    "2) ไปที่ **Tab 2** ปรับแถว/หน้าเพื่อ **พรีวิว** ให้วางตำแหน่งพอดี\\n"
    "3) ไปที่ **Tab 3** นำเข้า/แก้ไข **Preset (Body/Cover)** แล้วลอง Export อีกรอบ"
)

st.markdown("---")
st.caption("CSV: No, Student ID, Name, Semester 1, Semester 2, Total, Rating, Grade, Year • Preset รวม (data_row_index=0) • ใช้ PDF เท่านั้น • โหลดอัตโนมัติ (Template + Preset + CSV)")
