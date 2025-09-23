# -*- coding: utf-8 -*-
# =============================================================
# Streamlit App: PDF Template Overlay + CSV -> Batch PDF Export (PDF-only)
# Features:
#   ✅ CSV เดียว (v2): No, Student ID, Name, Semester 1, Semester 2, Total, Rating, Grade, Year
#   ✅ รองรับแปลงไฟล์ v1 -> v2 (v1 = No, Student ID, Name - Surname, Idea, Pronunciation, Preparedness, Confidence, Total (50))
#   ✅ Template เฉพาะ PDF เท่านั้น (ไม่มีอัปโหลดรูป)
#   ✅ Cover ใช้ข้อมูล "แถวที่ 0 เสมอ" (บังคับ)
#   ✅ Preset (.json) รวมไฟล์เดียว (Body + Cover) — data_row_index=0 เสมอ
#   ✅ พรีวิวสด — ใช้ use_container_width
#   ✅ โหลดค่าเริ่มต้นเทมเพลต (Body/Cover) และ Preset จาก GitHub อัตโนมัติ พร้อมแสดงสถานะชัดเจน
#   ✅ ตารางข้อมูลแก้ไขได้ (ใช้ในการพรีวิว/ส่งออก)
#
# Install deps:
#   pip install streamlit pandas pillow pymupdf requests
# =============================================================

import io
import json
from typing import List, Optional, Tuple

import streamlit as st
import pandas as pd
import requests

# PDF dependency
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

from PIL import Image  # pixmap rendering

# ------------------ Default URLs ------------------
DEFAULT_COVER_URL = "https://github.com/firstnattapon/Canva/blob/main/Cover.pdf"
DEFAULT_BODY_URL  = "https://github.com/firstnattapon/Canva/blob/main/Template.pdf"
DEFAULT_PRESET_URL = "https://raw.githubusercontent.com/firstnattapon/Canva/refs/heads/main/layout_preset.json"

# ------------------ Canonical columns & defaults ------------------
CANONICAL_COLS = {
    # v2
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
    "Total (50)": "total50",
    "Total": "total",
    "Rating": "rating",
    "Grade": "grade",
    "Year": "year",
    # v1 component fields
    "Idea": "idea",
    "Pronunciation": "pronunciation",
    "Preparedness": "preparedness",
    "Confidence": "confidence",
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
        # Normalize header whitespace e.g. ' Total  ' -> 'Total'
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
            elif "total(50)" in c2 or c2 == "total50":
                key = "total50"
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
            elif "idea" in c2:
                key = "idea"
            elif "pronun" in c2:
                key = "pronunciation"
            elif "prepared" in c2:
                key = "preparedness"
            elif "confid" in c2:
                key = "confidence"
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


# ------------------ v1 -> v2 Converter ------------------

def detect_v1_schema(df: pd.DataFrame) -> bool:
    cols = set(df.columns.str.lower())
    # v1 has these signatures
    return ("name - surname".lower() in cols) or ("idea" in cols) or ("total (50)".lower() in cols)


def convert_v1_to_v2(df_v1_raw: pd.DataFrame) -> Tuple[pd.DataFrame, dict]:
    """Convert v1 sheet to v2 dataframe.
       v1 cols: No, Student ID, Name - Surname, Idea, Pronunciation, Preparedness, Confidence, Total (50)
       v2 cols: No, Student ID, Name, Semester 1, Semester 2, Total, Rating, Grade, Year
    """
    info = {"used_sum_components": False, "notes": []}
    d = canonicalize_columns(df_v1_raw.copy())

    # Make sure required columns exist
    if "name" not in d.columns and "name - surname" in df_v1_raw.columns:
        d["name"] = df_v1_raw["Name - Surname"]

    # Compute total50 if missing
    if "total50" not in d.columns:
        comps = [c for c in ["idea", "pronunciation", "preparedness", "confidence"] if c in d.columns]
        if comps:
            d["total50"] = d[comps].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
            info["used_sum_components"] = True
            info["notes"].append("คำนวณ Total(50) จาก Idea+Pronunciation+Preparedness+Confidence")
        else:
            d["total50"] = pd.NA
            info["notes"].append("ไม่พบ Total(50) และไม่มีองค์ประกอบเพียงพอ")

    # Build v2 frame
    out_cols = ["no", "student_id", "name", "sem1", "sem2", "total", "rating", "grade", "year"]
    out = pd.DataFrame(columns=out_cols)

    out["no"] = pd.to_numeric(d.get("no", pd.Series(range(1, len(d)+1))), errors="coerce")
    out["student_id"] = d.get("student_id", "")
    out["name"] = d.get("name", "")

    out["sem1"] = pd.to_numeric(d.get("total50", pd.NA), errors="coerce")
    out["sem2"] = pd.NA  # unknown from v1
    # If sem2 is NaN, we can copy sem1 to total for convenience
    out["total"] = out["sem1"]
    out["rating"] = ""
    out["grade"] = ""
    out["year"] = ""

    # Clean types
    for c in ["no", "sem1", "sem2", "total"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    return out[out_cols], info


# ------------------ Streamlit UI ------------------

st.set_page_config(page_title="PDF Layout Editor — CSV (Unified) → Batch PDF [PDF-only]", layout="wide")
st.title("🖨️ PDF Layout Editor — CSV เดียว → Batch PDF (PDF-only, ปกครั้งเดียว)")
st.caption("Cover ใช้ข้อมูลชุดเดียวกับ Body แต่มี Layout แยก • ปกอยู่หน้าแรก 1 ครั้ง • Preset .json รวม Body/Cover • รองรับ **PDF เท่านั้น** • ปกใช้ข้อมูล \"แถว 0\" เสมอ • มีตัวช่วยแปลง v1 → v2 • โหลดค่าเริ่มต้นจาก GitHub อัตโนมัติ")

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

    st.header("📥 ข้อมูล (CSV v2)")
    csv_main = st.file_uploader("CSV หลัก (v2 schema)", type=["csv", "xlsx", "xls"])

    st.header("🔁 แปลงข้อมูล v1 → v2")
    csv_v1 = st.file_uploader("CSV แบบ v1 (legacy)", type=["csv", "xlsx", "xls"], key="csv_v1")
    do_convert = st.button("แปลงไฟล์ v1 → v2")

# Auto-fetch defaults if not uploaded
default_body_bytes = None
default_cover_bytes = None
body_source = "uploaded" if tpl_pdf is not None else "github"  # tentative
cover_source = "uploaded" if tpl_cover_pdf is not None else "github"  # tentative
if tpl_pdf is None:
    default_body_bytes = fetch_default_pdf(DEFAULT_BODY_URL)
    if default_body_bytes is None:
        body_source = "missing"
if cover_active and tpl_cover_pdf is None:
    default_cover_bytes = fetch_default_pdf(DEFAULT_COVER_URL)
    if default_cover_bytes is None:
        cover_source = "missing"

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

with colL:
    # Convert v1 -> v2 if requested
    converted_df = None
    convert_info = None
    if do_convert and csv_v1 is not None:
        df_v1 = try_read_table(csv_v1)
        if df_v1.empty:
            st.error("อ่านไฟล์ v1 ไม่ได้")
        else:
            converted_df, convert_info = convert_v1_to_v2(df_v1)
            st.success("แปลง v1 → v2 สำเร็จ")
            st.caption("นำไปใช้ต่อได้ทันที หรือดาวน์โหลดเป็น CSV")
            st.dataframe(converted_df.head(12), use_container_width=True)
            # Download converted CSV
            buf = io.StringIO()
            converted_df.to_csv(buf, index=False, encoding="utf-8-sig")
            st.download_button("⬇️ ดาวน์โหลด CSV v2 (จาก v1)", data=buf.getvalue().encode("utf-8-sig"),
                               file_name="converted_v2.csv", mime="text/csv")

            if convert_info:
                with st.expander("รายละเอียดการแปลง"):
                    st.write(convert_info)

            # Allow adopting converted data as active_df
            if st.button("ใช้ข้อมูล v2 ที่แปลงแล้วกับแอปนี้"):
                st.session_state["active_df"] = converted_df.copy()
                st.session_state["csv_fingerprint"] = ("converted_v2_inline", tuple(converted_df.columns), len(converted_df))
                st.success("ตั้งค่าข้อมูลสำหรับแอปเป็น v2 ที่แปลงแล้ว")

    # Load data — prefer session_state active_df if set by conversion
    if "active_df" in st.session_state and st.session_state["active_df"] is not None and len(st.session_state["active_df"])>0:
        df = st.session_state["active_df"].copy()
    else:
        df = canonicalize_columns(try_read_table(csv_main))

    # Auto-detect v1 and offer auto-convert hint
    if not df.empty and detect_v1_schema(df):
        st.warning("ตรวจพบว่าไฟล์นี้อาจเป็นรูปแบบ v1 — คุณสามารถใช้ตัวช่วยแปลง v1 → v2 ที่ Sidebar")

    if df.empty:
        st.warning("อัปโหลด CSV v2 หรือแปลงจาก v1 ก่อน")
        st.stop()

    # Ensure important v2 columns exist
    for c in ["no", "student_id", "name", "sem1", "sem2", "total", "rating", "grade", "year"]:
        if c not in df.columns:
            df[c] = ""

    # Reorder
    pref = ["no", "student_id", "name", "sem1", "sem2", "total", "rating", "grade", "year"]
    ordered = [c for c in pref if c in df.columns] + [c for c in df.columns if c not in pref]
    active_df = df[ordered]

    # Dtype fixes for numeric columns
    for _c in ["no", "sem1", "sem2", "total"]:
        if _c in active_df.columns:
            active_df[_c] = pd.to_numeric(active_df[_c], errors="coerce")

    st.subheader("📚 ข้อมูล (preview) — แก้ไขได้ตรงนี้")
    # Initialize or refresh session_state active_df when CSV changes (by basic fingerprint)
    csv_name = getattr(csv_main, "name", None)
    fingerprint = (csv_name, tuple(active_df.columns), len(active_df)) if csv_name else st.session_state.get("csv_fingerprint", None)
    if ("active_df" not in st.session_state) or (st.session_state.get("csv_fingerprint") != fingerprint):
        st.session_state["active_df"] = active_df.copy()
        st.session_state["csv_fingerprint"] = fingerprint

    edited = st.data_editor(
        st.session_state["active_df"],
        use_container_width=True,
        num_rows="dynamic",
        hide_index=False,
        column_config={
            "no": st.column_config.NumberColumn("No", step=1),
            "student_id": st.column_config.TextColumn("Student ID"),
            "name": st.column_config.TextColumn("Name"),
            "sem1": st.column_config.NumberColumn("Semester 1"),
            "sem2": st.column_config.NumberColumn("Semester 2"),
            "total": st.column_config.NumberColumn("Total"),
            "rating": st.column_config.TextColumn("Rating"),
            "grade": st.column_config.TextColumn("Grade"),
            "year": st.column_config.TextColumn("Year"),
        },
        key="editable_data",
    )
    # Persist edits for preview/export
    st.session_state["active_df"] = edited
    active_df = st.session_state["active_df"]
    st.caption("การแก้ไขในตารางนี้จะถูกใช้ทั้งพรีวิวและส่งออก PDF")

with colR:
    st.subheader("🧩 Preset (.json) — รวม Body + Cover (cover ใช้แถว 0 เสมอ)")

    # Init session states
    if "fields_df" not in st.session_state:
        st.session_state["fields_df"] = build_field_df(active_df.columns.tolist(), DEFAULT_FIELDS)
    if "cover_fields_df" not in st.session_state:
        st.session_state["cover_fields_df"] = build_field_df(active_df.columns.tolist(), DEFAULT_COVER_FIELDS)
    if "preset_loaded" not in st.session_state:
        st.session_state["preset_loaded"] = False

    with st.expander("Import / Export", expanded=False):
        col_i, col_e = st.columns(2)
        with col_i:
            preset_json = st.file_uploader("นำเข้า Preset (.json)", type=["json"], key="unified_preset_upload")
            preset_url = st.text_input("หรือระบุ URL (raw GitHub) สำหรับ Preset", value=DEFAULT_PRESET_URL)
            load_from_url = st.button("⬇️ โหลด Preset จาก URL")

            def _apply_unified_preset_bytes(preset_bytes: bytes):
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
                        st.success("นำเข้า Preset (Body + Cover) สำเร็จ")
                except Exception as e:
                    st.error(f"อ่านไฟล์/URL Preset ไม่ได้: {e}")

            if preset_json is not None:
                _apply_unified_preset_bytes(preset_json.read())

            if load_from_url:
                b = fetch_default_json(preset_url)
                if b:
                    _apply_unified_preset_bytes(b)

            # Auto-load from DEFAULT_PRESET_URL once if nothing uploaded yet
            if not st.session_state["preset_loaded"] and preset_json is None and not load_from_url:
                auto_b = fetch_default_json(DEFAULT_PRESET_URL)
                if auto_b:
                    _apply_unified_preset_bytes(auto_b)
                    st.info(f"Preset: โหลดจาก GitHub อัตโนมัติ\n{DEFAULT_PRESET_URL}")

        with col_e:
            try:
                payload = {
                    "version": 11,
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

try:
    if page_type == "Body":
        body_src = tpl_pdf.getvalue() if tpl_pdf is not None else default_body_bytes
        if body_src is not None:
            st.image(
                render_preview_with_pymupdf(body_src, st.session_state["fields_df"], record_body, 2.0),
                caption=f"Body — {get_record_display(record_body)}",
                use_container_width=True,
            )
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
                    # Overlay cover fields with row 0
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

            # Insert body pages per student
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
            total_pages = len(active_df) + (1 if (cover_active and (tpl_cover_pdf is not None or default_cover_bytes is not None)) else 0)
            st.success(f"เสร็จแล้ว: {total_pages} หน้า (ปก 1 + เนื้อหา {len(active_df)})")
            st.download_button("⬇️ ดาวน์โหลด PDF", data=pdf_bytes,
                               file_name="exported_batch_with_global_cover.pdf", mime="application/pdf")
    except Exception as e:
        st.error(f"ส่งออกไม่สำเร็จ: {e}")

st.markdown("---")
st.caption("CSV v2: No, Student ID, Name, Semester 1, Semester 2, Total, Rating, Grade, Year • Preset รวม (data_row_index=0) • PDF-only • มีตัวช่วยแปลง v1→v2 และสถานะโหลด GitHub ชัดเจน")
