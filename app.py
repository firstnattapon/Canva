# -*- coding: utf-8 -*-
# =============================================================
# Streamlit App: PDF Template Overlay + CSV -> Batch PDF Export (PDF-only)
# Spec:
#   ✅ รับไฟล์ CSV เดียว: No, Student ID, Name, Semester 1, Semester 2, Total, Rating, Grade, Year
#   ✅ Template เฉพาะ PDF เท่านั้น (ตัดโหมดภาพทิ้ง)
#   ✅ Cover ใช้ข้อมูลเหมือน Body แต่ Layout แยก (ปกหน้าแรกครั้งเดียว)
#   ✅ Preset (.json) รวมไฟล์เดียว (Body + Cover + cover_data_index)
#   ✅ พรีวิวสด (Body/Cover) — ใช้ use_container_width
#
# Install deps:
#   pip install streamlit pandas pillow pymupdf
# =============================================================

import io
import json
from typing import List

import streamlit as st
import pandas as pd

# PDF dependency
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

from PIL import Image  # pillow required by Streamlit image machinery (not used for uploads)

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

st.set_page_config(page_title="PDF Layout Editor — CSV (Unified) → Batch PDF [PDF-only]", layout="wide")
st.title("🖨️ PDF Layout Editor — CSV เดียว → Batch PDF (PDF-only, ปกครั้งเดียว)")
st.caption("Cover ใช้ข้อมูลชุดเดียวกับ Body แต่มี Layout แยก • ปกอยู่หน้าแรก 1 ครั้ง • Preset .json รวม Body/Cover • รองรับ **PDF เท่านั้น**")

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
    cover_data_index = st.number_input("ข้อมูลสำหรับปก: ใช้แถวที่", min_value=0, value=0, step=1)

    st.header("📥 ข้อมูล (CSV เดียว)")
    csv_main = st.file_uploader("CSV หลัก (ตามสคีมาใหม่)", type=["csv", "xlsx", "xls"])

with colL:
    # Load data — single CSV
    df = canonicalize_columns(try_read_table(csv_main))

    if df.empty:
        st.warning("อัปโหลด CSV ตามสคีมาใหม่ก่อน")
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
    st.subheader("🧩 Preset (.json) — รวม Body + Cover")

    # Init session states
    if "fields_df" not in st.session_state:
        st.session_state["fields_df"] = build_field_df(active_df.columns.tolist(), DEFAULT_FIELDS)
    if "cover_fields_df" not in st.session_state:
        st.session_state["cover_fields_df"] = build_field_df(active_df.columns.tolist(), DEFAULT_COVER_FIELDS)

    with st.expander("Import / Export", expanded=False):
        col_i, col_e = st.columns(2)
        with col_i:
            preset_json = st.file_uploader("นำเข้า Preset (.json)", type=["json"], key="unified_preset_upload")
            if preset_json is not None:
                try:
                    raw = json.load(preset_json)
                    # Back-compat: list/fields => Body only
                    if isinstance(raw, list) or "fields" in raw:
                        fields_list = raw.get("fields", raw if isinstance(raw, list) else [])
                        new_df = pd.DataFrame(fields_list)
                        req = ["field_key", "label", "active", "x", "y", "font", "size", "transform", "align"]
                        missing = [c for c in req if c not in new_df.columns]
                        if missing:
                            raise ValueError(f"ขาดคีย์ใน JSON: {missing}")
                        st.session_state["fields_df"] = new_df[req]
                        st.info("โหลดเฉพาะ Body (legacy) แล้ว")
                    else:
                        body = raw.get("body", {})
                        cover = raw.get("cover", {})
                        if "fields" in body:
                            st.session_state["fields_df"] = pd.DataFrame(body["fields"])
                        if "fields" in cover:
                            st.session_state["cover_fields_df"] = pd.DataFrame(cover["fields"])
                        if "data_row_index" in cover:
                            st.session_state["cover_data_index_from_preset"] = int(cover["data_row_index"])
                        st.success("นำเข้า Preset (Body + Cover) สำเร็จ")
                except Exception as e:
                    st.error(f"อ่านไฟล์ JSON ไม่ได้: {e}")
        with col_e:
            try:
                payload = {
                    "version": 7,
                    "body": {"fields": st.session_state["fields_df"].to_dict(orient="records")},
                    "cover": {
                        "fields": st.session_state["cover_fields_df"].to_dict(orient="records"),
                        "data_row_index": int(cover_data_index),
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

# Cover record (global)
if 'cover_data_index_from_preset' in st.session_state:
    cov_idx = int(min(max(0, st.session_state['cover_data_index_from_preset']), len(active_df)-1))
    st.info(f"Cover ใช้ index จาก preset: {cov_idx} (ปรับที่ Sidebar เพื่อเปลี่ยน)")
else:
    cov_idx = int(min(max(0, cover_data_index), len(active_df)-1))

record_cover = active_df.iloc[cov_idx]

page_type = st.radio("หน้าไหน", ["Body", "Cover"], index=0, horizontal=True)

try:
    if page_type == "Body":
        if tpl_pdf is not None:
            st.image(
                render_preview_with_pymupdf(tpl_pdf.getvalue(), st.session_state["fields_df"], record_body, 2.0),
                caption=f"Body — {get_record_display(record_body)}",
                use_container_width=True,
            )
            st.caption("Body: หน่วย X/Y = จุด (pt) — มุมซ้ายบนคือ (0,0)")
        else:
            st.info("อัปโหลด Template PDF ของ Body ก่อน")
    else:  # Cover
        if cover_active:
            if tpl_cover_pdf is not None:
                st.image(
                    render_preview_with_pymupdf(tpl_cover_pdf.getvalue(), st.session_state["cover_fields_df"], record_cover, 2.0),
                    caption=f"Cover — ใช้ข้อมูลแถวที่ {cov_idx} ({get_record_display(record_cover)})",
                    use_container_width=True,
                )
                st.caption("Cover: หน่วย X/Y = จุด (pt) — มุมซ้ายบนคือ (0,0) • ใช้ข้อมูลจากแถวที่ระบุ")
            else:
                st.info("เปิด Active ปก แล้ว—อัปโหลด Cover Template PDF")
        else:
            st.info("ยังไม่ได้เปิด Active ปก (หน้าแรกครั้งเดียว)")
except Exception as e:
    st.error(f"พรีวิวผิดพลาด: {e}")

st.divider()
st.subheader("📦 ส่งออก PDF ทั้งชุด (ปกหน้าแรก 1 ครั้ง + Body 1 หน้า/คน)")

if st.button("🚀 Export PDF"):
    try:
        if tpl_pdf is None:
            st.error("กรุณาอัปโหลด Template PDF ของ Body ก่อน")
        else:
            body_tpl_bytes = tpl_pdf.getvalue()
            out = fitz.open()

            # Insert global cover once using selected record
            if cover_active and (tpl_cover_pdf is not None):
                t_cover = fitz.open(stream=tpl_cover_pdf.getvalue(), filetype="pdf")
                out.insert_pdf(t_cover, from_page=0, to_page=0)
                page0 = out[-1]
                # Overlay cover fields with chosen record
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

            # Insert body pages per student
            for _, rec in active_df.iterrows():
                t_body = fitz.open(stream=body_tpl_bytes, filetype="pdf")
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
            total_pages = len(active_df) + (1 if (cover_active and tpl_cover_pdf is not None) else 0)
            st.success(f"เสร็จแล้ว: {total_pages} หน้า (ปก 1 + เนื้อหา {len(active_df)})")
            st.download_button("⬇️ ดาวน์โหลด PDF", data=pdf_bytes,
                               file_name="exported_batch_with_global_cover.pdf", mime="application/pdf")
    except Exception as e:
        st.error(f"ส่งออกไม่สำเร็จ: {e}")

st.markdown("---")
st.caption("รับ CSV เดียวคอลัมน์: No, Student ID, Name, Semester 1, Semester 2, Total, Rating, Grade, Year • Preset รวม: { version, body.fields[], cover.fields[], cover.data_row_index } • รองรับไฟล์ PDF เท่านั้น • ใช้ use_container_width เสมอ")
