
# -*- coding: utf-8 -*-
# =========================
# Streamlit App: PDF Template Overlay + CSV -> Batch PDF Export
# (with Layout Editor JSON import/export)
# =========================

import io
import json
from typing import List

import streamlit as st
import pandas as pd

# Optional deps
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

from PIL import Image, ImageDraw, ImageFont

try:
    from pypdf import PdfReader, PdfWriter, PageObject
except Exception:
    PdfReader = None
    PdfWriter = None
    PageObject = None

try:
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
except Exception:
    rl_canvas = None

# ------------------ Helpers ------------------

CANONICAL_COLS = {
    "No": "no",
    "Student ID": "student_id",
    "StudentID": "student_id",
    "ID": "student_id",
    "Name - Surname": "name",
    "Name": "name",
    "Idea": "idea",
    "Pronunciation": "pronunciation",
    "Preparedness": "preparedness",
    "Confidence": "confidence",
    "Total (50)": "total",
    "Total": "total",
}

DEFAULT_FIELDS = [
    # field_key, label, active, x, y, font, size, transform, align
    ("name", "Name", True, 140.0, 160.0, "helv", 12, "none", "left"),
    ("student_id", "Student ID", True, 140.0, 180.0, "helv", 11, "none", "left"),
    ("idea", "Idea", False, 400.0, 220.0, "helv", 12, "none", "left"),
    ("pronunciation", "Pronunciation", False, 460.0, 220.0, "helv", 12, "none", "left"),
    ("preparedness", "Preparedness", False, 520.0, 220.0, "helv", 12, "none", "left"),
    ("confidence", "Confidence", False, 580.0, 220.0, "helv", 12, "none", "left"),
    ("total", "Total (50)", True, 640.0, 220.0, "helv", 14, "none", "left"),
]

STD_FONTS = ["helv", "times", "cour"]  # Built-in fonts for PyMuPDF

def try_read_table(uploaded_file) -> pd.DataFrame:
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
            c2 = c.strip().lower().replace(" ", "").replace("-", "").replace("_", "")
            if c2 in ["studentid", "id"]:
                key = "student_id"
            elif c2 in ["name", "namesurname", "namesurmane"]:
                key = "name"
            elif "total" in c2:
                key = "total"
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

def build_field_df(existing_cols: List[str]) -> pd.DataFrame:
    rows = []
    existing = set(existing_cols)
    known = set()
    for k, label, active, x, y, font, size, transform, align in DEFAULT_FIELDS:
        rows.append({
            "field_key": k,
            "label": label,
            "active": active if k in existing or k in ["name","student_id","total"] else False,
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

def apply_transform(text: str, mode: str) -> str:
    if text is None:
        return ""
    if mode == "upper":
        return str(text).upper()
    if mode == "lower":
        return str(text).lower()
    if mode == "title":
        return str(text).title()
    return str(text)

def get_record_display(rec: pd.Series, key_cols=("student_id","name")) -> str:
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
        if key not in record:
            continue
        val = record[key]
        if pd.isna(val):
            continue
        text = apply_transform(val, row["transform"])
        x, y = float(row["x"]), float(row["y"])
        font = row.get("font", "helv")
        size = float(row.get("size", 12))
        try:
            p.insert_text((x, y), text, fontname=font if font in STD_FONTS else "helv",
                          fontsize=size, color=(0,0,0))
        except Exception:
            p.insert_text((x, y), text, fontname="helv", fontsize=size, color=(0,0,0))

    mat = fitz.Matrix(scale, scale)
    pix = p.get_pixmap(matrix=mat, alpha=False)
    from PIL import Image
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    td.close()
    newdoc.close()
    return img

from PIL import Image

def draw_on_image(img: Image.Image, fields_df: pd.DataFrame, record: pd.Series) -> Image.Image:
    im = img.copy()
    draw = ImageDraw.Draw(im)
    for _, row in fields_df.iterrows():
        if not row["active"]:
            continue
        key = row["field_key"]
        if key not in record:
            continue
        val = record[key]
        if pd.isna(val):
            continue
        text = apply_transform(val, row["transform"])
        x, y = int(row["x"]), int(row["y"])
        size = int(row.get("size", 12))
        try:
            font = ImageFont.truetype("DejaVuSans.ttf", size=size)
        except Exception:
            font = ImageFont.load_default()
        draw.text((x, y), str(text), font=font, fill=(0,0,0))
    return im

def export_batch_pdf_with_pymupdf(template_bytes: bytes,
                                  fields_df: pd.DataFrame,
                                  df: pd.DataFrame) -> bytes:
    if fitz is None:
        raise RuntimeError("ต้องติดตั้ง PyMuPDF (fitz) เพื่อส่งออก PDF")
    tdoc = fitz.open(stream=template_bytes, filetype="pdf")
    if tdoc.page_count < 1:
        raise RuntimeError("Template PDF ว่าง")
    out = fitz.open()
    for _, rec in df.iterrows():
        out.insert_pdf(tdoc, from_page=0, to_page=0)
        page = out[-1]
        for _, row in fields_df.iterrows():
            if not row["active"]:
                continue
            key = row["field_key"]
            if key not in rec:
                continue
            val = rec[key]
            if pd.isna(val):
                continue
            text = apply_transform(val, row["transform"])
            x, y = float(row["x"]), float(row["y"])
            font = row.get("font", "helv")
            size = float(row.get("size", 12))
            try:
                page.insert_text((x, y), str(text),
                                 fontname=font if font in STD_FONTS else "helv",
                                 fontsize=size, color=(0,0,0))
            except Exception:
                page.insert_text((x, y), str(text),
                                 fontname="helv", fontsize=size, color=(0,0,0))
    pdf_bytes = out.tobytes()
    out.close()
    tdoc.close()
    return pdf_bytes

# ------------------ Streamlit UI ------------------

st.set_page_config(page_title="PDF Layout Editor — CSV → Batch PDF", layout="wide")
st.title("🖨️ PDF Layout Editor — CSV → Batch PDF (1 หน้า/คน)")
st.caption("อัปโหลด **Template PDF (Canva หน้าเดียว)** + CSV เทอม 1/2 → ตั้งค่า X/Y ต่อฟิลด์ → พรีวิวสด → ส่งออก PDF ทั้งชุด")

colL, colR = st.columns([1.2, 1.0], gap="large")

with st.sidebar:
    st.header("📄 เทมเพลต")
    tpl_pdf = st.file_uploader("อัปโหลด Template PDF (หน้าเดียว)", type=["pdf"])
    tpl_img = st.file_uploader("หรืออัปโหลดภาพ (PNG/JPG) แทน", type=["png","jpg","jpeg"])

    if tpl_pdf is None and tpl_img is None:
        st.info("แนะนำให้อัปโหลด **PDF หน้าเดียว** จาก Canva")
    if tpl_pdf is not None and fitz is None:
        st.warning("ติดตั้ง `pymupdf` เพื่อพรีวิว/ส่งออกจาก PDF\n\n`pip install pymupdf`")

    st.header("📥 ข้อมูลนักเรียน")
    csv_t1 = st.file_uploader("CSV เทอม 1", type=["csv","xlsx","xls"])
    csv_t2 = st.file_uploader("CSV เทอม 2 (ไม่บังคับ)", type=["csv","xlsx","xls"])

    st.divider()
    join_key = st.selectbox("คีย์สำหรับจับคู่เทอม 1/2", ["student_id", "name"], index=0,
                            help="ต้องมีคอลัมน์นี้ในทั้งสองไฟล์")
    use_term = st.radio("เลือกชุดคะแนนสำหรับแสดง/พิมพ์", ["เทอม 1", "เทอม 2 (ถ้ามี)", "เฉลี่ย (Total เท่านั้น)"], index=0)

with colL:
    df1 = canonicalize_columns(try_read_table(csv_t1))
    df2 = canonicalize_columns(try_read_table(csv_t2))

    if df1.empty:
        st.warning("อัปโหลด CSV เทอม 1 ก่อน")
        st.stop()

    needed = ["student_id", "name"]
    for c in needed:
        if c not in df1.columns:
            df1[c] = ""

    active_df = df1.copy()
    if not df2.empty and join_key in df2.columns:
        merged = pd.merge(df1, df2, how="inner", on=join_key, suffixes=("_t1", "_t2"))
        if use_term.startswith("เทอม 1"):
            cols = []
            for c in df1.columns:
                if c + "_t1" in merged.columns:
                    cols.append(c + "_t1")
                elif c in merged.columns:
                    cols.append(c)
            active_df = merged[cols].copy()
            active_df.columns = [c.replace("_t1", "") for c in active_df.columns]
        elif use_term.startswith("เทอม 2"):
            cols = []
            base_cols = set(df1.columns).union(df2.columns)
            for c in base_cols:
                if c + "_t2" in merged.columns:
                    cols.append(c + "_t2")
                elif c in merged.columns:
                    cols.append(c)
            active_df = merged[cols].copy()
            active_df.columns = [c.replace("_t2", "") for c in active_df.columns]
        else:
            active_df = merged.copy()
            def pick(col):
                if col + "_t1" in merged and col + "_t2" in merged:
                    return (merged[col + "_t1"] + merged[col + "_t2"]) / 2.0
                if col + "_t1" in merged:
                    return merged[col + "_t1"]
                if col + "_t2" in merged:
                    return merged[col + "_t2"]
                return merged[col] if col in merged else None
            base_cols = ["student_id","name","idea","pronunciation","preparedness","confidence","total"]
            data = {}
            for c in base_cols:
                series = pick(c)
                if series is not None:
                    data[c] = series
            active_df = pd.DataFrame(data)
    else:
        active_df = df1.copy()

    if "no" not in active_df.columns and "No" in df1.columns:
        active_df["no"] = df1["No"]
    pref = ["no","student_id","name","idea","pronunciation","preparedness","confidence","total"]
    ordered = [c for c in pref if c in active_df.columns] + [c for c in active_df.columns if c not in pref]
    active_df = active_df[ordered]

    st.subheader("📚 ข้อมูล (preview)")
    st.dataframe(active_df.head(10), use_container_width=True)

with colR:
    st.subheader("⚙️ Layout Editor")

    if "fields_df" not in st.session_state:
        st.session_state["fields_df"] = build_field_df(active_df.columns.tolist())

    # --------- Import/Export JSON for Layout Editor ---------
    with st.expander("🧩 Presets (Import/Export .json)", expanded=False):
        col_i, col_e = st.columns(2)
        with col_i:
            layout_json = st.file_uploader("นำเข้า Layout (.json)", type=["json"], key="layout_json_upload")
            if layout_json is not None:
                try:
                    raw = json.load(layout_json)
                    fields_list = raw.get("fields", raw if isinstance(raw, list) else None)
                    if not isinstance(fields_list, list):
                        raise ValueError("โครงสร้าง JSON ต้องเป็น { 'fields': [...] } หรือเป็นลิสต์ของคอนฟิกฟิลด์")
                    new_df = pd.DataFrame(fields_list)
                    req = ["field_key","label","active","x","y","font","size","transform","align"]
                    missing = [c for c in req if c not in new_df.columns]
                    if missing:
                        raise ValueError(f"ขาดคีย์ใน JSON: {missing}")
                    # เลือกเก็บเฉพาะคอลัมน์ที่จำเป็น เผื่อมีของแถมมา
                    new_df = new_df[req]
                    st.session_state["fields_df"] = new_df
                    st.success("นำเข้า Layout สำเร็จ")
                except Exception as e:
                    st.error(f"อ่านไฟล์ JSON ไม่ได้: {e}")

        with col_e:
            try:
                payload = {
                    "version": 1,
                    "fields": st.session_state["fields_df"].to_dict(orient="records")
                }
                buf = io.StringIO()
                json.dump(payload, buf, ensure_ascii=False, indent=2)
                st.download_button(
                    "⬇️ Export Layout (.json)",
                    data=buf.getvalue().encode("utf-8"),
                    file_name="layout_preset.json",
                    mime="application/json"
                )
            except Exception as e:
                st.error(f"Export JSON ผิดพลาด: {e}")
    # --------------------------------------------------------

    edited = st.data_editor(
        st.session_state["fields_df"],
        use_container_width=True,
        hide_index=True,
        column_config={
            "field_key": st.column_config.TextColumn("field_key", disabled=True),
            "label": st.column_config.TextColumn("Label"),
            "active": st.column_config.CheckboxColumn("Active"),
            "x": st.column_config.NumberColumn("X", step=1, format="%.1f"),
            "y": st.column_config.NumberColumn("Y", step=1, format="%.1f"),
            "font": st.column_config.SelectboxColumn("Font", options=STD_FONTS),
            "size": st.column_config.NumberColumn("Size (pt)", step=1, format="%.0f"),
            "transform": st.column_config.SelectboxColumn("Case", options=["none","upper","lower","title"]),
            "align": st.column_config.SelectboxColumn("Align", options=["left","center","right"]),
        },
        key="fields_editor"
    )
    st.session_state["fields_df"] = edited

    st.subheader("🔎 พรีวิว")
    idx_options = list(range(len(active_df)))
    if len(idx_options) == 0:
        st.stop()
    rec_idx = st.number_input("แถวที่ต้องการพรีวิว (index)", min_value=0, max_value=len(idx_options)-1, value=0, step=1)
    record = active_df.iloc[int(rec_idx)]

    preview_img = None
    try:
        if tpl_pdf is not None and fitz is not None:
            template_bytes = tpl_pdf.read()
            preview_img = render_preview_with_pymupdf(template_bytes, edited, record, scale=2.0)
            st.image(preview_img, caption=get_record_display(record), use_column_width=True)
            st.caption("หน่วย X/Y = จุด (pt) — มุมซ้ายบนคือ (0,0)")
        elif tpl_img is not None:
            img = Image.open(tpl_img).convert("RGB")
            preview_img = draw_on_image(img, edited, record)
            st.image(preview_img, caption=get_record_display(record), use_column_width=True)
            st.caption("หน่วย X/Y = พิกเซล (px) — มุมซ้ายบนคือ (0,0)")
        else:
            st.info("อัปโหลดเทมเพลตก่อน แล้วพรีวิวจะขึ้นอัตโนมัติ")
    except Exception as e:
        st.error(f"พรีวิวผิดพลาด: {e}")

    st.divider()
    st.subheader("📦 ส่งออก PDF ทั้งชุด (1 หน้า/คน)")

    if st.button("🚀 Export PDF"):
        try:
            if tpl_pdf is not None and fitz is not None:
                template_bytes = tpl_pdf.getvalue() if hasattr(tpl_pdf, "getvalue") else tpl_pdf.read()
                pdf_bytes = export_batch_pdf_with_pymupdf(template_bytes, edited, active_df)
                st.success(f"เสร็จแล้ว: {len(active_df)} หน้า")
                st.download_button("⬇️ ดาวน์โหลด PDF", data=pdf_bytes, file_name="exported_batch.pdf", mime="application/pdf")
            elif tpl_img is not None:
                pages = []
                base_img = Image.open(tpl_img).convert("RGB")
                for _, rec in active_df.iterrows():
                    page = draw_on_image(base_img, edited, rec)
                    pages.append(page)
                buf = io.BytesIO()
                if len(pages) == 1:
                    pages[0].save(buf, format="PDF")
                else:
                    pages[0].save(buf, format="PDF", save_all=True, append_images=pages[1:])
                pdf_bytes = buf.getvalue()
                st.success(f"เสร็จแล้ว: {len(pages)} หน้า (จากภาพ)")
                st.download_button("⬇️ ดาวน์โหลด PDF", data=pdf_bytes, file_name="exported_batch.pdf", mime="application/pdf")
            else:
                st.error("กรุณาอัปโหลดเทมเพลตก่อน (PDF หรือ PNG/JPG)")
        except Exception as e:
            st.error(f"ส่งออกไม่สำเร็จ: {e}")

st.markdown("---")
st.caption("ทริก: ใน Canva ส่งออกเทมเพลตเป็น **PDF Standard (หน้าเดียว)** เพื่อความคมชัด • ฟอนต์ที่ใช้ในการพิมพ์คือ standard fonts ของ PDF: helv / times / cour")
