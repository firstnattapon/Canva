# streamlit run app.py
# -*- coding: utf-8 -*-

import io
import os
from datetime import datetime
import math

import streamlit as st
import pandas as pd

# PDF rendering
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

from pypdf import PdfReader, PdfWriter, PageObject

from PIL import Image

# =========================
# CONFIG หน้า & ฟอนต์
# =========================
st.set_page_config(page_title="Conversation Test Result Grade 2/5 — PDF Export", layout="wide")

DEFAULT_CSV_PATH = "/mnt/data/สเปรดชีตไม่มีชื่อ - P.2.csv"  # ใช้ได้ในสภาพแวดล้อมที่รองรับเท่านั้น
OUTPUT_NAME = "Conversation_Test_Result_Grade_2-5.pdf"

# พยายามหา NotoSansThai ก่อน ถ้าไม่เจอจะ fallback ไป DejaVuSans
FALLBACK_FONT = "DejaVuSans"
FONT_NORMAL_NAME = "NotoSansThai-Regular"
FONT_BOLD_NAME = "NotoSansThai-Bold"

# =========================
# ฟังก์ชันช่วยงาน
# =========================
def try_register_thai_fonts(uploaded_regular: bytes | None, uploaded_bold: bytes | None):
    """ลงทะเบียนฟอนต์ไทย (NotoSansThai) ถ้ามี; ไม่งั้น fallback เป็น DejaVuSans"""
    registered = False
    # 1) จากไฟล์อัปโหลด
    if uploaded_regular:
        try:
            pdfmetrics.registerFont(TTFont(FONT_NORMAL_NAME, io.BytesIO(uploaded_regular)))
            registered = True
        except Exception:
            pass
    if uploaded_bold:
        try:
            pdfmetrics.registerFont(TTFont(FONT_BOLD_NAME, io.BytesIO(uploaded_bold)))
        except Exception:
            # ไม่มี Bold ก็ไม่เป็นไร ใช้ Regular แทน
            pass

    # 2) จากระบบ (ถ้าไม่ได้อัปโหลด)
    search_candidates = [
        "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansThai-Regular.otf",
    ]
    if not registered:
        for p in search_candidates:
            if os.path.exists(p):
                try:
                    pdfmetrics.registerFont(TTFont(FONT_NORMAL_NAME, p))
                    registered = True
                    break
                except Exception:
                    continue

    # 3) ลงทะเบียน fallback
    if not registered:
        # Fallback จะใช้ชื่อ DejaVuSans ซึ่งโดยปกติ reportlab มากับระบบหลายที่
        try:
            pdfmetrics.getFont(FALLBACK_FONT)
        except Exception:
            # ถ้าไม่มีจริงๆ ให้บอกผู้ใช้ให้อัปโหลดฟอนต์
            st.warning("หา NotoSansThai/DejaVuSans ไม่เจอ โปรดอัปโหลดฟอนต์ .ttf อย่างน้อย 1 ไฟล์")
        return FALLBACK_FONT, FALLBACK_FONT

    # ถ้าได้ Regular แล้ว แต่ยังไม่มี Bold ให้ผูก Bold = Regular
    try:
        pdfmetrics.getFont(FONT_BOLD_NAME)
    except Exception:
        return FONT_NORMAL_NAME, FONT_NORMAL_NAME
    return FONT_NORMAL_NAME, FONT_BOLD_NAME


def guess_col(df: pd.DataFrame, candidates: list[str], default=None):
    cols = [c.lower().strip() for c in df.columns]
    for name in candidates:
        if name.lower() in cols:
            return df.columns[cols.index(name.lower())]
    # ลองแบบ contains
    for c in df.columns:
        cl = c.lower().replace(" ", "")
        for name in candidates:
            if name.lower().replace(" ", "") in cl:
                return c
    return default


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # เดาแมปคอลัมน์หลัก ๆ แบบยืดหยุ่น (รองรับไทย/อังกฤษ/สะกดหลากหลาย)
    mapping = {}

    mapping["no"] = guess_col(df, ["no", "ลำดับ", "ลำดับที่", "เลขที่", "#"])
    mapping["studentId"] = guess_col(df, ["studentid", "student id", "id", "รหัส", "เลขประจำตัว"])
    mapping["name"] = guess_col(df, ["name", "ชื่อ-สกุล", "ชื่อ", "student name", "นักเรียน"])

    mapping["idea"] = guess_col(df, ["idea", "ไอเดีย", "ความคิด", "creativity", "ความคิดสร้างสรรค์"])
    mapping["pronunciation"] = guess_col(df, ["pronunciation", "การออกเสียง", "ออกเสียง"])
    mapping["preparedness"] = guess_col(df, ["preparedness", "ความพร้อม", "การเตรียมตัว"])
    mapping["confidence"] = guess_col(df, ["confidence", "ความมั่นใจ"])
    mapping["total"] = guess_col(df, ["total", "คะแนนรวม", "รวม", "finaltotal"])
    mapping["grade"] = guess_col(df, ["grade", "เกรด", "ผลการเรียน", "ระดับ"])

    # สร้าง df ใหม่ตามลำดับคอลัมน์มาตรฐาน
    out = pd.DataFrame()
    for k in ["no", "studentId", "name", "idea", "pronunciation", "preparedness", "confidence", "total", "grade"]:
        if mapping.get(k) in df.columns:
            out[k] = df[mapping[k]]
        else:
            out[k] = ""

    # แปลงชนิดข้อมูลบางคอลัมน์ให้ปลอดภัย
    for k in ["idea", "pronunciation", "preparedness", "confidence", "total"]:
        try:
            out[k] = pd.to_numeric(out[k], errors="coerce")
        except Exception:
            pass
    # เติมลำดับถ้ายังว่าง
    if out["no"].isna().all() or (out["no"] == "").all():
        out["no"] = range(1, len(out) + 1)

    # ตัดสเปซชื่อ
    out["name"] = out["name"].astype(str).str.strip()

    return out


def read_background_as_pdf(base_file: bytes, filename: str) -> PdfReader:
    """รับไฟล์เทมเพลตจาก Canva ที่เป็น PDF หรือ PNG แล้วคืนเป็น PdfReader 1 หน้า
       - ถ้าเป็น PNG/JPG: แปลงเป็น PDF ขนาดภาพ 1:1
       - ถ้าเป็น PDF: อ่านตรง ๆ
    """
    name_lower = filename.lower()
    if name_lower.endswith(".pdf"):
        return PdfReader(io.BytesIO(base_file))
    else:
        # แปลงรูปเป็น PDF 1 หน้า
        img = Image.open(io.BytesIO(base_file)).convert("RGB")
        w, h = img.size  # พิกเซล
        # สร้าง PDF ขนาดเท่ารูป โดย 1 px ~ 1 pt (พอใช้ได้สำหรับ overlay ตรงตำแหน่ง)
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(w, h))
        c.drawImage(ImageReader(img), 0, 0, width=w, height=h, mask='auto')
        c.showPage()
        c.save()
        buf.seek(0)
        return PdfReader(buf)


def make_overlay_page(width_pt, height_pt, rows, start_index, rows_per_page,
                      cols_config, row_height, top_y, font_regular, font_bold, font_size, df):
    """สร้างหน้า overlay PDF หนึ่งหน้า (เฉพาะตัวหนังสือในตาราง) แล้วคืนเป็น PdfReader/Page"""
    page_buf = io.BytesIO()
    canv = rl_canvas.Canvas(page_buf, pagesize=(width_pt, height_pt))

    # เลือกฟอนต์
    canv.setFont(font_regular, font_size)

    # เขียนแต่ละแถว
    for i in range(rows_per_page):
        row_idx = start_index + i
        if row_idx >= rows:
            break
        y = top_y - i * row_height

        # ดึงค่าจาก df
        rec = df.iloc[row_idx]
        for col in cols_config:
            key = col["key"]
            x = col["x"]
            align = col.get("align", "left")
            max_width = col.get("max_width", None)
            text = "" if pd.isna(rec.get(key, "")) else str(rec.get(key, ""))

            # จัดรูปแบบตัวเลขสวย ๆ
            if key in ["idea", "pronunciation", "preparedness", "confidence", "total"]:
                try:
                    v = float(rec.get(key))
                    # ถ้าเป็นจำนวนเต็ม ไม่ต้องมีทศนิยม
                    text = f"{int(v)}" if abs(v - int(v)) < 1e-9 else f"{v:.2f}"
                except Exception:
                    pass

            # ย่อชื่อถ้ายาวเกิน
            if max_width:
                # วัดความกว้างคร่าว ๆ โดยอาศัย stringWidth
                while pdfmetrics.stringWidth(text, font_regular, font_size) > max_width and len(text) > 0:
                    text = text[:-1]
                # ใส่ … ถ้าตัด
                # (ถ้าถูกตัดจนหมดจะไม่ใส่)
                if pdfmetrics.stringWidth(text, font_regular, font_size) > max_width and len(text) > 1:
                    text = text[:-1] + "…"

            # วางข้อความ
            if align == "center":
                canv.drawCentredString(x, y, text)
            elif align == "right":
                canv.drawRightString(x, y, text)
            else:
                canv.drawString(x, y, text)

    canv.showPage()
    canv.save()
    page_buf.seek(0)
    overlay_reader = PdfReader(page_buf)
    return overlay_reader.pages[0]


def merge_background_and_overlay(bg_reader: PdfReader, overlay_pages: list[PageObject]) -> bytes:
    """ซ้อน overlay ทุกหน้าไปบนพื้นหลัง (ถ้าพื้นหลังมีหน้าน้อย จะวนใช้หน้าแรกซ้ำ)"""
    out = PdfWriter()
    bg_pages = len(bg_reader.pages)
    for i, ov in enumerate(overlay_pages):
        bg_page = bg_reader.pages[i % bg_pages]
        # ทำสำเนา ไม่แก้ต้นฉบับ
        new_page = PageObject.create_blank_page(width=bg_page.mediabox.width, height=bg_page.mediabox.height)
        new_page.merge_page(bg_page)    # วาง background ก่อน
        new_page.merge_page(ov)         # ซ้อน overlay
        out.add_page(new_page)

    buf = io.BytesIO()
    out.write(buf)
    buf.seek(0)
    return buf.read()


# =========================
# UI
# =========================
st.title("Conversation Test Result Grade 2/5 → PDF (Overlay on Canva Template)")

left, right = st.columns([2, 1])

with left:
    st.subheader("1) เลือกไฟล์ข้อมูล")
    csv_file = None
    default_exists = os.path.exists(DEFAULT_CSV_PATH)
    src = st.radio("แหล่งข้อมูล", ["อัปโหลด CSV", "ใช้ไฟล์ดีฟอลต์ในระบบ" if default_exists else "อัปโหลด CSV"], horizontal=True)
    if src == "อัปโหลด CSV" or not default_exists:
        csv_upl = st.file_uploader("อัปโหลด CSV (UTF-8) จาก Google Sheets/Excel", type=["csv"])
        if csv_upl is not None:
            csv_file = io.BytesIO(csv_upl.read())
    else:
        csv_file = DEFAULT_CSV_PATH

    st.subheader("2) อัปโหลดเทมเพลต (พื้นหลังจาก Canva)")
    bg_upl = st.file_uploader("รองรับ PDF/PNG/JPG — แนะนำ: Export จาก Canva เป็น PDF (พรินต์คุณภาพสูง)", type=["pdf", "png", "jpg", "jpeg"])

    st.subheader("3) (ทางเลือก) อัปโหลดฟอนต์ไทย .ttf")
    colf1, colf2 = st.columns(2)
    font_regular_upl = colf1.file_uploader("NotoSansThai-Regular.ttf", type=["ttf"])
    font_bold_upl = colf2.file_uploader("NotoSansThai-Bold.ttf", type=["ttf"])

with right:
    st.subheader("ตัวเลือกการจัดวาง (ปรับให้ตรงกับเส้นตารางใน Canva)")
    rows_per_page = st.number_input("จำนวนแถวต่อหน้า", min_value=5, max_value=40, value=20, step=1)
    top_y = st.number_input("ตำแหน่ง Y ของแถวบนสุด (pt)", min_value=100, max_value=1500, value=620, step=1)
    row_height = st.number_input("ความสูงแต่ละแถว (pt)", min_value=10, max_value=50, value=22, step=1)
    font_size = st.number_input("ขนาดฟอนต์ (pt)", min_value=8, max_value=16, value=10, step=1)

    st.markdown("**ตำแหน่งคอลัมน์ (x pt จากซ้าย):**")
    # กำหนดตำแหน่งคอลัมน์พื้นฐานให้ใกล้เคียง A4 แนวตั้ง
    x_no = st.number_input("x: No.", value=40, step=1)
    x_id = st.number_input("x: Student ID", value=75, step=1)
    x_name = st.number_input("x: Name", value=160, step=1)
    x_idea = st.number_input("x: Idea", value=355, step=1)
    x_pro = st.number_input("x: Pronunciation", value=408, step=1)
    x_pre = st.number_input("x: Preparedness", value=470, step=1)
    x_conf = st.number_input("x: Confidence", value=535, step=1)
    x_total = st.number_input("x: Total", value=600, step=1)
    x_grade = st.number_input("x: Grade", value=640, step=1)

    maxw_name = st.number_input("Max width: Name (pt)", min_value=0, max_value=500, value=180, step=5)

st.divider()

go = st.button("🚀 สร้าง PDF เหมือนต้นฉบับ")

if go:
    # 0) โหลดฟอนต์
    font_regular, font_bold = try_register_thai_fonts(
        font_regular_upl.read() if font_regular_upl else None,
        font_bold_upl.read() if font_bold_upl else None
    )

    # 1) อ่าน CSV
    if csv_file is None:
        st.error("ยังไม่ได้เลือกไฟล์ CSV")
        st.stop()
    try:
        if isinstance(csv_file, str):  # path
            df = pd.read_csv(csv_file)
        else:
            df = pd.read_csv(csv_file)
    except UnicodeDecodeError:
        # รองรับกรณีเป็น Excel แอบเปลี่ยน encoding
        df = pd.read_csv(csv_file, encoding="utf-8-sig")
    except Exception as e:
        st.error(f"อ่าน CSV ไม่สำเร็จ: {e}")
        st.stop()

    df_norm = normalize_dataframe(df)

    st.success(f"โหลดข้อมูล {len(df_norm)} แถว เรียบร้อย")

    # 2) เทมเพลตพื้นหลัง
    if bg_upl is None:
        st.error("โปรดอัปโหลดพื้นหลังจาก Canva (PDF/PNG/JPG) เพื่อให้ตรง 100%")
        st.stop()

    try:
        bg_reader = read_background_as_pdf(bg_upl.read(), bg_upl.name)
    except Exception as e:
        st.error(f"อ่านเทมเพลตไม่สำเร็จ: {e}")
        st.stop()

    # ใช้หน้าพื้นหลังหน้าแรกเพื่ออ้างอิงขนาดเพจ
    w = float(bg_reader.pages[0].mediabox.width)
    h = float(bg_reader.pages[0].mediabox.height)

    # 3) เตรียมคอนฟิกคอลัมน์
    cols_cfg = [
        {"key": "no", "x": x_no, "align": "center"},
        {"key": "studentId", "x": x_id, "align": "left"},
        {"key": "name", "x": x_name, "align": "left", "max_width": maxw_name},
        {"key": "idea", "x": x_idea, "align": "center"},
        {"key": "pronunciation", "x": x_pro, "align": "center"},
        {"key": "preparedness", "x": x_pre, "align": "center"},
        {"key": "confidence", "x": x_conf, "align": "center"},
        {"key": "total", "x": x_total, "align": "center"},
        {"key": "grade", "x": x_grade, "align": "center"},
    ]

    # 4) เรนเดอร์ overlay ทีละหน้า
    total_rows = len(df_norm)
    pages = math.ceil(total_rows / rows_per_page)
    overlay_pages = []
    for p in range(pages):
        start_idx = p * rows_per_page
        ov = make_overlay_page(
            width_pt=w,
            height_pt=h,
            rows=total_rows,
            start_index=start_idx,
            rows_per_page=rows_per_page,
            cols_config=cols_cfg,
            row_height=row_height,
            top_y=top_y,
            font_regular=font_regular,
            font_bold=font_bold,
            font_size=font_size,
            df=df_norm,
        )
        overlay_pages.append(ov)

    # 5) Merge กับพื้นหลัง
    try:
        merged_pdf_bytes = merge_background_and_overlay(bg_reader, overlay_pages)
    except Exception as e:
        st.error(f"รวมเลเยอร์ไม่สำเร็จ: {e}")
        st.stop()

    # 6) ส่งออก
    # ตั้งชื่อไฟล์ตามต้นฉบับ
    dt = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_name = f"{OUTPUT_NAME[:-4]}_{dt}.pdf"

    st.success("✅ สร้าง PDF เรียบร้อย (ซ้อนทับกับเทมเพลต Canva เพื่อความเหมือน 100%)")
    st.download_button("⬇️ ดาวน์โหลดไฟล์ PDF", data=merged_pdf_bytes, file_name=out_name, mime="application/pdf")

    with st.expander("พรีวิวแถวต้น ๆ ของข้อมูล (ตรวจความถูกต้องเร็ว ๆ)"):
        st.dataframe(df_norm.head(20), use_container_width=True)

    st.info(
        "ถ้าอักษรเหลื่อมเส้นตาราง: ปรับค่าพิกัดในแถบขวา (Top Y, Row Height, x ของแต่ละคอลัมน์, ขนาดฟอนต์).\n"
        "หลักการคือเราใช้หน้าออกจาก Canva เป็นพื้นหลังทั้งหมด แล้วพิมพ์เฉพาะตัวเลข/ชื่อทับลงไป."
    )
