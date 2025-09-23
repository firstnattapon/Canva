# -*- coding: utf-8 -*-
# Streamlit app: พรีวิว PDF แบบ Real-time ปรับ X/Y แล้วเห็นผลสด
# 1 หน้า/นักเรียน 1 คน, รองรับ CSV เทอมเดียวหรือสองเทอม, ใส่ Total(50) ลงเทมเพลต Canva
# แก้ปัญหาตัวหนังสือซ้อนกัน ด้วยการวาด "พื้นขาว" ใต้ข้อความแบบไดนามิก

import io
import base64
import importlib.util
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Conversation → PDF (Realtime Preview)", layout="wide")
st.title("📄 Conversation Result → PDF (Realtime X/Y Preview)")
st.caption("อัปโหลด Template PDF (Canva หน้าเดียว) + CSV เทอม 1/2 → พิมพ์ค่า Total ลงช่องที่เว้นไว้ • พรีวิวสด • กันซ้อนด้วยพื้นขาว")

# ========= Sidebar controls =========
st.sidebar.header("🔧 ตำแหน่งข้อความ (หน่วย pt)")
name_x = st.sidebar.number_input("Name X", 0, 2000, 140, step=1)
name_y = st.sidebar.number_input("Name Y", 0, 2000, 160, step=1)
id_x   = st.sidebar.number_input("Student ID X", 0, 2000, 140, step=1)
id_y   = st.sidebar.number_input("Student ID Y", 0, 2000, 190, step=1)
s1_x   = st.sidebar.number_input("Total S1 X", 0, 2000, 430, step=1)
s1_y   = st.sidebar.number_input("Total S1 Y", 0, 2000, 300, step=1)
s2_x   = st.sidebar.number_input("Total S2 X", 0, 2000, 430, step=1)
s2_y   = st.sidebar.number_input("Total S2 Y", 0, 2000, 340, step=1)

st.sidebar.header("🅰️ ฟอนต์")
font_size = st.sidebar.number_input("ขนาดตัวอักษร", 6, 72, 16, step=1)
bold = st.sidebar.checkbox("ตัวหนา (Helvetica-Bold)", value=True)

st.sidebar.header("🧭 ระบบแกน Y")
top_left_mode = st.sidebar.checkbox("ใช้ Y แบบจากด้านบนลงล่าง (Top-Left Origin)", value=False)
st.sidebar.caption("ปกติ PDF ใช้มุมล่างซ้ายเป็นจุดกำเนิด (y เพิ่มขึ้น = ขึ้นด้านบน). ถ้าติ๊กตัวนี้ y จะวัดจากด้านบนลงล่างแทน")

st.sidebar.header("🧻 กันซ้อน (White-out)")
whiteout = st.sidebar.checkbox("วาดพื้นขาวใต้ข้อความ (กันซ้อน)", value=True)
pad_x = st.sidebar.number_input("Padding X (pt)", 0, 40, 4, 1)
pad_y = st.sidebar.number_input("Padding Y (pt)", 0, 40, 2, 1)
st.sidebar.caption("พื้นขาวจะครอบพื้นที่ตัวหนังสืออัตโนมัติตามความยาวข้อความ + padding")

st.sidebar.header("🔗 การแม็ปข้อมูล")
join_key = st.sidebar.selectbox("คีย์จับคู่ 2 เทอม", ["Student ID", "Name - Surname"], index=0)
when_single = st.sidebar.selectbox("ถ้าอัปโหลด CSV เทอมเดียว ให้ใส่ลงช่อง", ["S1", "S2"], index=0)

st.sidebar.header("🔤 ฟอนต์ไทย (ทางเลือก)")
font_file = st.sidebar.file_uploader("อัปโหลด .ttf/.otf (สำหรับข้อความไทย/ฟอนต์เฉพาะ)", type=["ttf","otf"])
font_bytes = font_file.getvalue() if font_file else None

# ========= Uploaders =========
with st.expander("วิธีใช้ (ย่อ)"):
    st.markdown("""
1) อัปโหลด **Template PDF** จาก Canva (ต้องหน้าเดียว)  
2) อัปโหลด **CSV เทอม 1** และ/หรือ **CSV เทอม 2**  
   - คอลัมน์: `No, Student ID, Name - Surname, Idea, Pronunciation, Preparedness, Confidence, Total (50)`  
   - รองรับแถว `Score` และบรรทัดว่าง (ระบบคัดทิ้งให้อัตโนมัติ)  
3) เลือกคีย์จับคู่ (ID หรือ Name) + ตั้งว่าจะใส่เทอมเดียวลง S1/S2  
4) เลือกแถวพรีวิว → ขยับ X/Y → พรีวิวสด  
5) Export ทั้งชุด → PDF รวม (1 หน้า/คน)
""")

tpl_file = st.file_uploader("อัปโหลด Template PDF (Canva / หน้าเดียว)", type=["pdf"])
csv_s1 = st.file_uploader("อัปโหลด CSV เทอม 1", type=["csv"])
csv_s2 = st.file_uploader("อัปโหลด CSV เทอม 2 (ถ้ามี)", type=["csv"])

REQUIRED_COLS = ["No","Student ID","Name - Surname","Idea","Pronunciation","Preparedness","Confidence","Total (50)"]

# ========= CSV parser (robust) =========
def _decode_csv_bytes(b: bytes) -> str:
    for enc in ("utf-8-sig","utf-8","cp874","latin-1"):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return b.decode("utf-8", errors="ignore")

def parse_csv_bytes(b: bytes) -> pd.DataFrame | None:
    if b is None: return None
    txt = _decode_csv_bytes(b)
    df_raw = pd.read_csv(io.StringIO(txt), header=None, dtype=str)
    header_idx = None
    for i in range(min(20, len(df_raw))):
        row = df_raw.iloc[i].fillna("").astype(str).tolist()
        if "No" in row and "Student ID" in row:
            header_idx = i; break
    if header_idx is None:
        df = pd.read_csv(io.StringIO(txt), dtype=str).fillna("")
    else:
        headers = df_raw.iloc[header_idx].fillna("").astype(str).tolist()
        df = df_raw.iloc[header_idx+1:].copy(); df.columns = headers; df = df.fillna("")
    cols = [c for c in df.columns if c in REQUIRED_COLS]
    df = df[cols]
    if "No" in df.columns:
        # ใช้ .str.strip() เสมอเมื่อจัดการ Series
        mask_score = df["No"].astype(str).str.strip().str.lower().eq("score")
        df = df[~mask_score]
        df = df[df["No"].astype(str).str.strip() != ""]
    df.columns = [c.strip() for c in df.columns]
    return df.reset_index(drop=True)

# ========= Merge 2 semesters =========
def coalesce(a,b):
    a = "" if pd.isna(a) else str(a)
    b = "" if pd.isna(b) else str(b)
    return a if a.strip() else b

def merge_semesters(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
                    key: str, when_single: str) -> pd.DataFrame | None:
    if df1 is not None and df2 is not None:
        merged = pd.merge(df1, df2, on=key, how="outer", suffixes=("_S1","_S2"))
        merged["Name"] = [coalesce(a,b) for a,b in zip(merged.get("Name - Surname_S1",""), merged.get("Name - Surname_S2",""))]
        merged["StudentID"] = [coalesce(a,b) for a,b in zip(merged.get("Student ID_S1",""), merged.get("Student ID_S2",""))]
        merged["Total_S1"] = merged.get("Total (50)_S1","")
        merged["Total_S2"] = merged.get("Total (50)_S2","")
        out = merged[["Name","StudentID","Total_S1","Total_S2"]].copy()
        if key == "Student ID":
            out = out.sort_values(by=["StudentID","Name"], kind="stable")
        else:
            out = out.sort_values(by=["Name","StudentID"], kind="stable")
        return out.reset_index(drop=True)
    df = df1 if df1 is not None else df2
    if df is None: return None
    out = pd.DataFrame()
    out["Name"] = df.get("Name - Surname","")
    out["StudentID"] = df.get("Student ID","")
    if when_single == "S1":
        out["Total_S1"] = df.get("Total (50)",""); out["Total_S2"] = ""
    else:
        out["Total_S1"] = ""; out["Total_S2"] = df.get("Total (50)","")
    if key == "Student ID":
        out = out.sort_values(by=["StudentID","Name"], kind="stable")
    else:
        out = out.sort_values(by=["Name","StudentID"], kind="stable")
    return out.reset_index(drop=True)

# ========= Overlay builder (ReportLab) + merge (PyPDF) =========
def build_one_page_overlay_pdf(page_w: float, page_h: float,
                               name: str, sid: str, total_s1: str, total_s2: str,
                               font_size: float, bold: bool,
                               name_xy, id_xy, s1_xy, s2_xy,
                               top_left_mode: bool,
                               font_bytes: bytes | None,
                               whiteout: bool, pad_x: float, pad_y: float) -> bytes:
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import black, white
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))

    # ฟอนต์
    fontname = "Helvetica-Bold" if bold else "Helvetica"
    if font_bytes:
        try:
            pdfmetrics.registerFont(TTFont("CustomFont", io.BytesIO(font_bytes)))
            fontname = "CustomFont"
        except Exception:
            pass
    c.setFont(fontname, float(font_size))

    def put(x, y, text):
        s = "" if text is None else str(text).strip()
        if not s:
            return
        yy = (page_h - y) if top_left_mode else y  # แปลงแกนถ้าจำเป็น

        if whiteout:
            # คำนวณความกว้างข้อความ → วาดสี่เหลี่ยมพื้นขาว + padding
            try:
                from reportlab.pdfbase.pdfmetrics import stringWidth
                tw = stringWidth(s, fontname, float(font_size))
            except Exception:
                tw = len(s) * (font_size * 0.55)  # fallback ประมาณการ
            rect_x = float(x) - float(pad_x)
            rect_y = float(yy) - float(pad_y)
            rect_w = float(tw) + float(pad_x) * 2.0
            rect_h = float(font_size) + float(pad_y) * 2.0
            c.setFillColor(white)
            c.rect(rect_x, rect_y, rect_w, rect_h, fill=1, stroke=0)
            c.setFillColor(black)

        # วาดตัวอักษร
        c.drawString(float(x), float(yy), s)

    put(name_xy[0], name_xy[1], name)
    put(id_xy[0],   id_xy[1],   sid)
    put(s1_xy[0],   s1_xy[1],   total_s1)
    put(s2_xy[0],   s2_xy[1],   total_s2)

    c.showPage(); c.save()
    return buf.getvalue()

# def merge_overlay_on_template(template_pdf_bytes: bytes, overlay_pdf_bytes_list: list[bytes]) -> bytes:
#     from pypdf import PdfReader, PdfWriter
#     writer = PdfWriter()
#     for ov_bytes in overlay_pdf_bytes_list:
#         tpl_reader = PdfReader(io.BytesIO(template_pdf_bytes))
#         base_page = tpl_reader.pages[0]
#         ov_reader = PdfReader(io.BytesIO(ov_bytes))
#         ov_page = ov_reader.pages[0]
#         base_page.merge_page(ov_page)  # วางทับด้วยพื้นขาวรองแล้วค่อยตัวหนังสือ
#         writer.add_page(base_page)
#     out = io.BytesIO()
#     writer.write(out); out.seek(0)
#     return out.getvalue()

# def make_preview_pdf(template_pdf_bytes: bytes, rec: dict,
#                      font_size: float, bold: bool,
#                      name_xy, id_xy, s1_xy, s2_xy,
#                      top_left_mode: bool,
#                      font_bytes: bytes | None,
#                      whiteout: bool, pad_x: float, pad_y: float) -> bytes:
#     from pypdf import PdfReader
#     reader = PdfReader(io.BytesIO(template_pdf_bytes))
#     pg = reader.pages[0]
#     page_w = float(pg.mediabox.width); page_h = float(pg.mediabox.height)
#     overlay = build_one_page_overlay_pdf(
#         page_w, page_h,
#         rec.get("Name",""), rec.get("StudentID",""),
#         rec.get("Total_S1",""), rec.get("Total_S2",""),
#         font_size, bold,
#         name_xy, id_xy, s1_xy, s2_xy,
#         top_left_mode,
#         font_bytes,
#         whiteout, pad_x, pad_y
#     )
#     return merge_overlay_on_template(template_pdf_bytes, [overlay])

def make_full_pdf(template_pdf_bytes: bytes, records: pd.DataFrame,
                  font_size: float, bold: bool,
                  name_xy, id_xy, s1_xy, s2_xy,
                  top_left_mode: bool,
                  font_bytes: bytes | None,
                  whiteout: bool, pad_x: float, pad_y: float) -> bytes:
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(template_pdf_bytes))
    pg = reader.pages[0]
    page_w = float(pg.mediabox.width); page_h = float(pg.mediabox.height)
    overlays = []
    for _, r in records.iterrows():
        overlays.append(build_one_page_overlay_pdf(
            page_w, page_h,
            r.get("Name",""), r.get("StudentID",""),
            r.get("Total_S1",""), r.get("Total_S2",""),
            font_size, bold,
            name_xy, id_xy, s1_xy, s2_xy,
            top_left_mode,
            font_bytes,
            whiteout, pad_x, pad_y
        ))
    return merge_overlay_on_template(template_pdf_bytes, overlays)

# ========= Preview renderers =========
def render_preview_as_image(pdf_bytes: bytes, zoom_dpi: int = 160):
    """ใช้ PyMuPDF (ถ้ามี) เรนเดอร์หน้าแรกเป็น PNG แล้วโชว์ → ไม่โดนบล็อก iframe"""
    spec = importlib.util.find_spec("pymupdf")
    if spec is None:
        return False
    import pymupdf as fitz
    doc = fitz.open("pdf", pdf_bytes)
    page = doc[0]
    pix = page.get_pixmap(dpi=zoom_dpi)
    st.image(pix.tobytes("png"), caption=f"Preview @ {zoom_dpi} dpi", use_container_width=True)
    return True

def render_preview_as_pdf(pdf_bytes: bytes, height: int = 820):
    """ถ้าไม่มี PyMuPDF → ใช้ st.pdf (Streamlit ใหม่) หรือ fallback data:URL"""
    if hasattr(st, "pdf"):
        st.pdf(pdf_bytes, height=height)
        return
    b64 = base64.b64encode(pdf_bytes).decode("utf-8")
    html = f'''
    <iframe src="data:application/pdf;base64,{b64}"
            width="100%" height="{height}" style="border:1px solid #444; border-radius:8px;">
    </iframe>
    '''
    st.components.v1.html(html, height=height+20)

# ========= Main flow =========
if tpl_file is None:
    st.info("อัปโหลด **Template PDF (หน้าเดียว)** ก่อน"); st.stop()

df1 = parse_csv_bytes(csv_s1.getvalue()) if csv_s1 is not None else None
df2 = parse_csv_bytes(csv_s2.getvalue()) if csv_s2 is not None else None
if (df1 is None) and (df2 is None):
    st.warning("ยังไม่พบ CSV คะแนน กรุณาอัปโหลดอย่างน้อย 1 เทอม"); st.stop()

c1, c2 = st.columns(2)
if df1 is not None:
    with c1:
        st.subheader("CSV เทอม 1 (ตัวอย่าง 10 แถว)")
        st.dataframe(df1.head(10), use_container_width=True)
if df2 is not None:
    with c2:
        st.subheader("CSV เทอม 2 (ตัวอย่าง 10 แถว)")
        st.dataframe(df2.head(10), use_container_width=True)

key = "Student ID" if join_key == "Student ID" else "Name - Surname"
merged = merge_semesters(df1, df2, key, when_single)
if merged is None or merged.empty:
    st.error("รวมข้อมูลไม่สำเร็จ (เช็กคอลัมน์/คีย์จับคู่ใน CSV)"); st.stop()

st.success(f"พร้อมแปลง {len(merged)} คน (1 หน้า/คน)")

left, right = st.columns([1,2])
with left:
    st.subheader("เลือกคนสำหรับพรีวิวสด")
    options = []
    for _, r in merged.iterrows():
        sid = str(r.get("StudentID","")).strip()
        nm  = str(r.get("Name","")).strip()
        if sid and nm: options.append(f"{sid} — {nm}")
        elif nm:       options.append(nm)
        else:          options.append(sid or "—")
    idx = st.slider("แถวพรีวิว", 0, len(merged)-1, 0, 1)
    st.text(f"กำลังพรีวิว: {options[idx]}")

with right:
    st.subheader("พรีวิวแบบ Real-time (ขยับ X/Y แล้วอัปเดตทันที)")
    tpl_bytes = tpl_file.getvalue()
    rec = merged.iloc[int(idx)].to_dict()
    try:
        preview_pdf = make_preview_pdf(
            tpl_bytes, rec,
            font_size, bold,
            (name_x, name_y), (id_x, id_y), (s1_x, s1_y), (s2_x, s2_y),
            top_left_mode,
            font_bytes,
            whiteout, pad_x, pad_y
        )
        shown = render_preview_as_image(preview_pdf, zoom_dpi=st.sidebar.slider("Preview DPI", 120, 220, 160, 10))
        if not shown:
            render_preview_as_pdf(preview_pdf, height=820)
        st.download_button("⬇️ ดาวน์โหลดพรีวิว (PDF 1 หน้า)", preview_pdf, file_name="preview_1page.pdf")
    except Exception as e:
        st.error(f"เรนเดอร์พรีวิวล้มเหลว: {e}")

st.markdown("---")
st.subheader("📦 ส่งออก PDF ทั้งชุด (1 หน้า/คน)")
if st.button("Export ทั้งชุด", type="primary"):
    with st.spinner("กำลังเรนเดอร์…"):
        try:
            full_pdf = make_full_pdf(
                tpl_bytes, merged,
                font_size, bold,
                (name_x, name_y), (id_x, id_y), (s1_x, s1_y), (s2_x, s2_y),
                top_left_mode,
                font_bytes,
                whiteout, pad_x, pad_y
            )
            st.success("สำเร็จ! ดาวน์โหลดได้เลย")
            st.download_button("⬇️ ดาวน์โหลดไฟล์รวม (PDF)", full_pdf, file_name="Conversation_PerStudent_Output.pdf")
        except Exception as e:
            st.error(f"ส่งออกล้มเหลว: {e}")
