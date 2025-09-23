# -*- coding: utf-8 -*-
# Streamlit app: พรีวิว PDF แบบ Real-time ปรับ X/Y แล้วเห็นผลทันที
# โจทย์: 1 หน้า/นักเรียน 1 คน (จาก CSV 1 หรือ 2 เทอม), เติม Total(50) ลงช่องที่เว้นในเทมเพลต Canva

import io
import base64
import streamlit as st
import pandas as pd

# ===== UI HEAD =====
st.set_page_config(page_title="Conversation → PDF (Realtime Preview)", layout="wide")
st.title("📄 Conversation Result → PDF (Realtime X/Y Preview)")
st.caption("อัปโหลดเทมเพลต PDF (Canva หน้าเดียว) + CSV เทอม 1/2 → พิมพ์ค่า Total ลงช่องที่เว้นไว้ • พรีวิวสดเมื่อปรับ X/Y")

# ===== Sidebar: Controls =====
st.sidebar.header("🔧 ตำแหน่งข้อความ (หน่วย pt)")
# ค่าเริ่มต้นกะคร่าว ๆ — จูนจากพรีวิวให้ตรงแผ่นจริง
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
st.sidebar.caption("ปกติ PDF มีจุดกำเนิดที่มุมล่างซ้าย (y เพิ่มขึ้น = ขึ้นด้านบน). ถ้าติ๊กตัวนี้ y จะวัดจากด้านบนลงล่างแทน")

st.sidebar.header("🔗 การแม็ปข้อมูล")
join_key = st.sidebar.selectbox("คีย์จับคู่ 2 เทอม", ["Student ID", "Name - Surname"], index=0)
when_single = st.sidebar.selectbox("ถ้าอัปโหลด CSV เทอมเดียว ให้ใส่ลงช่อง", ["S1", "S2"], index=0)

st.sidebar.header("🔤 ฟอนต์ไทย (ทางเลือก)")
font_file = st.sidebar.file_uploader("อัปโหลด .ttf/.otf (ถ้าจะพิมพ์ไทยให้สวย)", type=["ttf","otf"])

# ===== Uploaders =====
with st.expander("วิธีใช้ (ย่อ)"):
    st.markdown("""
1) อัปโหลด **Template PDF** จาก Canva (ต้องหน้าเดียว)  
2) อัปโหลด **CSV เทอม 1** และ/หรือ **CSV เทอม 2**  
   - คอลัมน์: `No, Student ID, Name - Surname, Idea, Pronunciation, Preparedness, Confidence, Total (50)`  
   - รับแถว `Score` และบรรทัดว่าง (ระบบคัดทิ้งให้อัตโนมัติ)  
3) เลือกคีย์จับคู่ (ID หรือ Name) และกำหนดว่าจะใส่แถวเทอมเดียวลง S1/S2  
4) เลือกนักเรียนจากรายการ → ปรับ X/Y → พรีวิว PDF แบบ Real-time  
5) กด Export ทั้งชุด → ได้ไฟล์ PDF รวม (1 หน้า/คน)
""")

tpl_file = st.file_uploader("อัปโหลด Template PDF (Canva / หน้าเดียว)", type=["pdf"])
csv_s1 = st.file_uploader("อัปโหลด CSV เทอม 1", type=["csv"])
csv_s2 = st.file_uploader("อัปโหลด CSV เทอม 2 (ถ้ามี)", type=["csv"])

REQUIRED_COLS = ["No","Student ID","Name - Surname","Idea","Pronunciation","Preparedness","Confidence","Total (50)"]

# ===== CSV Parser (Robust) =====
def _decode_csv_bytes(b: bytes) -> str:
    for enc in ("utf-8-sig","utf-8","cp874","latin-1"):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return b.decode("utf-8", errors="ignore")

def parse_csv_bytes(b: bytes) -> pd.DataFrame | None:
    if b is None:
        return None
    txt = _decode_csv_bytes(b)
    df_raw = pd.read_csv(io.StringIO(txt), header=None, dtype=str)
    header_idx = None
    for i in range(min(20, len(df_raw))):
        row = df_raw.iloc[i].fillna("").astype(str).tolist()
        if "No" in row and "Student ID" in row:
            header_idx = i
            break
    if header_idx is None:
        df = pd.read_csv(io.StringIO(txt), dtype=str).fillna("")
    else:
        headers = df_raw.iloc[header_idx].fillna("").astype(str).tolist()
        df = df_raw.iloc[header_idx+1:].copy()
        df.columns = headers
        df = df.fillna("")
    cols = [c for c in df.columns if c in REQUIRED_COLS]
    df = df[cols]
    if "No" in df.columns:
        mask_score = df["No"].astype(str).str.strip().str.lower().eq("score")
        df = df[~mask_score]
        df = df[df["No"].astype(str).str.strip() != ""]
    df.columns = [c.strip() for c in df.columns]
    return df.reset_index(drop=True)

# ===== Merge 2 semesters =====
def coalesce(a,b):
    a = "" if pd.isna(a) else str(a)
    b = "" if pd.isna(b) else str(b)
    return a if a.strip() else b

def merge_semesters(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
                    key: str, when_single: str) -> pd.DataFrame | None:
    if df1 is not None and df2 is not None:
        merged = pd.merge(df1, df2, on=key, how="outer", suffixes=("_S1","_S2"))
        merged["Name"] = [coalesce(a,b) for a,b in zip(merged.get("Name - Surname_S1",""),
                                                       merged.get("Name - Surname_S2",""))]
        merged["StudentID"] = [coalesce(a,b) for a,b in zip(merged.get("Student ID_S1",""),
                                                            merged.get("Student ID_S2",""))]
        merged["Total_S1"] = merged.get("Total (50)_S1","")
        merged["Total_S2"] = merged.get("Total (50)_S2","")
        out = merged[["Name","StudentID","Total_S1","Total_S2"]].copy()
        if key == "Student ID":
            out = out.sort_values(by=["StudentID","Name"], kind="stable")
        else:
            out = out.sort_values(by=["Name","StudentID"], kind="stable")
        return out.reset_index(drop=True)

    df = df1 if df1 is not None else df2
    if df is None:
        return None
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

# ===== PDF overlay backend (ReportLab + PyPDF) =====
def build_one_page_overlay_pdf(page_w: float, page_h: float,
                               name: str, sid: str, total_s1: str, total_s2: str,
                               font_size: float, bold: bool,
                               name_xy, id_xy, s1_xy, s2_xy,
                               top_left_mode: bool,
                               font_file_bytes: bytes | None) -> bytes:
    """
    สร้าง PDF 1 หน้า ที่มีแค่ตัวหนังสือ (overlay) ขนาดเท่ากับหน้าเทมเพลต
    """
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import black
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    c.setFillColor(black)

    # ฟอนต์
    fontname = "Helvetica-Bold" if bold else "Helvetica"
    if font_file_bytes:
        try:
            # ลงทะเบียนฟอนต์อัปโหลดเป็นชื่อ 'CustomFont' / 'CustomFont-Bold'
            # (ReportLab ไม่รู้จัก Bold ออโต้ของฟอนต์ไทยเสมอไป — ใช้ไฟล์เดียวก็ตั้งเป็นชื่อเดียวพอ)
            tmp = io.BytesIO(font_file_bytes)
            pdfmetrics.registerFont(TTFont("CustomFont", tmp))
            fontname = "CustomFont"
        except Exception:
            pass

    c.setFont(fontname, float(font_size))

    def put(x, y, text):
        s = "" if text is None else str(text).strip()
        if not s:
            return
        yy = page_h - y if top_left_mode else y   # ถ้า top-left mode, แปลง y: บนลงล่าง → bottom-left
        c.drawString(float(x), float(yy), s)

    x, y = name_xy;   put(x, y, name)
    x, y = id_xy;     put(x, y, sid)
    x, y = s1_xy;     put(x, y, total_s1)
    x, y = s2_xy;     put(x, y, total_s2)

    c.showPage()
    c.save()
    return buf.getvalue()

def merge_overlay_on_template(template_pdf_bytes: bytes, overlay_pdf_bytes_list: list[bytes]) -> bytes:
    """
    นำ overlay แต่ละหน้ามาวางทับบนหน้าแรกของ template ซ้ำ ๆ (1 หน้า/คน)
    """
    from pypdf import PdfReader, PdfWriter

    writer = PdfWriter()
    for overlay_bytes in overlay_pdf_bytes_list:
        # อ่าน template และ overlay ใหม่ทุกครั้งเพื่อเลี่ยงการแก้ไข obj ร่วม
        tpl_reader = PdfReader(io.BytesIO(template_pdf_bytes))
        base_page = tpl_reader.pages[0]

        ov_reader = PdfReader(io.BytesIO(overlay_bytes))
        ov_page = ov_reader.pages[0]
        base_page.merge_page(ov_page)  # วางทับ

        writer.add_page(base_page)

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.getvalue()

def make_preview_pdf(template_pdf_bytes: bytes, rec: dict,
                     font_size: float, bold: bool,
                     name_xy, id_xy, s1_xy, s2_xy,
                     top_left_mode: bool,
                     font_file_bytes: bytes | None) -> bytes:
    # อ่านขนาดหน้าเทมเพลต
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(template_pdf_bytes))
    pg = reader.pages[0]
    page_w = float(pg.mediabox.width)
    page_h = float(pg.mediabox.height)

    overlay = build_one_page_overlay_pdf(
        page_w, page_h,
        rec.get("Name",""), rec.get("StudentID",""),
        rec.get("Total_S1",""), rec.get("Total_S2",""),
        font_size, bold,
        name_xy, id_xy, s1_xy, s2_xy,
        top_left_mode,
        font_file.read() if font_file else None
    )
    return merge_overlay_on_template(template_pdf_bytes, [overlay])

def make_full_pdf(template_pdf_bytes: bytes, records: pd.DataFrame,
                  font_size: float, bold: bool,
                  name_xy, id_xy, s1_xy, s2_xy,
                  top_left_mode: bool,
                  font_file_bytes: bytes | None) -> bytes:
    # อ่านขนาดหน้าเทมเพลต
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(template_pdf_bytes))
    pg = reader.pages[0]
    page_w = float(pg.mediabox.width)
    page_h = float(pg.mediabox.height)

    overlays = []
    for _, r in records.iterrows():
        overlays.append(build_one_page_overlay_pdf(
            page_w, page_h,
            r.get("Name",""), r.get("StudentID",""),
            r.get("Total_S1",""), r.get("Total_S2",""),
            font_size, bold,
            name_xy, id_xy, s1_xy, s2_xy,
            top_left_mode,
            font_file_bytes
        ))
    return merge_overlay_on_template(template_pdf_bytes, overlays)

def embed_pdf(pdf_bytes: bytes, height: int = 820):
    b64 = base64.b64encode(pdf_bytes).decode("utf-8")
    html = f'''
    <iframe
        src="data:application/pdf;base64,{b64}"
        width="100%" height="{height}" style="border:1px solid #444; border-radius:8px;">
    </iframe>
    '''
    st.components.v1.html(html, height=height+20)

# ===== Main Logic =====
if tpl_file is None:
    st.info("อัปโหลด **Template PDF (หน้าเดียว)** ก่อน")
    st.stop()

# อ่าน CSVs
df1 = parse_csv_bytes(csv_s1.getvalue()) if csv_s1 is not None else None
df2 = parse_csv_bytes(csv_s2.getvalue()) if csv_s2 is not None else None

if (df1 is None) and (df2 is None):
    st.warning("ยังไม่พบ CSV คะแนน กรุณาอัปโหลดอย่างน้อย 1 เทอม")
    st.stop()

# แสดงตัวอย่างข้อมูล
c1, c2 = st.columns(2)
if df1 is not None:
    with c1:
        st.subheader("CSV เทอม 1 (ตัวอย่าง 10 แถว)")
        st.dataframe(df1.head(10), use_container_width=True)
if df2 is not None:
    with c2:
        st.subheader("CSV เทอม 2 (ตัวอย่าง 10 แถว)")
        st.dataframe(df2.head(10), use_container_width=True)

# รวมข้อมูล
key = "Student ID" if join_key == "Student ID" else "Name - Surname"
merged = merge_semesters(df1, df2, key, when_single)
if merged is None or merged.empty:
    st.error("รวมข้อมูลไม่สำเร็จ (เช็กคอลัมน์/คีย์จับคู่ใน CSV)")
    st.stop()

st.success(f"พร้อมแปลง {len(merged)} คน (1 หน้า/คน)")

# ===== Selector for Realtime Preview =====
left, right = st.columns([1,2])

with left:
    st.subheader("เลือกคนที่จะแสดงพรีวิวสด")
    # ถ้ามีคอลัมน์ชื่อ/ID ให้โชว์รายการเลือก
    display_options = []
    for _, r in merged.iterrows():
        label = f'{r.get("StudentID","")}'.strip()
        name = f'{r.get("Name","")}'.strip()
        if label and name:
            display_options.append(f'{label} — {name}')
        elif name:
            display_options.append(name)
        else:
            display_options.append(label or "—")

    idx = st.number_input("ลำดับ (index)", min_value=0, max_value=len(merged)-1, value=0, step=1)
    st.caption("หรือเลือกด้วยสไลเดอร์ด้านล่าง")
    idx = st.slider("แถวที่ต้องการพรีวิว", 0, len(merged)-1, idx, 1)
    st.text(f"กำลังพรีวิว: {display_options[idx]}")

with right:
    st.subheader("พรีวิว PDF (อัปเดตทันทีเมื่อปรับ X/Y)")
    # สร้างพรีวิวทีละหน้า (คนเดียว) โดยใช้ค่าปัจจุบันทั้งหมด
    tpl_bytes = tpl_file.getvalue()
    rec = merged.iloc[int(idx)].to_dict()
    try:
        preview_bytes = make_preview_pdf(
            tpl_bytes, rec,
            font_size, bold,
            (name_x, name_y), (id_x, id_y), (s1_x, s1_y), (s2_x, s2_y),
            top_left_mode,
            font_file.read() if font_file else None
        )
        embed_pdf(preview_bytes, height=820)
        st.download_button("⬇️ ดาวน์โหลดพรีวิว (PDF 1 หน้า)", preview_bytes, file_name="preview_1page.pdf")
    except Exception as e:
        st.error(f"เรนเดอร์พรีวิวล้มเหลว: {e}")

st.markdown("---")
st.subheader("📦 ส่งออก PDF ทั้งชุด (1 หน้า/คน)")
export_btn = st.button("Export ทั้งชุด", type="primary")
if export_btn:
    with st.spinner("กำลังเรนเดอร์ PDF ทั้งชุด…"):
        try:
            full_bytes = make_full_pdf(
                tpl_bytes, merged,
                font_size, bold,
                (name_x, name_y), (id_x, id_y), (s1_x, s1_y), (s2_x, s2_y),
                top_left_mode,
                font_file.read() if font_file else None
            )
            st.success("สำเร็จ! ดาวน์โหลดได้เลยด้านล่าง")
            st.download_button("⬇️ ดาวน์โหลดไฟล์รวม (PDF)", full_bytes, file_name="Conversation_PerStudent_Output.pdf")
        except Exception as e:
            st.error(f"ส่งออกล้มเหลว: {e}")
