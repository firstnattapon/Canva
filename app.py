# -*- coding: utf-8 -*-
# Streamlit app: พรีวิว PDF แบบ Real-time + "ทุกตัวมี X/Y ของตัวเอง"
# ✅ 1 หน้า/นักเรียน 1 คน, รองรับ CSV เทอมเดียวหรือสองเทอม
# ✅ วางได้หลายฟิลด์แบบยืดหยุ่น: แต่ละฟิลด์กำหนด x, y, ขนาด, หนา/ไม่หนา, ชิดซ้าย/กลาง/ขวา ได้เอง
# ✅ บันทึก/โหลดเลย์เอาต์ (.json)
# ✅ ใส่ Total(50) ลงเทมเพลต Canva ได้ และจะวางค่าอื่น ๆ ได้ตามที่ตั้งใน Layout Editor

import io
import json
import base64
import importlib.util
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Conversation → PDF (Realtime Preview)", layout="wide")
st.title("📄 Conversation Result → PDF (Realtime X/Y Preview)")
st.caption("อัปโหลดเทมเพลต PDF (Canva หน้าเดียว) + CSV เทอม 1/2 → เลือกฟิลด์ที่จะวาง แล้วกำหนด X/Y เป็นรายฟิลด์ • พรีวิวสดเมื่อปรับค่า")

# ========= Sidebar: Global controls =========
st.sidebar.header("🧭 ระบบแกน & ฟอนต์เริ่มต้น")
top_left_mode = st.sidebar.checkbox("ใช้ Y แบบจากด้านบนลงล่าง (Top-Left Origin)", value=False)
st.sidebar.caption("ปกติ PDF ใช้มุมล่างซ้ายเป็นจุดกำเนิด (y เพิ่มขึ้น = ขึ้นด้านบน). ถ้าติ๊กตัวนี้ y จะวัดจากด้านบนลงล่างแทน")

def_font_size = st.sidebar.number_input("ขนาดตัวอักษรเริ่มต้น", 6, 72, 16, step=1)
def_bold = st.sidebar.checkbox("ตัวหนาเริ่มต้น (Helvetica-Bold)", value=True)

st.sidebar.header("🔤 ฟอนต์ไทย (ตัวเลือก)")
font_file = st.sidebar.file_uploader("อัปโหลด .ttf/.otf (สำหรับข้อความไทย/ฟอนต์เฉพาะ)", type=["ttf","otf"])
font_bytes = font_file.getvalue() if font_file else None

# ========= Uploaders =========
with st.expander("วิธีใช้ (ย่อ)"):
    st.markdown(
        """
1) อัปโหลด **Template PDF** จาก Canva (ต้องหน้าเดียว)  
2) อัปโหลด **CSV เทอม 1** และ/หรือ **CSV เทอม 2**  
   - คอลัมน์: `No, Student ID, Name - Surname, Idea, Pronunciation, Preparedness, Confidence, Total (50)`  
   - รองรับแถว `Score` และบรรทัดว่าง (ระบบคัดทิ้งให้อัตโนมัติ)  
3) เลือกคีย์จับคู่ (ID หรือ Name) + ตั้งว่าจะใส่เทอมเดียวลง S1/S2  
4) **Layout Editor**: เพิ่มแถวฟิลด์ที่อยากวาง → ตั้ง `source, x, y, font_size, bold, align` ทีละตัว  
5) เลือกแถวพรีวิว → ขยับค่าใน Layout Editor → พรีวิวสด  
6) Export ทั้งชุด → PDF รวม (1 หน้า/คน)  
        """
    )

st.sidebar.header("🔗 การแม็ปข้อมูล (รวมเทอม)")
join_key = st.sidebar.selectbox("คีย์จับคู่ 2 เทอม", ["Student ID", "Name - Surname"], index=0)
when_single = st.sidebar.selectbox("ถ้าอัปโหลด CSV เทอมเดียว ให้ใส่ลงช่อง", ["S1", "S2"], index=0)

# ========= File inputs =========
tpl_file = st.file_uploader("อัปโหลด Template PDF (Canva / หน้าเดียว)", type=["pdf"])
csv_s1 = st.file_uploader("อัปโหลด CSV เทอม 1", type=["csv"])
csv_s2 = st.file_uploader("อัปโหลด CSV เทอม 2 (ถ้ามี)", type=["csv"])

REQUIRED_COLS = [
    "No","Student ID","Name - Surname","Idea","Pronunciation","Preparedness","Confidence","Total (50)"
]

# ========= CSV parser (robust) =========
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
        # ตัดแถวสรุปคะแนน/ว่าง
        mask_score = df["No"].astype(str).str.strip().str.lower().eq("score")
        df = df[~mask_score]
        df = df[df["No"].astype(str).str.strip() != ""]
    df.columns = [c.strip() for c in df.columns]
    return df.reset_index(drop=True)

# ========= Merge 2 semesters =========
def coalesce(a, b):
    a = "" if pd.isna(a) else str(a)
    b = "" if pd.isna(b) else str(b)
    return a if a.strip() else b

def merge_semesters(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
                    key: str, when_single: str) -> pd.DataFrame | None:
    if df1 is not None and df2 is not None:
        merged = pd.merge(df1, df2, on=key, how="outer", suffixes=("_S1","_S2"))
        merged["Name"] = [coalesce(a,b) for a,b in zip(merged.get("Name - Surname_S1",""), merged.get("Name - Surname_S2",""))]
        merged["StudentID"] = [coalesce(a,b) for a,b in zip(merged.get("Student ID_S1",""), merged.get("Student ID_S2",""))]
        merged["Idea_S1"] = merged.get("Idea_S1", "");         merged["Idea_S2"] = merged.get("Idea_S2", "")
        merged["Pronunciation_S1"] = merged.get("Pronunciation_S1", ""); merged["Pronunciation_S2"] = merged.get("Pronunciation_S2", "")
        merged["Preparedness_S1"] = merged.get("Preparedness_S1", ""); merged["Preparedness_S2"] = merged.get("Preparedness_S2", "")
        merged["Confidence_S1"] = merged.get("Confidence_S1", "");   merged["Confidence_S2"] = merged.get("Confidence_S2", "")
        merged["Total_S1"] = merged.get("Total (50)_S1","")
        merged["Total_S2"] = merged.get("Total (50)_S2","")
        out = merged[[
            "Name","StudentID",
            "Idea_S1","Pronunciation_S1","Preparedness_S1","Confidence_S1","Total_S1",
            "Idea_S2","Pronunciation_S2","Preparedness_S2","Confidence_S2","Total_S2",
        ]].copy()
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
    # เติมช่อง S1/S2 ตาม when_single
    if when_single == "S1":
        out["Idea_S1"] = df.get("Idea", ""); out["Idea_S2"] = ""
        out["Pronunciation_S1"] = df.get("Pronunciation", ""); out["Pronunciation_S2"] = ""
        out["Preparedness_S1"] = df.get("Preparedness", ""); out["Preparedness_S2"] = ""
        out["Confidence_S1"] = df.get("Confidence", ""); out["Confidence_S2"] = ""
        out["Total_S1"] = df.get("Total (50)", ""); out["Total_S2"] = ""
    else:
        out["Idea_S1"] = ""; out["Idea_S2"] = df.get("Idea", "")
        out["Pronunciation_S1"] = ""; out["Pronunciation_S2"] = df.get("Pronunciation", "")
        out["Preparedness_S1"] = ""; out["Preparedness_S2"] = df.get("Preparedness", "")
        out["Confidence_S1"] = ""; out["Confidence_S2"] = df.get("Confidence", "")
        out["Total_S1"] = ""; out["Total_S2"] = df.get("Total (50)", "")

    if key == "Student ID":
        out = out.sort_values(by=["StudentID","Name"], kind="stable")
    else:
        out = out.sort_values(by=["Name","StudentID"], kind="stable")
    return out.reset_index(drop=True)

# ========= PDF overlay primitives =========
def build_one_page_overlay_pdf(page_w: float, page_h: float,
                               record: dict,
                               layout_rows: list[dict],
                               def_font_size: float, def_bold: bool,
                               top_left_mode: bool,
                               font_bytes: bytes | None) -> bytes:
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import black
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    c.setFillColor(black)

    # font setup
    base_regular = "Helvetica"
    base_bold = "Helvetica-Bold"
    custom_name = None
    if font_bytes:
        try:
            pdfmetrics.registerFont(TTFont("CustomFont", io.BytesIO(font_bytes)))
            custom_name = "CustomFont"
        except Exception:
            custom_name = None

    def draw_text(x: float, y: float, text: str, align: str, fs: float, bold: bool):
        if text is None:
            return
        s = str(text).strip()
        if not s:
            return
        yy = (page_h - y) if top_left_mode else y
        # choose font
        fontname = custom_name or (base_bold if bold else base_regular)
        c.setFont(fontname, float(fs))
        # align draw
        a = (align or "left").lower()
        if a.startswith("cen"):
            c.drawCentredString(float(x), float(yy), s)
        elif a.startswith("right") or a == "r":
            c.drawRightString(float(x), float(yy), s)
        else:
            c.drawString(float(x), float(yy), s)

    for row in layout_rows:
        src = (row.get("source") or "").strip()
        x = float(row.get("x") or 0)
        y = float(row.get("y") or 0)
        fs = float(row.get("font_size") or def_font_size)
        bd = bool(row.get("bold")) if row.get("bold") is not None else def_bold
        align = (row.get("align") or "left").lower()
        val = record.get(src, "") if src else ""
        draw_text(x, y, val, align, fs, bd)

    c.showPage(); c.save()
    return buf.getvalue()


def merge_overlay_on_template(template_pdf_bytes: bytes, overlay_pdf_bytes_list: list[bytes]) -> bytes:
    from pypdf import PdfReader, PdfWriter
    writer = PdfWriter()
    for ov_bytes in overlay_pdf_bytes_list:
        tpl_reader = PdfReader(io.BytesIO(template_pdf_bytes))
        base_page = tpl_reader.pages[0]
        ov_reader = PdfReader(io.BytesIO(ov_bytes))
        ov_page = ov_reader.pages[0]
        base_page.merge_page(ov_page)
        writer.add_page(base_page)
    out = io.BytesIO()
    writer.write(out); out.seek(0)
    return out.getvalue()


def make_preview_pdf(template_pdf_bytes: bytes, record: dict,
                     layout_rows: list[dict],
                     def_font_size: float, def_bold: bool,
                     top_left_mode: bool,
                     font_bytes: bytes | None) -> bytes:
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(template_pdf_bytes))
    pg = reader.pages[0]
    page_w = float(pg.mediabox.width); page_h = float(pg.mediabox.height)
    overlay = build_one_page_overlay_pdf(
        page_w, page_h, record, layout_rows, def_font_size, def_bold, top_left_mode, font_bytes
    )
    return merge_overlay_on_template(template_pdf_bytes, [overlay])


def make_full_pdf(template_pdf_bytes: bytes, records: pd.DataFrame,
                  layout_rows: list[dict],
                  def_font_size: float, def_bold: bool,
                  top_left_mode: bool,
                  font_bytes: bytes | None) -> bytes:
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(template_pdf_bytes))
    pg = reader.pages[0]
    page_w = float(pg.mediabox.width); page_h = float(pg.mediabox.height)
    overlays = []
    for _, r in records.iterrows():
        overlays.append(build_one_page_overlay_pdf(
            page_w, page_h, r.to_dict(), layout_rows, def_font_size, def_bold, top_left_mode, font_bytes
        ))
    return merge_overlay_on_template(template_pdf_bytes, overlays)

# ========= Preview renderers =========
def render_preview_as_image(pdf_bytes: bytes, zoom_dpi: int = 150):
    """ใช้ PyMuPDF (ถ้ามี) เรนเดอร์หน้าแรกเป็น PNG แล้วโชว์"""
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
    """ใช้ st.pdf ถ้ามี (Streamlit ใหม่), ถ้าไม่มีก็ fallback data: URL iframe"""
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

# Read CSVs
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

# Merge
key = "Student ID" if join_key == "Student ID" else "Name - Surname"
merged = merge_semesters(df1, df2, key, when_single)
if merged is None or merged.empty:
    st.error("รวมข้อมูลไม่สำเร็จ (เช็กคอลัมน์/คีย์จับคู่ใน CSV)"); st.stop()

st.success(f"พร้อมแปลง {len(merged)} คน (1 หน้า/คน)")

# ========= Layout Editor =========
st.markdown("---")
st.subheader("🧩 Layout Editor — ตั้งค่า X/Y ต่อฟิลด์ (ทุกตัวมี X/Y ของตัวเอง)")

# Default candidate sources
sources = [
    "Name","StudentID",
    "Idea_S1","Pronunciation_S1","Preparedness_S1","Confidence_S1","Total_S1",
    "Idea_S2","Pronunciation_S2","Preparedness_S2","Confidence_S2","Total_S2",
]

# Seed default layout rows once
DEFAULT_LAYOUT = [
    {"field_label":"Name","source":"Name","x":140,"y":160,"font_size":def_font_size,"bold":True,"align":"left"},
    {"field_label":"StudentID","source":"StudentID","x":140,"y":190,"font_size":def_font_size-1,"bold":False,"align":"left"},
    {"field_label":"Total_S1","source":"Total_S1","x":430,"y":300,"font_size":def_font_size,"bold":True,"align":"center"},
    {"field_label":"Total_S2","source":"Total_S2","x":430,"y":340,"font_size":def_font_size,"bold":True,"align":"center"},
]

if "layout_df" not in st.session_state:
    st.session_state.layout_df = pd.DataFrame(DEFAULT_LAYOUT)

# Import layout JSON
imp_col, exp_col = st.columns([1,1])
with imp_col:
    imp = st.file_uploader("โหลด Layout (.json)", type=["json"], key="impjson")
    if imp is not None:
        try:
            data = json.loads(imp.getvalue().decode("utf-8"))
            st.session_state.layout_df = pd.DataFrame(data)
            st.success("โหลด Layout สำเร็จ")
        except Exception as e:
            st.error(f"โหลด Layout ไม่สำเร็จ: {e}")
with exp_col:
    if st.download_button(
        "💾 บันทึก Layout (.json)",
        data=json.dumps(st.session_state.layout_df.to_dict(orient="records"), ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="layout.json",
        mime="application/json",
    ):
        pass

# Editor — allow add/delete rows
edited = st.data_editor(
    st.session_state.layout_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "field_label": st.column_config.TextColumn("Label", help="ชื่อเรียกเพื่อจำได้"),
        "source": st.column_config.SelectboxColumn("source", options=sources, help="คอลัมน์ข้อมูลที่จะวาง"),
        "x": st.column_config.NumberColumn("x", min_value=0, max_value=2000, step=1),
        "y": st.column_config.NumberColumn("y", min_value=0, max_value=2000, step=1),
        "font_size": st.column_config.NumberColumn("font_size", min_value=6, max_value=72, step=1, help="เว้นว่างจะใช้ค่ากลาง"),
        "bold": st.column_config.CheckboxColumn("bold"),
        "align": st.column_config.SelectboxColumn("align", options=["left","center","right"], help="ตำแหน่งอิง x")
    },
    hide_index=True,
)

st.session_state.layout_df = edited.copy()

# ========= Preview =========
left, right = st.columns([1,2])
with left:
    st.subheader("เลือกคนสำหรับพรีวิวสด")
    options = []
    for _, r in merged.iterrows():
        sid = str(r.get("StudentID",""))
        nm  = str(r.get("Name",""))
        if sid and nm:
            options.append(f"{sid} — {nm}")
        elif nm:
            options.append(nm)
        else:
            options.append(sid or "—")
    idx = st.slider("แถวพรีวิว", 0, len(merged)-1, 0, 1)
    st.text(f"กำลังพรีวิว: {options[idx]}")

with right:
    st.subheader("พรีวิวแบบ Real-time (ขยับ Layout แล้วอัปเดตทันที)")
    tpl_bytes = tpl_file.getvalue()
    rec = merged.iloc[int(idx)].to_dict()
    try:
        layout_rows = st.session_state.layout_df.fillna("").to_dict(orient="records")
        preview_pdf = make_preview_pdf(
            tpl_bytes, rec, layout_rows, def_font_size, def_bold, top_left_mode, font_bytes
        )
        shown = render_preview_as_image(preview_pdf, zoom_dpi=st.sidebar.slider("Preview DPI", 120, 220, 160, 10))
        if not shown:
            render_preview_as_pdf(preview_pdf, height=820)
        st.download_button("⬇️ ดาวน์โหลดพรีวิว (PDF 1 หน้า)", preview_pdf, file_name="preview_1page.pdf")
    except Exception as e:
        st.error(f"เรนเดอร์พรีวิวล้มเหลว: {e}")

# ========= Export =========
st.markdown("---")
st.subheader("📦 ส่งออก PDF ทั้งชุด (1 หน้า/คน)")
if st.button("Export ทั้งชุด", type="primary"):
    with st.spinner("กำลังเรนเดอร์…"):
        try:
            layout_rows = st.session_state.layout_df.fillna("").to_dict(orient="records")
            full_pdf = make_full_pdf(
                tpl_bytes, merged, layout_rows, def_font_size, def_bold, top_left_mode, font_bytes
            )
            st.success("สำเร็จ! ดาวน์โหลดได้เลย")
            st.download_button("⬇️ ดาวน์โหลดไฟล์รวม (PDF)", full_pdf, file_name="Conversation_PerStudent_Output.pdf")
        except Exception as e:
            st.error(f"ส่งออกล้มเหลว: {e}")
