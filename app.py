
import io
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Conversation Result (Per Student) — Dual Backend", layout="wide")
st.title("📄 Conversation Result → 1 หน้า/1 คน (PyMuPDF หรือ Fallback ReportLab+PyPDF)")
st.caption("พื้นหลังจาก Canva (PDF หน้าเดียว) + CSV คะแนน → แปลงเป็น PDF รายคน รองรับ 1–2 เทอม")

# ===== Sidebar =====
st.sidebar.header("🔧 ตำแหน่งข้อความ (หน่วย pt)")
name_x = st.sidebar.number_input("Name X", 0, 2000, 140)
name_y = st.sidebar.number_input("Name Y", 0, 2000, 160)
# id_x   = st.sidebar.number_input("Student ID X", 0, 2000, 140)
# id_y   = st.sidebar.number_input("Student ID Y", 0, 2000, 190)
# s1_x   = st.sidebar.number_input("Total S1 X", 0, 2000, 430)
# s1_y   = st.sidebar.number_input("Total S1 Y", 0, 2000, 300)
# s2_x   = st.sidebar.number_input("Total S2 X", 0, 2000, 430)
# s2_y   = st.sidebar.number_input("Total S2 Y", 0, 2000, 340)

st.sidebar.header("🅰️ ฟอนต์")
font_size = st.sidebar.number_input("ขนาดตัวอักษร", 6, 64, 16)
bold = st.sidebar.checkbox("ตัวหนา (Helvetica Bold)", value=True)

st.sidebar.header("🔗 การแม็ปข้อมูล")
join_key = st.sidebar.selectbox("คีย์จับคู่ข้อมูล", ["Student ID", "Name - Surname"], index=0)
when_single = st.sidebar.selectbox("ถ้าอัปโหลด CSV แค่ 1 เทอม ใส่ลงช่องไหน?", ["S1", "S2"], index=0)

# ===== Uploaders =====
tpl_file = st.file_uploader("อัปโหลด Template PDF (พื้นหลัง Canva / หน้าเดียว)", type=["pdf"])
csv_s1 = st.file_uploader("อัปโหลด CSV เทอม 1", type=["csv"])
csv_s2 = st.file_uploader("อัปโหลด CSV เทอม 2 (ถ้ามี)", type=["csv"])

REQUIRED_COLS = ["No","Student ID","Name - Surname","Idea","Pronunciation","Preparedness","Confidence","Total (50)"]

# ===== CSV parsing (robust) =====
def _decode_csv_bytes(b: bytes) -> str:
    for enc in ("utf-8-sig","utf-8","cp874","latin-1"):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return b.decode("utf-8", errors="ignore")

def parse_csv_bytes(b: bytes) -> pd.DataFrame:
    if b is None: return None
    text = _decode_csv_bytes(b)
    import pandas as pd, io
    df_raw = pd.read_csv(io.StringIO(text), header=None, dtype=str)
    header_idx = None
    for i in range(min(20, len(df_raw))):
        row = df_raw.iloc[i].fillna("").astype(str).tolist()
        if "No" in row and "Student ID" in row:
            header_idx = i
            break
    if header_idx is None:
        df = pd.read_csv(io.StringIO(text), dtype=str).fillna("")
    else:
        headers = df_raw.iloc[header_idx].fillna("").astype(str).tolist()
        df = df_raw.iloc[header_idx+1:].copy()
        df.columns = headers
        df = df.fillna("")
    cols = [c for c in df.columns if c in REQUIRED_COLS]
    df = df[cols]
    if "No" in df.columns:
        mask = df["No"].astype(str).str.strip().str.lower().eq("score")
        df = df[~mask]
        df = df[df["No"].astype(str).str.strip() != ""]
    df.columns = [c.strip() for c in df.columns]
    return df.reset_index(drop=True)

def coalesce(a,b):
    a = "" if pd.isna(a) else str(a); b = "" if pd.isna(b) else str(b)
    return a if a.strip() else b

def merge_semesters(df1, df2, key, when_single):
    import pandas as pd
    if df1 is not None and df2 is not None:
        merged = pd.merge(df1, df2, on=key, how="outer", suffixes=("_S1","_S2"))
        merged["Name"] = [coalesce(a,b) for a,b in zip(merged.get("Name - Surname_S1",""), merged.get("Name - Surname_S2",""))]
        merged["StudentID"] = [coalesce(a,b) for a,b in zip(merged.get("Student ID_S1",""), merged.get("Student ID_S2",""))]
        merged["Total_S1"] = merged.get("Total (50)_S1","")
        merged["Total_S2"] = merged.get("Total (50)_S2","")
        out = merged[["Name","StudentID","Total_S1","Total_S2"]].copy()
        out = out.sort_values(by=[("StudentID" if key=="Student ID" else "Name")], kind="stable").reset_index(drop=True)
        return out
    df = df1 if df1 is not None else df2
    if df is None: return None
    out = pd.DataFrame()
    out["Name"] = df.get("Name - Surname","")
    out["StudentID"] = df.get("Student ID","")
    if when_single == "S1":
        out["Total_S1"] = df.get("Total (50)",""); out["Total_S2"] = ""
    else:
        out["Total_S1"] = ""; out["Total_S2"] = df.get("Total (50)","")
    out = out.sort_values(by=[("StudentID" if key=="Student ID" else "Name")], kind="stable").reset_index(drop=True)
    return out

# ===== Backends =====
def build_pdf_pymupdf(tpl_bytes: bytes, records: pd.DataFrame, font_size=16, bold=True) -> bytes:
    import pymupdf as fitz
    tpl = fitz.open("pdf", tpl_bytes)
    w, h = tpl[0].rect.width, tpl[0].rect.height
    out = fitz.open()
    fontname = "helvb" if bold else "helv"
    for _, r in records.iterrows():
        page = out.new_page(width=w, height=h)
        page.show_pdf_page(page.rect, tpl, 0)
        def put(x,y,text):
            s = "" if text is None else str(text).strip()
            if not s: return
            page.insert_text((x,y), s, fontsize=font_size, fontname=fontname, fill=(0,0,0))
        put(name_x, name_y, r.get("Name",""))
        put(id_x,   id_y,   r.get("StudentID",""))
        put(s1_x,   s1_y,   r.get("Total_S1",""))
        put(s2_x,   s2_y,   r.get("Total_S2",""))
    return out.tobytes()

def build_pdf_overlay(tpl_bytes: bytes, records: pd.DataFrame, font_size=16, bold=True) -> bytes:
    # Fallback: ReportLab + pypdf
    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.utils import ImageReader
    from reportlab.lib.colors import black
    import tempfile, os

    # Save template to temp file (PdfReader needs a path or stream)
    tpl_path = tempfile.mktemp(suffix=".pdf")
    with open(tpl_path, "wb") as f:
        f.write(tpl_bytes)
    reader = PdfReader(tpl_path)
    page = reader.pages[0]
    w = float(page.mediabox.width)
    h = float(page.mediabox.height)

    writer = PdfWriter()
    fontname = "Helvetica-Bold" if bold else "Helvetica"

    # For each record, create an overlay page with ReportLab and merge onto template
    for _, r in records.iterrows():
        overlay_path = tempfile.mktemp(suffix=".pdf")
        c = canvas.Canvas(overlay_path, pagesize=(w, h))
        c.setFillColor(black)
        c.setFont(fontname, float(font_size))

        def put(x,y,text):
            s = "" if text is None else str(text).strip()
            if not s: return
            # ReportLab origin is bottom-left; our coordinates assume top-left? We assume PDF coords bottom-left.
            # If user coordinates were tuned for PyMuPDF (same coords), both use bottom-left, so ok.
            c.drawString(float(x), float(y), s)

        put(name_x, name_y, r.get("Name",""))
        put(id_x,   id_y,   r.get("StudentID",""))
        put(s1_x,   s1_y,   r.get("Total_S1",""))
        put(s2_x,   s2_y,   r.get("Total_S2",""))
        c.showPage()
        c.save()

        # Merge overlay onto template page copy
        base_page = reader.pages[0]
        from pypdf import PdfReader as _PdfReader
        overlay_reader = _PdfReader(overlay_path)
        overlay_page = overlay_reader.pages[0]
        base_page.merge_page(overlay_page)  # draw overlay on top
        writer.add_page(base_page)

        # cleanup overlay temp
        try: os.remove(overlay_path)
        except: pass

    # Write writer to bytes
    out_bytes = io.BytesIO()
    writer.write(out_bytes)
    out_bytes.seek(0)

    try: os.remove(tpl_path)
    except: pass

    return out_bytes.getvalue()

# ===== Main =====
if tpl_file is None:
    st.info("อัปโหลด **Template PDF (หน้าเดียว)** ก่อน")
else:
    tpl_bytes = tpl_file.getvalue()
    df1 = parse_csv_bytes(csv_s1.getvalue()) if csv_s1 is not None else None
    df2 = parse_csv_bytes(csv_s2.getvalue()) if csv_s2 is not None else None

    if (df1 is None) and (df2 is None):
        st.warning("ยังไม่พบ CSV คะแนน กรุณาอัปโหลดอย่างน้อย 1 เทอม")
    else:
        if df1 is not None:
            st.subheader("CSV เทอม 1 (ตัวอย่าง 10 แถว)"); st.dataframe(df1.head(10), use_container_width=True)
        if df2 is not None:
            st.subheader("CSV เทอม 2 (ตัวอย่าง 10 แถว)"); st.dataframe(df2.head(10), use_container_width=True)

        key = "Student ID" if join_key == "Student ID" else "Name - Surname"
        merged = merge_semesters(df1, df2, key, when_single)
        st.success(f"รวมข้อมูลพร้อมแปลง: {len(merged)} คน (1 หน้า/คน)")

        # Choose backend
        backend = st.radio("เลือกเอนจินเรนเดอร์", ["Auto (PyMuPDF ถ้ามี)", "Fallback (ReportLab + PyPDF)"], index=0, horizontal=True)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("👀 พรีวิวหน้าแรก"):
                one = merged.head(1)
                try:
                    if backend.startswith("Auto"):
                        pdf_bytes = build_pdf_pymupdf(tpl_bytes, one, font_size=font_size, bold=bold)
                    else:
                        raise ImportError("force-fallback")
                except Exception:
                    pdf_bytes = build_pdf_overlay(tpl_bytes, one, font_size=font_size, bold=bold)
                st.download_button("ดาวน์โหลดพรีวิว (PDF 1 หน้า)", pdf_bytes, file_name="preview_1page.pdf")
        with col2:
            if st.button("📦 Export ทั้งชุด (PDF รวมทุกคน)"):
                try:
                    if backend.startswith("Auto"):
                        pdf_bytes = build_pdf_pymupdf(tpl_bytes, merged, font_size=font_size, bold=bold)
                    else:
                        raise ImportError("force-fallback")
                except Exception:
                    pdf_bytes = build_pdf_overlay(tpl_bytes, merged, font_size=font_size, bold=bold)
                st.download_button("⬇️ ดาวน์โหลด PDF รวม", pdf_bytes, file_name="Conversation_PerStudent_Output.pdf")
                st.balloons()
