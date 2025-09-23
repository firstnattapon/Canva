import io
import streamlit as st
import pandas as pd
import PyMuPDF as fitz  # <<< ใช้ pymupdf เพื่อเลี่ยงแพ็กเกจชื่อซ้ำ "fitz"

st.set_page_config(page_title="Conversation Result (Per Student Pages)", layout="wide")
st.title("📄 Conversation Result → 1 หน้า/1 คน (รองรับ 1–2 เทอม)")
st.caption("พื้นหลังจาก Canva (อัปโหลด PDF หนึ่งหน้า) + CSV คะแนน → แปลงเป็น PDF รายคนแบบ 35 หน้า")

# ========== Sidebar: Layout & Options ==========
st.sidebar.header("🔧 ตำแหน่งข้อความ (หน่วย pt)")
name_x = st.sidebar.number_input("Name X", 0, 2000, 140)
name_y = st.sidebar.number_input("Name Y", 0, 2000, 160)
id_x   = st.sidebar.number_input("Student ID X", 0, 2000, 140)
id_y   = st.sidebar.number_input("Student ID Y", 0, 2000, 190)
s1_x   = st.sidebar.number_input("Total S1 X", 0, 2000, 430)
s1_y   = st.sidebar.number_input("Total S1 Y", 0, 2000, 300)
s2_x   = st.sidebar.number_input("Total S2 X", 0, 2000, 430)
s2_y   = st.sidebar.number_input("Total S2 Y", 0, 2000, 340)

st.sidebar.header("🅰️ ฟอนต์")
font_size = st.sidebar.number_input("ขนาดตัวอักษร", 6, 64, 16)
bold = st.sidebar.checkbox("ตัวหนา (Helvetica Bold)", value=True)

st.sidebar.header("🔗 การแม็ปข้อมูลระหว่างเทอม")
join_key = st.sidebar.selectbox("คีย์จับคู่ข้อมูล", ["Student ID", "Name - Surname"], index=0)
when_single = st.sidebar.selectbox("ถ้าอัปโหลด CSV แค่ 1 เทอม ใส่ลงช่องไหน?", ["S1", "S2"], index=0)

st.sidebar.caption("A4 แนวตั้ง ≈ 595×842 pt • ปรับตำแหน่งแล้วกดพรีวิวให้เข้าที่ 100%")

# ========== Uploaders ==========
with st.expander("วิธีใช้แบบเร็ว"):
    st.markdown("""
1) อัปโหลด **Template PDF** จาก Canva (ต้องเป็นหน้าเดียว)  
2) อัปโหลด **CSV เทอม 1** และ/หรือ **CSV เทอม 2** โครงสร้าง:
   `No, Student ID, Name - Surname, Idea, Pronunciation, Preparedness, Confidence, Total (50)`  
   - รองรับแถว `Score` และบรรทัดว่าง (ระบบคัดทิ้งให้เอง)  
3) ตั้งพิกัด **Name / Student ID / Total S1 / Total S2**  
4) เลือกคีย์จับคู่ระหว่างเทอม (ID หรือ Name)  
5) พรีวิว 1 หน้า → Export ทั้งชุด (หนึ่งหน้า/หนึ่งคน)
""")

tpl_file = st.file_uploader("อัปโหลด Template PDF (พื้นหลัง Canva / หน้าเดียว)", type=["pdf"])
csv_s1 = st.file_uploader("อัปโหลด CSV เทอม 1", type=["csv"])
csv_s2 = st.file_uploader("อัปโหลด CSV เทอม 2 (ถ้ามี)", type=["csv"])

REQUIRED_COLS = ["No","Student ID","Name - Surname","Idea","Pronunciation","Preparedness","Confidence","Total (50)"]

# ========== CSV Parsing (robust) ==========
def _decode_csv_bytes(b: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp874", "latin-1"):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return b.decode("utf-8", errors="ignore")

def parse_csv_bytes(b: bytes) -> pd.DataFrame:
    """อ่าน CSV ที่อาจมีบรรทัดว่าง/หัวตารางไม่ชัด • คืน DataFrame เฉพาะคอลัมน์ที่ต้องใช้"""
    if b is None:
        return None
    text = _decode_csv_bytes(b)
    # อ่านแบบไม่มี header ก่อน เพื่อหาแถวหัวตารางเอง
    df_raw = pd.read_csv(io.StringIO(text), header=None, dtype=str)
    header_idx = None
    max_scan = min(20, len(df_raw))
    for i in range(max_scan):
        row_vals = df_raw.iloc[i].fillna("").astype(str).tolist()
        if "No" in row_vals and "Student ID" in row_vals:
            header_idx = i
            break
    if header_idx is None:
        # ลองอ่านแบบมี header ปกติ
        df = pd.read_csv(io.StringIO(text), dtype=str).fillna("")
    else:
        headers = df_raw.iloc[header_idx].fillna("").astype(str).tolist()
        df = df_raw.iloc[header_idx+1:].copy()
        df.columns = headers
        df = df.fillna("")
    # เก็บเฉพาะคอลัมน์ที่ต้องใช้
    cols = [c for c in df.columns if c in REQUIRED_COLS]
    df = df[cols]
    # ตัดแถว Score และแถวว่าง
    if "No" in df.columns:
        mask_score = df["No"].astype(str).str.strip().str.lower().eq("score")
        df = df[~mask_score]
        df = df[df["No"].astype(str).str.strip() != ""]
    # strip header ให้สะอาด
    df.columns = [c.strip() for c in df.columns]
    return df.reset_index(drop=True)

# ========== Merge Semesters ==========
def coalesce(a, b):
    a = "" if pd.isna(a) else str(a)
    b = "" if pd.isna(b) else str(b)
    return a if a.strip() else b

def merge_semesters(df1: pd.DataFrame, df2: pd.DataFrame, key: str, when_single: str) -> pd.DataFrame:
    """รวม 2 เทอมด้วย key ที่เลือก → คอลัมน์ Name / StudentID / Total_S1 / Total_S2"""
    if df1 is not None and df2 is not None:
        merged = pd.merge(df1, df2, on=key, how="outer", suffixes=("_S1","_S2"))
        # ชื่อ / รหัส
        name = []
        sid  = []
        for _, r in merged.iterrows():
            name.append(coalesce(r.get("Name - Surname_S1",""), r.get("Name - Surname_S2","")))
            sid.append(coalesce(r.get("Student ID_S1",""), r.get("Student ID_S2","")))
        merged["Name"] = name
        merged["StudentID"] = sid
        merged["Total_S1"] = merged.get("Total (50)_S1", "")
        merged["Total_S2"] = merged.get("Total (50)_S2", "")
        out = merged[["Name","StudentID","Total_S1","Total_S2"]].copy()
        # เรียงเพื่อให้อ่านง่าย
        if key == "Student ID":
            out = out.sort_values(by=["StudentID","Name"], kind="stable")
        else:
            out = out.sort_values(by=["Name","StudentID"], kind="stable")
        out = out.reset_index(drop=True)
        return out

    # ถ้ามีแค่ไฟล์เดียว
    df = df1 if df1 is not None else df2
    if df is None:
        return None
    out = pd.DataFrame()
    out["Name"] = df.get("Name - Surname", "")
    out["StudentID"] = df.get("Student ID", "")
    if when_single == "S1":
        out["Total_S1"] = df.get("Total (50)", "")
        out["Total_S2"] = ""
    else:
        out["Total_S1"] = ""
        out["Total_S2"] = df.get("Total (50)", "")
    # เรียง
    if key == "Student ID":
        out = out.sort_values(by=["StudentID","Name"], kind="stable")
    else:
        out = out.sort_values(by=["Name","StudentID"], kind="stable")
    return out.reset_index(drop=True)

# ========== Drawing ==========
def draw_one(page, name, sid, total_s1, total_s2, font_size=16, bold=True):
    fontname = "helvb" if bold else "helv"
    def put(x, y, text):
        if text is None: 
            return
        s = str(text).strip()
        if not s: 
            return
        page.insert_text((x, y), s, fontsize=font_size, fontname=fontname, fill=(0,0,0))
    put(name_x, name_y, name)
    put(id_x,   id_y,   sid)
    put(s1_x,   s1_y,   total_s1)
    put(s2_x,   s2_y,   total_s2)

def build_pdf(template_bytes: bytes, records: pd.DataFrame, font_size=16, bold=True) -> bytes:
    tpl = fitz.open("pdf", template_bytes)
    if tpl.page_count < 1:
        raise ValueError("Template PDF ต้องมีอย่างน้อย 1 หน้า")
    w, h = tpl[0].rect.width, tpl[0].rect.height
    out = fitz.open()
    for _, r in records.iterrows():
        page = out.new_page(width=w, height=h)
        page.show_pdf_page(page.rect, tpl, 0)  # วางพื้นหลัง Canva
        draw_one(
            page,
            r.get("Name",""),
            r.get("StudentID",""),
            r.get("Total_S1",""),
            r.get("Total_S2",""),
            font_size=font_size,
            bold=bold
        )
    return out.tobytes()

# ========== Main ==========
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
            st.subheader("CSV เทอม 1 (ตัวอย่าง 10 แถว)")
            st.dataframe(df1.head(10), use_container_width=True)
        if df2 is not None:
            st.subheader("CSV เทอม 2 (ตัวอย่าง 10 แถว)")
            st.dataframe(df2.head(10), use_container_width=True)

        key = "Student ID" if join_key == "Student ID" else "Name - Surname"
        merged = merge_semesters(df1, df2, key, when_single)
        st.success(f"รวมข้อมูลพร้อมแปลง: {len(merged)} คน (1 หน้า/คน)")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("👀 พรีวิวหน้าแรก"):
                one = merged.head(1)
                pdf_bytes = build_pdf(tpl_bytes, one, font_size=font_size, bold=bold)
                st.download_button("ดาวน์โหลดพรีวิว (PDF 1 หน้า)", pdf_bytes, file_name="preview_1page.pdf")
        with c2:
            if st.button("📦 Export ทั้งชุด (PDF รวมทุกคน)"):
                pdf_bytes = build_pdf(tpl_bytes, merged, font_size=font_size, bold=bold)
                st.download_button("⬇️ ดาวน์โหลด PDF รวม", pdf_bytes, file_name="Conversation_PerStudent_Output.pdf")
                st.balloons()
