import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
from datetime import datetime

# --- Configuration ---
st.set_page_config(page_title="Work Report System 2026", layout="wide")

# เชื่อมต่อกับ Google Sheets (ต้องเอา URL ไปใส่ใน .streamlit/secrets.toml หรือใส่ตรงๆ ในคอนเน็คชั่น)
# ในขั้นตอน Deploy บน Cloud คุณนิวต้องเอา URL ใส่ใน Secrets ของ Streamlit นะครับ
conn = st.connection("gsheets", type=GSheetsConnection)

def get_users():
    return conn.read(worksheet="users")

def get_all_reports():
    return conn.read(worksheet="reports")

def save_user(new_user_df):
    existing_users = get_users()
    updated_users = pd.concat([existing_users, new_user_df], ignore_index=True)
    conn.update(worksheet="users", data=updated_users)

def save_report(new_report_df):
    existing_reports = get_all_reports()
    updated_reports = pd.concat([existing_reports, new_report_df], ignore_index=True)
    conn.update(worksheet="reports", data=updated_reports)

# --- ฟังก์ชันอื่นๆ (set_font, format_thai_date, generate_word) ใช้ของเดิมได้เลยครับ ---
def set_font(run, size=16):
    run.font.name = 'TH SarabunIT๙'
    run.font.size = Pt(size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH SarabunIT๙')
    run._element.rPr.rFonts.set(qn('w:ascii'), 'TH SarabunIT๙')
    run._element.rPr.rFonts.set(qn('w:hAnsi'), 'TH SarabunIT๙')

def format_thai_date(date_str):
    try:
        day, month, year = date_str.split('/')
        thai_months_short = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
        return f"{int(day)} {thai_months_short[int(month)-1]} {year}"
    except: return date_str

def generate_word(u_info, filtered_df):
    doc = Document()
    thai_months_full = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    now = datetime.now()
    current_month = thai_months_full[now.month - 1]
    current_year_be = now.year + 543

    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(0.5), Inches(0.8)
        section.left_margin, section.right_margin = Inches(1.0), Inches(0.5)

    # หัวกระดาษ
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(f"(ตัวอย่าง)\n(ร่าง) รายงานการทำงานจ้าง ประจำเดือน {current_month} พ.ศ. {current_year_be}")
    run.bold = True
    set_font(run, 18)

    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    full_name = f"{u_info['nametitle']}{u_info['name']}"
    position_text = f"{full_name}\nลูกจ้างเหมาบริการสนับสนุนการขับเคลื่อนงานนโยบายของรัฐบาลและ\nการให้บริการประชาชนในพื้นที่ (ดำเนินงานฝ่าย{u_info['position']})"
    run = info.add_run(position_text)
    set_font(run, 16)

    # ตาราง
    table = doc.add_table(rows=2, cols=8); table.style = 'Table Grid'
    headers = ['วัน เดือน ปี', 'งานที่ทำ', 'จำนวน\n(เรื่อง/ชิ้น)', 'ผลการดำเนินงาน', '', '', 'ระยะเวลา\nดำเนินงาน', 'หมายเหตุ']
    for i, h in enumerate(headers):
        p = table.rows[0].cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h); run.bold = True; set_font(run, 14)

    table.rows[0].cells[3].merge(table.rows[0].cells[4]).merge(table.rows[0].cells[5])
    sub_headers = {3: 'เสร็จ', 4: 'ไม่เสร็จ', 5: 'ส่งแก้ไข'}
    for col_idx, text in sub_headers.items():
        p = table.rows[1].cells[col_idx].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(p.add_run(text), 14)

    for col in [0, 1, 2, 6, 7]:
        table.rows[0].cells[col].merge(table.rows[1].cells[col])
        table.rows[0].cells[col].vertical_alignment = 1

    for _, r in filtered_df.iterrows():
        row_cells = table.add_row().cells
        data_list = [format_thai_date(r['date']), r['task'], r['amount'], r['done'], r['pending'], r['edit'], r['duration'], r['remark']]
        for i, val in enumerate(data_list):
            p = row_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT
            set_font(p.add_run(str(val)), 14)

    # ท้ายกระดาษ
    footer = doc.sections[0].footer
    footer_para = footer.paragraphs[0]; footer_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = footer_para.add_run(f"(ลงชื่อ)............................................................ผู้รับจ้าง")
    set_font(run, 16)

    bio = BytesIO(); doc.save(bio); return bio.getvalue()

# --- UI & Logic ---
if "logged_in" not in st.session_state: st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("📱 ระบบรายงานการจ้างเหมา (Google Sheets)")
    tab1, tab2 = st.tabs(["🔐 Login", "📝 Register"])
    
    with tab1:
        u_input = st.text_input("Username", key="l_u")
        p_input = st.text_input("Password", type="password", key="l_p")
        if st.button("เข้าสู่ระบบ"):
            users_df = get_users()
            user_data = users_df[users_df['username'] == u_input]
            if not user_data.empty and str(user_data.iloc[0]['password']) == p_input:
                st.session_state.logged_in = True
                st.session_state.username = u_input
                st.session_state.user_info = user_data.iloc[0].to_dict()
                st.rerun()
            else: st.error("User หรือ Password ไม่ถูกต้อง")

    with tab2:
        t = st.selectbox("คำนำหน้าชื่อ", ["นาย", "นาง", "นางสาว"], key="r_t")
        nu = st.text_input("Username", key="r_u")
        nn = st.text_input("ชื่อ-สกุล", key="r_n")
        np = st.text_input("Password", type="password", key="r_p")
        npos = st.text_input("ฝ่าย (เช่น บริหารทั่วไป)", key="r_pos")
        if st.button("สมัครสมาชิก"):
            new_user = pd.DataFrame([{"nametitle": t, "name": nn, "position": npos, "password": np, "username": nu}])
            save_user(new_user); st.success("สมัครสำเร็จ! ไปที่หน้า Login ได้เลย")

else:
    # หน้าแสดงผลหลังจาก Login (เหมือนเดิม แต่เปลี่ยนตอน Save ให้เรียก save_report())
    user = st.session_state.user_info
    st.sidebar.write(f"สวัสดีคุณ: {user['name']}")
    if st.sidebar.button("Logout"): st.session_state.logged_in = False; st.rerun()

    if st.session_state.username == "admin":
        st.title("👨‍💼 Admin Panel")
        all_rep = get_all_reports()
        # ส่วน Admin สำหรับโหลด Word (ทำเหมือนเดิม)
    else:
        st.title("📝 บันทึกงาน")
        # ส่วน data_editor (ทำเหมือนเดิม)
        # เมื่อกดปุ่ม "บันทึกรายงาน" ให้ใช้:
        # new_report_df = pd.DataFrame(new_data)
        # save_report(new_report_df)