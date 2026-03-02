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

# เชื่อมต่อ Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

# --- Functions จัดการข้อมูล (บังคับไม่อ่าน Cache เพื่อป้องกันข้อมูลหาย) ---
def get_users():
    return conn.read(worksheet="users", ttl=0)

def get_all_reports():
    return conn.read(worksheet="reports", ttl=0)

def save_user(new_user_df):
    st.cache_data.clear() # ล้างแคชก่อนอ่านข้อมูลเดิม
    existing_users = get_users()
    updated_users = pd.concat([existing_users, new_user_df], ignore_index=True)
    conn.update(worksheet="users", data=updated_users)
    st.cache_data.clear() # ล้างแคชหลังบันทึก

def save_report(new_report_df):
    st.cache_data.clear()
    existing_reports = get_all_reports()
    updated_reports = pd.concat([existing_reports, new_report_df], ignore_index=True)
    conn.update(worksheet="reports", data=updated_reports)
    st.cache_data.clear()

# --- ฟังก์ชันจัดการฟอนต์และวันที่ ---
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

# --- ฟังก์ชันสร้างไฟล์ Word (ดึงข้อมูลจาก Google Sheet) ---
def generate_word(u_info, filtered_df):
    doc = Document()
    thai_months_full = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    now = datetime.now()
    current_month = thai_months_full[now.month - 1]
    current_year_be = now.year + 543

    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(0.5), Inches(0.8)
        section.left_margin, section.right_margin = Inches(1.0), Inches(0.5)

    # ส่วนหัวกระดาษ
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(f"รายงานการทำงานจ้าง ประจำเดือน {current_month} พ.ศ. {current_year_be}")
    run.bold = True
    set_font(run, 18)

    # ข้อมูลพนักงาน (ตำแหน่งแบบยาว)
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    full_name = f"{u_info.get(ชื่อ ,'nametitle', ' ')}{u_info.get('name', '')}"
    pos_text = (f"{full_name}\nตำแหน่ง ลูกจ้างเหมาบริการสนับสนุนการขับเคลื่อนงานนโยบายของรัฐบาลและ\n"
                f"การให้บริการประชาชนในพื้นที่ (ดำเนินงานฝ่าย{u_info.get('position', '')})")
    run = info.add_run(pos_text)
    set_font(run, 16)

    # สร้างตาราง
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

    # วนลูปใส่ข้อมูลงาน
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
    run = footer_para.add_run(f"(นาย/นาง/นางสาว)............................................................ผู้รับจ้าง")
    set_font(run, 16)

    bio = BytesIO(); doc.save(bio); return bio.getvalue()

# --- ระบบ Login / Register ---
if "logged_in" not in st.session_state: st.session_state.logged_in = False

users_df = get_users()

# สร้าง Admin อัตโนมัติใน Sheet ถ้ายังไม่มี
if users_df.empty or "admin" not in users_df['username'].values:
    admin_setup = pd.DataFrame([{"nametitle": "นาย", "name": "Administrator", "position": "บริหารทั่วไป", "password": "admin", "username": "admin"}])
    save_user(admin_setup)
    st.rerun()

if not st.session_state.logged_in:
    st.title("📱 ระบบรายงานการจ้างเหมา (Google Sheets)")
    tab1, tab2 = st.tabs(["🔐 Login", "📝 Register"])
    
    with tab1:
        u_in = st.text_input("Username", key="login_u")
        p_in = st.text_input("Password", type="password", key="login_p")
        if st.button("เข้าสู่ระบบ"):
            u_match = users_df[users_df['username'] == u_in]
            if not u_match.empty and str(u_match.iloc[0]['password']) == p_in:
                st.session_state.logged_in = True
                st.session_state.username = u_in
                st.session_state.user_info = u_match.iloc[0].to_dict()
                st.rerun()
            else: st.error("Username หรือ Password ไม่ถูกต้อง")

    with tab2:
        t = st.selectbox("คำนำหน้าชื่อ", ["นาย", "นาง", "นางสาว"], key="reg_t")
        nu = st.text_input("Username", key="reg_u")
        nn = st.text_input("ชื่อ-สกุล", key="reg_n")
        np = st.text_input("Password", type="password", key="reg_p")
        npos = st.text_input("ฝ่าย", key="reg_pos")
        if st.button("สมัครสมาชิก"):
            if nu in users_df['username'].values: st.error("Username นี้มีผู้ใช้แล้ว")
            else:
                new_u = pd.DataFrame([{"nametitle": t, "name": nn, "position": npos, "password": np, "username": nu}])
                save_user(new_u)
                st.success("สมัครสมาชิกสำเร็จ! กรุณาเข้าสู่ระบบ")
                st.rerun()

else:
    # --- หน้าหลักหลัง Login ---
    curr_u = st.session_state.username
    user = st.session_state.user_info
    
    with st.sidebar:
        st.write(f"ผู้ใช้งาน: {user['name']}")
        if st.button("Logout"): 
            st.session_state.logged_in = False
            st.cache_data.clear()
            st.rerun()

    # --- ส่วน Admin: สร้างไฟล์ Word ---
    if curr_u == "admin":
        st.title("👨‍💼 แผงควบคุม Admin (สร้างไฟล์รายงาน)")
        all_rep_df = get_all_reports()
        staff_list = users_df[users_df['username'] != 'admin']['username'].tolist()
        
        if staff_list:
            target = st.selectbox("เลือกพนักงาน", staff_list)
            df_target = all_rep_df[all_rep_df['username'] == target].copy()
            
            if not df_target.empty:
                df_target['dt'] = pd.to_datetime(df_target['date'], format='%d/%m/%Y')
                dr = st.date_input("เลือกช่วงวันที่", value=(df_target['dt'].min().date(), df_target['dt'].max().date()))
                
                if st.button("📥 สร้างไฟล์และดาวน์โหลด Word"):
                    mask = (df_target['dt'].dt.date >= dr[0]) & (df_target['dt'].dt.date <= dr[1])
                    t_info = users_df[users_df['username'] == target].iloc[0].to_dict()
                    word_data = generate_word(t_info, df_target.loc[mask].sort_values('dt'))
                    st.download_button(f"โหลดไฟล์ {target}", word_data, f"Report_{target}.docx")
            else: st.info("พนักงานคนนี้ยังไม่มีการบันทึกงาน")
        else: st.info("ยังไม่มีสมาชิกในระบบ")

    # --- ส่วน User: บันทึกรายงาน ---
    else:
        st.title("📝 บันทึกงานประจำวัน")
        init_data = pd.DataFrame({'วันที่': [datetime.now().date()], 'งานที่ทำ': [""], 'จำนวนรวม': [0], 'เสร็จ': [1], 'ไม่เสร็จ': [0], 'ส่งแก้ไข': [0], 'ระยะเวลา': ["1 วัน"], 'หมายเหตุ': [""]})
        ed_df = st.data_editor(init_data, num_rows="dynamic", use_container_width=True, column_config={"วันที่": st.column_config.DateColumn(format="DD/MM/YYYY")})
        
        if st.button("🚀 บันทึกรายงาน"):
            new_reps = []
            for _, r in ed_df.iterrows():
                new_reps.append({
                    "username": curr_u,
                    "date": r['วันที่'].strftime("%d/%m/%Y"),
                    "task": r['งานที่ทำ'], "amount": r['จำนวนรวม'], "done": r['เสร็จ'], 
                    "pending": r['ไม่เสร็จ'], "edit": r['ส่งแก้ไข'], "duration": r['ระยะเวลา'], "remark": r['หมายเหตุ']
                })
            save_report(pd.DataFrame(new_reps))
            st.success("บันทึกข้อมูลเรียบร้อยแล้ว!")

            st.balloons()

