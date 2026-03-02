import streamlit as st
import json
import os
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
from datetime import datetime

# --- Configuration & Styling ---
st.set_page_config(page_title="Work Report System 2026", layout="wide")

USER_FILE = "users.json"
REPORT_FILE = "all_reports.json"

def load_data(file):
    if os.path.exists(file):
        with open(file, "r", encoding="utf-8") as f: return json.load(f)
    return {}

def save_data(file, data):
    with open(file, "w", encoding="utf-8") as f: json.dump(data, f, indent=4, ensure_ascii=False)

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

# --- Word Generation Engine ---
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
    full_name = f"{u_info.get('nametitle', '')}{u_info['name']}"
    
    position_text = (
        f"{full_name}\n"
        f"ลูกจ้างเหมาบริการสนับสนุนการขับเคลื่อนงานนโยบายของรัฐบาลและ\n"
        f"การให้บริการประชาชนในพื้นที่ (ดำเนินงานฝ่าย{u_info['position']})"
    )
    run = info.add_run(position_text)
    set_font(run, 16)

    # ตาราง
    table = doc.add_table(rows=2, cols=8); table.style = 'Table Grid'
    headers = ['วัน เดือน ปี', 'งานที่ทำ', 'จำนวน\n(เรื่อง/ชิ้น)', 'ผลการดำเนินงาน', '', '', 'ระยะเวลา\nดำเนินงาน', 'หมายเหตุ']
    
    for i, h in enumerate(headers):
        p = table.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h); run.bold = True; set_font(run, 14)

    table.rows[0].cells[3].merge(table.rows[0].cells[4]).merge(table.rows[0].cells[5])
    sub_headers = {3: 'เสร็จ', 4: 'ไม่เสร็จ', 5: 'ส่งแก้ไข'}
    for col_idx, text in sub_headers.items():
        p = table.rows[1].cells[col_idx].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
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
            p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
            set_font(p.add_run(str(val)), 14)

    # ท้ายกระดาษ (ซ้ายล่างสุดเสมอ)
    footer = doc.sections[0].footer
    footer_para = footer.paragraphs[0]; footer_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = footer_para.add_run(f"(ลงชื่อ)............................................................ผู้รับจ้าง")
    set_font(run, 16)

    bio = BytesIO(); doc.save(bio); return bio.getvalue()

# --- Authentication & UI ---
users = load_data(USER_FILE)
if "admin" not in users:
    users["admin"] = {"nametitle": "นาย", "name": "Administrator", "position": "บริหารทั่วไป", "password": "admin"}
    save_data(USER_FILE, users)

if "logged_in" not in st.session_state: st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("📱 ระบบรายงานการจ้างเหมา")
    tab1, tab2 = st.tabs(["🔐 Login", "📝 Register"])
    with tab1:
        # แก้ไข Duplicate ID โดยการเพิ่ม key
        u = st.text_input("User", key="login_user")
        p = st.text_input("Pass", type="password", key="login_pass")
        if st.button("เข้าสู่ระบบ", key="login_btn"):
            users = load_data(USER_FILE)
            if u in users and users[u]["password"] == p:
                st.session_state.logged_in, st.session_state.username, st.session_state.user_info = True, u, users[u]
                st.rerun()
            else: st.error("ข้อมูลไม่ถูกต้อง")
    with tab2:
        # แก้ไข Duplicate ID โดยการเพิ่ม key
        t = st.selectbox("คำนำหน้าชื่อ", ["นาย", "นาง", "นางสาว"], key="reg_title")
        nu = st.text_input("User", key="reg_user")
        nn = st.text_input("ชื่อ-สกุล", key="reg_name")
        np = st.text_input("Pass", type="password", key="reg_pass")
        npos = st.text_input("ฝ่าย (เช่น บริหารทั่วไป)", key="reg_pos")
        if st.button("สมัครสมาชิก", key="reg_btn"):
            users = load_data(USER_FILE)
            if nu in users: st.error("Username นี้มีผู้ใช้แล้ว")
            else:
                users[nu] = {"nametitle": t, "name": nn, "position": npos, "password": np}
                save_data(USER_FILE, users); st.success("สมัครสำเร็จ! กรุณาไปที่หน้า Login")
else:
    curr_u, user = st.session_state.username, st.session_state.user_info
    with st.sidebar:
        st.write(f"สวัสดีคุณ: {user['name']}")
        if st.button("Logout"): st.session_state.logged_in = False; st.rerun()
        
        # เพิ่มปุ่มเคลียร์ข้อมูลสำหรับ Admin
        if curr_u == "admin":
            st.divider()
            if st.button("🗑️ ลบรายงานทั้งหมด (Reset)"):
                if os.path.exists(REPORT_FILE):
                    os.remove(REPORT_FILE)
                    st.success("ล้างข้อมูลเรียบร้อย!")
                    st.rerun()

    if curr_u == "admin":
        st.title("👨‍💼 แผงควบคุม Admin")
        all_rep = load_data(REPORT_FILE)
        all_u = load_data(USER_FILE)
        staff_list = [k for k in all_rep.keys() if k != "admin"]
        if staff_list:
            target = st.selectbox("เลือกพนักงาน", staff_list)
            df = pd.DataFrame(all_rep[target])
            df['dt'] = pd.to_datetime(df['date'], format='%d/%m/%Y')
            dr = st.date_input("ช่วงวันที่", value=(df['dt'].min().date(), df['dt'].max().date()))
            if len(dr) == 2 and st.button("📥 ดาวน์โหลดไฟล์ Word"):
                mask = (df['dt'].dt.date >= dr[0]) & (df['dt'].dt.date <= dr[1])
                st.download_button(f"โหลดไฟล์ {target}", generate_word(all_u[target], df.loc[mask]), f"Report_{target}.docx")
        else: st.info("ยังไม่มีข้อมูลรายงานในระบบ")
    else:
        st.title("📝 บันทึกงาน")
        init_df = pd.DataFrame({'วันที่': [datetime.now().date()], 'งานที่ทำ': [""], 'จำนวนรวม': [0], 'เสร็จ': [0], 'ไม่เสร็จ': [0], 'ส่งแก้ไข': [0], 'ระยะเวลา': ["1 วัน"], 'หมายเหตุ': [""]})
        ed_df = st.data_editor(init_df, num_rows="dynamic", use_container_width=True, column_config={"วันที่": st.column_config.DateColumn(format="DD/MM/YYYY")})
        if st.button("🚀 บันทึกรายงาน"):
            all_rep = load_data(REPORT_FILE)
            new_data = [{"date": r['วันที่'].strftime("%d/%m/%Y"), "task": r['งานที่ทำ'], "amount": r['จำนวนรวม'], "done": r['เสร็จ'], "pending": r['ไม่เสร็จ'], "edit": r['ส่งแก้ไข'], "duration": r['ระยะเวลา'], "remark": r['หมายเหตุ']} for _, r in ed_df.iterrows()]
            if curr_u not in all_rep: all_rep[curr_u] = []
            all_rep[curr_u].extend(new_data); save_data(REPORT_FILE, all_rep); st.success("บันทึกแล้ว!"); st.balloons()