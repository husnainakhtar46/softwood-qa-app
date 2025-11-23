import streamlit as st
import pandas as pd
import json
import uuid
from datetime import date, datetime
from fpdf import FPDF
from PIL import Image
import gspread
from google.oauth2.service_account import Credentials
import io

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
SHEET_NAME = "Softwood_QA_Index" 

st.set_page_config(page_title="Softwood QA App", page_icon="🧵", layout="wide")


@st.cache_resource
def get_google_connection():
    try:

        if "gcp_service_account" in st.secrets:
            creds_dict = st.secrets["gcp_service_account"]

            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)

        else:
            creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
            
        gc = gspread.authorize(creds)
        return gc
    except Exception as e:
        st.error(f"Detailed Connection Error: {e}")
        return None

@st.cache_data(ttl=60) 
def fetch_sheet_data(sheet_name, tab_name):
    gc = get_google_connection()
    if not gc: return []
    try:
        sh = gc.open(sheet_name)
        ws = sh.worksheet(tab_name)
        return ws.get_all_records()
    except:
        return []

class GoogleManager:
    def __init__(self):
        self.gc = get_google_connection()
        if not self.gc:
            st.error("Connection Error. Please check Secrets or local JSON file.")
            st.stop()
        try:
            self.sh = self.gc.open(SHEET_NAME)
        except:
            st.error(f"Could not find Sheet '{SHEET_NAME}'.")
            st.stop()
        self._ensure_worksheet("Templates")
        self._ensure_worksheet("Reports")
        self._ensure_worksheet("Customers")

    def _ensure_worksheet(self, title):
        try:
            self.sh.worksheet(title)
        except:
            if title == "Customers":
                self.sh.add_worksheet(title=title, rows=100, cols=3)
                ws = self.sh.worksheet(title)
                ws.append_row(["ID", "Customer_Name", "Emails_JSON"])
            else:
                self.sh.add_worksheet(title=title, rows=100, cols=4)
                ws = self.sh.worksheet(title)
                ws.append_row(["ID", "Name", "Date_Created", "JSON_Data"])

    def save_entry(self, type_sheet, name, data):
        ws = self.sh.worksheet(type_sheet)
        unique_id = str(uuid.uuid4())
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        json_dump = json.dumps(data)
        
        if type_sheet == "Customers":
            ws.append_row([unique_id, name, json_dump])
        else:
            ws.append_row([unique_id, name, current_time, json_dump])
        st.cache_data.clear()

    def get_list(self, type_sheet):
        return fetch_sheet_data(SHEET_NAME, type_sheet)

    def delete_entry(self, type_sheet, unique_id):
        ws = self.sh.worksheet(type_sheet)
        cell = ws.find(unique_id)
        if cell:
            ws.delete_rows(cell.row)
            st.cache_data.clear()

if 'google_db' not in st.session_state:
    st.session_state.google_db = GoogleManager()

db = st.session_state.google_db


def send_email_with_pdf(recipient_list, subject, body, pdf_bytes, pdf_filename):

    try:
        SENDER_EMAIL = st.secrets["email"]["address"]
        SENDER_PASSWORD = st.secrets["email"]["password"]
    except:
        return False, "Email secrets not found in Config!"

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(recipient_list)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    part = MIMEApplication(pdf_bytes, Name=pdf_filename)
    part['Content-Disposition'] = f'attachment; filename="{pdf_filename}"'
    msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        text = msg.as_string()
        server.sendmail(SENDER_EMAIL, recipient_list, text)
        server.quit()
        return True, "Email sent successfully!"
    except Exception as e:
        return False, str(e)


class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, "SOFTWOOD TEXTILES", 0, 1, 'C')
        self.set_font('Arial', 'I', 10)
        self.cell(0, 10, 'Document SW/QA-004 | Sample Evaluation Report', 0, 1, 'C')
        self.ln(15) 
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')


def main_app():
    st.sidebar.title("Navigation")
    mode = st.sidebar.radio("Go to:", ["New Inspection", "Admin (Templates)", "Admin (Customers)"])


    if mode == "Admin (Templates)":
        st.title("🛠️ Template Builder")
        templates_list = db.get_list("Templates")
        template_options = {t['Name']: t for t in templates_list}
        
        with st.expander("➕ Create New Template"):
            new_temp_name = st.text_input("Template Name")
            points_input = st.text_area("Measurement Points (One per line)")
            if st.button("Save Template"):
                if new_temp_name and points_input:
                    point_list = [p.strip() for p in points_input.split('\n') if p.strip()]
                    db.save_entry("Templates", new_temp_name, point_list)
                    st.success("Saved!")
                    st.rerun()
        
        st.subheader("Existing Templates")
        if template_options:
            selected_del = st.selectbox("Delete Template", ["None"] + list(template_options.keys()))
            if selected_del != "None" and st.button("Confirm Delete"):
                db.delete_entry("Templates", template_options[selected_del]['ID'])
                st.success("Deleted.")
                st.rerun()


    elif mode == "Admin (Customers)":
        st.title("👥 Customer & Email Management")
        customers_list = db.get_list("Customers")
        
        with st.expander("➕ Add New Customer List", expanded=True):
            cust_name = st.text_input("Customer Name")
            emails_input = st.text_area("Enter Emails (Comma separated)")
            if st.button("Save Customer"):
                if cust_name and emails_input:
                    email_list = [e.strip() for e in emails_input.replace('\n', ',').split(',') if e.strip()]
                    db.save_entry("Customers", cust_name, email_list)
                    st.success(f"Saved {cust_name}.")
                    st.rerun()

        st.subheader("Existing Customer Lists")
        if customers_list:
            for c in customers_list:
                try:
                    e_list = json.loads(c['Emails_JSON'])
                    with st.expander(f"{c['Customer_Name']} ({len(e_list)} emails)"):
                        st.write(", ".join(e_list))
                        if st.button(f"Delete {c['Customer_Name']}", key=c['ID']):
                            db.delete_entry("Customers", c['ID'])
                            st.rerun()
                except: continue


    elif mode == "New Inspection":
        st.title("📋 New Quality Report")
        
        templates_list = db.get_list("Templates")
        template_map = {t['Name']: json.loads(t['JSON_Data']) for t in templates_list if 'JSON_Data' in t}
        
        customers_list = db.get_list("Customers")
        customer_map = {c['Customer_Name']: json.loads(c['Emails_JSON']) for c in customers_list if 'Emails_JSON' in c}

        reports_list = db.get_list("Reports")
        reports_list.sort(key=lambda x: x['Date_Created'], reverse=True)
        report_options = {f"{r['Name']} ({r['Date_Created']})": r for r in reports_list}
        
        st.sidebar.markdown("---")
        selected_report = st.sidebar.selectbox("📂 Load / Duplicate", ["None"] + list(report_options.keys()))
        
        d_style, d_color, d_stage, d_temp, d_meas = "", "", "PPS", None, {}
        
        if selected_report != "None":
            r_data = json.loads(report_options[selected_report]['JSON_Data'])
            d_style = r_data.get('style', "")
            d_color = r_data.get('color', "")
            d_stage = r_data.get('stage', "PPS")
            d_temp = r_data.get('template_name', "")
            d_meas = r_data.get('measurements', {})
            st.info(f"Loaded: {selected_report}")

        with st.expander("1. Job Details", expanded=True):
            c1, c2, c3 = st.columns(3)
            style = c1.text_input("Style #", value=d_style)
            color = c2.text_input("Color", value=d_color)
            stage = c3.selectbox("Stage", ["Proto", "PPS", "Production"], index=["Proto", "PPS", "Production"].index(d_stage) if d_stage in ["Proto", "PPS", "Production"] else 1)
            c4, c5 = st.columns(2)
            sel_cust = c4.selectbox("Select Customer (For Emailing)", ["None"] + list(customer_map.keys()))
            avail_temps = list(template_map.keys())
            if not avail_temps:
                st.warning("No templates found.")
                st.stop()
            idx = avail_temps.index(d_temp) if d_temp in avail_temps else 0
            sel_temp = c5.selectbox("Template", avail_temps, index=idx)

        st.subheader("2. Measurements")
        current_meas = {}
        report_rows = []
        
        if sel_temp:
            points = template_map[sel_temp]
            cols = st.columns([2, 1, 1, 1, 1, 1, 1])
            headers = ["POM", "Tol", "Std", "S1", "S2", "S3", "Stat"]
            for i, h in enumerate(headers): cols[i].markdown(f"**{h}**")
            
            for p in points:
                v = d_meas.get(p, {})
                c = st.columns([2, 1, 1, 1, 1, 1, 1])
                c[0].write(p)
                tol = c[1].number_input(f"t_{p}", value=v.get('tol', 0.5), step=0.1, label_visibility="collapsed")
                std = c[2].number_input(f"s_{p}", value=v.get('std', 0.0), step=0.5, label_visibility="collapsed")
                s1 = c[3].number_input(f"1_{p}", value=v.get('s1', 0.0), step=0.5, label_visibility="collapsed")
                s2 = c[4].number_input(f"2_{p}", value=v.get('s2', 0.0), step=0.5, label_visibility="collapsed")
                s3 = c[5].number_input(f"3_{p}", value=v.get('s3', 0.0), step=0.5, label_visibility="collapsed")
                
                current_meas[p] = {'tol': tol, 'std': std, 's1': s1, 's2': s2, 's3': s3}
                status = "OK"
                samples = [x for x in [s1, s2, s3] if x != 0]
                if samples and std != 0:
                    for s in samples:
                        if abs(s - std) > tol: status = "FAIL"
                if status == "FAIL": c[6].error("FAIL")
                elif not samples: c[6].write("-")
                else: c[6].success("OK")
                report_rows.append({"p":p, "tol":tol, "std":std, "s1":s1, "s2":s2, "s3":s3, "stat":status})

        st.subheader("3. Actions & Visuals")
        col_img1, col_img2 = st.columns(2)
        f_img = col_img1.file_uploader("Front", type=['jpg','png','jpeg'])
        d_img = col_img2.file_uploader("Defect", type=['jpg','png','jpeg'])
        rem = st.text_area("Remarks")
        dec = st.selectbox("Decision", ["Approved", "Rejected"])
        
        def save_image_safe(uploaded_file, filename):
            try:
                image = Image.open(uploaded_file)
                if image.mode != 'RGB': image = image.convert('RGB')
                image.save(filename, "JPEG")
                return True
            except: return False

        def generate_pdf_bytes():
            pdf = PDF()
            pdf.add_page()
            pdf.set_font("Arial", '', 12)
            pdf.cell(0, 10, f"Style: {style} | {color} | {stage}", 0, 1)
            pdf.cell(0, 10, f"Result: {dec}", 0, 1)
            pdf.ln(5)
            pdf.set_font("Arial", 'B', 8)
            pdf.cell(50, 8, "Point", 1)
            pdf.cell(15, 8, "Tol", 1)
            pdf.cell(15, 8, "Std", 1)
            pdf.cell(15, 8, "S1", 1)
            pdf.cell(15, 8, "S2", 1)
            pdf.cell(15, 8, "S3", 1)
            pdf.cell(20, 8, "Status", 1)
            pdf.ln()
            pdf.set_font("Arial", '', 8)
            for r in report_rows:
                if r['stat'] == "FAIL": pdf.set_text_color(255,0,0)
                else: pdf.set_text_color(0,0,0)
                pdf.cell(50, 8, r['p'], 1)
                pdf.cell(15, 8, str(r['tol']), 1)
                pdf.cell(15, 8, str(r['std']), 1)
                pdf.cell(15, 8, str(r['s1']), 1)
                pdf.cell(15, 8, str(r['s2']), 1)
                pdf.cell(15, 8, str(r['s3']), 1)
                pdf.cell(20, 8, r['stat'], 1)
                pdf.ln()
            pdf.set_text_color(0,0,0)
            pdf.ln(5)
            pdf.multi_cell(0, 10, f"Remarks: {rem}")
            if f_img:
                pdf.add_page()
                pdf.cell(0, 10, "Front View", 0, 1)
                if save_image_safe(f_img, "temp_f.jpg"): pdf.image("temp_f.jpg", x=10, y=40, w=90)
            if d_img:
                if not f_img: 
                    pdf.add_page()
                    pdf.cell(0, 10, "Defect View", 0, 1)
                x_pos = 110 if f_img else 10
                if save_image_safe(d_img, "temp_d.jpg"): pdf.image("temp_d.jpg", x=x_pos, y=40, w=90)
            return pdf.output(dest='S').encode('latin-1')

        col_act1, col_act2, col_act3 = st.columns(3)
        with col_act1:
            if st.button("💾 Save to Cloud"):
                if not style: st.error("Style # required")
                else:
                    save_name = f"{style} - {color}"
                    save_data = {"style": style, "color": color, "stage": stage, "template_name": sel_temp, "measurements": current_meas}
                    with st.spinner("Saving..."):
                        db.save_entry("Reports", save_name, save_data)
                    st.success("Saved!")
        with col_act2:
            if st.button("📄 Preview/Download PDF"):
                pdf_bytes = generate_pdf_bytes()
                st.download_button("Click to Download", pdf_bytes, f"{style}.pdf", "application/pdf")
        with col_act3:
            if st.button("📧 Send Email to Customer", type="primary"):
                if sel_cust == "None": st.error("Please select a Customer first!")
                else:
                    recipients = customer_map[sel_cust]
                    pdf_bytes = generate_pdf_bytes()
                    email_body = f"Dear Team,\n\nPlease find attached the QA Report for Style: {style}, Color: {color}.\n\nResult: {dec}\n\nBest Regards,\nSoftwood QA Team"
                    with st.spinner("Sending Email..."):
                        success, msg = send_email_with_pdf(recipients, f"QA Report: {style} - {color}", email_body, pdf_bytes, f"{style}.pdf")
                    if success: st.success("Email Sent Successfully!")
                    else: st.error(f"Failed: {msg}")

if __name__ == "__main__":
    main_app()