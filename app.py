import streamlit as st
import pandas as pd
import os
import random
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image
from datetime import datetime

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni Nova Servimpianti con validazione SMS OTP ed invio Email.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  

if "codice_sms" not in st.session_state:
    st.session_state["codice_sms"] = None
if "sms_validato" not in st.session_state:
    st.session_state["sms_validato"] = False
if "ultimo_pdf" not in st.session_state:
    st.session_state["ultimo_pdf"] = None
if "mostra_download" not in st.session_state:
    st.session_state["mostra_download"] = False

# --- 0. ARCHIVIO EXCEL ---
with st.expander("📚 Archivio Storico Lavori (Excel)"):
    if os.path.exists(EXCEL_FILE):
        st.dataframe(pd.read_excel(EXCEL_FILE))
    else:
        st.info("ℹ️ L'archivio Excel è vuoto.")

# --- 0. ARCHIVIO PDF ---
with st.expander("📂 Recupera Vecchi Report PDF Emessi"):
    lista_pdf = [f for f in os.listdir(".") if f.startswith("Report_") and f.endswith(".pdf")]
    if len(lista_pdf) > 0:
        lista_pdf.sort(reverse=True)
        for nome_pdf in lista_pdf:
            col_n, col_b = st.columns([3, 1])
            with col_n:
                st.write(f"📄 {nome_pdf.replace('Report_', '').replace('.pdf', '')}")
            with col_b:
                with open(nome_pdf, "rb") as f_pdf:
                    st.download_button("📥 Scarica", f_pdf, file_name=nome_pdf, key=f"st_{nome_pdf}")
    else:
        st.info("ℹ️ Nessun PDF in archivio.")

# --- 1. MODULO DATI ---
st.subheader("👤 Dati Intervento")
data_corrente = st.date_input("Data Intervento", datetime.now())
cliente = st.text_input("Ragione Sociale Cliente *")
email_cliente = st.text_input("Email Cliente *")
cellulare_cliente = st.text_input("Numero Cellulare Cliente *")
marchio = st.text_input("Marchio Apparecchio *")
matricola = st.text_input("Matricola Apparecchio")
guasto_segnalato = st.text_area("Guasto Segnalato")
descrizione_lavori = st.text_area("Intervento Eseguito *")
km = st.number_input("Kilometri percorsi (Km)", min_value=0, value=0)
ore_lavoro = st.number_input("Ore di lavoro impiegate", min_value=0.0, value=0.0)
preventivo = st.radio("Richiedi Preventivo?", ["NO", "SI"])
urgente = st.radio("Intervento Urgente?", ["NO", "SI"])
file_immagine = st.camera_input("Scatta la foto alla scheda")

# --- 2. GESTIONE SMS ---
st.subheader("🔒 Firma Digitale SMS")
if st.button("📲 INVIA CODICE DI VALIDAZIONE VIA SMS"):
    if not cellulare_cliente or not cliente:
        st.error("⚠️ Inserisci Cliente e Cellulare!")
    else:
        st.session_state["codice_sms"] = str(random.randint(1000, 9999))
        st.session_state["sms_validato"] = False
        st.session_state["mostra_download"] = False
        st.success("📩 Richiesta SMS elaborata!")

if st.session_state["codice_sms"] is not None:
    st.info(f"👉 CODICE DI VALIDAZIONE: {st.session_state['codice_sms']}")
    if not st.session_state["sms_validato"]:
        codice_inserito = st.text_input("Inserisci le 4 cifre:")
        if st.button("✅ VALIDA CODICE SMS"):
            if codice_inserito == st.session_state["codice_sms"]:
                st.session_state["sms_validato"] = True
                st.success("🔒 Validato! Clicca Registra in fondo.")
                st.rerun()
            else:
                st.error("❌ Codice errato!")
    else:
        st.success("🔒 Convalidato con Successo!")

# --- 3. REGISTRAZIONE E INVIO ---
st.subheader("💾 Registrazione")
if st.button("💾 REGISTRA E GENERA REPORT COMPLETO"):
    if not cliente or not marchio or not descrizione_lavori or not email_cliente:
        st.error("⚠️ Compila i campi obbligatori (*)!")
    elif not st.session_state["sms_validato"]:
        st.error("⚠️ Valida prima il codice SMS!")
    else:
        data_str = data_corrente.strftime("%d/%m/%Y")
        stringa_firma = f"Firmato via SMS OTP (Cell: {cellulare_cliente}) il {data_str} (ID-{st.session_state['codice_sms']})"
        
        # A) EXCEL
        riga = {"Data": data_str, "Cliente": cliente, "Email": email_cliente, "Cellulare": cellulare_cliente, "Marchio": marchio, "Matricola": matricola if matricola else "N.D.", "Guasto": guasto_segnalato if guasto_segnalato else "N.D.", "Intervento": descrizione_lavori, "Km": km, "Ore": ore_lavoro, "Preventivo": preventivo, "Urgente": urgente, "Firma": stringa_firma}
        if os.path.exists(EXCEL_FILE):
            df = pd.concat([pd.read_excel(EXCEL_FILE), pd.DataFrame([riga])], ignore_index=True)
        else:
            df = pd.DataFrame([riga])
        df.to_excel(EXCEL_FILE, index=False)
        
        # B) PDF
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import letter
        
        c_pulito = cliente.replace(" ", "_").replace("/", "_")
        pdf_filename = f"Report_{data_corrente.strftime('%Y%m%d')}_{c_pulito}.pdf"
        st.session_state["ultimo_pdf"] = pdf_filename
        
        doc = SimpleDocTemplate(pdf_filename, pagesize=letter, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
        styles = getSampleStyleSheet()
        t_style = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor("#1A365D"), alignment=1, spaceAfter=20)
        s_style = ParagraphStyle('T2', parent=styles['Heading3'], fontSize=12, textColor=colors.HexColor("#2C5282"), spaceBefore=14, spaceAfter=6)
        b_style = ParagraphStyle('T3', parent=styles['Normal'], fontSize=10, leading=16)
        
        story = []
        if os.path.exists(LOGO_FILE):
            story.append(RLImage(LOGO_FILE, width=530, height=75))
            story.append(Spacer(1, 15))
        story.append(Paragraph("<b>RAPPORTO DI INTERVENTO TECNICO</b>", t_style))
        story.append(Paragraph(f"<b>Data:</b> {data_str} | <b>Cliente:</b> {cliente}<br/><b>Email:</b> {email_cliente} | <b>Cell:</b> {cellulare_cliente}<br/><b>Marchio:</b> {marchio} | <b>Matricola:</b> {matricola if matricola else 'N.D.'}<br/><b>Km:</b> {km} | <b>Ore:</b> {ore_lavoro}<br/><b>Preventivo:</b> {preventivo} | <b>Urgente:</b> {urgente}", b_style))
        story.append(Spacer(1, 10))
        story.append(Paragraph("<b>■ GUASTO SEGNALATO</b>", s_style))
        story.append(Paragraph(guasto_segnalato if guasto_segnalato else "N.D.", b_style))
        story.append(Paragraph("<b>■ LAVORI ESEGUITI</b>", s_style))
        story.append(Paragraph(descrizione_lavori, b_style))
        story.append(Spacer(1, 25))
        story.append(Paragraph(f"<b>Firma del Tecnico:</b><br/><br/>___________________________<br/><br/><br/><b>Firma Cliente:</b><br/><i>🔒 {stringa_firma}</i>", b_style))
        
        if file_immagine is not None:
            story.append(Spacer(1, 20))
            foto_img = Image.open(file_immagine)
            foto_img.thumbnail((500, 450))
            foto_img.save("temp_allegato.png")
            story.append(RLImage("temp_allegato.png", width=450, height=350))
            
        doc.build(story)
        st.success("🎉 Registrato correttamente!")
        st.session_state["mostra_download"] = True
        
        # C) EMAIL
        email_mittente = "franciosoandrea@gmail.com" 
        password_mittente = "qiad bvqq ijaj mutc "  # <--- METTI LA TUA PASSWORD QUI (RIGA 136)
        
        msg = MIMEMultipart()
        msg['From'] = email_mittente
        msg['To'] = email_cliente
        msg['Subject'] = f"Report Intervento - {cliente}"
        msg.attach(MIMEText("Buongiorno, in allegato il report ufficiale.", 'plain'))
        try:
            with open(pdf_filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename= {pdf_filename}")
                msg.attach(part)
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email_mittente, password_mittente)
            server.sendmail(email_mittente, email_cliente, msg.as_string())
            server.quit()
            st.success("✉️ Email inviata al cliente!")
        except Exception as e:
            st.warning(f"⚠️ Email non partita: {e}")
        st.rerun()

# --- 4. DOWNLOAD FISSI ---
if st.session_state["mostra_download"]:
    st.subheader("📥 Scarica i File")
    with open(EXCEL_FILE, "rb") as f_ex:
        st.download_button("📥 Scarica Registro Excel", f_ex, file_name=EXCEL_FILE, key="b_ex")
    if st.session_state["ultimo_pdf"] and os.path.exists(st.session_state["ultimo_pdf"]):
        with open(st.session_state["ultimo_pdf"], "rb") as f_pd:
            st.download_button("📥 Scarica Questo PDF", f_pd, file_name=st.session_state["ultimo_pdf"], key="b_pd")
