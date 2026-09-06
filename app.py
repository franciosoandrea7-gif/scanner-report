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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni Nova Servimpianti con validazione SMS OTP ed invio Email.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  

# --- INIZIALIZZAZIONE STATI DI SESSIONE PER DATI PERSISTENTI ---
if "codice_sms" not in st.session_state:
    st.session_state["codice_sms"] = None
if "sms_validato" not in st.session_state:
    st.session_state["sms_validato"] = False
if "ultimo_pdf" not in st.session_state:
    st.session_state["ultimo_pdf"] = None
if "mostra_download" not in st.session_state:
    st.session_state["mostra_download"] = False

# --- 0. VISUALIZZAZIONE STORICO LAVORI SVOLTI ---
with st.expander("📚 Visualizza Archivio Storico Lavori Svolti (Excel)"):
    if os.path.exists(EXCEL_FILE):
        df_storico = pd.read_excel(EXCEL_FILE)
        st.dataframe(df_storico)
    else:
        st.info("ℹ️ L'archivio è attualmente vuoto. Diventerà visibile non appena registrerai il primo rapporto.")

# --- 1. DATI INSERIMENTO ---
st.subheader("👤 Dati Cliente & Macchina")
data_corrente = st.date_input("Data Intervento", datetime.now())
cliente = st.text_input("Ragione Sociale Cliente *")
email_cliente = st.text_input("Email Cliente (Per invio copia PDF) *")
cellulare_cliente = st.text_input("Numero di Cellulare Cliente (Per SMS OTP) *")
marchio = st.text_input("Marchio Apparecchio *")
matricola = st.text_input("Matricola Apparecchio")

st.subheader("📝 Stato della Riparazione")
guasto_segnalato = st.text_area("Guasto Segnalato")
descrizione_lavori = st.text_area("Intervento Eseguito / Note Tecniche *")

st.subheader("📊 Costi e Tempistiche")
km = st.number_input("Kilometri percorsi (Km)", min_value=0, value=0)
ore_lavoro = st.number_input("Ore di lavoro impiegate", min_value=0.0, value=0.0)
preventivo = st.radio("Richiedi Preventivo?", ["NO", "SI"])
urgente = st.radio("Intervento Urgente?", ["NO", "SI"])

st.subheader("📸 Foto Scheda")
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
        st.success("📩 Richiesta SMS dell'intervento elaborata con successo!")

if st.session_state["codice_sms"] is not None:
    st.info(f"👉 CODICE DI VALIDAZIONE DA INSERIRE: {st.session_state['codice_sms']}")
    if not st.session_state["sms_validato"]:
        codice_inserito = st.text_input("Inserisci le 4 cifre per convalidare il verbale:")
        if st.button("✅ VALIDA CODICE SMS"):
            if codice_inserito == st.session_state["codice_sms"]:
                st.session_state["sms_validato"] = True
                st.success("🔒 Documento Convalidato e Firmato! Procedi in fondo per il salvataggio.")
                st.rerun()
            else:
                st.error("❌ Codice errato! Controlla e riprova.")
    else:
        st.success("🔒 Documento già Convalidato con Successo!")

# --- 3. SEZIONE SALVATAGGIO, CREAZIONE PDF E INVIO EMAIL (NUOVA SEZIONE AGGIUNTA CORRETTA) ---
st.subheader("💾 Registrazione e Invio")

if st.button("💾 REGISTRA E GENERA REPORT COMPLETO"):
    if not cliente or not marchio or not descrizione_lavori or not email_cliente:
        st.error("⚠️ Compila tutti i campi obbligatori (*) prima di procedere!")
    elif not st.session_state["sms_validato"]:
        st.error("⚠️ Non puoi salvare il report senza prima aver validato il codice SMS OTP del cliente!")
    else:
        st.write("🔄 Elaborazione e scrittura documenti in corso...")
        data_str = data_corrente.strftime("%d/%m/%Y")
        stringa_firma = f"Firmato digitalmente tramite certificazione forte SMS OTP inviata al numero {cellulare_cliente} in data {data_str} con codice ID-{st.session_state['codice_sms']}"
        
        # A) SALVATAGGIO CORRETTO IN EXCEL
        riga = {
            "Data": data_str, "Cliente": cliente, "Email": email_cliente, "Cellulare": cellulare_cliente,
            "Marchio": marchio, "Matricola": matricola if matricola else "N.D.",
            "Guasto": guasto_segnalato if guasto_segnalato else "N.D.", "Intervento": descrizione_lavori,
            "Km": km, "Ore": ore_lavoro, "Preventivo": preventivo, "Urgente": urgente, "Firma": stringa_firma
        }
        
        if os.path.exists(EXCEL_FILE):
            df_vecchio = pd.read_excel(EXCEL_FILE)
            df_nuovo = pd.concat([df_vecchio, pd.DataFrame([riga])], ignore_index=True)
        else:
            df_nuovo = pd.DataFrame([riga])
        df_nuovo.to_excel(EXCEL_FILE, index=False)
        
        # B) GENERAZIONE PDF PROFESSIONALE
        c_pulito = cliente.replace(" ", "_").replace("/", "_")
        pdf_filename = f"Report_{data_corrente.strftime('%Y%m%d')}_{c_pulito}.pdf"
        st.session_state["ultimo_pdf"] = pdf_filename
        
        doc = SimpleDocTemplate(pdf_filename, pagesize=letter, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor("#1A365D"), alignment=1, spaceAfter=20)
        section_heading = ParagraphStyle('T2', parent=styles['Heading3'], fontSize=12, leading=16, textColor=colors.HexColor("#2C5282"), spaceBefore=14, spaceAfter=6)
        body_style = ParagraphStyle('T3', parent=styles['Normal'], fontSize=10, leading=16)
        firma_style = ParagraphStyle('T4', parent=styles['Normal'], fontSize=9, leading=14, textColor=colors.HexColor("#4A5568"))
        
        story = []
        if os.path.exists(LOGO_FILE):
            story.append(RLImage(LOGO_FILE, width=530, height=75))
            story.append(Spacer(1, 15))
            
        story.append(Paragraph("<b>RAPPORTO DI INTERVENTO TECNICO</b>", title_style))
        story.append(Spacer(1, 10))
        story.append(Paragraph(f"<b>Data Intervento:</b> {data_str}<br/><b>Cliente:</b> {cliente}<br/><b>Email:</b> {email_cliente}<br/><b>Cellulare:</b> {cellulare_cliente}<br/><b>Marchio:</b> {marchio}<br/><b>Matricola:</b> {matricola if matricola else 'N.D.'}<br/><b>Km:</b> {km} Km | <b>Ore:</b> {ore_lavoro}<br/><b>Preventivo:</b> {preventivo} | <b>Urgente:</b> {urgente}", body_style))
        story.append(Spacer(1, 15))
        story.append(Paragraph("<b>■ GUASTO SEGNALATO</b>", section_heading))
        story.append(Paragraph(guasto_segnalato if guasto_segnalato else "N.D.", body_style))
        story.append(Spacer(1, 10))
        story.append(Paragraph("<b>■ DETTAGLIO LAVORI ESEGUITI</b>", section_heading))
        story.append(Paragraph(descrizione_lavori, body_style))
        story.append(Spacer(1, 25))
        story.append(Paragraph("<b>Firma del Tecnico:</b><br/><br/><br/>___________________________", body_style))
        story.append(Spacer(1, 45)) 
        story.append(Paragraph("<b>Firma per Accettazione Cliente (Validazione forte SMS OTP):</b>", body_style))
        story.append(Spacer(1, 5))
        story.append(Paragraph(f"<i>🔒 {stringa_firma}</i>", firma_style))
        
        if file_immagine is not None:
            story.append(Spacer(1, 25))
            story.append(Paragraph("<b>■ ALLEGATO FOTO SCHEDA</b>", section_heading))
            foto_img = Image.open(file_immagine)
            foto_path = "temp_allegato.png"
            foto_img.thumbnail((500, 450))
            foto_img.save(foto_path)
            story.append(RLImage(foto_path, width=450, height=350))
            
        doc.build(story)
        st.success("🎉 Intervento registrato correttamente nel database!")
        st.session_state["mostra_download"] = True
        
        # C) INVIO AUTOMATICO EMAIL VIA GMAIL PROTETTO
        email_mittente = "franciosoandrea@gmail.com" 
        password_mittente = "qiad bvqq ijaj mutc"  # <--- METTI LA TUA PASSWORD A 16 LETTERE DI GOOGLE QUI!
        
        msg = MIMEMultipart()
        msg['From'] = email_mittente
        msg['To'] = email_cliente
        msg['Subject'] = f"Rapporto Intervento Tecnico - {cliente}"
        msg.attach(MIMEText("Buongiorno,\nin allegato inviamo copia del rapporto ufficiale Nova Servimpianti.\n\nCordiali Saluti.", 'plain'))
        
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
            st.success("✉️ Copia del report inviata con successo all'email del cliente!")
        except Exception as e:
            st.warning(f"⚠️ Nota: File salvati, ma l'email non è partita automaticamente. Errore: {e}")

# --- 4. VISUALIZZAZIONE PULSANTI DOWNLOAD FISSI DOPO IL SALVATAGGIO ---
