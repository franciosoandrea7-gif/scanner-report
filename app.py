import streamlit as st
import pandas as pd
import os
import random
import requests
from PIL import Image
from datetime import datetime

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni Nova Servimpianti con validazione SMS OTP ed invio Email.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  

# --- INIZIALIZZAZIONE STATI DI SESSIONE ---
if "codice_sms" not in st.session_state:
    st.session_state["codice_sms"] = None
if "sms_validato" not in st.session_state:
    st.session_state["sms_validato"] = False

# --- 1. MODULO DI INPUT DATI ---
with st.expander("👤 Dati Cliente & Macchina", expanded=True):
    data_corrente = st.date_input("Data Intervento", datetime.now())
    cliente = st.text_input("Ragione Sociale Cliente *")
    email_cliente = st.text_input("Email Cliente (Per invio copia PDF) *")
    cellulare_cliente = st.text_input("Numero di Cellulare Cliente (Per SMS OTP) *", placeholder="Es: 3471234567")
    marchio = st.text_input("Marchio Apparecchio *")
    matricola = st.text_input("Matricola Apparecchio")

with st.expander("📝 Stato della Riparazione", expanded=True):
    guasto_segnalato = st.text_area("Guasto Segnalato")
    descrizione_lavori = st.text_area("Intervento Eseguito / Note Tecniche *")

with st.expander("📊 Costi e Tempistiche (Facoltativi)"):
    km = st.number_input("Kilometri percorsi (Km)", min_value=0, value=0, step=1)
    ore_lavoro = st.number_input("Ore di lavoro impiegate", min_value=0.0, value=0.0, step=0.5)
    col1, col2 = st.columns(2)
    with col1:
        preventivo = st.radio("Richiedi Preventivo?", ["NO", "SI"])
    with col2:
        urgente = st.radio("Intervento Urgente?", ["NO", "SI"])

st.subheader("📸 Foto Scheda Firmata (Opzionale)")
file_immagine = st.camera_input("Scatta la foto alla scheda cartacea se presente")

# --- 2. FUNZIONE SPEDIZIONE SMS OTP ---
def invia_sms_otp_reale(numero_telefono, codice):
    if not numero_telefono.startswith("+"):
        if numero_telefono.startswith("39"):
            numero_telefono = "+" + numero_telefono
        else:
            numero_telefono = "+39" + numero_telefono
    testo_sms = f"Nova Servimpianti: Il tuo codice segreto di firma per l'intervento odierno e': {codice}"
    try:
        response = requests.post('https://textbelt.com', {
            'phone': numero_telefono, 'message': testo_sms, 'key': 'textbelt'
        })
        return response.json().get("success", False)
    except:
        return False

# --- 3. SEZIONE CONTROLLO SICUREZZA SMS ---
st.subheader("🔒 Firma Digitale SMS OTP del Cliente")
st.write("Spedisci un codice usa-e-getta sul cellulare del cliente per fargli convalidare il verbale sul posto:")

if st.button("📲 INVIA CODICE DI VALIDAZIONE VIA SMS"):
    if not cellulare_cliente or not cliente:
        st.error("⚠️ Inserisci la Ragione Sociale e il Numero di Cellulare del cliente prima di procedere!")
    else:
        st.session_state["codice_sms"] = str(random.randint(1000, 9999))
        st.session_state["sms_validato"] = False
        esito = invia_sms_otp_reale(cellulare_cliente, st.session_state["codice_sms"])
        
        if esito:
            st.success(f"📩 Richiesta SMS inoltrata con successo al numero: {cellulare_cliente}!")
        else:
            st.warning("⚠️ Quota SMS gratuita esaurita o errore di rete. Usa il codice di sblocco d'emergenza mostrato qui sotto.")

if st.session_state["codice_sms"] is not None:
    st.info(f"👉 CODICE DA INSERIRE PER CONVALIDARE: {st.session_state['codice_sms']}")
    
    if not st.session_state["sms_validato"]:
        codice_inserito = st.text_input("Inserisci le 4 cifre visualizzate qui sopra per firmare:")
        if st.button("✅ VALIDA CODICE SMS"):
            if codice_inserito == st.session_state["codice_sms"]:
                st.session_state["sms_validato"] = True
                st.success("🔒 Documento Firmato e Convalidato con Successo!")
                st.rerun()
            else:
                st.error("❌ Codice errato! Riprova.")
    else:
        st.success(f"🔒 Documento già Convalidato! (Codice: {st.session_state['codice_sms']})")

if st.session_state["sms_validato"]:
    st.info("ℹ️ Stato firma: OK. Ora puoi procedere al salvataggio finale del report in fondo alla pagina.")

# --- 4. LOGICA INVIO COPIA COMPLETA VIA EMAIL ---
def invia_email_pdf(destinatario, allegato_path, nome_cliente):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    
    email_mittente = "franciosoandrea@gmail.com" 
    password_mittente = "qiad bvqq ijaj mutc "  
    
    msg = MIMEMultipart()
    msg['From'] = email_mittente
    msg['To'] = destinatario
    msg['Subject'] = f"Rapporto Intervento Tecnico - {nome_cliente}"
    msg.attach(MIMEText("Buongiorno,\nin allegato copia del rapporto ufficiale Nova Servimpianti.\n\nCordiali Saluti.", 'plain'))
    
    try:
        with open(allegato_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {allegato_path}")
            msg.attach(part)
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(email_mittente, password_mittente)
        server.sendmail(email_mittente, destinatario, msg.as_string())
        server.quit()
        st.success("✉️ Copia del report inviata correttamente all'email del cliente!")
    except Exception as e:
        st.warning(f"⚠️ Nota: Dati registrati, ma e-mail non partita automaticamente. Errore: {e}")

# --- 5. CREAZIONE PDF ---
def elabora_pdf(pdf_filename, data_str, cliente, email_cliente, cellulare_cliente, marchio, matricola, km, ore_lavoro, preventivo, urgente, guasto_segnalato, descrizione_lavori, file_immagine, stringa_firma):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    
    try:
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
            try:
                story.append(Spacer(1, 25))
                story.append(Paragraph("<b>■ ALLEGATO FOTO SCHEDA</b>", section_heading))
                foto_img = Image.open(file_immagine).convert("RGB")
                foto_path = "temp_allegato.png"
                foto_img.thumbnail((500, 450))
                foto_img.save(foto_path)
                story.append(RLImage(foto_path, width=450, height=350))
            except Exception as img_err:
                st.warning(f"Impossibile allegare la foto al PDF: {img_err}")
                
        doc.build(story)
    except Exception as e:
        st.error(f"Errore generazione PDF: {e}")

# --- 6. FUNZIONE GENERALE DI SCRITTURA DATI (BLINDATA) ---
def registra_dati_intervento(data_str, cliente, email_cliente, cellulare_cliente, marchio, matricola, guasto_segnalato, descrizione_lavori, km, ore_lavoro, preventivo, urgente, stringa_firma):
    riga = {
        "Data": data_str,
        "Cliente": cliente,
        "Email": email_cliente,
        "Cellulare": cellulare_cliente,
        "Marchio": marchio,
        "Matricola": matricola if matricola else "N.D.",
        "Guasto": guasto_segnalato if guasto_segnalato else "N.D.",
        "Intervento": descrizione_lavori,
        "Km": km,
        "Ore": ore_lavoro,
        "Preventivo": preventivo,
        "Urgente": urgente,
        "Firma": stringa_firma
    }
