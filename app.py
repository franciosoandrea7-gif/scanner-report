import streamlit as st
import pandas as pd
import os
import random
from PIL import Image
from datetime import datetime

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni Nova Servimpianti con validazione forte Email OTP.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  

# Inizializzazione variabili per il codice segreto OTP
if "codice_generato" not in st.session_state:
    st.session_state["codice_generato"] = None
if "otp_validato" not in st.session_state:
    st.session_state["otp_validato"] = False

# --- 1. MODULO DI INPUT DATI ---
with st.expander("👤 Dati Cliente & Macchina", expanded=True):
    data_corrente = st.date_input("Data Intervento", datetime.now())
    cliente = st.text_input("Ragione Sociale Cliente *")
    email_cliente = st.text_input("Email Cliente (Obbligatoria per Firma OTP) *")
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
        preventivo = st.radio("Richiega Preventivo?", ["NO", "SI"])
    with col2:
        urgente = st.radio("Intervento Urgente?", ["NO", "SI"])

st.subheader("📸 Foto Scheda Firmata (Opzionale)")
file_immagine = st.camera_input("Scatta la foto alla scheda cartacea se presente")

# --- 2. FUNZIONE SPEDIZIONE CODICE SEGRETO ---
def invia_codice_otp(destinatario, codice):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    
    email_mittente = "franciosoandrea@gmail.com" 
    password_mittente = "la-tua-password-di-16-lettere-di-google"  # <--- METTI LA TUA PASSWORD DI GOOGLE QUI!
    
    msg = MIMEMultipart()
    msg['From'] = email_mittente
    msg['To'] = destinatario
    msg['Subject'] = f"Codice di Sicurezza Intervento - Nova Servimpianti"
    
    corpo = f"Buongiorno,\nil codice segreto temporaneo per firmare e convalidare il rapporto tecnico dell'intervento odierno e':\n\n【 {codice} 】\n\nComunichi questo codice al tecnico per sigillare il documento."
    msg.attach(MIMEText(corpo, 'plain'))
    
    try:
        server = smtplib.SMTP("://gmail.com", 587)
        server.starttls()
        server.login(email_mittente, password_mittente)
        server.sendmail(email_mittente, destinatario, msg.as_string())
        server.quit()
        return True
    except:
        return False

# --- 3. SEZIONE COSTRUZIONE SICUREZZA OTP ---
st.subheader("🔒 Firma Digitale OTP del Cliente")
st.write("Invia un codice di sicurezza usa-e-getta all'email del cliente per fargli firmare il verbale:")

if st.button("📲 INVIA CODICE DI SIGN ALLA MAIL DEL CLIENTE"):
    if not email_cliente or not cliente:
        st.error("⚠️ Inserisci la Ragione Sociale e l'Email del cliente prima di inviare il codice!")
    else:
        # Genera numero casuale a 4 cifre
        st.session_state["codice_generato"] = str(random.randint(1000, 9999))
        st.session_state["otp_validato"] = False
        
        with st.spinner("Invio del codice segreto in corso..."):
            esito = invia_codice_otp(email_cliente, st.session_state["codice_generato"])
            if esito:
                st.success(f"📩 Codice inviato con successo alla casella: {email_cliente}!")
            else:
                st.error("❌ Impossibile inviare l'email. Controlla la password a 16 lettere inserita nel codice.")

# Box per inserire il codice dettato dal cliente
if st.session_state["codice_generato"] is not None and not st.session_state["otp_validato"]:
    codice_inserito = st.text_input("Inserisci qui il codice a 4 cifre che il cliente ha ricevuto via e-mail:")
    
    if st.button("✅ VALIDA CODICE"):
        if codice_inserito == st.session_state["codice_generato"]:
            st.session_state["otp_validato"] = True
            st.success("🔒 Documento Firmato e Validato Digitalmente con Successo!")
        else:
            st.error("❌ Codice errato! Riprova o richiedi un nuovo invio.")

# --- 4. LOGICA INVIO REPORT FINALE ---
def invia_email_pdf(destinatario, allegato_path, nome_cliente):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    email_mittente = "franciosoandrea@gmail.com" 
    password_mittente = "qiad bvqq ijaj mutc"  # <--- METTI LA TUA PASSWORD DI GOOGLE QUI!
    msg = MIMEMultipart()
    msg['From'] = email_mittente
    msg['To'] = destinatario
    msg['Subject'] = f"Rapporto Intervento Tecnico - {nome_cliente}"
    msg.attach(MIMEText("Buongiorno,\nin allegato il report ufficiale dell'intervento.", 'plain'))
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
    except:
        pass

# --- 5. CREAZIONE PDF ---
def elabora_pdf(pdf_filename, data_str, cliente, email_cliente, marchio, matricola, km, ore_lavoro, preventivo, urgente, guasto_segnalato, descrizione_lavori, file_immagine, stringa_firma):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
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
    story.append(Paragraph(f"<b>Data Intervento:</b> {data_str}<br/><b>Cliente:</b> {cliente}<br/><b>Email:</b> {email_cliente}<br/><b>Marchio:</b> {marchio}<br/><b>Matricola:</b> {matricola if matricola else 'N.D.'}<br/><b>Km:</b> {km} Km | <b>Ore:</b> {ore_lavoro}<br/><b>Preventivo:</b> {preventivo} | <b>Urgente:</b> {urgente}", body_style))
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>■ GUASTO SEGNALATO</b>", section_heading))
    story.append(Paragraph(guasto_segnalato if guasto_segnalato else "N.D.", body_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>■ DETTAGLIO LAVORI ESEGUITI</b>", section_heading))
    story.append(Paragraph(descrizione_lavori, body_style))
    story.append(Spacer(1, 25))
    story.append(Paragraph("<b>Firma del Tecnico:</b><br/><br/><br/>___________________________", body_style))
    story.append(Spacer(1, 40)) 
    story.append(Paragraph("<b>Firma per Accettazione Cliente (Validazione OTP):</b>", body_style))
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

# --- 6. BOTTONE DI SALVATAGGIO ---
if st.button("💾 REGISTRA E GENERA REPORT COMPLETO"):
    if not cliente or not marchio or not descrizione_lavori:
        st.error("⚠️ Compila i campi obbligatori (*)!")
    elif not st.session_state["otp_validato"]:
        st.error("⚠️ Il documento non puo' essere salvato senza la firma OTP valida del cliente!")
    else:
        with st.spinner("Salvataggio in corso..."):
            data_str = data_corrente.strftime("%d/%m/%Y")
            stringa_firma = f"Firmato digitalmente tramite validazione forte Email OTP in data {data_str} con codice ID-{st.session_state['codice_generato']}"
            
            riga = {"Data": data_str, "Cliente": cliente, "Email": email_cliente, "Marchio": marchio, "Matricola": matricola if matricola else "N.D.", "Guasto": guasto_segnalato if guasto_segnalato else "N.D.", "Intervento": descrizione_lavori, "Km": km, "Ore": ore_lavoro, "Preventivo": preventivo, "Urgente": urgente, "Firma": stringa_firma}
            
            if os.path.exists(EXCEL_FILE):
                df = pd.concat([pd.read_excel(EXCEL_FILE), pd.DataFrame([riga])], ignore_index=True)
            else:
                df = pd.DataFrame([riga])
            df.to_excel(EXCEL_FILE, index=False)
            
