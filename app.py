import streamlit as st
import pandas as pd
import os
import base64
from PIL import Image
from datetime import datetime
import io

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni Nova Servimpianti con firma ed invio PDF.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  

# --- SEZIONE 1: MODULO DI INPUT DATI ---
with st.expander("👤 Dati Cliente & Macchina", expanded=True):
    data_corrente = st.date_input("Data Intervento", datetime.now())
    cliente = st.text_input("Ragione Sociale Cliente *")
    email_cliente = st.text_input("Email Cliente (Per invio verbale)")
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

# --- SEZIONE 2: SCHERMO BIANCO NATIVIO PER FIRMA TOUCH ---
st.subheader("✍️ Firma Digitale del Cliente")
st.write("Fai firmare il cliente con il dito nel riquadro bianco qui sotto:")

# Componente HTML + Javascript per la firma touch a schermo intero o parziale
canvas_html = """
<div style="border:2px solid #CBD5E0; border-radius:8px; background-color:#ffffff; padding:5px;">
    <canvas id="signature-pad" width="450" height="150" style="width:100%; height:150px; cursor:crosshair; touch-action:none;"></canvas>
</div>
<div style="margin-top:10px;">
    <button id="clear-btn" style="background-color:#E2E8F0; border:none; padding:8px 15px; border-radius:5px; font-weight:bold; cursor:pointer;">Cancella Firma</button>
</div>

<script src="https://jsdelivr.net"></script>
<script>
    const canvas = document.getElementById('signature-pad');
    const signaturePad = new SignaturePad(canvas, {
        backgroundColor: 'rgb(255, 255, 255)',
        penColor: 'rgb(0, 0, 0)'
    });
    
    document.getElementById('clear-btn').addEventListener('click', () => {
        signaturePad.clear();
        window.parent.postMessage({type: 'streamlit:setComponentValue', value: ''}, '*');
    });

    canvas.addEventListener('touchend', sendData);
    canvas.addEventListener('mouseup', sendData);

    function sendData() {
        if (!signaturePad.isEmpty()) {
            const dataUrl = signaturePad.toDataURL('image/png');
            window.parent.postMessage({type: 'streamlit:setComponentValue', value: dataUrl}, '*');
        }
    }
</script>
"""

# Visualizza lo schermo bianco per firmare raccogliendo il risultato in Streamlit
import streamlit.components.v1 as components
firma_base64 = components.html(canvas_html, height=210)

# Tracciamento interno del valore della firma
if "valore_firma" not in st.session_state:
    st.session_state["valore_firma"] = None

# Messaggio di stato firma
st.info("💡 Fai firmare sul display prima di premere il tasto Registra in fondo.")

# --- SEZIONE 3: FUNZIONE INVIO EMAIL ---
def invia_email_pdf(destinatario, allegato_path, nome_cliente):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    email_mittente = "franciosoandrea@gmail.com" 
    password_mittente = "qiad bvqq ijaj mutc"  # <--- INSERISCI LA TUA PASSWORD QUI
    
    msg = MIMEMultipart()
    msg['From'] = email_mittente
    msg['To'] = destinatario
    msg['Subject'] = f"Rapporto Intervento Tecnico - {nome_cliente}"
    msg.attach(MIMEText("Buongiorno,\nin allegato inviamo il rapporto tecnico in formato PDF.\n\nCordiali Saluti.", 'plain'))
    
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
        st.success("✉️ Email inviata al cliente con successo!")
    except Exception as e:
        st.warning(f"⚠️ Nota: Dati salvati, ma l'email non è partita. Errore: {e}")

# --- SEZIONE 4: FUNZIONE CREAZIONE PDF ---
def elabora_pdf(pdf_filename, data_str, cliente, email_cliente, marchio, matricola, km, ore_lavoro, preventivo, urgente, guasto_segnalato, descrizione_lavori, file_immagine):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter

    doc = SimpleDocTemplate(pdf_filename, pagesize=letter, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('NewTitle', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor("#1A365D"), alignment=1, spaceAfter=20)
    section_heading = ParagraphStyle('SectionHeading', parent=styles['Heading3'], fontSize=12, leading=16, textColor=colors.HexColor("#2C5282"), spaceBefore=14, spaceAfter=6)
    body_style = ParagraphStyle('NewBody', parent=styles['Normal'], fontSize=10, leading=16)
    
    story = []
    if os.path.exists(LOGO_FILE):
        story.append(RLImage(LOGO_FILE, width=530, height=75))
        story.append(Spacer(1, 15))
    
    story.append(Paragraph("<b>RAPPORTO DI INTERVENTO TECNICO</b>", title_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"<b>Data Intervento:</b> {data_str}", body_style))
    story.append(Paragraph(f"<b>Ragione Sociale Cliente:</b> {cliente}", body_style))
    story.append(Paragraph(f"<b>Email Cliente:</b> {email_cliente if email_cliente else 'N.D.'}", body_style))
    story.append(Paragraph(f"<b>Apparecchio / Marchio:</b> {marchio}", body_style))
    story.append(Paragraph(f"<b>Matricola:</b> {matricola if matricola else 'N.D.'}", body_style))
    story.append(Paragraph(f"<b>Kilometri Percorsi:</b> {km} Km", body_style))
    story.append(Paragraph(f"<b>Ore Impiegate:</b> {ore_lavoro} ore", body_style))
    story.append(Paragraph(f"<b>Richiesta Preventivo:</b> {preventivo} | <b>Intervento Urgente:</b> {urgente}", body_style))
    
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>■ GUASTO SEGNALATO</b>", section_heading))
    story.append(Paragraph(guasto_segnalato if guasto_segnalato else "Nessuna segnalazione inserita.", body_style))
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>■ DETTAGLIO LAVORI ESEGUITI</b>", section_heading))
    story.append(Paragraph(descrizione_lavori, body_style))
    
    story.append(Spacer(1, 20))
    story.append(Paragraph("------------------------------------------------------------------------------------------------------------------------", body_style))
    story.append(Spacer(1, 10))
    
    story.append(Paragraph("<b>Firma del Tecnico:</b>", body_style))
    story.append(Spacer(1, 40)) 
    story.append(Paragraph("___________________________", body_style))
    
    story.append(Spacer(1, 50)) # AMPIO STACCO TRA LE DUE FIRME
    story.append(Paragraph("<b>Firma per Accettazione Cliente:</b>", body_style))
    story.append(Spacer(1, 40))
    story.append(Paragraph("___________________________", body_style))
    story.append(Spacer(1, 25))
    
    if file_immagine is not None:
        story.append(Paragraph("<b>■ ALLEGATO - SCHEDA CARTACEA CON FIRMA ORIGINALE</b>", section_heading))
        story.append(Spacer(1, 10))
        foto_img = Image.open(file_immagine)
        foto_path = "temp_allegato.png"
        foto_img.thumbnail((500, 450))
        foto_img.save(foto_path)
        story.append(RLImage(foto_path, width=450, height=350))
        
    doc.build(story)

# --- SEZIONE 5: LOGICA DI SALVATAGGIO GENERALE ---
if st.button("💾 REGISTRA E GENERA REPORT COMPLETO"):
    if not cliente or not marchio or not descrizione_lavori:
        st.error("⚠️ Inserisci almeno Cliente, Marchio e Descrizione Lavori!")
    else:
        with st.spinner("Salvataggio e generazione documenti..."):
            try:
                data_str = data_corrente.strftime("%d/%m/%Y")
                
                # A) SALVATAGGIO IN EXCEL
                riga = {
                    "Data": data_str, "Cliente": cliente, "Email Cliente": email_cliente if email_cliente else "N.D.",
                    "Marchio": marchio, "Matricola": matricola if matricola else "N.D.",
                    "Guasto Segnalato": guasto_segnalato if guasto_segnalato else "N.D.", "Intervento": descrizione_lavori,
                    "Km": km if km > 0 else "0", "Ore Lavoro": ore_lavoro if ore_lavoro > 0 else "0",
                    "Preventivo": preventivo, "Urgente": urgente
                }
                df_nuovo = pd.DataFrame([riga])
                if os.path.exists(EXCEL_FILE):
                    df_esistente = pd.read_excel(EXCEL_FILE)
                    df_finale = pd.concat([df_esistente, df_nuovo], ignore_index=True)
                else:
