import streamlit as st
import pandas as pd
import os
from PIL import Image
from datetime import datetime
from streamlit_canvas import st_canvas

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni Nova Servimpianti con firma ed invio PDF.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  

# --- 1. CONFIGURAZIONE SEZIONI DATI ---
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

# --- 2. ACQUISIZIONE MEDIA (FOTO E FIRMA TOUCH) ---
st.subheader("📸 Foto Scheda Firmata (Opzionale)")
file_immagine = st.camera_input("Scatta la foto alla scheda cartacea se presente")

st.subheader("✍️ Firma Digitale del Cliente (Sul display)")
st.write("Fai firmare il cliente direttamente qui sotto con il dito o un pennino:")
canvas_result = st_canvas(
    fill_color="rgba(255, 255, 255, 1)",
    stroke_width=3,
    stroke_color="#000000",
    background_color="#ffffff",
    height=150,
    width=400,
    drawing_mode="freedraw",
    key="canvas_firma",
)

# --- 3. LOGICA INVIO EMAIL CON GMAIL ---
def invia_email_pdf(destinatario, allegato_path, nome_cliente):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    email_mittente = "franciosoandrea@gmail.com" 
    password_mittente = "qiad bvqq ijaj mutc "  # <--- METTI LA TUA PASSWORD A 16 LETTERE DI GOOGLE QUI!
    
    msg = MIMEMultipart()
    msg['From'] = email_mittente
    msg['To'] = destinatario
    msg['Subject'] = f"Rapporto Intervento Tecnico - {nome_cliente}"
    
    corpo_testo = f"Buongiorno,\nin allegato inviamo il rapporto tecnico in formato PDF relativo all'intervento eseguito.\n\nCordiali Saluti."
    msg.attach(MIMEText(corpo_testo, 'plain'))
    
    with open(allegato_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {allegato_path}")
        msg.attach(part)
        
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(email_mittente, password_mittente)
        server.sendmail(email_mittente, destinatario, msg.as_string())
        server.quit()
        st.success("✉️ Email inviata al cliente con successo!")
    except Exception as e:
        st.warning(f"⚠️ Nota: Dati salvati, ma l'email non è partita automaticamente. Errore: {e}")

# --- 4. SALVATAGGIO ---
if st.button("💾 REGISTRA E GENERA REPORT COMPLETO"):
    if not cliente or not marchio or not descrizione_lavori:
        st.error("⚠️ Per completare la registrazione devi inserire almeno Cliente, Marchio e Descrizione Lavori!")
    else:
        with st.spinner("Elaborazione dati e generazione documenti..."):
            try:
                data_str = data_corrente.strftime("%d/%m/%Y")
                
                # A) AGGIORNAMENTO EXCEL CUMULATIVO
                riga = {
                    "Data": data_str,
                    "Cliente": cliente,
                    "Email Cliente": email_cliente if email_cliente else "Non inserita",
                    "Marchio": marchio,
                    "Matricola": matricola if matricola else "N.D.",
                    "Guasto Segnalato": guasto_segnalato if guasto_segnalato else "N.D.",
                    "Intervento": descrizione_lavori,
                    "Km": km if km > 0 else "0",
                    "Ore Lavoro": ore_lavoro if ore_lavoro > 0 else "0",
                    "Preventivo": preventivo,
                    "Urgente": urgente
                }
                df_nuovo = pd.DataFrame([riga])
                if os.path.exists(EXCEL_FILE):
                    df_esistente = pd.read_excel(EXCEL_FILE)
                    df_finale = pd.concat([df_esistente, df_nuovo], ignore_index=True)
                else:
                    df_finale = df_nuovo
                df_finale.to_excel(EXCEL_FILE, index=False)

                # B) CREAZIONE PDF PROFESSIONALE LINEARE
                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
                from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                from reportlab.lib import colors
                from reportlab.lib.pagesizes import letter
                
                cliente_pulito = cliente.replace(" ", "_").replace("/", "_")
                pdf_filename = f"Report_{data_corrente.strftime('%Y%m%d')}_{cliente_pulito}.pdf"
                
                doc = SimpleDocTemplate(pdf_filename, pagesize=letter, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
                styles = getSampleStyleSheet()
                
                title_style = ParagraphStyle(
                    'NewTitle', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor("#1A365D"), alignment=1, spaceAfter=20
                )
                section_heading = ParagraphStyle(
                    'SectionHeading', parent=styles['Heading3'], fontSize=12, leading=16, textColor=colors.HexColor("#2C5282"), spaceBefore=14, spaceAfter=6
                )
                body_style = ParagraphStyle(
                    'NewBody', parent=styles['Normal'], fontSize=10, leading=16
                )
                
                story = []
                
                # Intestazione con Logo per lungo
                if os.path.exists(LOGO_FILE):
                    story.append(RLImage(LOGO_FILE, width=530, height=75))
                    story.append(Spacer(1, 15))
                
                story.append(Paragraph("<b>RAPPORTO DI INTERVENTO TECNICO</b>", title_style))
                story.append(Spacer(1, 10))
                
                # Lista Dati Intervento
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
                
                # SEZIONE FIRME DISTANZIATE
                story.append(Paragraph("<b>Firma del Tecnico:</b>", body_style))
                story.append(Spacer(1, 40)) # Spazio vuoto per la firma manuale del tecnico
                story.append(Paragraph("___________________________", body_style))
                
                story.append(Spacer(1, 30)) # Grande distacco tra le due firme
                
                story.append(Paragraph("<b>Firma per Accettazione Cliente (Digitale):</b>", body_style))
                
                # Incolla la firma touch nel PDF se il cliente ha firmato
                if canvas_result.image_data is not None:
                    import numpy as np
                    firma_array = canvas_result.image_data.astype('uint8')
                    # Controlla se il canvas non è completamente vuoto/bianco
                    if np.any(firma_array[:, :, 3] > 0): 
                        firma_img = Image.fromarray(firma_array, 'RGBA')
                        firma_path = "temp_firma.png"
                        firma_img.save(firma_path)
                        story.append(Spacer(1, 5))
                        story.append(RLImage(firma_path, width=150, height=55))
                
                story.append(Paragraph("___________________________", body_style))
                story.append(Spacer(1, 35))
                
