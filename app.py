import streamlit as st
import pandas as pd
import os
from PIL import Image
from datetime import datetime
from streamlit_canvas import st_canvas
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni con firma, archivio PDF ed invio Email.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  # Carica un file chiamato logo.png su GitHub per vederlo nel PDF

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
        preventivo = st.radio("Richiede Preventivo?", ["NO", "SI"])
    with col2:
        urgente = st.radio("Intervento Urgente?", ["NO", "SI"])

# --- 2. ACQUISIZIONE MEDIA (FOTO E FIRMA) ---
st.subheader("📸 Archiviazione Foglio Cartaceo")
file_immagine = st.camera_input("Scatta una foto al foglio cartaceo (Opzionale)")

st.subheader("✍️ Firma Digitale del Cliente")
st.write("Fai firmare il cliente direttamente qui sotto con il dito:")
canvas_result = st_canvas(
    fill_color="rgba(255, 255, 255, 1)",
    stroke_width=3,
    stroke_color="#000000",
    background_color="#ffffff",
    height=150,
    width=350,
    drawing_mode="freedraw",
    key="canvas",
)

# --- 3. FUNZIONE INVIO EMAIL ---
def invia_email_pdf(destinatario, allegato_path, nome_cliente):
    # NOTA: Configura questi parametri con la tua email aziendale
    email_mittente = "franciosoandrea@gmail.com" 
    password_mittente = "la_tua_password_app_gmail" 
    
    msg = MIMEMultipart()
    msg['From'] = email_mittente
    msg['To'] = destinatario
    msg['Subject'] = f"Rapporto Intervento Tecnico - {nome_cliente}"
    
    corpo_testo = f"Buongiorno,\nin allegato inviamo il rapporto tecnico relativo all'intervento eseguito in data odierna.\n\nCordiali Saluti."
    msg.attach(MIMEText(corpo_testo, 'plain'))
    
    with open(allegato_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {allegato_path}")
        msg.attach(part)
        
    try:
        server = smtplib.SMTP("://gmail.com", 587)
        server.starttls()
        server.login(email_mittente, password_mittente)
        server.sendmail(email_mittente, destinatario, msg.as_string())
        server.quit()
        st.success("✉️ Email inviata al cliente con successo!")
    except Exception as e:
        st.warning(f"⚠️ Impossibile inviare l'email automaticamente: {e}. Controlla la configurazione SMTP nel codice.")

# --- 4. SALVATAGGIO ---
if st.button("💾 REGISTRA E GENERA REPORT COMPLETO"):
    if not cliente or not marchio or not descrizione_lavori:
        st.error("⚠️ Compila i campi obbligatori contrassegnati con l'asterisco (*)")
    else:
        with st.spinner("Elaborazione dati e generazione documenti..."):
            try:
                data_str = data_corrente.strftime("%d/%m/%Y")
                
                # A) AGGIORNAMENTO EXCEL CUMULATIVO
                riga = {
                    "Data": data_str,
                    "Cliente": cliente,
                    "Email Cliente": email_cliente,
                    "Marchio": marchio,
                    "Matricola": matricola,
                    "Guasto Segnalato": guasto_segnalato,
                    "Intervento": descrizione_lavori,
                    "Km": km if km > 0 else "N.D.",
                    "Ore Lavoro": ore_lavoro if ore_lavoro > 0 else "N.D.",
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

                # B) CREAZIONE PDF PROFESSIONALE
                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
                from reportlab.lib.styles import getSampleStyleSheet
                from reportlab.lib.pagesizes import letter
                
                cliente_pulito = cliente.replace(" ", "_").replace("/", "_")
                pdf_filename = f"Report_{data_corrente.strftime('%Y%m%d')}_{cliente_pulito}.pdf"
                
                doc = SimpleDocTemplate(pdf_filename, pagesize=letter)
                styles = getSampleStyleSheet()
                story = []
                
                # Inserimento Logo se presente
                if os.path.exists(LOGO_FILE):
                    story.append(RLImage(LOGO_FILE, width=120, height=50))
                    story.append(Spacer(1, 10))
                
                story.append(Paragraph(f"<b>RAPPORTO DI INTERVENTO TECNICO</b>", styles['Title']))
                story.append(Spacer(1, 15))
                story.append(Paragraph(f"<b>Data:</b> {data_str} | <b>Cliente:</b> {cliente} ({email_cliente})", styles['Normal']))
                story.append(Paragraph(f"<b>Apparecchio:</b> {marchio} | <b>Matricola:</b> {matricola if matricola else 'N.D.'}", styles['Normal']))
                story.append(Paragraph(f"<b>Km Percorsi:</b> {km} | <b>Ore Impiegate:</b> {ore_lavoro}", styles['Normal']))
                story.append(Paragraph(f"<b>Preventivo Richiesto:</b> {preventivo} | <b>Urgente:</b> {urgente}", styles['Normal']))
                story.append(Spacer(1, 10))
                story.append(Paragraph(f"<b>Guasto Segnalato:</b><br/>{guasto_segnalato}", styles['Normal']))
                story.append(Spacer(1, 10))
                story.append(Paragraph(f"<b>Dettaglio Lavori Eseguiti:</b><br/>{descrizione_lavori}", styles['Normal']))
                story.append(Spacer(1, 15))
                
                # Inserimento Firma Digitale nel PDF
                if canvas_result.image_data is not None:
                    firma_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                    firma_path = "temp_firma.png"
                    firma_img.save(firma_path)
                    story.append(Paragraph("<b>Firma del Cliente per Accettazione:</b>", styles['Normal']))
                    story.append(RLImage(firma_path, width=150, height=65))
                
                # Inserimento eventuale foto scattata
                if file_immagine is not None:
                    story.append(Spacer(1, 15))
                    story.append(Paragraph("<b>Allegato - Foto della Scheda Cartacea:</b>", styles['Normal']))
                    foto_img = Image.open(file_immagine)
                    foto_path = "temp_foto.png"
                    foto_img.save(foto_path)
                    story.append(RLImage(foto_path, width=400, height=300))
                
                doc.build(story)
                st.success("🎉 Dati salvati nell'Excel e PDF creato correttamente!")
                
                # C) INVIO EMAIL SE COMPILATA
                if email_cliente:
                    invia_email_pdf(email_cliente, pdf_filename, cliente)
                
                # Pulsanti di Download sul telefono
                with open(EXCEL_FILE, "rb") as f:
                    st.download_button("📥 Scarica Excel Completo", f, file_name=EXCEL_FILE)
                with open(pdf_filename, "rb") as f:
                    st.download_button("📥 Scarica PDF Report", f, file_name=pdf_filename)
                    
            except Exception as e:
                st.error(f"Errore durante l'elaborazione: {e}")

