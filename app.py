import streamlit as st
import pandas as pd
import os
from PIL import Image
from datetime import datetime

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni con archivio PDF ed invio automatico Email.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  # Se carichi un file logo.png su GitHub comparirà in testa al PDF

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

# --- 2. ACQUISIZIONE FOTO DEL FOGLIO FIRMATO ---
st.subheader("📸 Foto Scheda Firmata")
st.write("Inquadra e fotografa il foglio cartaceo compilato e firmato dal cliente:")
file_immagine = st.camera_input("Scatta la foto")

# --- 3. LOGICA INVIO EMAIL CON GMAIL ---
def invia_email_pdf(destinatario, allegato_path, nome_cliente):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    email_mittente = "franciosoandrea@gmail.com" 
    password_mittente = "xnqd msjk klrn gzlm"  # Assicurati che qui ci sia la tua password per le app di Google
    
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
        server = smtplib.SMTP("://gmail.com", 587)
        server.starttls()
        server.login(email_mittente, password_mittente)
        server.sendmail(email_mittente, destinatario, msg.as_string())
        server.quit()
        st.success("✉️ Email inviata al cliente con successo!")
    except Exception as e:
        st.warning(f"⚠️ Nota: Dati salvati, ma l'email non è partita automaticamente. Verifica la password delle app di Google. Errore: {e}")

# --- 4. SALVATAGGIO ---
if st.button("💾 REGISTRA E GENERA REPORT COMPLETO"):
    if not cliente or not marchio or not descrizione_lavori or file_immagine is None:
        st.error("⚠️ Per completare la registrazione devi inserire Cliente, Marchio, Descrizione Lavori e scattare la Foto del foglio!")
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

                # B) CONVERSIONE FOTO IN PDF PROFESSIONALE
                image = Image.open(file_immagine)
                cliente_pulito = cliente.replace(" ", "_").replace("/", "_")
                pdf_filename = f"Report_{data_corrente.strftime('%Y%m%d')}_{cliente_pulito}.pdf"
                
                # Salviamo l'immagine direttamente in un PDF pulito a piena pagina
                image.save(pdf_filename, "PDF", resolution=100.0)
                
                st.success("🎉 Dati salvati correttamente nel registro Excel!")
                
                # C) INVIO EMAIL SE COMPILATA
                if email_cliente:
                    invia_email_pdf(email_cliente, pdf_filename, cliente)
                
                # Pulsanti di Download sul telefono per sicurezza
                with open(EXCEL_FILE, "rb") as f:
                    st.download_button("📥 Scarica Excel Completo", f, file_name=EXCEL_FILE)
                with open(pdf_filename, "rb") as f:
                    st.download_button("📥 Scarica PDF Report", f, file_name=pdf_filename)
                    
            except Exception as e:
                st.error(f"Errore durante l'elaborazione: {e}")

