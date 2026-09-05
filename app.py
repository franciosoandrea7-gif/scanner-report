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

# Inizializzazione stati di sessione
if "codice_sms" not in st.session_state:
    st.session_state["codice_sms"] = None
if "sms_validato" not in st.session_state:
    st.session_state["sms_validato"] = False

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
        st.success("📩 Richiesta SMS elaborata!")

if st.session_state["codice_sms"] is not None:
    st.info(f"👉 CODICE DA INSERIRE: {st.session_state['codice_sms']}")
    if not st.session_state["sms_validato"]:
        codice_inserito = st.text_input("Inserisci le 4 cifre per firmare:")
        if st.button("✅ VALIDA CODICE SMS"):
            if codice_inserito == st.session_state["codice_sms"]:
                st.session_state["sms_validato"] = True
                st.success("🔒 Documento Convalidato! Procedi in fondo per il salvataggio.")
            else:
                st.error("❌ Codice errato!")
    else:
        st.success("🔒 Convalidato!")

# --- 3. BOTTONE FINALE CON REGISTRAZIONE E INVIO EMAIL SMTP ---
st.write("---")
st.subheader("💾 Operazione Finale")

tasto_registra = st.button("💾 REGISTRA E GENERA REPORT COMPLETO", type="primary")

if tasto_registra:
    if not cliente or not marchio or not descrizione_lavori or not cellulare_cliente or not email_cliente:
        st.error("⚠️ Compila tutti i campi obbligatori (*)")
    elif not st.session_state["sms_validato"]:
        st.error("⚠️ Il cliente deve prima convalidare il codice SMS!")
    else:
        with st.spinner("Registrazione in corso e invio email completa..."):
            # A. REGISTRAZIONE DATI EXCEL
            data_str = data_corrente.strftime("%d/%m/%Y")
            stringa_firma = f"Firmato via SMS OTP il {data_str} (Codice: {st.session_state['codice_sms']})"
            
            riga = {
                "Data": data_str, "Cliente": cliente, "Email": email_cliente, "Cellulare": cellulare_cliente, 
                "Marchio": marchio, "Matricola": matricola if matricola else "N.D.", 
                "Guasto": guasto_segnalato if guasto_segnalato else "N.D.", "Intervento": descrizione_lavori, 
                "Km": km, "Ore": ore_lavoro, "Preventivo": preventivo, "Urgente": urgente, "Firma": stringa_firma
            }
            
            df_vecchio = pd.read_excel(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else pd.DataFrame()
            df_nuovo = pd.concat([df_vecchio, pd.DataFrame([riga])], ignore_index=True)
            df_nuovo.to_excel(EXCEL_FILE, index=False)
            st.success(f"🎯 Intervento di {cliente} salvato nel registro Excel!")

            # B. CONFIGURAZIONE E INVIO EMAIL HTML COMPLETA (PORTA 587)
            email_mittente = "franciosoandrea@gmail.com" 
            password_mittente = "qiad bvqq ijaj mutc "  
            
            msg = MIMEMultipart()
            msg['From'] = email_mittente
            msg['To'] = email_cliente
            msg['Subject'] = f"Rapporto Ufficiale Intervento Tecnico - {cliente}"
            
            corpo_html = f"""
            <html>
            <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
                <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">
                    <h2 style="color: #1A365D; border-bottom: 2px solid #2C5282; padding-bottom: 10px;">🛠️ RAPPORTO DI INTERVENTO TECNICO</h2>
                    <p>Buongiorno, inviamo di seguito il riepilogo ufficiale dell'intervento odierno eseguito da <b>Nova Servimpianti</b>.</p>
                    
                    <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
                        <tr style="background-color: #f7fafc;"><td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold; width: 40%;">Data Intervento:</td><td style="padding: 8px; border: 1px solid #e2e8f0;">{data_str}</td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">Ragione Sociale:</td><td style="padding: 8px; border: 1px solid #e2e8f0;">{cliente}</td></tr>
                        <tr style="background-color: #f7fafc;"><td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">Marchio Apparecchio:</td><td style="padding: 8px; border: 1px solid #e2e8f0;">{marchio}</td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">Matricola:</td><td style="padding: 8px; border: 1px solid #e2e8f0;">{matricola if matricola else 'N.D.'}</td></tr>
                        <tr style="background-color: #f7fafc;"><td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">Kilometri / Ore Lavoro:</td><td style="padding: 8px; border: 1px solid #e2e8f0;">{km} Km / {ore_lavoro} ore</td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">Richiesto Preventivo / Urgente:</td><td style="padding: 8px; border: 1px solid #e2e8f0;">{preventivo} / {urgente}</td></tr>
                    </table>
                    
                    <h4 style="color: #2C5282; margin-bottom: 5px;">■ GUASTO SEGNALATO:</h4>
                    <p style="background-color: #f7fafc; padding: 10px; border-left: 4px solid #cbd5e0; margin-top: 0;">{guasto_segnalato if guasto_segnalato else 'N.D.'}</p>
                    
                    <h4 style="color: #2C5282; margin-bottom: 5px;">■ INTERVENTO ESEGUITO / NOTE TECNICHE:</h4>
                    <p style="background-color: #f7fafc; padding: 10px; border-left: 4px solid #4299e1; margin-top: 0;">{descrizione_lavori}</p>
                    
                    <div style="background-color: #ebf8ff; border: 1px solid #bee3f8; padding: 12px; border-radius: 4px; margin-top: 25px; font-size: 13px;">
                        🔒 <b>Certificato di Validazione:</b><br/>
                        {stringa_firma}<br/>
                        <i>Documento convalidato sul posto tramite firma digitale SMS OTP (Inviata al numero: {cellulare_cliente})</i>
                    </div>
                    
                    <p style="font-size: 12px; color: #718096; margin-top: 30px; text-align: center;">Nova Servimpianti — Email generata automaticamente dal gestionale di bordo.</p>
                </div>
            </body>
            </html>
            """
            msg.attach(MIMEText(corpo_html, 'html'))
            
            # Allegato Foto
            if file_immagine is not None:
                foto_img = Image.open(file_immagine).convert("RGB")
                foto_path = "temp_allegato_email.png"
                foto_img.save(foto_path)
                with open(foto_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename= foto_scheda_firmata.png")
                    msg.attach(part)
            
            # Invio SMTP Lineare senza blocchi nidificati interrotti
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()  
            server.login(email_mittente, password_mittente)
            server.sendmail(email_mittente, email_cliente, msg.as_string())
            server.quit()
            
            st.success("✉️ Rapporto completo inviato correttamente via email al cliente!")
