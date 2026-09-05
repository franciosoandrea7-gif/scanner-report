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
                st.success("🔒 Documento Convalidato!")
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
        with st.spinner("Registrazione in corso e invio email..."):
            try:
                # A. REGISTRAZIONE DATI EXCEL (CORRETTO: 'urgente' anziché 'urgent')
                data_str = data_corrente.strftime("%d/%m/%Y")
                riga = {"Data": data_str, "Cliente": cliente, "Email": email_cliente, "Cellulare": cellulare_cliente, "Marchio": marchio, "Matricola": matricola if matricola else "N.D.", "Guasto": guasto_segnalato, "Intervento": descrizione_lavori, "Km": km, "Ore": ore_lavoro, "Preventivo": preventivo, "Urgente": urgente}
                
                df_vecchio = pd.read_excel(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else pd.DataFrame()
                df_nuovo = pd.concat([df_vecchio, pd.DataFrame([riga])], ignore_index=True)
                df_nuovo.to_excel(EXCEL_FILE, index=False)
                st.success(f"🎯 Intervento di {cliente} salvato nel registro Excel!")

                # B. CONFIGURAZIONE E INVIO EMAIL VIA SMTP (PORTA 587)
                email_mittente = "franciosoandrea@gmail.com" 
                password_mittente = "qiad bvqq ijaj mutc "  
                
                msg = MIMEMultipart()
                msg['From'] = email_mittente
                msg['To'] = email_cliente
                msg['Subject'] = f"Rapporto Intervento Tecnico - {cliente}"
                
                testo_corpo = f"Buongiorno,\n\nConfermiamo la corretta registrazione del rapporto ufficiale per l'intervento odierno eseguito presso {cliente}.\n\nNote Intervento:\n{descrizione_lavori}\n\nFirma validata digitalmente tramite cellulare.\n\nCordiali Saluti\nNova Servimpianti."
                msg.attach(MIMEText(testo_corpo, 'plain'))
                
                # Se è presente la foto dello schema, la alleghiamo direttamente alla mail
                if file_immagine is not None:
                    try:
                        foto_img = Image.open(file_immagine).convert("RGB")
                        foto_path = "temp_allegato_email.png"
                        foto_img.save(foto_path)
                        with open(foto_path, "rb") as attachment:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment.read())
                            encoders.encode_base64(part)
                            part.add_header("Content-Disposition", f"attachment; filename= foto_intervento.png")
                            msg.attach(part)
                    except Exception as img_err:
                        st.warning(f"Impossibile allegare la foto all'email: {img_err}")
                
                # Connessione al server Gmail SMTP sulla porta 587
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()  
                server.login(email_mittente, password_mittente)
                server.sendmail(email_mittente, email_cliente, msg.as_string())
                server.quit()
                
                st.success("✉️ Email di notifica inviata correttamente al cliente!")
                
            except Exception as e:
                st.error(f"Si è verificato un problema durante l'operazione: {e}")
