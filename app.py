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
from openpyxl import load_workbook

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni Nova Servimpianti con validazione forte ed invio Email.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  

# --- CONFIGURAZIONE TEAM TECNICI E PIN SEGRETI ---
TECNICI = {
    "Andrea Francioso": "1974",
    "Daniele Gennari 1": "1990",
    "Lidia Distratis 2": "1977",
    "Nome Dipendente 3": "3415",
    "Nome Dipendente 4": "7712"
}

if "codice_sms" not in st.session_state:
    st.session_state["codice_sms"] = None
if "sms_validato" not in st.session_state:
    st.session_state["sms_validato"] = False
if "ultimo_pdf" not in st.session_state:
    st.session_state["ultimo_pdf"] = None
if "mostra_download" not in st.session_state:
    st.session_state["mostra_download"] = False

# --- RECUPERO DINAMICO LISTA CLIENTI DA EXCEL ---
lista_clienti_esistenti = []
if os.path.exists(EXCEL_FILE):
    try:
        wb = load_workbook(EXCEL_FILE, read_only=True)
        # Sostituiamo i trattini bassi con gli spazi per renderli leggibili nel menu
        lista_clienti_esistenti = [sheet.replace("_", " ") for sheet in wb.sheetnames if sheet != "Sheet1"]
        wb.close()
    except Exception:
        lista_clienti_esistenti = []

lista_clienti_esistenti.sort()
opzioni_menu_clienti = ["➕ AGGIUNGI NUOVO CLIENTE"] + lista_clienti_esistenti

# --- 0. ARCHIVIO EXCEL ---
with st.expander("📚 Archivio Storico Lavori (Excel)"):
    if os.path.exists(EXCEL_FILE):
        try:
            wb = load_workbook(EXCEL_FILE, read_only=True)
            fogli = wb.sheetnames
            wb.close()
            if fogli:
                foglio_scelto = st.selectbox("Seleziona il registro del cliente da visualizzare:", fogli)
                st.dataframe(pd.read_excel(EXCEL_FILE, sheet_name=foglio_scelto))
            else:
                st.info("ℹ️ L'archivio Excel non contiene fogli validi.")
        except Exception as e:
            st.error(f"Errore nella lettura del registro: {e}")
    else:
        st.info("ℹ️ L'archivio Excel è vuoto.")

# --- 0. ARCHIVIO PDF ---
with st.expander("📂 Recupera Vecchi Report PDF Emessi"):
    lista_pdf = [f for f in os.listdir(".") if f.startswith("Report_") and f.endswith(".pdf")]
    if len(lista_pdf) > 0:
        lista_pdf.sort(reverse=True)
        for nome_pdf in lista_pdf:
            col_n, col_b = st.columns(2)
            with col_n:
                st.write(f"📄 {nome_pdf.replace('Report_', '').replace('.pdf', '')}")
            with col_b:
                with open(nome_pdf, "rb") as f_pdf:
                    st.download_button("📥 Scarica", f_pdf, file_name=nome_pdf, key=f"st_{nome_pdf}")
    else:
        st.info("ℹ️ Nessun PDF in archivio.")

# --- SEZIONE SELEZIONE E FIRMA TECNICO ---
st.subheader("👨‍🔧 Responsabile Intervento")
tecnico_selezionato = st.selectbox("Seleziona il tuo nome dal personale Nova:", list(TECNICI.keys()))
pin_tecnico = st.text_input("Inserisci il tuo PIN Tecnico segreto per firmare:", type="password")

tecnico_autorizzato = False
if pin_tecnico:
    if TECNICI.get(tecnico_selezionato) == pin_tecnico:
        tecnico_autorizzato = True
        st.success(f"✍️ ID Tecnico Verificato: {tecnico_selezionato}")
    else:
        st.error("❌ PIN errato!")

# --- 1. MODULO DATI CLIENTE ---
st.subheader("👤 Dati Intervento")
data_corrente = st.date_input("Data Intervento", datetime.now())

# Menu a tendina intelligente per i Clienti
cliente_selezionato_menu = st.selectbox("Seleziona Cliente *", opzioni_menu_clienti)

if cliente_selezionato_menu == "➕ AGGIUNGI NUOVO CLIENTE":
    cliente = st.text_input("Inserisci Nuova Ragione Sociale Cliente *")
else:
    cliente = cliente_selezionato_menu
    st.info(f"📍 Cliente selezionato: **{cliente}**")

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
st.subheader("🔒 Firma Digitale SMS Cliente")
if st.button("📲 INVIA CODICE DI VALIDAZIONE VIA SMS"):
    if not tecnico_autorizzato:
        st.error("⛔ AZIONE BLOCCATA: Devi prima selezionare il tuo nome Tecnico e inserire il PIN corretto in alto!")
    elif not cellulare_cliente or not cliente:
        st.error("⚠️ Inserisci Cliente e Cellulare!")
    else:
        st.session_state["codice_sms"] = str(random.randint(1000, 9999))
        st.session_state["sms_validato"] = False
        st.session_state["mostra_download"] = False
        
        from twilio.rest import Client
        
        ACCOUNT_SID = st.secrets["TWILIO_ACCOUNT_SID"]
        AUTH_TOKEN = st.secrets["TWILIO_AUTH_TOKEN"]
        NUMERO_TWILIO = st.secrets["TWILIO_NUMBER"]
        
        num_destinatario = cellulare_cliente
        if not num_destinatario.startswith("+"):
            if num_destinatario.startswith("39"):
                num_destinatario = "+" + num_destinatario
            else:
                num_destinatario = "+39" + num_destinatario
                
        testo_messaggio = f"Nova Servimpianti: Il tuo codice segreto di firma per l'intervento odierno e': {st.session_state['codice_sms']}"
        
        try:
            twilio_client = Client(ACCOUNT_SID, AUTH_TOKEN)
            twilio_client.messages.create(body=testo_messaggio, from_=NUMERO_TWILIO, to=num_destinatario)
            st.success("📩 SMS inviato al telefono del cliente con successo!")
        except Exception as e:
            st.warning(f"⚠️ Nota: Richiesta elaborata. Se l'SMS non arriva, usa il codice mostrato qui sotto. Errore: {e}")

if st.session_state["codice_sms"] is not None:
    st.info(f"👉 CODICE DI VALIDAZIONE D'EMERGENZA: {st.session_state['codice_sms']}")
    if not st.session_state["sms_validato"]:
        codice_inserito = st.text_input("Inserisci le 4 cifre:", key="codice_verifica_sms")
        if st.button("✅ VALIDA CODICE SMS"):
            if codice_inserito == st.session_state["codice_sms"]:
                st.session_state["sms_validato"] = True
                st.success("🔒 Validato!")
                st.rerun()
            else:
                st.error("❌ Codice errato!")
    else:
        st.success("🔒 Convalidato con Successo!")

# --- 3. LOGICA INVIO COPIA EMAIL ---
def invia_email_pdf(destinatario, allegato_path, nome_cliente):
    email_mittente = "franciosoandrea@gmail.com" 
    password_mittente = st.secrets["GMAIL_PASSWORD"]
    
    msg_cli = MIMEMultipart()
    msg_cli['From'] = email_mittente
    msg_cli['To'] = destinatario
    msg_cli['Subject'] = f"Report Intervento - {nome_cliente}"
    msg_cli.attach(MIMEText("Buongiorno, in allegato copia del rapporto ufficiale Nova Servimpianti.\n\nCordiali Saluti.", 'plain'))
    
    msg_teco = MIMEMultipart()
    msg_teco['From'] = email_mittente
    msg_teco['To'] = "franciosoandrea@me.com"
    msg_teco['Subject'] = f"Nova Servimpianti - Backup Intervento: {nome_cliente}"
    msg_teco.attach(MIMEText(f"Rapporto registrato correttamente nel database.\nIn allegato trovi il PDF dell'intervento e il file Excel Generale aggiornato.", 'plain'))
    
    try:
        with open(allegato_path, "rb") as att_pdf:
            part_pdf = MIMEBase("application", "octet-stream")
            part_pdf.set_payload(att_pdf.read())
            encoders.encode_base64(part_pdf)
            part_pdf.add_header("Content-Disposition", f"attachment; filename= {allegato_path}")
            msg_cli.attach(part_pdf)
            
            part_pdf_teco = MIMEBase("application", "octet-stream")
            att_pdf.seek(0)
            part_pdf_teco.set_payload(att_pdf.read())
            encoders.encode_base64(part_pdf_teco)
            part_pdf_teco.add_header("Content-Disposition", f"attachment; filename= {allegato_path}")
            msg_teco.attach(part_pdf_teco)
            
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as att_ex:
                part_ex = MIMEBase("application", "octet-stream")
                part_ex.set_payload(att_ex.read())
                encoders.encode_base64(part_ex)
                part_ex.add_header("Content-Disposition", f"attachment; filename= {EXCEL_FILE}")
                msg_teco.attach(part_ex)
                
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(email_mittente, password_mittente)
        server.sendmail(email_mittente, destinatario, msg_cli.as_string())
        server.sendmail(email_mittente, "franciosoandrea@me.com", msg_teco.as_string())
        server.quit()
        st.success("✉️ Documenti inviati! PDF inviato al cliente, PDF + Excel Storico inviati a franciosoandrea@me.com")
    except Exception as e:
        st.warning(f"⚠️ Nota: File registrati, ma l'invio email ha riscontrato un problema: {e}")

# --- 4. CREAZIONE PDF ---
def elabora_pdf(pdf_filename, data_str, cliente, email_cliente, cellulare_cliente, marchio, matricola, km, ore_lavoro, preventivo, urgent, guasto_segnalato, descrizione_lavori, file_immagine, stringa_firma, firma_tecnico):
