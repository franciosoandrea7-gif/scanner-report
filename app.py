import streamlit as st
import pandas as pd
import os
from PIL import Image
from datetime import datetime

# Configurazione interfaccia per Smartphone
st.set_page_config(page_title="Gestione Report Nova", page_icon="📝", layout="centered")

st.title("📝 Registro Schede Tecniche")
st.write("Compila i dati rapidi e scatta la foto per generare il PDF e aggiornare l'Excel.")

EXCEL_FILE = "registro_riparazioni.xlsx"

# 1. INSERIMENTO DATI RAPIDI PER EXCEL
st.subheader("📊 Dati per l'Excel")
data_corrente = st.date_input("Data Intervento", datetime.now())
cliente = st.text_input("Ragione Sociale Cliente", placeholder="Es: Porto Taranto Edifici")
marchio = st.text_input("Marchio Apparecchio", placeholder="Es: Mitsubishi")

# NUOVO CAMPO AGGIUNTO
descrizione_lavori = st.text_area("Descrizione Intervento / Lavori Svolti", placeholder="Scrivi qui i lavori eseguiti (es. Sostituzione scheda, ricarica gas...)")

# 2. FOTO DELLA SCHEDA CARTACEA
st.subheader("📸 Archiviazione PDF")
file_immagine = st.camera_input("Scatta una foto alla scheda firmata")

if file_immagine is not None:
    if st.button("💾 Salva Intervento"):
        if not cliente or not marchio or not descrizione_lavori:
            st.error("⚠️ Per favore, compila Cliente, Marchio e Descrizione prima di salvare!")
        else:
            with st.spinner("Salvataggio in corso..."):
                try:
                    # Formatta la data per l'Excel
                    data_str = data_corrente.strftime("%d/%m/%Y")
                    
                    # A) SALVATAGGIO IN EXCEL
                    riga = {
                        "Data": data_str,
                        "Ragione Sociale Cliente": cliente,
                        "Marchio": marchio,
                        "Descrizione Lavori": descrizione_lavori  # Salvataggio nel foglio excel
                    }
                    df_nuovo = pd.DataFrame([riga])
                    
                    if os.path.exists(EXCEL_FILE):
                        df_esistente = pd.read_excel(EXCEL_FILE)
                        df_finale = pd.concat([df_esistente, df_nuovo], ignore_index=True)
                    else:
                        df_finale = df_nuovo
                        
                    df_finale.to_excel(EXCEL_FILE, index=False)
                    
                    # B) CONVERSIONE FOTO IN PDF
                    image = Image.open(file_immagine)
                    cliente_pulito = cliente.replace(" ", "_").replace("/", "_")
                    pdf_filename = f"Report_{data_corrente.strftime('%Y%m%d')}_{cliente_pulito}.pdf"
                    
                    image.save(pdf_filename, "PDF", resolution=100.0)
                    
                    st.success(f"🎉 Successo! Dati inseriti nell'Excel e PDF creato come: {pdf_filename}")
                    
                    # Pulsanti per scaricare i file sul telefono
                    with open(EXCEL_FILE, "rb") as f:
                        st.download_button("📥 Scarica Excel Aggiornato", f, file_name=EXCEL_FILE)
                        
                    with open(pdf_filename, "rb") as f:
                        st.download_button("📥 Scarica PDF Scheda", f, file_name=pdf_filename)
                        
                except Exception as e:
                    st.error(f"Errore durante il salvataggio: {e}")
