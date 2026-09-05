import streamlit as st
import pandas as pd
import os
from PIL import Image
from datetime import datetime

st.set_page_config(page_title="Nova Report Pro", page_icon="⚙️", layout="centered")
st.title("🛠️ Nova Report Pro")
st.write("Gestionale riparazioni Nova Servimpianti con firma ed invio PDF.")

EXCEL_FILE = "registro_riparazioni.xlsx"
LOGO_FILE = "logo.png"  

# --- 1. MODULO DI INPUT DATI ---
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

# --- 2. SCHERMO BIANCO NATIVO PER FIRMA TOUCH ---
st.subheader("✍️ Firma Digitale del Cliente")
st.write("Fai firmare il cliente con il dito nel riquadro bianco qui sotto:")

# Sistema di disegno nativo in JavaScript (senza librerie esterne)
canvas_html = """
<div style="border:2px solid #CBD5E0; border-radius:8px; background-color:#ffffff; padding:5px; touch-action:none;">
    <canvas id="paint" width="450" height="150" style="width:100%; height:150px; background-color:#ffffff; cursor:crosshair;"></canvas>
</div>
<div style="margin-top:10px;">
    <button id="clear" style="background-color:#E2E8F0; border:none; padding:8px 15px; border-radius:5px; font-weight:bold; cursor:pointer;">Cancella Disegno</button>
</div>

<script>
    var canvas = document.getElementById('paint');
    var ctx = canvas.getContext('2d');
    ctx.lineWidth = 3;
    ctx.lineJoin = 'round';
    ctx.lineCap = 'round';
    ctx.strokeStyle = '#000000';
    
    var drawing = false;

    // Mouse eventi
    canvas.addEventListener('mousedown', function(e) { drawing = true; draw(e); }, false);
    canvas.addEventListener('mousemove', draw, false);
    canvas.addEventListener('mouseup', function() { drawing = false; ctx.beginPath(); sendData(); }, false);

    // Touch eventi per iPad e Smartphone
    canvas.addEventListener('touchstart', function(e) { drawing = true; drawTouch(e); e.preventDefault(); }, false);
    canvas.addEventListener('touchmove', function(e) { drawTouch(e); e.preventDefault(); }, false);
    canvas.addEventListener('touchend', function(e) { drawing = false; ctx.beginPath(); sendData(); e.preventDefault(); }, false);

    function draw(e) {
        if (!drawing) return;
        var rect = canvas.getBoundingClientRect();
        var x = (e.clientX - rect.left) * (canvas.width / rect.width);
        var y = (e.clientY - rect.top) * (canvas.height / rect.height);
        ctx.lineTo(x, y);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(x, y);
    }

    function drawTouch(e) {
        if (!drawing) return;
        var rect = canvas.getBoundingClientRect();
        var touch = e.touches[0];
        var x = (touch.clientX - rect.left) * (canvas.width / rect.width);
        var y = (touch.clientY - rect.top) * (canvas.height / rect.height);
        ctx.lineTo(x, y);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(x, y);
    }

    document.getElementById('clear').addEventListener('click', function() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        document.getElementById('firma_input').value = "";
    }, false);

    function sendData() {
        var dataUrl = canvas.toDataURL('image/png');
        window.parent.postMessage({type: 'streamlit:setComponentValue', value: dataUrl}, '*');
    }
</script>
"""
import streamlit.components.v1 as components
# Crea il riquadro grafico nativo touch
firma_base64 = components.html(canvas_html, height=210)

st.info("💡 Fai firmare sul display dell'iPad/Telefono. Se non vedi il tratto, muovi il dito con decisione sul riquadro bianco.")

# --- 3. LOGICA INVIO EMAIL ---
def invia_email_pdf(destinatario, allegato_path, nome_cliente):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    email_mittente = "franciosoandrea@gmail.com" 
    password_mittente = "qiad bvqq ijaj mutc "  # <--- METTI LA TUA PASSWORD DI GOOGLE QUI!
    msg = MIMEMultipart()
    msg['From'] = email_mittente
    msg['To'] = destinatario
    msg['Subject'] = f"Rapporto Intervento Tecnico - {nome_cliente}"
    msg.attach(MIMEText("Buongiorno,\nin allegato il report PDF.\n\nCordiali Saluti.", 'plain'))
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
        st.warning(f"⚠️ Nota email non partita: {e}")

# --- 4. CREAZIONE PDF ---
def elabora_pdf(pdf_filename, data_str, cliente, email_cliente, marchio, matricola, km, ore_lavoro, preventivo, urgente, guasto_segnalato, descrizione_lavori, file_immagine):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    doc = SimpleDocTemplate(pdf_filename, pagesize=letter, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor("#1A365D"), alignment=1, spaceAfter=20)
    section_heading = ParagraphStyle('T2', parent=styles['Heading3'], fontSize=12, leading=16, textColor=colors.HexColor("#2C5282"), spaceBefore=14, spaceAfter=6)
    body_style = ParagraphStyle('T3', parent=styles['Normal'], fontSize=10, leading=16)
    story = []
    if os.path.exists(LOGO_FILE):
        story.append(RLImage(LOGO_FILE, width=530, height=75))
        story.append(Spacer(1, 15))
    story.append(Paragraph("<b>RAPPORTO DI INTERVENTO TECNICO</b>", title_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"<b>Data Intervento:</b> {data_str}<br/><b>Cliente:</b> {cliente}<br/><b>Email:</b> {email_cliente if email_cliente else 'N.D.'}<br/><b>Marchio:</b> {marchio}<br/><b>Matricola:</b> {matricola if matricola else 'N.D.'}<br/><b>Km:</b> {km} Km | <b>Ore:</b> {ore_lavoro}<br/><b>Preventivo:</b> {preventivo} | <b>Urgente:</b> {urgente}", body_style))
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>■ GUASTO SEGNALATO</b>", section_heading))
    story.append(Paragraph(guasto_segnalato if guasto_segnalato else "N.D.", body_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>■ DETTAGLIO LAVORI ESEGUITI</b>", section_heading))
    story.append(Paragraph(descrizione_lavori, body_style))
    
    story.append(Spacer(1, 25))
    story.append(Paragraph("<b>Firma del Tecnico:</b><br/><br/><br/>___________________________", body_style))
    story.append(Spacer(1, 55)) # SPAZIO DI SICUREZZA LARGHEZZA FIRME
    story.append(Paragraph("<b>Firma per Accettazione Cliente:</b><br/><br/><br/>___________________________", body_style))
    
    if file_immagine is not None:
        story.append(Spacer(1, 25))
        story.append(Paragraph("<b>■ ALLEGATO FOTO SCHEDA</b>", section_heading))
        foto_img = Image.open(file_immagine)
        foto_path = "temp_allegato.png"
        foto_img.thumbnail((500, 450))
        foto_img.save(foto_path)
        story.append(RLImage(foto_path, width=450, height=350))
    try:
        doc.build(story)
    except Exception as e:
        st.error(f"Errore build PDF: {e}")

# --- 5. BOTTONE DI SALVATAGGIO ---
if st.button("💾 REGISTRA E GENERA REPORT COMPLETO"):
    if not cliente or not marchio or not descrizione_lavori:
        st.error("⚠️ Compila i campi obbligatori (*)!")
    else:
        data_str = data_corrente.strftime("%d/%m/%Y")
        riga = {"Data": data_str, "Cliente": cliente, "Email": email_cliente if email_cliente else "N.D.", "Marchio": marchio, "Matricola": matricola if matricola else "N.D.", "Guasto": guasto_segnalato if guasto_segnalato else "N.D.", "Intervento": descrizione_lavori, "Km": km, "Ore": ore_lavoro, "Preventivo": preventivo, "Urgente": urgente}
        
        if os.path.exists(EXCEL_FILE):
            df = pd.concat([pd.read_excel(EXCEL_FILE), pd.DataFrame([riga])], ignore_index=True)
        else:
            df = pd.DataFrame([riga])
        df.to_excel(EXCEL_FILE, index=False)
        
        c_pulito = cliente.replace(" ", "_").replace("/", "_")
        pdf_filename = f"Report_{data_corrente.strftime('%Y%m%d')}_{c_pulito}.pdf"
