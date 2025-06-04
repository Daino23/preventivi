import streamlit as st
from docx import Document
from datetime import datetime
from io import BytesIO
from docx.shared import Inches

st.set_page_config(page_title="Preventivo Studio Dainotti Avanzato")
st.title("Preventivo Studio Dainotti - Versione Avanzata")

# Logo
st.sidebar.image("https://i.ibb.co/nb6JqGC/logo-studiodainotti.png", width=150)

# Intestazione
st.subheader("Dati intestazione")
data = st.date_input("Data", value=datetime.today())
numero = st.text_input("Numero preventivo", "88")
cliente = st.text_input("Cliente", "Gianfranco Consiglio")
oggetto = st.text_area("Oggetto", "Servizi strategici per il lancio e la promozione di due prodotti – Analisi + Funnel + Landing + Ads")
includi_iva = st.checkbox("Includi IVA (22%)", value=True)

# Voci
st.subheader("Voci di preventivo")
with st.form("form_voci"):
    voce = st.text_input("Voce")
    frequenza = st.text_input("Frequenza")
    descrizione = st.text_area("Descrizione")
    prezzo_reale = st.number_input("Prezzo reale (€)", min_value=0.0)
    prezzo_applicato = st.number_input("Prezzo applicato (€)", min_value=0.0)
    aggiungi = st.form_submit_button("Aggiungi voce")

if "lista_voci" not in st.session_state:
    st.session_state.lista_voci = []

if aggiungi:
    st.session_state.lista_voci.append({
        "voce": voce,
        "frequenza": frequenza,
        "descrizione": descrizione,
        "prezzo_reale": prezzo_reale,
        "prezzo_applicato": prezzo_applicato
    })

for riga in st.session_state.lista_voci:
    st.markdown(f"**{riga['voce']}** - {riga['frequenza']} - {riga['descrizione']}")
    st.markdown(f"Prezzo reale: €{riga['prezzo_reale']} | Prezzo applicato: €{riga['prezzo_applicato']}")

# Generazione documento
st.subheader("Generazione documento")

firma_finale = st.text_area("Testo firma finale", "Cordiali saluti,\nAndrea Dainotti\nStudio Dainotti")

if st.button("Genera Preventivo Word e PDF"):
    doc = Document()

    # Logo
    header = doc.sections[0].header
    try:
        from urllib.request import urlopen
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
            tmpfile.write(urlopen("https://i.ibb.co/nb6JqGC/logo-studiodainotti.png").read())
            tmpfile.flush()
            doc.add_picture(tmpfile.name, width=Inches(1.5))
    except:
        pass

    doc.add_heading("Preventivo – Analisi e Strategia Marketing", level=1)
    doc.add_paragraph("Studio Dainotti")
    doc.add_paragraph(f"Preventivo n. {numero}")
    doc.add_paragraph(f"Data: {data.strftime('%d/%m/%Y')}")
    doc.add_paragraph(f"Cliente: {cliente}")
    doc.add_paragraph(f"Oggetto: {oggetto}")

    doc.add_heading("Dettaglio Servizi e Valore", level=2)
    table = doc.add_table(rows=1, cols=5)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Voce'
    hdr_cells[1].text = 'Frequenza'
    hdr_cells[2].text = 'Descrizione'
    hdr_cells[3].text = 'Prezzo reale (€)'
    hdr_cells[4].text = 'Prezzo applicato (€)'

    totale_reale = 0
    totale_applicato = 0

    for voce in st.session_state.lista_voci:
        row_cells = table.add_row().cells
        row_cells[0].text = voce['voce']
        row_cells[1].text = voce['frequenza']
        row_cells[2].text = voce['descrizione']
        row_cells[3].text = f"{voce['prezzo_reale']:.2f}"
        row_cells[4].text = f"{voce['prezzo_applicato']:.2f}"
        totale_reale += voce['prezzo_reale']
        totale_applicato += voce['prezzo_applicato']

    sconto = totale_reale - totale_applicato
    percentuale_sconto = (sconto / totale_reale * 100) if totale_reale else 0

    doc.add_paragraph("")
    doc.add_paragraph(f"Valore complessivo dei servizi: €{totale_reale:.2f} + IVA")
    doc.add_paragraph(f"Totale applicato: €{totale_applicato:.2f} + IVA")
    doc.add_paragraph(f"Sconto applicato: –€{sconto:.2f} (–{percentuale_sconto:.1f}%)")

    if includi_iva:
        totale_ivato = totale_applicato * 1.22
        doc.add_paragraph(f"Totale comprensivo di IVA (22%): €{totale_ivato:.2f}")

    doc.add_heading("Condizioni", level=2)
    doc.add_paragraph("Consegna: entro 15-20 giorni lavorativi dalla conferma")
    doc.add_paragraph("Pagamento: 50% alla conferma – 50% alla consegna")
    doc.add_paragraph("Modalità: Bonifico Bancario intestato a: Dainotti Srls, via Roma, 52, 21010 Porto Valtravaglia (VA)")
    doc.add_paragraph("IBAN: IT06S0538750401000003855981")
    doc.add_paragraph(f"Causale: Preventivo n. {numero} del {data.strftime('%d/%m/%Y')}")
    doc.add_paragraph("\nAttenzione: Il preventivo ha validità 7 giorni dalla data di emissione")

    doc.add_paragraph("\n" + firma_finale)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    file_name = f"Preventivo_{cliente.replace(' ', '_')}_{data.strftime('%d-%m-%Y')}.docx"

    st.download_button(
        label="📄 Scarica Preventivo Word",
        data=buffer,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
