import streamlit as st
import pandas as pd
from io import BytesIO

# Titolo dell'app
st.title("Identifica Duplicati nelle Matricole")

# Carica il file Excel
uploaded_file = st.file_uploader("Carica il file Excel con le matricole", type=["xlsx"])

if uploaded_file:
    # Leggi il file Excel
    df = pd.read_excel(uploaded_file)

    # Verifica se la colonna 'Matricola' esiste
    if 'MATRICOLA' in df.columns:
        # Normalizza le matricole rimuovendo caratteri speciali
        df['Matricola_normalizzata'] = df['MATRICOLA'].astype(str).str.replace(r'\W+', '', regex=True)

        # Trova i duplicati
        duplicati = df[df.duplicated('Matricola_normalizzata', keep=False)]

        # Mostra i duplicati
        st.write("Duplicati trovati:")
        st.write(duplicati)

        # Creare un file Excel in memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            duplicati.to_excel(writer, index=False, sheet_name='Duplicati')
        output.seek(0)  # Riportare il puntatore all'inizio del file

        # Consenti di scaricare il file con i duplicati
        st.download_button(label="Scarica i duplicati in Excel", 
                           data=output,
                           file_name="duplicati.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("La colonna 'MATRICOLA' non è presente nel file caricato.")