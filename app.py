import streamlit as st
import pandas as pd
import numpy as np
import io

st.title("🔍 Analisi Servizi IT")
st.write("Carica un file Excel per generare un report con i dettagli dei servizi IT.")

# Caricamento del file
uploaded_file = st.file_uploader("📂 Carica il file Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Lista unica dei servizi IT
    servizi_it = df['Servizio IT'].unique()

    # Creazione del file Excel in memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for servizio in servizi_it:
            df_servizio = df[df['Servizio IT'] == servizio]
            aree_utenti = df_servizio.groupby(['Area', 'Member Name'])['Group Name'].apply(list).reset_index()
            gruppi_servizio = df_servizio['Group Name'].unique()

            presenza_df = pd.DataFrame(columns=['Area', 'Utenza'] + list(gruppi_servizio))
            presenza_df[['Area', 'Utenza']] = aree_utenti[['Area', 'Member Name']]

            for idx, row in aree_utenti.iterrows():
                gruppi_utente = row['Group Name']
                for gruppo in gruppi_servizio:
                    presenza_df.at[idx, gruppo] = 'PRESENTE' if gruppo in gruppi_utente else 'NON PRESENTE'

            sheet_name = servizio[:31]

            count = 1
            while sheet_name in writer.book.sheetnames:
                sheet_name = f"{servizio[:28]}_{count}"
                count += 1

            presenza_df.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)

    # Pulsante per scaricare il file generato
    st.download_button(
        label="📥 Scarica il file Excel generato",
        data=output,
        file_name="analisi_servizi_it.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
