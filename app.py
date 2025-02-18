import streamlit as st
import pandas as pd
import numpy as np
import io


st.title("🔍 Analisi IT: Servizi e Famiglie")
st.write("Carica un file Excel e scegli il tipo di analisi da eseguire.")

# Caricamento del file
uploaded_file = st.file_uploader("📂 Carica il file Excel", type=["xlsx"])

# Selezione dell'analisi
do_servizi_it = st.checkbox("📊 Analisi Servizi IT")
do_famiglie = st.checkbox("🏠 Analisi Famiglie")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Aggiungi il logo al file Excel
        workbook = writer.book
        worksheet = workbook.add_worksheet('Logo')
        logo_path = r'C:\Users\luca.spagnuolo.ext\Downloads\logoConsip.jpg'
        
        # Aggiungi l'immagine nella cella A1
        worksheet.insert_image('A1', logo_path)

        # Analisi Servizi IT
        if do_servizi_it:
            servizi_it = df['Servizio IT'].unique()
            for servizio in servizi_it:
                df_servizio = df[df['Servizio IT'] == servizio]
                aree_utenti = df_servizio.groupby(['Area', 'Member Name'])['Group Name'].apply(list).reset_index()
                gruppi_servizio = df_servizio['Group Name'].unique()
                
                presenza_df = pd.DataFrame(columns=['Area', 'Utenza'] + list(gruppi_servizio))
                presenza_df[['Area', 'Utenza']] = aree_utenti[['Area', 'Member Name']]
                
                for idx, row in aree_utenti.iterrows():
                    gruppi_utente = row['Group Name']
                    for gruppo in gruppi_servizio:
                        presenza_df.at[idx, gruppo] = '1' if gruppo in gruppi_utente else ''
                
                sheet_name = servizio[:31]
                presenza_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Analisi Famiglie
        if do_famiglie:
            def gruppi_comuni_per_divisione(df):
                esclusi = {'guglielmo.gardenghi', 'mariella.ventriglia', 'benedetta.patane', 'alfredo.monaco'}
                df = df[~df['Member Name'].isin(esclusi)]
                divisioni = df['Area'].unique()
                risultato = []
                
                for divisione in divisioni:
                    df_div = df[df['Area'] == divisione]
                    num_membri = df_div['Member Name'].nunique()
                    gruppi_comuni = set(df_div['Group Name'])
                    
                    for membro in df_div['Member Name'].unique():
                        gruppi_membro = set(df_div[df_div['Member Name'] == membro]['Group Name'])
                        gruppi_comuni &= gruppi_membro
                    
                    gruppi_non_comuni = set(df_div['Group Name']) - gruppi_comuni
                    membri_possessori = {gruppo: list(df_div[df_div['Group Name'] == gruppo]['Member Name']) for gruppo in gruppi_non_comuni}
                    membri_mancanti = {gruppo: list(set(df_div['Member Name']) - set(membri_possessori.get(gruppo, []))) for gruppo in gruppi_non_comuni}
                    
                    colonne_famiglia = {}
                    for famiglia in df['Famiglia'].unique():
                        gruppi_comuni_famiglia = [g for g in gruppi_comuni if famiglia in g]
                        gruppi_non_comuni_famiglia = [g for g in gruppi_non_comuni if famiglia in g]
                        
                        colonne_famiglia[f'{famiglia} comuni'] = ', '.join(gruppi_comuni_famiglia) if gruppi_comuni_famiglia else '-'
                        colonne_famiglia[f'{famiglia} non comuni'] = ', '.join(f'{g} ({len(membri_possessori.get(g, []))})' for g in gruppi_non_comuni_famiglia) if gruppi_non_comuni_famiglia else '-'
                        colonne_famiglia[f'{famiglia} Dettaglio Utenti non comuni'] = '; '.join(f'{g}: {", ".join(m)}' for g, m in membri_possessori.items() if g in gruppi_non_comuni_famiglia) if gruppi_non_comuni_famiglia else '-'
                        colonne_famiglia[f'{famiglia} Dettaglio Utenti Mancanti non comuni'] = '; '.join(f'{g}: {", ".join(m)}' for g, m in membri_mancanti.items() if g in gruppi_non_comuni_famiglia) if gruppi_non_comuni_famiglia else '-'
                    
                    risultato.append({'Area': divisione, 'Numero Utenti': num_membri} | colonne_famiglia)
                
                return pd.DataFrame(risultato)
            
            df_famiglie = gruppi_comuni_per_divisione(df)
            for famiglia in df['Famiglia'].unique():
                df_famiglia = df_famiglie[['Area', 'Numero Utenti'] + [col for col in df_famiglie.columns if famiglia in col]]
                df_famiglia.to_excel(writer, sheet_name=famiglia[:31], index=False)
    
    output.seek(0)
    
    st.download_button(
        label="📥 Scarica il file Excel generato",
        data=output,
        file_name="analisi_IT_con_logo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
