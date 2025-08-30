import streamlit as st
import pandas as pd
import os
import re
import gspread
import gspread_dataframe as gd
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build

# =========================
# CONFIGURAZIONE GENERALE
# =========================
REQUIRED_COLUMNS = ["Docente", "Giorno", "Ora", "Classe", "Tipo", "Escludi"]
ORARIO_SHEET_NAME = "OrarioSostituzioni"

# =========================
# FUNZIONI DI SUPPORTO PER GESTIONE DATI
# =========================
def get_gdrive_client():
    gdrive_credentials = st.secrets["gdrive"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(gdrive_credentials, scope)
    client = gspread.authorize(creds)
    return client

def carica_orario():
    try:
        client = get_gdrive_client()
        sheet = client.open(ORARIO_SHEET_NAME).worksheet("orario")
        df = gd.get_as_dataframe(sheet)
        df = df.dropna(how='all')
        # Pulisco e converto i tipi di colonna per maggiore robustezza
        df['Escludi'] = df['Escludi'].astype(bool)
        df['Tipo'] = df['Tipo'].astype(str).str.strip()
        df['Docente'] = df['Docente'].astype(str).str.strip()
        df['Giorno'] = df['Giorno'].astype(str).str.strip()
        df['Ora'] = df['Ora'].astype(str).str.strip()
        df['Classe'] = df['Classe'].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Errore nel caricamento dell'orario da Google Sheets: {e}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

def salva_orario(df):
    try:
        client = get_gdrive_client()
        sheet = client.open(ORARIO_SHEET_NAME).worksheet("orario")
        gd.set_with_dataframe(sheet, df)
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio dell'orario su Google Sheets: {e}")
        return False

def carica_statistiche():
    try:
        client = get_gdrive_client()
        storico_sheet = client.open(ORARIO_SHEET_NAME).worksheet("storico")
        df_storico = gd.get_as_dataframe(storico_sheet)
        assenze_sheet = client.open(ORARIO_SHEET_NAME).worksheet("assenze")
        df_assenze = gd.get_as_dataframe(assenze_sheet)
        return df_storico, df_assenze
    except Exception as e:
        st.error(f"Errore nel caricamento delle statistiche da Google Sheets: {e}")
        return pd.DataFrame(), pd.DataFrame()

def salva_sostituzione(data, giorno, sostituzioni, assenze):
    try:
        client = get_gdrive_client()
        storico_sheet = client.open(ORARIO_SHEET_NAME).worksheet("storico")
        assenze_sheet = client.open(ORARIO_SHEET_NAME).worksheet("assenze")
        
        # Salva sostituzioni
        storico_data = [[data, giorno, r["Sostituto"], 1] for _, r in sostituzioni.iterrows() if r["Sostituto"] != "Nessuno"]
        storico_sheet.append_rows(storico_data)
        
        # Salva assenze
        assenze_data = [[data, giorno, r['Docente Assente'], r['Ora'], r['Classe']] for _, r in sostituzioni.iterrows()]
        assenze_sheet.append_rows(assenze_data)
        
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio dei dati su Google Sheets: {e}")
        return False

def vista_pivot_docenti(df, mode="docenti"):
    df_temp = df.copy()
    if df_temp.empty:
        st.warning("Nessun orario disponibile.")
        return
    df_temp["Info"] = df_temp.apply(lambda row: f"[S] {row['Docente']}" if "Sostegno" in str(row["Tipo"]) else row["Docente"], axis=1)
    if mode == "docenti":
        pivot = df_temp.pivot_table(
            index="Ora",
            columns="Giorno",
            values="Info",
            aggfunc=lambda x: " / ".join(x)
        ).fillna("-")
    elif mode == "classi":
        pivot = df_temp.pivot_table(
            index=["Giorno", "Ora"],
            columns="Classe",
            values="Info",
            aggfunc=lambda x: " / ".join(sorted(list(dict.fromkeys(x)), key=lambda s: 0 if not "[S]" in s else 1))
        ).fillna("-")
    def color_cells(val):
        text = str(val)
        if "[S]" in text:
            return "color: green; font-weight: bold;"
        elif text.strip() != "":
            return "color: blue;"
        return ""
    styled = pivot.style.set_properties(**{
            "text-align": "center",
            "font-size": "8pt"
        }).map(color_cells)
    st.dataframe(styled, use_container_width=True)

def download_orario(df):
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Scarica Orario (CSV)",
        data=csv,
        file_name="orario_aggiornato.csv",
        mime="text/csv",
    )

# =========================
# LOGICA APPLICAZIONE
# =========================
st.title("📚 Gestione orari e Sostituzioni Docenti")
orario_df = carica_orario()

menu = st.sidebar.radio(
    "Naviga",
    ["Inserisci/Modifica Orario", "Gestione Assenze", "Visualizza Orario", "Statistiche"]
)

# --- INSERIMENTO/MODIFICA ORARIO ---
if menu == "Inserisci/Modifica Orario":
    st.header("➕ Inserisci o modifica l'orario")
    uploaded_file = st.file_uploader("Carica un nuovo orario (CSV)", type="csv")
    if uploaded_file:
        df_tmp = pd.read_csv(uploaded_file)
        if all(col in df_tmp.columns for col in REQUIRED_COLUMNS):
            orario_df = df_tmp
            if salva_orario(orario_df):
                st.success("Orario caricato e salvato su Google Sheets ✅")
        else:
            st.error(f"CSV non valido. Deve contenere: {REQUIRED_COLUMNS}")

    with st.expander("Aggiungi una nuova lezione"):
        docente = st.selectbox(
            "Nome docente",
            ["➕ Nuovo docente"] + (sorted(orario_df["Docente"].unique()) if not orario_df.empty else []),
            key="docente_input"
        )
        if docente == "➕ Nuovo docente":
            docente = st.text_input("Inserisci nuovo docente")
        giorno = st.selectbox("Giorno", ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"])
        ora = st.selectbox("Ora", ["I", "II", "III", "IV", "V", "VI"])
        classe = st.selectbox(
            "Classe",
            ["➕ Nuova classe"] + (sorted(orario_df["Classe"].unique()) if not orario_df.empty else []),
            key="classe_input"
        )
        if classe == "➕ Nuova classe":
            classe = st.text_input("Inserisci nuova classe")
        tipo = st.selectbox("Tipo", ["Lezione", "Sostegno", "Altro"])
        escludi = st.checkbox("Escludi da sostituzioni")
        if st.button("Aggiungi", key="add_lesson"):
            if docente and giorno and ora and classe and tipo:
                row = {
                    "Docente": docente,
                    "Giorno": giorno,
                    "Ora": ora,
                    "Classe": classe,
                    "Tipo": tipo,
                    "Escludi": escludi
                }
                nuovo = pd.DataFrame([row])
                conflitto = not orario_df[
                    (orario_df["Docente"] == docente) &
                    (orario_df["Giorno"] == giorno) &
                    (orario_df["Ora"] == ora)
                ].empty
                if conflitto:
                    st.error("Conflitto: questo docente è già assegnato a un'altra classe in questa ora.")
                else:
                    orario_df = pd.concat([orario_df, nuovo], ignore_index=True)
                    if salva_orario(orario_df):
                        st.success("Lezione aggiunta all'orario e salvata su Google Sheets ✅")
            else:
                st.error("Compila tutti i campi per aggiungere una lezione.")

    st.subheader("📝 Modifica orario attuale")
    if not orario_df.empty:
        col_order = REQUIRED_COLUMNS
        df_edit = orario_df[col_order]
        df_edit = df_edit.sort_values(by=["Docente"])
        edited_df = st.data_editor(df_edit, use_container_width=True, hide_index=True)
        if st.button("Salva modifiche"):
            orario_df = edited_df
            if salva_orario(orario_df):
                st.success("Orario modificato e salvato su Google Sheets ✅")
    download_orario(orario_df)

# --- GESTIONE ASSENZE ---
elif menu == "Gestione Assenze":
    st.header("🚨 Gestione Assenze")
    if orario_df.empty:
        st.warning("Non hai ancora caricato nessun orario.")
    else:
        docenti_assenti = st.multiselect("Seleziona docenti assenti", sorted(orario_df["Docente"].unique()))
        data_sostituzione = st.date_input("Data della sostituzione")
        giorno_assente = data_sostituzione.strftime("%A")
        traduzione_giorni = {
            "Monday": "Lunedì", "Tuesday": "Martedì", "Wednesday": "Mercoledì", "Thursday": "Giovedì",
            "Friday": "Venerdì", "Saturday": "Sabato", "Sunday": "Domenica"
        }
        giorno_assente = traduzione_giorni.get(giorno_assente, giorno_assente)
        
        if not docenti_assenti:
            st.info("Seleziona almeno un docente per continuare.")
        else:
            assenze = orario_df[
                (orario_df["Docente"].isin(docenti_assenti)) & 
                (orario_df["Giorno"] == giorno_assente) &
                (~orario_df["Escludi"])
            ]
            
            if assenze.empty:
                st.info("I docenti selezionati non hanno lezioni in questo giorno.")
            else:
                st.subheader("📌 Ore scoperte")
                st.dataframe(assenze[["Docente", "Ora", "Classe"]].drop_duplicates(), use_container_width=True, hide_index=True)
                
                sostituzioni = []
                st.subheader("🔄 Proponi Sostituzioni")
                
                # 🔹 Cerca i docenti di sostegno
                docenti_sostegno = orario_df[orario_df["Tipo"] == "Sostegno"]["Docente"].unique().tolist()
                docenti_curricolari = [d for d in orario_df["Docente"].unique().tolist() if d not in docenti_sostegno]
                
                for _, r in assenze.iterrows():
                    classe = r["Classe"]
                    ora = r["Ora"]
                    docente_assente = r["Docente"]
                    tipo_lezione = r["Tipo"]
                    
                    # 🔹 Trova i docenti già occupati in quell'ora
                    docenti_occupati = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == ora) &
                        (orario_df["Docente"] != docente_assente)
                    ]["Docente"].tolist()
                    
                    # 🔹 Sostituti disponibili
                    disponibili = [d for d in orario_df["Docente"].unique().tolist() if d not in docenti_occupati and d not in docenti_assenti]
                    
                    # 🔹 Proposta iniziale
                    proposta_sostituto = "Nessuno"
                    if "Sostegno" in str(tipo_lezione):
                        # Priorità 1: Sostegno libero della stessa classe
                        sostegno_stessa_classe = orario_df[
                            (orario_df["Classe"] == classe) &
                            (orario_df["Tipo"].str.contains("Sostegno", case=False, na=False)) &
                            (orario_df["Docente"].isin(disponibili))
                        ]
                        if not sostegno_stessa_classe.empty:
                            proposta_sostituto = sostegno_stessa_classe.iloc[0]["Docente"]
                        else:
                            # Priorità 2: Altri docenti di sostegno liberi
                            altri_sostegno_disponibili = [d for d in docenti_sostegno if d in disponibili]
                            if altri_sostegno_disponibili:
                                proposta_sostituto = altri_sostegno_disponibili[0]
                    else:
                        # Priorità 3: Docenti curricolari liberi
                        curricolari_disponibili = [d for d in docenti_curricolari if d in disponibili]
                        if curricolari_disponibili:
                            proposta_sostituto = curricolari_disponibili[0]
                            
                    opzioni_sostituto = ["Nessuno"] + sorted(disponibili)
                    
                    sostituto_scelto = st.selectbox(
                        f"Sostituto per **{docente_assente}** in **{classe}** alla **{ora}** ora:",
                        options=opzioni_sostituto,
                        index=opzioni_sostituto.index(proposta_sostituto) if proposta_sostituto in opzioni_sostituto else 0,
                        key=f"select_{docente_assente}_{classe}_{ora}"
                    )
                    
                    sostituzioni.append({
                        "Docente Assente": docente_assente,
                        "Classe": classe,
                        "Ora": ora,
                        "Sostituto": sostituto_scelto
                    })

                sostituzioni_df = pd.DataFrame(sostituzioni)

                if st.button("Conferma e salva sostituzioni"):
                    if salva_sostituzione(pd.Timestamp.now().strftime("%Y-%m-%d"), giorno_assente, sostituzioni_df, assenze):
                        st.success("Sostituzioni salvate con successo su Google Sheets! ✅")

# --- VISUALIZZA ORARIO ---
elif menu == "Visualizza Orario":
    st.header("📅 Orario completo")
    if orario_df.empty:
        st.warning("Nessun orario disponibile.")
    else:
        vista_pivot_docenti(orario_df, mode="classi")
        download_orario(orario_df)

# --- STATISTICHE ---
elif menu == "Statistiche":
    st.header("📊 Statistiche Sostituzioni e Assenze")
    df_storico, df_assenze = carica_statistiche()
    if not df_storico.empty:
        st.subheader("Statistiche Sostituzioni")
        st.dataframe(df_storico, use_container_width=True, hide_index=True)
        st.bar_chart(df_storico.groupby("docente")["ore"].sum())
    else:
        st.info("Nessuna sostituzione registrata.")

    if not df_assenze.empty:
        st.subheader("Statistiche Assenze")
        st.dataframe(df_assenze, use_container_width=True, hide_index=True)
        st.bar_chart(df_assenze.groupby("docente")["totale_ore_assenti"].sum())
    else:
        st.info("Nessuna assenza registrata.")
