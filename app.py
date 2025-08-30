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
        
        if not df_storico.empty:
            df_storico['data'] = pd.to_datetime(df_storico['data'])
            df_storico['data'] = df_storico['data'].dt.strftime('%Y-%m-%d')
            df_storico['docente'] = df_storico['docente'].astype(str).str.strip()
        if not df_assenze.empty:
            df_assenze['data'] = pd.to_datetime(df_assenze['data'])
            df_assenze['data'] = df_assenze['data'].dt.strftime('%Y-%m-%d')
            df_assenze['docente'] = df_assenze['docente'].astype(str).str.strip()
            
        return df_storico, df_assenze
    except Exception as e:
        st.error(f"Errore nel caricamento delle statistiche da Google Sheets: {e}")
        return pd.DataFrame(columns=['data', 'giorno', 'docente', 'ore']), pd.DataFrame(columns=['data', 'giorno', 'docente', 'ora', 'classe'])

def salva_storico_assenze(data_sostituzione, giorno_assente, sostituzioni_df, ore_assenti):
    try:
        client = get_gdrive_client()
        storico_sheet = client.open(ORARIO_SHEET_NAME).worksheet("storico")
        assenze_sheet = client.open(ORARIO_SHEET_NAME).worksheet("assenze")
        
        storico_data = [[
            str(data_sostituzione),
            giorno_assente,
            row["Sostituto"],
            1
        ] for _, row in sostituzioni_df.iterrows() if row["Sostituto"] != "Nessuno"]
        
        if storico_data:
            storico_sheet.append_rows(storico_data)
        
        assenze_data = [[
            str(data_sostituzione),
            giorno_assente,
            row["Docente Assente"],
            row["Ora"],
            row["Classe"]
        ] for _, row in ore_assenti.iterrows()]
        
        if assenze_data:
            assenze_sheet.append_rows(assenze_data)
            
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio dei dati su Google Sheets: {e}")
        return False
        
def clear_sheet_content(sheet_name):
    try:
        client = get_gdrive_client()
        sheet = client.open(ORARIO_SHEET_NAME).worksheet(sheet_name)
        sheet.clear()
        if sheet_name == "storico":
            sheet.append_row(["data", "giorno", "docente", "ore"])
        elif sheet_name == "assenze":
            sheet.append_row(["data", "giorno", "docente", "ora", "classe"])
        return True
    except Exception as e:
        st.error(f"Errore nell'azzeramento del foglio {sheet_name}: {e}")
        return False

def download_orario(df):
    if not df.empty:
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="⬇️ Scarica orario in CSV",
            data=csv,
            file_name="orario.csv",
            mime="text/csv"
        )

def vista_pivot_docenti(df, mode="classi"):
    if df.empty:
        st.warning("Nessun orario disponibile.")
        return

    dfp = df.copy()
    def format_cell(row):
        base = row["Docente"]
        return f"[S] {base}" if "Sostegno" in str(row["Tipo"]) else base

    if mode == "docenti":
        dfp["Info"] = dfp.apply(format_cell, axis=1)
        pivot = dfp.pivot_table(
            index="Ora",
            columns="Giorno",
            values="Info",
            aggfunc=lambda x: " / ".join(x)
        ).fillna("")
        styled = pivot.style.set_properties(**{"text-align": "center"}).map(lambda val: "color: green; font-weight: bold;" if "[S]" in str(val) else "color: blue;" if str(val).strip() != "" else "")
        st.dataframe(styled, use_container_width=True)

    elif mode == "classi":
        dfp["Info"] = dfp["Docente"]
        pivot = dfp.pivot_table(
            index=["Giorno", "Ora"],
            columns="Classe",
            values="Info",
            aggfunc=lambda x: " / ".join(list(dict.fromkeys(x)))
        ).fillna("-")

        def sort_classi(classe):
            m = re.match(r"(\d+)\s*([A-Za-zÀ-ÖØ-öø-ÿ]+)", str(classe))
            if m:
                return (int(m.group(1)), m.group(2).upper())
            return (10**9, str(classe))
        pivot = pivot.reindex(sorted(pivot.columns, key=sort_classi), axis=1)

        ordine_giorni = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]
        ordine_ore = ["I", "II", "III", "IV", "V", "VI"]
        pivot = pivot.reindex(pd.MultiIndex.from_product([ordine_giorni, ordine_ore], names=["Giorno", "Ora"]))

        pivot = pivot.reset_index()

        new_rows = []
        last_giorno = None
        for _, row in pivot.iterrows():
            if row["Giorno"] != last_giorno and last_giorno is not None:
                empty_row = {col: "" for col in pivot.columns}
                empty_row["Giorno"] = "──"
                new_rows.append(empty_row)
            new_rows.append(row.to_dict())
            last_giorno = row["Giorno"]

        pivot = pd.DataFrame(new_rows, columns=pivot.columns)

        last_giorno = None
        for i in range(len(pivot)):
            if pivot.loc[i, "Giorno"] == last_giorno and pivot.loc[i, "Giorno"] not in ["", "──"]:
                pivot.loc[i, "Giorno"] = ""
            else:
                last_giorno = pivot.loc[i, "Giorno"]

        def color_cells_classi(val):
            if isinstance(val, str) and "[S]" in val:
                return "color: green; font-weight: bold;"
            return ""

        styled = pivot.style.set_properties(**{"text-align": "center"}).map(color_cells_classi)
        st.dataframe(styled, use_container_width=True, hide_index=True)

# =========================
# AVVIO APP
# =========================
st.title("📚 Gestione orari e Sostituzioni Docenti")
orario_df = carica_orario()

# =========================
# MENU PRINCIPALE
# =========================
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
                st.success("Orario caricato con successo ✅")
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
            ore_assenti = orario_df[
                (orario_df["Docente"].isin(docenti_assenti)) & 
                (orario_df["Giorno"] == giorno_assente) &
                (~orario_df["Escludi"])
            ]
            
            if ore_assenti.empty:
                st.info("I docenti selezionati non hanno lezioni in questo giorno.")
            else:
                st.subheader("📌 Ore scoperte")
                st.dataframe(ore_assenti[["Docente", "Ora", "Classe"]].drop_duplicates(), use_container_width=True, hide_index=True)
                
                sostituzioni_proposte = []
                st.subheader("🔄 Proponi Sostituzioni")
                
                for _, r in ore_assenti.iterrows():
                    classe = r["Classe"]
                    ora = r["Ora"]
                    docente_assente = r["Docente"]
                    
                    docenti_occupati_in_ora = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == ora)
                    ]["Docente"].tolist()
                    
                    proposto = "Nessuno"
                    
                    # Logica originale per la proposta del sostituto
                    prioritari = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == ora) &
                        (orario_df["Tipo"] == "Sostegno") &
                        (orario_df["Classe"] == classe)
                    ]
                    
                    if not prioritari.empty:
                        proposto = prioritari["Docente"].iloc[0]
                    else:
                        sostegni_disponibili = orario_df[
                            (orario_df["Tipo"] == "Sostegno")
                        ]["Docente"].unique()
                        
                        liberi_in_ora = [d for d in orario_df["Docente"].unique() if d not in docenti_occupati_in_ora]
                        
                        candidati = [d for d in sostegni_disponibili if d in liberi_in_ora and d != docente_assente]
                        
                        if candidati:
                            proposto = candidati[0]
                        else:
                            proposto = "Nessuno"

                    # Costruisci opzioni con prefisso [S] per i sostegni
                    sostegni = orario_df[orario_df["Tipo"] == "Sostegno"]["Docente"].unique()
                    opzioni_validi = [d for d in orario_df["Docente"].unique() if d not in docenti_assenti and d not in docenti_occupati_in_ora]
                    
                    opzioni = ["Nessuno"] + sorted(
                        opzioni_validi,
                        key=lambda x: (0 if x in sostegni else 1, x)
                    )

                    sostituto_scelto = st.selectbox(
                        f"Sostituto per **{docente_assente}** in **{classe}** alla **{ora}** ora:",
                        options=opzioni,
                        index=opzioni.index(proposto) if proposto in opzioni else 0,
                        key=f"select_{docente_assente}_{classe}_{ora}"
                    )
                    
                    sostituzioni_proposte.append({
                        "Docente Assente": docente_assente,
                        "Classe": classe,
                        "Ora": ora,
                        "Sostituto": sostituto_scelto
                    })

                sostituzioni_df = pd.DataFrame(sostituzioni_proposte)
                st.dataframe(sostituzioni_df, use_container_width=True, hide_index=True)

                if st.button("Conferma e salva sostituzioni"):
                    if salva_storico_assenze(data_sostituzione, giorno_assente, sostituzioni_df, ore_assenti):
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
        df_storico_agg = df_storico.groupby("docente")["ore"].sum().reset_index()
        df_storico_agg = df_storico_agg.rename(columns={"ore": "Totale Ore Sostituite"})
        st.dataframe(df_storico_agg, use_container_width=True, hide_index=True)
        st.bar_chart(df_storico_agg, x="docente", y="Totale Ore Sostituite")
    else:
        st.info("Nessuna sostituzione registrata.")

    if not df_assenze.empty:
        st.subheader("Statistiche Assenze")
        df_assenze_agg = df_assenze.groupby("docente")["ora"].count().reset_index()
        df_assenze_agg = df_assenze_agg.rename(columns={"ora": "Totale Ore Assenti"})
        st.dataframe(df_assenze_agg, use_container_width=True, hide_index=True)
        st.bar_chart(df_assenze_agg, x="docente", y="Totale Ore Assenti")
    else:
        st.info("Nessuna assenza registrata.")
        
    st.subheader("⚠️ Azzeramento dati storici")
    conferma_storico = st.checkbox("Confermo di voler cancellare definitivamente lo storico delle sostituzioni")
    conferma_assenze = st.checkbox("Confermo di voler cancellare definitivamente lo storico delle assenze")
    
    if st.button("Elimina storico"):
        if conferma_storico:
            if clear_sheet_content("storico"):
                st.success("Storico delle sostituzioni eliminato ✅")
        if conferma_assenze:
            if clear_sheet_content("assenze"):
                st.success("Storico delle assenze eliminato ✅")
        if not conferma_storico and not conferma_assenze:
            st.warning("Devi spuntare almeno una delle conferme per eliminare i dati.")
