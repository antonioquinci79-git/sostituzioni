import streamlit as st
import pandas as pd
import os
import re
import io
import zipfile
import gspread
import gspread_dataframe as gd
from google.oauth2.service_account import Credentials
from datetime import datetime

# =========================
# STILI PERSONALIZZATI
# =========================
st.markdown("""
<style>
/* Bottoni più grandi */
.stButton button {
    width: 100%;
    padding: 0.8em;
    font-size: 1.05em;
    border-radius: 12px;
}

/* Tabelle: font più piccolo e leggibile */
.stDataFrame, .stDataEditor {
    font-size: 0.9em !important;
}

/* Input: allarga i selectbox */
.stSelectbox, .stTextInput, .stDateInput, .stMultiSelect {
    width: 100% !important;
}

/* Margini orizzontali per respirare su mobile */
.block-container {
    padding-left: 1rem;
    padding-right: 1rem;
}

/* Pills: dimensione tocco comoda su mobile */
[data-testid="stPills"] button {
    font-size: 1em !important;
    padding: 0.5em 0.9em !important;
    border-radius: 20px !important;
    margin: 3px !important;
    min-height: 2.4em !important;
}
</style>
""", unsafe_allow_html=True)

# =========================
# CONFIGURAZIONE FILE / SHEETS
# =========================
REQUIRED_COLUMNS = ["Docente", "Giorno", "Ora", "Classe", "Tipo", "Escludi"]
SPREADSHEET_NAME = "OrarioSostituzioni"
ORARIO_SHEET = "orario"
STORICO_SHEET = "storico"
ASSENZE_SHEET = "assenze"
GIORNI_SETTIMANA = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]
ORE_LEZIONE = ["I", "II", "III", "IV", "V", "VI"]
TIPI_LEZIONE = ["Lezione", "Sostegno", "Altro"]

# =========================
# CLIENT GOOGLE DRIVE
# =========================
SCOPE = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

@st.cache_resource(show_spinner=False)
def get_gdrive_client():
    gdrive_credentials = st.secrets["gdrive"]
    creds = Credentials.from_service_account_info(gdrive_credentials, scopes=SCOPE)
    client = gspread.authorize(creds)
    return client

@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    """Apre (o crea) lo spreadsheet UNA SOLA VOLTA per la vita dell'app
    (cache_resource), invece di rifare la ricerca per nome ad ogni rerun."""
    client = get_gdrive_client()
    try:
        return client.open(SPREADSHEET_NAME)
    except gspread.SpreadsheetNotFound:
        # Creazione: l'utente dovrà eventualmente condividere il foglio
        # con l'email del service account se vuole accedervi anche da browser.
        return client.create(SPREADSHEET_NAME)

@st.cache_resource(show_spinner=False)
def get_worksheet(sheet_name: str):
    """Restituisce l'handle del worksheet, creandolo con l'header corretto
    se non esiste. Cachato: una volta risolto l'handle, i rerun successivi
    non fanno più alcuna chiamata 'metadata' a Google, solo letture/scritture
    sui valori quando effettivamente richieste."""
    headers_by_sheet = {
        ORARIO_SHEET: REQUIRED_COLUMNS,
        STORICO_SHEET: ["data", "giorno", "docente", "ore"],
        ASSENZE_SHEET: ["data", "giorno", "docente", "ora", "classe"],
    }
    sh = get_spreadsheet()
    try:
        return sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows="200", cols="20")
        header_df = pd.DataFrame(columns=headers_by_sheet.get(sheet_name, []))
        gd.set_with_dataframe(ws, header_df, include_index=False, include_column_header=True)
        return ws

# =========================
# INIZIALIZZAZIONE FOGLI (se mancanti creali con header corretti)
# =========================
def ensure_sheets_exist():
    """Pre-carica gli handle dei worksheet. Grazie al cache_resource su
    get_worksheet, dal secondo rerun in poi questa funzione non genera
    alcuna chiamata di rete."""
    for nome_foglio in (ORARIO_SHEET, STORICO_SHEET, ASSENZE_SHEET):
        get_worksheet(nome_foglio)

# =========================
# CARICAMENTO / SALVATAGGIO ORARIO
# =========================
@st.cache_data(ttl=300, show_spinner=False)
def carica_orario():
    try:
        ws = get_worksheet(ORARIO_SHEET)
        df = gd.get_as_dataframe(ws, evaluate_formulas=True, header=0)
        # rimuovo colonne totalmente vuote
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        # riempio eventuali colonne mancanti con default
        for col in REQUIRED_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        # drop rows completamente vuote
        df = df.dropna(how='all').copy()
        # normalizzazione tipi
        if "Escludi" in df.columns:
            # Google Sheets può portare True/False o stringhe
            df["Escludi"] = df["Escludi"].replace({pd.NA: False, "": False}).astype(bool)
        else:
            df["Escludi"] = False
        for col in ["Tipo", "Docente", "Giorno", "Ora", "Classe"]:
            df[col] = df[col].astype(str).str.strip().fillna("")
        # mantieni solo le colonne richieste
        df = df.loc[:, REQUIRED_COLUMNS]
        return df
    except Exception as e:
        st.error(f"Errore nel caricamento dell'orario da Google Sheets: {e}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

def salva_orario(df):
    try:
        ws = get_worksheet(ORARIO_SHEET)
        # assicurati che le colonne siano quelle giuste e in ordine
        df_to_save = df.copy()
        for col in REQUIRED_COLUMNS:
            if col not in df_to_save.columns:
                df_to_save[col] = ""
        df_to_save = df_to_save[REQUIRED_COLUMNS]
        gd.set_with_dataframe(ws, df_to_save, include_index=False, include_column_header=True)
        carica_orario.clear()  # invalida la cache: i dati appena salvati saranno subito visibili
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio dell'orario su Google Sheets: {e}")
        return False

# =========================
# CARICAMENTO / SALVATAGGIO STATISTICHE (storico + assenze)
# =========================
@st.cache_data(ttl=300, show_spinner=False)
def carica_statistiche():
    try:
        ws_storico = get_worksheet(STORICO_SHEET)
        ws_assenze = get_worksheet(ASSENZE_SHEET)
        df_storico = gd.get_as_dataframe(ws_storico, header=0).dropna(how='all')
        df_assenze = gd.get_as_dataframe(ws_assenze, header=0).dropna(how='all')
        # Normalizza nomi e tipi
        if not df_storico.empty:
            if "data" in df_storico.columns:
                df_storico["data"] = pd.to_datetime(df_storico["data"], errors="coerce")
                df_storico["data"] = df_storico["data"].dt.strftime("%Y-%m-%d")

            if "ore" in df_storico.columns:
                df_storico["ore"] = pd.to_numeric(df_storico["ore"], errors="coerce").fillna(0).astype(int)

            for c in ["docente", "giorno"]:
                if c in df_storico.columns:
                    df_storico[c] = df_storico[c].astype(str).str.strip().str.lower()
        if not df_assenze.empty:
            if "data" in df_assenze.columns:
                df_assenze["data"] = pd.to_datetime(df_assenze["data"], errors="coerce")
                df_assenze["data"] = df_assenze["data"].dt.strftime("%Y-%m-%d")
            for c in ["docente", "giorno", "ora", "classe"]:
                if c in df_assenze.columns:
                    df_assenze[c] = df_assenze[c].astype(str).str.strip()
        # Se i fogli sono vuoti, restituisci DataFrame con le colonne attese
        if df_storico.empty:
            df_storico = pd.DataFrame(columns=["data", "giorno", "docente", "ore"])
        if df_assenze.empty:
            df_assenze = pd.DataFrame(columns=["data", "giorno", "docente", "ora", "classe"])
        return df_storico, df_assenze
    except Exception as e:
        st.error(f"Errore nel caricamento delle statistiche da Google Sheets: {e}")
        return pd.DataFrame(columns=["data", "giorno", "docente", "ore"]), pd.DataFrame(columns=["data", "giorno", "docente", "ora", "classe"])


def salva_storico_assenze(data_sostituzione, giorno_assente, sostituzioni_df, ore_assenti):
    try:
        ws_storico = get_worksheet(STORICO_SHEET)
        ws_assenze = get_worksheet(ASSENZE_SHEET)

        # Filtra solo le sostituzioni effettive (esclude "Nessuno")
        sostituzioni_effettive = sostituzioni_df[
            sostituzioni_df["Sostituto"].notna() &
            (sostituzioni_df["Sostituto"].str.strip() != "") &
            (sostituzioni_df["Sostituto"].str.strip().str.lower() != "nessuno")
        ].copy()

        # Prepara i dati per lo storico: ogni sostituzione effettiva vale 1 ora
        storico_data = [
            [str(data_sostituzione), giorno_assente, row["Sostituto"], 1]
            for _, row in sostituzioni_effettive.iterrows()
        ]

        if storico_data:
            ws_storico.append_rows(storico_data, value_input_option="USER_ENTERED")

        # Filtra le assenze effettive: includi solo le ore dove il docente è stato effettivamente sostituito
        ore_effettivamente_assenti = ore_assenti[
            ore_assenti["Ora"].isin(sostituzioni_effettive["Ora"])
        ].copy()

        assenze_data = [
            [str(data_sostituzione), giorno_assente, row["Docente"], row["Ora"], row["Classe"]]
            for _, row in ore_effettivamente_assenti.iterrows()
        ]

        if assenze_data:
            ws_assenze.append_rows(assenze_data, value_input_option="USER_ENTERED")

        carica_statistiche.clear()  # invalida la cache: storico/assenze aggiornati saranno subito visibili
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio dei dati su Google Sheets: {e}")
        return False

def clear_sheet_content(sheet_name):
    try:
        ws = get_worksheet(sheet_name)
        ws.clear()
        if sheet_name == STORICO_SHEET:
            ws.append_row(["data", "giorno", "docente", "ore"])
        elif sheet_name == ASSENZE_SHEET:
            ws.append_row(["data", "giorno", "docente", "ora", "classe"])
        if sheet_name in (STORICO_SHEET, ASSENZE_SHEET):
            carica_statistiche.clear()
        elif sheet_name == ORARIO_SHEET:
            carica_orario.clear()
        return True
    except Exception as e:
        st.error(f"Errore nell'azzeramento del foglio {sheet_name}: {e}")
        return False

# =========================
# FUNZIONE PER IL BACKUP CORRETTA
# =========================
def create_backup():
    try:
        # Esporta il foglio "orario" in CSV
        ws_orario = get_worksheet(ORARIO_SHEET)
        df_orario = gd.get_as_dataframe(ws_orario, evaluate_formulas=True, header=0).dropna(how='all')

        # Esporta il foglio "storico" in CSV
        ws_storico = get_worksheet(STORICO_SHEET)
        df_storico = gd.get_as_dataframe(ws_storico, evaluate_formulas=True, header=0).dropna(how='all')

        # Esporta il foglio "assenze" in CSV
        ws_assenze = get_worksheet(ASSENZE_SHEET)
        df_assenze = gd.get_as_dataframe(ws_assenze, evaluate_formulas=True, header=0).dropna(how='all')

        # Comprimi i DataFrame in un file ZIP in memoria
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            # Aggiungi i file CSV al buffer
            zip_file.writestr("orario.csv", df_orario.to_csv(index=False).encode('utf-8'))
            zip_file.writestr("storico.csv", df_storico.to_csv(index=False).encode('utf-8'))
            zip_file.writestr("assenze.csv", df_assenze.to_csv(index=False).encode('utf-8'))

        # Restituisci il contenuto del file ZIP
        zip_buffer.seek(0)
        return zip_buffer
    except Exception as e:
        st.error(f"Errore durante la creazione del backup: {e}")
        return None

# =========================
# UTILITA' PER DOWNLOAD e PIVOT
# =========================
def build_docente_tipo_map(df):
    """Precalcola una mappa {docente: tipo} una sola volta, invece di
    rifare un filtro su orario_df per ogni docente dentro i loop."""
    if df.empty:
        return {}
    tmp = df[df["Tipo"].astype(str).str.strip() != ""]
    if tmp.empty:
        return {}
    return tmp.groupby("Docente")["Tipo"].first().to_dict()

def trova_conflitti_orario(df):
    """Ritorna le combinazioni (Docente, Giorno, Ora) che compaiono più di
    una volta in df, cioè un docente assegnato a due classi nella stessa ora."""
    if df.empty:
        return []
    conteggi = df.groupby(["Docente", "Giorno", "Ora"]).size()
    return [idx for idx, n in conteggi.items() if n > 1 and idx[0].strip() != ""]

def download_orario(df):
    if not df.empty:
        st.download_button(
            "⬇️ Scarica orario in CSV",
            data=df.to_csv(index=False),
            file_name="orario.csv",
            mime="text/csv"
        )

def vista_pivot_docenti(df, mode="docenti"):
    # Copiato e adattato da app.py per mantenere la stessa UX
    if df.empty:
        st.warning("Nessun orario disponibile.")
        return

    if mode == "docenti":
        dfp = df.copy()
        def format_cell(row):
            base = f"{row['Docente']} ({row['Classe']})"
            return f"[S] {base}" if "Sostegno" in str(row["Tipo"]) else base
        dfp["Info"] = dfp.apply(format_cell, axis=1)

        pivot = dfp.pivot_table(
            index="Ora",
            columns="Giorno",
            values="Info",
            aggfunc=lambda x: " / ".join(x)
        ).fillna("")

        def color_cells(val):
            text = str(val)
            if "[S]" in text:
                return "color: green; font-weight: bold;"
            elif text.strip() != "":
                return "color: blue;"
            return ""

        styled = pivot.style.map(color_cells)
        st.dataframe(styled, use_container_width=True)

    elif mode == "classi":
        dfp = df.copy()
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

        ordine_giorni = GIORNI_SETTIMANA
        ordine_ore = ORE_LEZIONE
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

        styled = pivot.style.set_properties(**{"text-align": "center"})
        st.dataframe(styled, use_container_width=True, hide_index=True)

# =========================
# AVVIO APP
# =========================
st.title("📚Sostituzioni docenti")
# assicurati che i fogli esistano con le intestazioni
try:
    with st.spinner('Caricamento dati...'):
        ensure_sheets_exist()
except Exception as e:
    st.error(f"Impossibile inizializzare i fogli Google: {e}")

with st.spinner('Caricamento orario...'):
    orario_df = carica_orario()

# =========================
# MENU PRINCIPALE (mobile-friendly)
# =========================
menu = st.segmented_control(
"Navigazione",
["Inserisci/Modifica Orario", "Gestione Assenze", "Visualizza Orario", "Statistiche"],
selection_mode="single"
)



# --- INSERIMENTO/MODIFICA ORARIO ---
if menu == "Inserisci/Modifica Orario":
    st.header("➕ Inserisci o modifica l'orario")

    uploaded_file = st.file_uploader("Carica un nuovo orario (CSV)", type="csv")
    if uploaded_file:
        df_tmp = pd.read_csv(uploaded_file)
        if not all(col in df_tmp.columns for col in REQUIRED_COLUMNS):
            st.error(f"CSV non valido. Deve contenere: {REQUIRED_COLUMNS}")
        else:
            df_nuovo = df_tmp[REQUIRED_COLUMNS].copy()
            for col in ["Docente", "Giorno", "Ora", "Classe", "Tipo"]:
                df_nuovo[col] = df_nuovo[col].astype(str).str.strip()
            valori_non_validi = df_nuovo[
                ~df_nuovo["Giorno"].isin(GIORNI_SETTIMANA) | ~df_nuovo["Ora"].isin(ORE_LEZIONE)
            ]
            conflitti_csv = trova_conflitti_orario(df_nuovo)
            if not valori_non_validi.empty:
                st.error(
                    f"Il CSV contiene valori di Giorno/Ora non validi (attesi {GIORNI_SETTIMANA} e {ORE_LEZIONE}). "
                    "Correggi il file e ricaricalo."
                )
                st.dataframe(valori_non_validi, use_container_width=True, hide_index=True)
            elif conflitti_csv:
                st.error("Il CSV contiene docenti assegnati a più classi nella stessa ora:")
                for docente, giorno, ora in conflitti_csv:
                    st.write(f"- {docente}: {giorno} ora {ora}")
            else:
                orario_df = df_nuovo
                if salva_orario(orario_df):
                    st.success("Orario caricato con successo ✅")
                    st.rerun()

    # Inserimento nuova lezione
    with st.expander("Aggiungi una nuova lezione"):
        docente = st.selectbox(
            "Nome docente",
            ["➕ Nuovo docente"] + (sorted(orario_df["Docente"].unique()) if not orario_df.empty else []),
            key="docente_input"
        )
        if docente == "➕ Nuovo docente":
            docente = st.text_input("Inserisci nuovo docente")
        giorno = st.selectbox("Giorno", GIORNI_SETTIMANA)
        ora = st.selectbox("Ora", ORE_LEZIONE)
        classe = st.selectbox(
            "Classe",
            ["➕ Nuova classe"] + (sorted(orario_df["Classe"].unique()) if not orario_df.empty else []),
            key="classe_input"
        )
        if classe == "➕ Nuova classe":
            classe = st.text_input("Inserisci nuova classe")
        tipo = st.selectbox("Tipo", TIPI_LEZIONE)
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
                        st.rerun()
            else:
                st.error("Compila tutti i campi per aggiungere una lezione.")

    # Modifica orario esistente
    st.subheader("📝 Modifica orario attuale")
    if not orario_df.empty:
        col_order = REQUIRED_COLUMNS
        df_edit = orario_df[col_order].copy()
        df_edit = df_edit.sort_values(by=["Docente"])
        edited_df = st.data_editor(
            df_edit,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Giorno": st.column_config.SelectboxColumn("Giorno", options=GIORNI_SETTIMANA, required=True),
                "Ora": st.column_config.SelectboxColumn("Ora", options=ORE_LEZIONE, required=True),
                "Tipo": st.column_config.SelectboxColumn("Tipo", options=TIPI_LEZIONE, required=True),
                "Escludi": st.column_config.CheckboxColumn("Escludi"),
                "Docente": st.column_config.TextColumn("Docente", required=True),
                "Classe": st.column_config.TextColumn("Classe", required=True),
            }
        )
        if st.button("Salva modifiche"):
            df_da_salvare = edited_df.copy()
            for col in ["Docente", "Giorno", "Ora", "Classe", "Tipo"]:
                df_da_salvare[col] = df_da_salvare[col].astype(str).str.strip()
            # scarta righe vuote (es. nuove righe aggiunte e non compilate)
            df_da_salvare = df_da_salvare[df_da_salvare["Docente"] != ""].copy()

            conflitti_edit = trova_conflitti_orario(df_da_salvare)
            if conflitti_edit:
                st.error("Impossibile salvare: lo stesso docente è assegnato a più classi nella stessa ora:")
                for docente, giorno, ora in conflitti_edit:
                    st.write(f"- {docente}: {giorno} ora {ora}")
            else:
                orario_df = df_da_salvare
                if salva_orario(orario_df):
                    st.success("Orario modificato e salvato su Google Sheets ✅")
                    st.rerun()
    download_orario(orario_df)

# ... (codice precedente) ...

# --- GESTIONE ASSENZE ---
elif menu == "Gestione Assenze":
    st.header("🚨 Gestione Assenze")

    if orario_df.empty:
        st.warning("Non hai ancora caricato nessun orario.")
    else:
        docenti_assenti = st.multiselect("Seleziona docenti assenti", sorted(orario_df["Docente"].unique()))
        data_sostituzione = st.date_input("Data della sostituzione")

        # Giorno calcolato automaticamente dalla data (in italiano)
        giorno_assente = data_sostituzione.strftime("%A")
        traduzione_giorni = {
            "Monday": "Lunedì", "Tuesday": "Martedì", "Wednesday": "Mercoledì",
            "Thursday": "Giovedì", "Friday": "Venerdì", "Saturday": "Sabato", "Sunday": "Domenica"
        }
        giorno_assente = traduzione_giorni.get(giorno_assente, giorno_assente)

        if giorno_assente not in GIORNI_SETTIMANA:
            st.warning(f"Hai selezionato {giorno_assente}, un giorno non presente nell'orario scolastico (Lun-Ven).")

        # =========================
        # CLASSI IN USCITA DIDATTICA (libera i curricolari di quelle classi)
        # =========================
        classi_uscita_per_ora = {}  # {ora: set(classi in uscita in quell'ora)}
        with st.expander("🚌 Classi in uscita didattica (oggi)"):
            st.caption(
                "Se una o più classi sono in uscita, i docenti curricolari [C] che in "
                "quell'ora avrebbero lezione con quella classe risultano liberi e "
                "selezionabili come sostituti, senza generare un conflitto alla conferma."
            )
            classi_disponibili = sorted(orario_df["Classe"].unique()) if not orario_df.empty else []
            classi_uscita_selezionate = st.multiselect(
                "Classi in uscita",
                classi_disponibili,
                key="classi_uscita_multiselect"
            )
            for classe_u in classi_uscita_selezionate:
                ore_classe_u = [o for o in ORE_LEZIONE if not orario_df[
                    (orario_df["Classe"] == classe_u) &
                    (orario_df["Giorno"] == giorno_assente) &
                    (orario_df["Ora"] == o)
                ].empty]
                if not ore_classe_u:
                    st.caption(f"⚠️ {classe_u} non ha lezioni previste {giorno_assente}.")
                    continue
                ore_scelte_u = st.multiselect(
                    f"Ore in uscita per {classe_u} (default: tutta la giornata)",
                    ore_classe_u,
                    default=ore_classe_u,
                    key=f"ore_uscita_{classe_u}"
                )
                for ora_u in ore_scelte_u:
                    classi_uscita_per_ora.setdefault(ora_u, set()).add(classe_u)

        if not docenti_assenti:
            st.info("Seleziona almeno un docente per continuare.")
        else:
            # Ore scoperte dell'assente (mostriamo anche se l'assente ha Escludi True)
            ore_assenti = orario_df[
                (orario_df["Docente"].isin(docenti_assenti)) &
                (orario_df["Giorno"] == giorno_assente)
            ].copy()

            if ore_assenti.empty:
                st.info("I docenti selezionati non hanno lezioni in quel giorno.")
            else:
                st.subheader("📌 Ore scoperte")

                # --- Aggiunta: calcolo sostegni in servizio per ogni ora scoperta ---
                ore_assenti_display = ore_assenti.copy()
                sostegni_presenti = []

                for _, r in ore_assenti.iterrows():
                    ora = r["Ora"]
                    classe = r["Classe"]

                    # Cerca docenti di sostegno in quella classe-ora
                    sost_df = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == ora) &
                        (orario_df["Classe"] == classe) &
                        (orario_df["Tipo"].str.lower() == "sostegno")
                    ]

                    if sost_df.empty:
                        sostegni_presenti.append("—")
                    else:
                        # Elenco nomi separati da virgola
                        lista = ", ".join(sorted(sost_df["Docente"].unique()))
                        sostegni_presenti.append(lista)

                # Colonna aggiuntiva
                ore_assenti_display["Sostegni in servizio"] = sostegni_presenti

                # Ordina in modo naturale
                ore_assenti_display = ore_assenti_display[["Docente", "Ora", "Classe", "Tipo", "Sostegni in servizio"]]

                st.dataframe(ore_assenti_display, use_container_width=True, hide_index=True)


                st.subheader("🔄 Possibili sostituti")
                sostituzioni = []

                # Precalcolo: tutti i docenti (escludiamo definitivamente chi ha Escludi=True)
                tutti_docenti = sorted(orario_df["Docente"].unique())
                escludi_docenti = set(orario_df[orario_df["Escludi"]]["Docente"].unique())
                # Tutti i docenti assenti oggi (non solo quello della singola ora): nessuno di
                # loro può comparire come possibile sostituto, in nessuna ora.
                docenti_assenti_set = set(docenti_assenti)

                # Mappa docente -> tipo, calcolata UNA volta sola (prima costava un filtro
                # su tutto orario_df per ogni singolo docente, ad ogni ora scoperta)
                docente_tipo_map = build_docente_tipo_map(orario_df)
                def tipo_docente(d):
                    return docente_tipo_map.get(d, "")

                # Per ogni ora scoperta costruisco la lista di opzioni con l'ordine richiesto
                for _, row in ore_assenti.iterrows():
                    ora = row["Ora"]
                    classe = row["Classe"]
                    assente = row["Docente"]

                    # Docenti presenti in quell'ora (escludiamo chi ha Escludi=True e
                    # chiunque sia stato segnato assente oggi, non solo l'assente di questa riga)
                    presenti_ora_df = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == ora) &
                        (~orario_df["Escludi"]) &
                        (~orario_df["Docente"].isin(docenti_assenti_set))
                    ].copy()

                    presenti_ora = list(dict.fromkeys(presenti_ora_df["Docente"].tolist()))  # mantiene ordine stabile

                    added = set()
                    options = ["Nessuno"]

                    # 1) Sostegni della stessa classe in quell'ora -> [S]
                    same_class_sost = presenti_ora_df[
                        (presenti_ora_df["Tipo"].str.lower() == "sostegno") &
                        (presenti_ora_df["Classe"] == classe) &
                        (presenti_ora_df["Docente"] != assente)
                    ]["Docente"].unique().tolist()
                    for d in sorted(same_class_sost):
                        label = f"[S] {d}"
                        if d not in added:
                            options.append(label); added.add(d)

                    # 2) Altri sostegni presenti in quell'ora -> [S]
                    other_sost = presenti_ora_df[
                        (presenti_ora_df["Tipo"].str.lower() == "sostegno") &
                        (presenti_ora_df["Docente"] != assente)
                    ]["Docente"].unique().tolist()
                    for d in sorted(other_sost):
                        if d in added: 
                            continue
                        label = f"[S] {d}"
                        options.append(label); added.add(d)

                    # 3) Curricolari presenti in quell'ora -> [C], a meno che la loro
                    # classe non sia in uscita didattica in quell'ora (in tal caso sono
                    # liberi -> [C] [USCITA])
                    curricolari_presenti_df = presenti_ora_df[
                        (presenti_ora_df["Tipo"].str.lower() != "sostegno") &
                        (presenti_ora_df["Docente"] != assente)
                    ]
                    uscita_classi_ora = classi_uscita_per_ora.get(ora, set())

                    curricolari_liberi_uscita = curricolari_presenti_df[
                        curricolari_presenti_df["Classe"].isin(uscita_classi_ora)
                    ]["Docente"].unique().tolist()

                    curricolari_occupati = curricolari_presenti_df[
                        ~curricolari_presenti_df["Classe"].isin(uscita_classi_ora)
                    ]["Docente"].unique().tolist()

                    # 3a) Curricolari liberi perché la classe è in uscita -> [C] [USCITA]
                    for d in sorted(curricolari_liberi_uscita):
                        if d in added:
                            continue
                        label = f"[C] [USCITA] {d}"
                        options.append(label); added.add(d)

                    # 3b) Curricolari realmente occupati in quell'ora -> [C]
                    for d in sorted(curricolari_occupati):
                        if d in added:
                            continue
                        label = f"[C] {d}"
                        options.append(label); added.add(d)

                    # 4) Docenti che NON compaiono in quell'ora -> [S] [NP] o [C] [NP]
                    # candidati NP: tutti i docenti presenti nell'orario generale MA non in 'presenti_ora',
                    # non assenti oggi (nessuno di loro), non escludi
                    presenti_ora_set = set(presenti_ora)
                    np_candidates = [
                        d for d in tutti_docenti
                        if d not in presenti_ora_set
                        and d not in docenti_assenti_set
                        and d not in escludi_docenti
                    ]

                    # separo NP in S e C (in base al "Tipo" trovato nell'orario)
                    np_sost = []
                    np_curr = []
                    for d in np_candidates:
                        t = tipo_docente(d).lower()
                        if t == "sostegno":
                            np_sost.append(d)
                        else:
                            np_curr.append(d)

                    # Aggiungo prima NP sostegni, poi NP curricolari
                    for d in sorted(np_sost):
                        if d in added:
                            continue
                        label = f"[S] [NP] {d}"
                        options.append(label); added.add(d)

                    for d in sorted(np_curr):
                        if d in added:
                            continue
                        label = f"[C] [NP] {d}"
                        options.append(label); added.add(d)

                    # Rimuovo eventuali docenti esclusi (già filtrati) e assente (già escluso)
                    # --- suggerimento automatico (come prima gerarchia, con i liberi per
                    # uscita subito dopo i sostegni e prima dei curricolari davvero occupati) ---
                    proposto_display = "Nessuno"
                    if len(same_class_sost) > 0:
                        proposto_display = f"[S] {sorted(same_class_sost)[0]}"
                    elif len(other_sost) > 0:
                        proposto_display = f"[S] {sorted(other_sost)[0]}"
                    elif len(curricolari_liberi_uscita) > 0:
                        proposto_display = f"[C] [USCITA] {sorted(curricolari_liberi_uscita)[0]}"
                    elif len(curricolari_occupati) > 0:
                        proposto_display = f"[C] {sorted(curricolari_occupati)[0]}"
                    elif len(np_sost) > 0:
                        proposto_display = f"[S] [NP] {sorted(np_sost)[0]}"
                    elif len(np_curr) > 0:
                        proposto_display = f"[C] [NP] {sorted(np_curr)[0]}"

                    default_index = options.index(proposto_display) if proposto_display in options else 0

                    scelta = st.selectbox(
                        f"Sostituto per {ora} ora - Classe {classe} (assente {assente})",
                        options,
                        index=default_index,
                        key=f"sost_{assente}_{ora}_{classe}"
                    )

                    # pulisco il nome per lo storico (rimuovo prefissi tipo "[S] [NP] " ecc.)
                    if scelta == "Nessuno":
                        nome_pulito = "Nessuno"
                    else:
                        nome_pulito = (
                            scelta.replace("[S] [NP] ", "")
                                  .replace("[C] [NP] ", "")
                                  .replace("[C] [USCITA] ", "")
                                  .replace("[S] ", "")
                                  .replace("[C] ", "")
                                  .strip()
                        )

                    sostituzioni.append({
                        "Ora": ora,
                        "Classe": classe,
                        "Assente": assente,
                        "Sostituto_display": scelta,
                        "Sostituto": nome_pulito
                    })

                sostituzioni_df = pd.DataFrame(sostituzioni)

                # Ordina per ora
                ordine_ore = ORE_LEZIONE
                if not sostituzioni_df.empty:
                    sostituzioni_df["Ora"] = pd.Categorical(sostituzioni_df["Ora"], categories=ordine_ore, ordered=True)
                    sostituzioni_df = sostituzioni_df.sort_values("Ora").reset_index(drop=True)

                # --- VISTA TABELLA (mostra la label con [S]/[C]/[NP]) ---
                st.subheader("📋 Sostituzioni in tabella")
                tabella_df = sostituzioni_df[["Ora", "Classe", "Assente", "Sostituto_display"]].copy()
                tabella_df = tabella_df.rename(columns={"Sostituto_display": "Sostituzione"})
                tabella_df["Ora"] = pd.Categorical(tabella_df["Ora"], categories=ordine_ore, ordered=True)
                tabella_df = tabella_df.sort_values(["Ora", "Classe"]).reset_index(drop=True)

                styled_tabella = (
                    tabella_df.style
                        .set_table_styles([
                            {"selector": "th.col0", "props": [("width", "50px")]},
                            {"selector": "th.col1", "props": [("width", "80px")]},
                            {"selector": "th", "props": [("background-color", "#f0f0f0"),
                                                         ("font-weight", "bold"),
                                                         ("text-align", "center"),
                                                         ("font-size", "16px")]},
                            {"selector": "td", "props": [("text-align", "center"),
                                                         ("padding", "6px 12px"),
                                                         ("font-size", "16px")]}
                        ])
                        .apply(lambda x: ['background-color: #f9f9f9' if i % 2 else 'background-color: white'
                                          for i in range(len(x))], axis=0)
                )
                html = styled_tabella.hide(axis="index").to_html()
                st.markdown(html, unsafe_allow_html=True)

                # --- VISTA TESTUALE ---
                st.subheader("📝 Sostituzioni in formato testo (mobile/copincolla)")
                testo_output = "Buongiorno, supplenze:\n\n"

                # uso sostituzioni_df che ha sia display che nome pulito
                for ora, gruppo in sostituzioni_df.groupby("Ora"):
                    if not gruppo.empty:
                        testo_output += f"🕐 *{ora} ORA*\n"
                        for _, r in gruppo.iterrows():
                            sost_pulito = r['Sostituto'] if r['Sostituto'] not in ["Nessuno", "", "—"] else "—"
                            testo_output += f"Classe {r['Classe']}\n"
                            testo_output += f"👩‍🏫 Assente: {r['Assente']}\n"
                            testo_output += f"✅ Sostituzione: {sost_pulito}\n\n"

                testo_strip = testo_output.strip()
                st.text_area("Testo pronto da copiare", value=testo_strip, height=300)
                # Bottone "Copia negli appunti" via JavaScript
                testo_js = testo_strip.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$")
                st.components.v1.html(f"""
<button onclick="navigator.clipboard.writeText(`{testo_js}`).then(()=>{{
    this.innerText='✅ Copiato!'; setTimeout(()=>this.innerText='📋 Copia negli appunti', 2000);
}})" style="
    width:100%; padding:0.7em; font-size:1em; font-weight:bold;
    background:#0068c9; color:white; border:none; border-radius:10px; cursor:pointer;
">📋 Copia negli appunti</button>
""", height=55)

                # --- Step 1: conferma (controllo conflitti) ---
                if st.button("✅ Conferma tabella (non salva ancora)"):
                    check_df = sostituzioni_df.copy()
                    conflitti = []
                    conflitti_orario = []

                    for ora_val in check_df["Ora"].unique():
                        assegnazioni = check_df[check_df["Ora"] == ora_val]
                        sostituti = [s for s in assegnazioni["Sostituto"] if s not in ["Nessuno", "", "—"]]

                        # 1) Conflitto: stesso docente scelto su più classi nella stessa ora
                        duplicati = [s for s in sostituti if sostituti.count(s) > 1]
                        if duplicati:
                            conflitti.append((ora_val, list(set(duplicati))))

                        # 2) Conflitto: docente scelto ma già in orario in quell’ora (SOLO
                        # curricolari, e solo se la sua classe in quell’ora NON è in uscita)
                        for s in sostituti:
                            # recupero il tipo del docente dalla mappa precalcolata
                            tipo = docente_tipo_map.get(s, "").lower()
                            if tipo != "sostegno":
                                lezione_in_quell_ora = orario_df[
                                    (orario_df["Docente"] == s) &
                                    (orario_df["Giorno"] == giorno_assente) &
                                    (orario_df["Ora"] == ora_val)
                                ]
                                if not lezione_in_quell_ora.empty:
                                    classe_lezione = lezione_in_quell_ora.iloc[0]["Classe"]
                                    classi_in_uscita_ora = classi_uscita_per_ora.get(ora_val, set())
                                    if classe_lezione not in classi_in_uscita_ora:
                                        conflitti_orario.append((ora_val, s))

                    if conflitti or conflitti_orario:
                        if conflitti:
                            st.error("⚠️ Errore: lo stesso docente è stato assegnato a più classi nella stessa ora:")
                            for ora_c, docs in conflitti:
                                st.write(f"- Ora {ora_c}: {', '.join(docs)}")
                        if conflitti_orario:
                            st.error("⚠️ Errore: alcuni docenti curricolari scelti come supplenti hanno già lezione in quell’ora:")
                            for ora_c, docente in conflitti_orario:
                                st.write(f"- Ora {ora_c}: {docente} è già impegnato in orario")
                        st.stop()

                    # se tutto ok salvo in session_state
                    st.session_state["sostituzioni_confermate"] = sostituzioni_df.copy()
                    st.session_state["ore_assenti_confermate"] = ore_assenti.copy()
                    st.session_state["data_sostituzione_tmp"] = data_sostituzione
                    st.session_state["giorno_assente_tmp"] = giorno_assente

                    st.success("Tabella confermata ✅ Ora puoi salvarla nello storico.")


                # --- Step 2: Salva nello storico ---
                if st.session_state.get("sostituzioni_confermate") is not None:
                    if st.button("💾 Salva nello storico", key="save_storico_main"):
                        sost_df = st.session_state.get("sostituzioni_confermate")
                        ore_assenti_session = st.session_state.get("ore_assenti_confermate")
                        data_tmp = st.session_state.get("data_sostituzione_tmp")
                        giorno_tmp = st.session_state.get("giorno_assente_tmp")

                        if sost_df is not None and ore_assenti_session is not None:
                            # salva_storico_assenze si aspetta la colonna "Sostituto" con il nome pulito
                            if salva_storico_assenze(data_tmp, giorno_tmp, sost_df, ore_assenti_session):
                                st.success("Assenze e sostituzioni salvate nello storico ✅")
                                for k in ["sostituzioni_confermate", "ore_assenti_confermate",
                                          "data_sostituzione_tmp", "giorno_assente_tmp"]:
                                    st.session_state.pop(k, None)
                                try:
                                    st.rerun()
                                except Exception:
                                    pass





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
    st.header("📊 Statistiche Sostituzioni")
    with st.spinner('Caricamento statistiche...'):
        df_storico, df_assenze = carica_statistiche()

    def render_cards(titolo, righe, colore):
        """Renderizza un blocco card colorato con titolo e lista di (nome, valore)."""
        medaglie = ["🥇", "🥈", "🥉"]
        items_html = ""
        for i, (nome, val) in enumerate(righe):
            medaglia = medaglie[i] if i < 3 else ""
            items_html += f"""
  <div style="display:flex;justify-content:space-between;align-items:center;
              padding:8px 0;border-bottom:1px solid rgba(255,255,255,0.25);">
    <span style="font-size:1.05em;">{medaglia} {nome.title()}</span>
    <span style="font-size:1.1em;font-weight:bold;white-space:nowrap;margin-left:12px;">{val} ore</span>
  </div>"""
        st.markdown(f"""
<div style="background:{colore};border-radius:14px;padding:16px 20px;color:white;margin-bottom:16px;">
  <div style="font-size:0.85em;font-weight:600;opacity:0.85;margin-bottom:6px;">{titolo}</div>
  {items_html}
</div>""", unsafe_allow_html=True)

    if df_storico.empty:
        st.info("Nessuna statistica disponibile. Registra prima delle sostituzioni.")
    else:
        df_sum = df_storico.groupby("docente")["ore"].sum().reset_index()
        df_sum = df_sum.rename(columns={"ore": "Totale Ore Sostituite"})
        df_sorted = df_sum.sort_values("Totale Ore Sostituite", ascending=False).reset_index(drop=True)

        top3 = [(r["docente"], int(r["Totale Ore Sostituite"])) for _, r in df_sorted.head(3).iterrows()]

        # Per i "meno sostituzioni" filtro solo i docenti di sostegno [S],
        # inclusi quelli a zero ore (non presenti nello storico)
        docenti_sostegno = set(
            orario_df[
                (orario_df["Tipo"].str.lower() == "sostegno") &
                (~orario_df["Escludi"])
            ]["Docente"].str.lower().unique()
        )
        df_tutti_sost = pd.DataFrame({"docente": sorted(docenti_sostegno)})
        df_tutti_sost = df_tutti_sost.merge(
            df_sum.assign(docente=df_sum["docente"].str.lower()),
            on="docente", how="left"
        ).fillna(0)
        df_tutti_sost["Totale Ore Sostituite"] = df_tutti_sost["Totale Ore Sostituite"].astype(int)
        df_tutti_sost = df_tutti_sost.sort_values("Totale Ore Sostituite").reset_index(drop=True)
        bot3 = [(r["docente"], int(r["Totale Ore Sostituite"])) for _, r in df_tutti_sost.head(3).iterrows()]

        col_top, col_bot = st.columns(2)
        with col_top:
            render_cards("🟢 Più sostituzioni", top3, "#2e7d32")
        with col_bot:
            render_cards("🟠 Meno sostituzioni [S]", bot3, "#e65100")

        st.dataframe(df_sorted, use_container_width=True, hide_index=True)
        st.bar_chart(df_sorted.set_index("docente"))

    st.subheader("⚠️ Azzeramento storico sostituzioni")
    conferma = st.checkbox("Confermo di voler cancellare definitivamente lo storico delle sostituzioni", key="conf_storico")
    if st.button("Elimina storico sostituzioni"):
        if conferma:
            if clear_sheet_content(STORICO_SHEET):
                st.success("Storico delle sostituzioni eliminato ✅")
        else:
            st.warning("Devi spuntare la conferma prima di cancellare lo storico delle sostituzioni.")

    # STATISTICHE ASSENZE
    st.header("📊 Statistiche Assenze")
    if df_assenze.empty:
        st.info("Nessuna assenza registrata.")
    else:
        # Ore totali per docente
        df_ore = df_assenze.groupby("docente")["ora"].count().reset_index().rename(columns={"ora": "Totale Ore Assenti"})
        # Giorni effettivi: date distinte per docente
        df_giorni = df_assenze.groupby("docente")["data"].nunique().reset_index().rename(columns={"data": "Giorni Assenti"})
        df_assenze_agg = df_ore.merge(df_giorni, on="docente")
        df_assenze_agg = df_assenze_agg.sort_values("Totale Ore Assenti", ascending=False).reset_index(drop=True)

        # Card assenze con ore + giorni
        medaglie = ["🥇", "🥈", "🥉"]
        items_html = ""
        for i, (_, r) in enumerate(df_assenze_agg.head(3).iterrows()):
            medaglia = medaglie[i] if i < 3 else ""
            items_html += f"""
  <div style="display:flex;justify-content:space-between;align-items:center;
              padding:8px 0;border-bottom:1px solid rgba(255,255,255,0.25);">
    <span style="font-size:1.05em;">{medaglia} {r['docente'].title()}</span>
    <span style="font-size:0.95em;font-weight:bold;white-space:nowrap;margin-left:12px;">
      {int(r['Totale Ore Assenti'])} ore &nbsp;·&nbsp; {int(r['Giorni Assenti'])} giorni
    </span>
  </div>"""
        st.markdown(f"""
<div style="background:#e65100;border-radius:14px;padding:16px 20px;color:white;margin-bottom:16px;">
  <div style="font-size:0.85em;font-weight:600;opacity:0.85;margin-bottom:6px;">🟠 Più assenze</div>
  {items_html}
</div>""", unsafe_allow_html=True)

        st.dataframe(df_assenze_agg, use_container_width=True, hide_index=True)
        st.bar_chart(df_assenze_agg.set_index("docente")[["Totale Ore Assenti"]])

    st.subheader("⚠️ Azzeramento storico assenze")
    conferma_assenze = st.checkbox("Confermo di voler cancellare definitivamente lo storico delle assenze", key="conf_assenze")
    if st.button("Elimina storico assenze"):
        if conferma_assenze:
            if clear_sheet_content(ASSENZE_SHEET):
                st.success("Storico delle assenze eliminato ✅")
        else:
            st.warning("Devi spuntare la conferma prima di cancellare lo storico delle assenze.")

    st.subheader("Cloud Backup")
    st.info("Scarica un backup compresso dei dati dei fogli Orario, Storico e Assenze.")
    backup_file = create_backup()
    if backup_file:
        st.download_button(
            label="⬇️ Scarica Backup (ZIP)",
            data=backup_file,
            file_name=f"{SPREADSHEET_NAME}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
            help="Crea un backup in formato .zip e lo scarica direttamente."
        )
