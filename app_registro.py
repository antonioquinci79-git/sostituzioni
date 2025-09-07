import streamlit as st
import pandas as pd
import gspread
import gspread_dataframe as gd
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# =========================
# CONFIG
# =========================
SPREADSHEET_NAME = "RegistroAlunni"
SEGNALAZIONI_SHEET = "segnalazioni"

REQUIRED_COLUMNS = ["Nome", "Materia", "Criticità", "Data", "Docente", "Note"]

# =========================
# GOOGLE SHEETS CLIENT
# =========================
@st.cache_resource
def get_gdrive_client():
    gdrive_credentials = st.secrets["gdrive"]
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/spreadsheets'
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(gdrive_credentials, scope)
    client = gspread.authorize(creds)
    return client

def ensure_sheet_exist():
    client = get_gdrive_client()
    try:
        sh = client.open(SPREADSHEET_NAME)
    except:
        sh = client.create(SPREADSHEET_NAME)
    existing = {ws.title for ws in sh.worksheets()}
    if SEGNALAZIONI_SHEET not in existing:
        ws = sh.add_worksheet(title=SEGNALAZIONI_SHEET, rows="1000", cols="20")
        header_df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        gd.set_with_dataframe(ws, header_df, include_index=False, include_column_header=True)

def carica_segnalazioni():
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(SEGNALAZIONI_SHEET)
        df = gd.get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
        if not df.empty:
            df = df[REQUIRED_COLUMNS]
        else:
            df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        return df
    except Exception as e:
        st.error(f"Errore nel caricamento segnalazioni: {e}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

def salva_segnalazione(nome, materia, criticita, data, docente, note):
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(SEGNALAZIONI_SHEET)
        new_row = [nome, materia, criticita, data, docente, note]
        ws.append_row(new_row, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio: {e}")
        return False

# =========================
# STREAMLIT APP
# =========================
st.title("📘 Registro Elettronico Alunni")

# inizializza foglio
ensure_sheet_exist()

menu = st.sidebar.selectbox(
    "📌 Naviga",
    ["➕ Inserisci Segnalazione", "📜 Storico", "📊 Statistiche"]
)

# --- INSERISCI SEGNALAZIONE ---
if menu == "➕ Inserisci Segnalazione":
    st.header("Aggiungi una nuova segnalazione")

    with st.form("segnalazione_form"):
        nome = st.text_input("👤 Nome alunno")
        materia = st.text_input("📚 Materia")
        criticita = st.selectbox(
            "⚠️ Criticità",
            ["", "Non ha portato i compiti", "Non ha il materiale didattico", "Disturbo in classe", "Altro"]
        )
        data = st.date_input("📅 Data", datetime.today()).strftime("%Y-%m-%d")
        docente = st.text_input("👩‍🏫 Docente")
        note = st.text_area("📝 Note aggiuntive")

        submitted = st.form_submit_button("💾 Salva segnalazione")

        if submitted:
            if nome and materia and criticita:
                ok = salva_segnalazione(nome, materia, criticita, data, docente, note)
                if ok:
                    st.success("✅ Segnalazione salvata correttamente")
            else:
                st.error("Compila almeno Nome, Materia e Criticità.")

# --- STORICO ---
elif menu == "📜 Storico":
    st.header("Storico segnalazioni")

    df = carica_segnalazioni()
    if df.empty:
        st.info("Nessuna segnalazione presente.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            filtro_nome = st.selectbox("Filtra per alunno", [""] + sorted(df["Nome"].dropna().unique().tolist()))
        with col2:
            filtro_materia = st.selectbox("Filtra per materia", [""] + sorted(df["Materia"].dropna().unique().tolist()))

        if filtro_nome:
            df = df[df["Nome"] == filtro_nome]
        if filtro_materia:
            df = df[df["Materia"] == filtro_materia]

        st.dataframe(df, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Scarica CSV filtrato",
            data=df.to_csv(index=False),
            file_name="storico_segnalazioni.csv",
            mime="text/csv"
        )

# --- STATISTICHE ---
elif menu == "📊 Statistiche":
    st.header("Statistiche segnalazioni")

    df = carica_segnalazioni()
    if df.empty:
        st.info("Nessuna segnalazione presente.")
    else:
        # Conteggio per alunno
        st.subheader("Segnalazioni per alunno")
        conteggio_alunni = df.groupby("Nome")["Criticità"].count().reset_index(name="Totale")
        st.dataframe(conteggio_alunni, use_container_width=True, hide_index=True)
        st.bar_chart(conteggio_alunni.set_index("Nome"))

        # Conteggio per materia
        st.subheader("Segnalazioni per materia")
        conteggio_materie = df.groupby("Materia")["Criticità"].count().reset_index(name="Totale")
        st.dataframe(conteggio_materie, use_container_width=True, hide_index=True)
        st.bar_chart(conteggio_materie.set_index("Materia"))
