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
ALUNNI_SHEET = "alunni"
MATERIE_SHEET = "materie"
CRITICITA_SHEET = "criticita"

REQUIRED_COLUMNS = ["Nome", "Classe", "Materia", "Criticità", "Data", "Docente", "Note"]

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

# ===== Helper per parsing date robusto =====
def parse_date_series(series: pd.Series) -> pd.Series:
    """
    Pulisce la serie rimuovendo apostrofi/virgolette iniziali e prova a parsare:
    1) ISO %Y-%m-%d
    2) %d/%m/%Y
    3) fallback con dayfirst=True
    Restituisce Series di datetime64[ns] (NaT se non parsabile).
    """
    s = series.astype(str).str.strip()
    # rimuovi apostrofi o virgolette iniziali (Google Sheets può mostrarli)
    s = s.str.replace(r"^['\"]+", "", regex=True)

    # prova ISO YYYY-MM-DD
    parsed = pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")

    # fallback dd/mm/yyyy
    mask = parsed.isna()
    if mask.any():
        parsed2 = pd.to_datetime(s[mask], format="%d/%m/%Y", errors="coerce")
        parsed.loc[mask] = parsed2

    # ultimo fallback: parsing generico cercando dayfirst
    mask = parsed.isna()
    if mask.any():
        parsed3 = pd.to_datetime(s[mask], errors="coerce", dayfirst=True)
        parsed.loc[mask] = parsed3

    return parsed

def carica_segnalazioni():
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(SEGNALAZIONI_SHEET)
        df = gd.get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
        if not df.empty:
            # garantisco colonne e ordine
            df = df.reindex(columns=REQUIRED_COLUMNS)
            # formatta/parsifica la colonna Data in modo robusto (datetime dtype)
            if "Data" in df:
                df["Data"] = parse_date_series(df["Data"])
        else:
            df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        return df
    except Exception as e:
        st.error(f"Errore nel caricamento segnalazioni: {e}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

def salva_segnalazione(nome, classe, materia, criticita, data, docente, note):
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(SEGNALAZIONI_SHEET)
        new_row = [nome or "", classe or "", materia or "", criticita or "", data or "", docente or "", note or ""]
        # USER_ENTERED permette a Google Sheets di interpretare la stringa della data come vero tipo data
        ws.append_row(new_row, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio: {e}")
        return False

# =========================
# FOGLI DI SUPPORTO
# =========================
def carica_alunni():
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(ALUNNI_SHEET)
        df = gd.get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
        if not df.empty and "Nome" in df and "Classe" in df:
            return df
    except:
        return pd.DataFrame(columns=["Nome", "Classe"])
    return pd.DataFrame(columns=["Nome", "Classe"])

def carica_materie():
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(MATERIE_SHEET)
        df = gd.get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
        if not df.empty and "Materia" in df:
            return df["Materia"].dropna().unique().tolist()
    except:
        return []
    return []

def carica_criticita():
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(CRITICITA_SHEET)
        df = gd.get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
        if not df.empty and "Criticità" in df:
            return df["Criticità"].dropna().unique().tolist()
    except:
        return []
    return []

# =========================
# STREAMLIT APP
# =========================
st.title("📘 Registro Elettronico Alunni")

# inizializza foglio principale
ensure_sheet_exist()

menu = st.sidebar.selectbox(
    "📌 Naviga",
    ["➕ Inserisci Segnalazione", "📜 Storico", "📊 Statistiche"]
)

# --- INSERISCI SEGNALAZIONE ---
if menu == "➕ Inserisci Segnalazione":
    st.header("Aggiungi una nuova segnalazione")

    alunni_df = carica_alunni()
    materie_list = carica_materie()
    criticita_list = carica_criticita()

    with st.form("segnalazione_form"):
        nome = st.selectbox("👤 Nome alunno", [""] + alunni_df["Nome"].dropna().tolist())
        classe = ""
        if nome:
            classe_row = alunni_df[alunni_df["Nome"] == nome]
            if not classe_row.empty:
                classe = classe_row.iloc[0]["Classe"]

        materia = st.selectbox("📚 Materia", [""] + materie_list)
        criticita = st.selectbox("⚠️ Criticità", [""] + criticita_list)

        # uso formato ISO per storage, Google Sheets lo interpreterà come data grazie a USER_ENTERED
        date_obj = st.date_input("📅 Data", datetime.today())
        data = date_obj.strftime("%Y-%m-%d")

        docente = st.text_input("👩‍🏫 Docente", value="A.Q.")
        note = st.text_area("📝 Note aggiuntive")

        submitted = st.form_submit_button("💾 Salva segnalazione")

        if submitted:
            if nome and materia and criticita:
                ok = salva_segnalazione(nome, classe, materia, criticiti if (criticiti := criticita) else criticita, data, docente, note)
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
        col1, col2, col3 = st.columns(3)
        with col1:
            filtro_nome = st.selectbox("Filtra per alunno", [""] + sorted(df["Nome"].dropna().unique().tolist()))
        with col2:
            filtro_classe = st.selectbox("Filtra per classe", [""] + sorted(df["Classe"].dropna().unique().tolist()))
        with col3:
            filtro_materia = st.selectbox("Filtra per materia", [""] + sorted(df["Materia"].dropna().unique().tolist()))

        if filtro_nome:
            df = df[df["Nome"] == filtro_nome]
        if filtro_classe:
            df = df[df["Classe"] == filtro_classe]
        if filtro_materia:
            df = df[df["Materia"] == filtro_materia]

        # ordinamento per data (più recente in alto)
        if "Data" in df:
            # Data è già datetime dopo carica_segnalazioni; in caso contrario forziamo
            df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
            df = df.sort_values("Data", ascending=False)

        # creiamo una copia per la visualizzazione (formattata dd/mm/YYYY)
        df_display = df.copy()
        if "Data" in df_display:
            df_display["Data"] = df_display["Data"].dt.strftime("%d/%m/%Y").fillna("")

        st.dataframe(df_display, use_container_width=True, hide_index=True)

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
        criticita_list = carica_criticita()
        if not criticita_list:
            st.info("Nessun elenco di criticità disponibile.")
        else:
            # Conta per criticità (solo quelle definite nel foglio)
            st.subheader("Distribuzione delle criticità")
            conteggio_criticita = df[df["Criticità"].isin(criticita_list)] \
                                  .groupby("Criticità")["Nome"] \
                                  .count().reset_index(name="Totale")

            st.bar_chart(conteggio_criticita.set_index("Criticità"))

        # Trend temporale delle segnalazioni
        st.subheader("Andamento segnalazioni nel tempo")
        if "Data" in df:
            df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
            trend = df.dropna(subset=["Data"]).groupby(df["Data"].dt.to_period("M"))["Nome"].count()
            trend.index = trend.index.to_timestamp()
            st.line_chart(trend)
        else:
            st.info("Nessuna informazione sulla data disponibile per il trend.")
