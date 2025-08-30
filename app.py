import streamlit as st
import pandas as pd
import sqlite3
import os
import re
import bcrypt
import gspread
import gspread_dataframe as gd
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build

# =========================
# CONFIGURAZIONE GENERALE
# =========================
REQUIRED_COLUMNS = ["Docente", "Giorno", "Ora", "Classe", "Tipo", "Escludi"]
DB_FILE = "sostituzioni.db"

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
        ORARIO_SHEET_NAME = "OrarioSostituzioni"
        sheet = client.open(ORARIO_SHEET_NAME).worksheet("orario")
        df = gd.get_as_dataframe(sheet)
        df = df.dropna(how='all')
        return df
    except Exception as e:
        st.error(f"Errore nel caricamento dell'orario da Google Sheets: {e}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

def salva_orario(df):
    try:
        client = get_gdrive_client()
        ORARIO_SHEET_NAME = "OrarioSostituzioni"
        sheet = client.open(ORARIO_SHEET_NAME).worksheet("orario")
        gd.set_with_dataframe(sheet, df)
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio dell'orario su Google Sheets: {e}")
        return False

def salva_sostituzione(data, giorno, docente, ore_sostituzione, assenze):
    try:
        client = get_gdrive_client()
        ORARIO_SHEET_NAME = "OrarioSostituzioni"

        storico_sheet = client.open(ORARIO_SHEET_NAME).worksheet("storico")
        storico_sheet.append_rows([[data, giorno, docente, ore_sostituzione]])

        assenze_sheet = client.open(ORARIO_SHEET_NAME).worksheet("assenze")
        assenze_data = [[data, giorno, d[0], d[1], d[2]] for d in assenze]
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

def carica_statistiche():
    try:
        client = get_gdrive_client()
        ORARIO_SHEET_NAME = "OrarioSostituzioni"

        storico_sheet = client.open(ORARIO_SHEET_NAME).worksheet("storico")
        df_storico = gd.get_as_dataframe(storico_sheet)

        assenze_sheet = client.open(ORARIO_SHEET_NAME).worksheet("assenze")
        df_assenze = gd.get_as_dataframe(assenze_sheet)

        return df_storico, df_assenze
    except Exception as e:
        st.error(f"Errore nel caricamento delle statistiche da Google Sheets: {e}")
        return pd.DataFrame(), pd.DataFrame()

# =========================
# FUNZIONI DI SUPPORTO PER GESTIONE UTENTI (Google Sheets)
# =========================
def hash_password(password):
    hashed = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())
    return hashed.decode('utf-8')

def check_password(password, hashed_password):
    return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))

def get_users_df():
    try:
        client = get_gdrive_client()
        ORARIO_SHEET_NAME = "OrarioSostituzioni"
        sheet = client.open(ORARIO_SHEET_NAME).worksheet("utenti")
        df = gd.get_as_dataframe(sheet)
        df = df.dropna(how='all')
        if df.empty or 'username' not in df.columns or 'password_hash' not in df.columns:
            return pd.DataFrame(columns=['username', 'password_hash'])
        return df
    except gspread.exceptions.WorksheetNotFound:
        st.error("Il foglio 'utenti' non è stato trovato. Per favore crealo nel foglio di calcolo.")
        return pd.DataFrame(columns=['username', 'password_hash'])
    except Exception as e:
        st.error(f"Errore nel caricamento degli utenti da Google Sheets: {e}")
        return pd.DataFrame(columns=['username', 'password_hash'])

def add_user(username, password):
    try:
        client = get_gdrive_client()
        ORARIO_SHEET_NAME = "OrarioSostituzioni"
        sheet = client.open(ORARIO_SHEET_NAME).worksheet("utenti")
        hashed_password = hash_password(password)
        sheet.append_row([username, hashed_password])
        return True
    except Exception as e:
        st.error(f"Errore nell'aggiunta dell'utente: {e}")
        return False

# =========================
# LOGICA APPLICAZIONE
# =========================
st.title("📚 Gestione orari e Sostituzioni Docenti")

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# Blocco per la schermata di login
if not st.session_state["logged_in"]:
    st.sidebar.header("Login")
    username = st.sidebar.text_input("Username")
    password = st.sidebar.text_input("Password", type="password")

    # Inizializza l'utente admin se non esiste
    users_df = get_users_df()
    if users_df.empty or "admin" not in users_df["username"].values:
        default_password = st.secrets.get("admin_password", "password_di_default")
        add_user("admin", default_password)
        st.info("Utente di default 'admin' creato. Usa la password specificata nel file secrets (o 'password_di_default').")

    if st.sidebar.button("Accedi"):
        users_df = get_users_df()
        user_row = users_df[users_df["username"] == username]
        
        if not user_row.empty:
            hashed_password = user_row.iloc[0]["password_hash"]
            if check_password(password, hashed_password):
                st.session_state["logged_in"] = True
                st.session_state["username"] = username
                st.experimental_rerun()
            else:
                st.error("Username o password errati.")
        else:
            st.error("Username o password errati.")

else:
    st.sidebar.success(f"Benvenuto, {st.session_state['username']}!")
    if st.sidebar.button("Logout"):
        st.session_state["logged_in"] = False
        st.experimental_rerun()

    orario_df = carica_orario()

    # --- CARICAMENTO ORARIO ---
    st.header("🔄 Carica e Salva Orario")
    uploaded_file = st.file_uploader("Carica un nuovo file .csv (l'orario attuale verrà sovrascritto)", type="csv")
    if uploaded_file:
        df_new = pd.read_csv(uploaded_file)
        if salva_orario(df_new):
            st.success("Orario caricato e salvato correttamente ✅")

    if not orario_df.empty:
        download_orario(orario_df)
    
    # --- VISUALIZZA ORARIO ---
    st.header("📅 Orario completo")
    if not orario_df.empty:
        vista_pivot_docenti(orario_df, mode="classi")
    
    # --- GESTIONE ASSENZE E SOSTITUZIONI ---
    st.header("📝 Gestione Sostituzioni")
    docenti = sorted(orario_df["Docente"].unique().tolist())
    giorni = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]
    ore = ["I", "II", "III", "IV", "V", "VI"]

    st.subheader("1. Docenti Assenti")
    docenti_assenti = st.multiselect("Seleziona uno o più docenti assenti", docenti)
    giorno_assenza = st.selectbox("Seleziona il giorno dell'assenza", giorni)

    if docenti_assenti:
        st.subheader("2. Ore scoperte")
        assenze = orario_df[(orario_df["Docente"].isin(docenti_assenti)) & (orario_df["Giorno"] == giorno_assenza)]
        if assenze.empty:
            st.info("Nessuna ora di lezione scoperta per i docenti assenti in questo giorno.")
        else:
            st.dataframe(assenze[["Docente", "Ora", "Classe"]].drop_duplicates(), use_container_width=True)

            sostituzioni = []
            st.subheader("3. Proponi Sostituzioni")
            for _, r in assenze.iterrows():
                classe = r["Classe"]
                ora = r["Ora"]
                docente_assente = r["Docente"]
                tipo_lezione = r["Tipo"]
                
                altri_docenti_liberi = orario_df[
                    (orario_df["Giorno"] == giorno_assenza) &
                    (orario_df["Ora"] == ora) &
                    (orario_df["Docente"] != docente_assente)
                ].index
                
                docenti_curriculari_disponibili = [d for d in docenti if d not in orario_df.loc[altri_docenti_liberi, "Docente"].tolist()]
                
                proposta_sostituto = "Da Assegnare"
                if "Sostegno" in str(tipo_lezione):
                    sostegno_in_classe = orario_df[
                        (orario_df["Classe"] == classe) &
                        (orario_df["Tipo"].str.contains("Sostegno", case=False, na=False)) &
                        (orario_df["Docente"].isin(docenti_curriculari_disponibili))
                    ]
                    if not sostegno_in_classe.empty:
                        proposta_sostituto = sostegno_in_classe.iloc[0]["Docente"]
                
                sostituto_scelto = st.selectbox(
                    f"Sostituzione per **{docente_assente}** in **{classe}** alla **{ora}** ora:",
                    options=["Da Assegnare"] + docenti_curriculari_disponibili,
                    index=0 if "Da Assegnare" not in docenti_curriculari_disponibili else docenti_curriculari_disponibili.index(proposta_sostituto) + 1 if proposta_sostituto in docenti_curriculari_disponibili else 0
                )
                
                sostituzioni.append({
                    "Docente Assente": docente_assente,
                    "Classe": classe,
                    "Ora": ora,
                    "Sostituto": sostituto_scelto
                })
            
            if st.button("Conferma Sostituzioni"):
                if salva_sostituzione(pd.Timestamp.now().strftime("%Y-%m-%d"), giorno_assenza, docenti_assenti[0], len(assenze), assenze[["Docente", "Ora", "Classe"]].values.tolist()):
                    st.success("Sostituzioni salvate con successo!")
                    
    # --- STATISTICHE ---
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
