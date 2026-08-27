import os
import io
import datetime
from typing import List, Dict, Tuple, Optional
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# CONSTANTS & CONFIGURATION
# ==========================================
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "Database_Plesso")
CREDENTIALS_FILE = "credentials.json"

STORICO_SHEET = "Storico_Complessivo"
ASSENZE_SHEET = "Registro_Assenze"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

STATI_VALIDI = [
    "Presente",
    "Assente",
    "Ritardato",
    "Uscita Anticipata",
    "Giustificato",
    "Non Giustificato",
]

PLESSI_E_CLASSI = {
    "Plesso Centrale": ["1A", "1B", "2A", "2B", "3A", "3B", "4A", "4B", "5A", "5B"],
    "Plesso Nord": ["1C", "2C", "3C", "4C", "5C"],
    "Plesso Sud": ["1D", "2D", "3D", "4D", "5D"],
}

PAGE_CONFIG = {
    "page_title": "Gestione Plesso Scolastico",
    "page_icon": "🏫",
    "layout": "wide",
    "initial_sidebar_state": "expanded",
}

st.set_page_config(**PAGE_CONFIG)

# ==========================================
# GOOGLE SHEETS CONNECTION & CACHING
# ==========================================
@st.cache_resource
def get_gspread_client() -> gspread.Client:
    """Autentica e restituisce il client gspread."""
    if os.path.exists(CREDENTIALS_FILE):
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    else:
        # Supporto alle Secrets di Streamlit Cloud
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def get_spreadsheet() -> gspread.Spreadsheet:
    """Ottiene lo Spreadsheet principale."""
    client = get_gspread_client()
    return client.open(SPREADSHEET_NAME)

def get_worksheet(sheet_name: str) -> gspread.Worksheet:
    """Ottiene un foglio specifico, creandolo se non esiste."""
    sh = get_spreadsheet()
    try:
        return sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        if sheet_name == STORICO_SHEET:
            ws = sh.add_worksheet(title=sheet_name, rows=100, cols=10)
            ws.append_row(["ID", "Data", "Plesso", "Classe", "Studente", "Stato", "Note"])
            return ws
        elif sheet_name == ASSENZE_SHEET:
            ws = sh.add_worksheet(title=sheet_name, rows=100, cols=10)
            ws.append_row(["ID", "Data", "Plesso", "Classe", "Studente", "Motivo", "Giustificato"])
            return ws
        else:
            raise

def clear_sheet_content(sheet_name: str):
    """Pulisce il contenuto di un foglio mantenendo la riga di intestazione."""
    ws = get_worksheet(sheet_name)
    headers = ws.row_values(1)
    ws.clear()
    if headers:
        ws.append_row(headers)

# ==========================================
# BUSINESS LOGIC & DATA OPERATIONS
# ==========================================
@st.cache_data(ttl=60)
def carica_statistiche() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Carica i dati dai fogli Google per l'analisi."""
    ws_storico = get_worksheet(STORICO_SHEET)
    ws_assenze = get_worksheet(ASSENZE_SHEET)
    
    df_storico = pd.DataFrame(ws_storico.get_all_records())
    df_assenze = pd.DataFrame(ws_assenze.get_all_records())
    
    return df_storico, df_assenze

def registra_presenze(data: datetime.date, plesso: str, classe: str, registrazioni: List[Dict]):
    """Registra o aggiorna le presenze nel foglio Storico."""
    ws = get_worksheet(STORICO_SHEET)
    dati_esistenti = ws.get_all_records()
    
    # Genera nuovo ID univoco se necessario
    last_id = max([r.get("ID", 0) for r in dati_esistenti], default=0) if dati_esistenti else 0
    
    nuove_righe = []
    data_str = data.strftime("%Y-%m-%d")
    
    for reg in registrazioni:
        last_id += 1
        nuove_righe.append([
            last_id,
            data_str,
            plesso,
            classe,
            reg["studente"],
            reg["stato"],
            reg.get("note", "")
        ])
    
    if nuove_righe:
        ws.append_rows(nuove_righe, value_input_option="USER_ENTERED")
        carica_statistiche.clear()

def archivia_anno_scolastico(anno: str) -> bool:
    """
    Copia i dati correnti in nuovi fogli di archivio per l'anno scolastico specificato 
    e resetta i fogli correnti.
    """
    try:
        sh = get_spreadsheet()
        suffisso = anno.replace("/", "-")
        nomi_archivio = {
            STORICO_SHEET: f"archivio_storico_{suffisso}",
            ASSENZE_SHEET: f"archivio_assenze_{suffisso}",
        }

        # Verfica preventiva dell'esistenza degli archivi
        for sheet_src, nome_dest in nomi_archivio.items():
            try:
                sh.worksheet(nome_dest)
                st.error(f"Esiste già un archivio per l'anno {anno} (`{nome_dest}`). Scegli un nome o un anno diverso.")
                return False
            except gspread.WorksheetNotFound:
                pass

        # Copia dei dati e creazione dei nuovi fogli
        for sheet_src, nome_dest in nomi_archivio.items():
            ws_src = get_worksheet(sheet_src)
            dati = ws_src.get_all_values()
            
            # Calcolo dinamico di righe e colonne necessarie per evitare errori Out Of Bounds
            num_rows = max(len(dati) + 10, 50)
            num_cols = max([len(riga) for riga in dati], default=10) if dati else 10

            ws_dest = sh.add_worksheet(title=nome_dest, rows=num_rows, cols=num_cols)
            if dati:
                ws_dest.update(values=dati, value_input_option="USER_ENTERED")

        # Resetta i fogli di lavoro correnti
        clear_sheet_content(STORICO_SHEET)
        clear_sheet_content(ASSENZE_SHEET)
        carica_statistiche.clear()
        return True
    except Exception as e:
        st.error(f"Si è verificato un errore durante l'archiviazione: {e}")
        return False

# ==========================================
# STREAMLIT USER INTERFACE
# ==========================================
def main():
    st.title("🏫 Sistema Gestione Plesso Scolastico")

    # Menu di navigazione laterale
    st.sidebar.title("Navigazione")
    menu = st.sidebar.radio(
        "Seleziona Funzionalità",
        ["Inserimento Presenze", "Gestione Assenze", "Statistiche & Archiviazione"]
    )

    # --------------------------------------
    # MENU 1: INSERIMENTO PRESENZE
    # --------------------------------------
    if menu == "Inserimento Presenze":
        st.header("📋 Registrazione Presenze Giornaliere")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            data_reg = st.date_input("Data", datetime.date.today())
        with col2:
            plesso_sel = st.selectbox("Plesso", list(PLESSI_E_CLASSI.keys()))
        with col3:
            classe_sel = st.selectbox("Classe", PLESSI_E_CLASSI[plesso_sel])

        st.subheader(f"Studenti - Classe {classe_sel} ({plesso_sel})")
        
        # Simula o recupera una lista studenti per dimostrazione
        studenti_demo = [f"Studente {i}" for i in range(1, 11)]
        
        with st.form("form_presenze"):
            registrazioni = []
            for st_name in studenti_demo:
                c_nome, c_stato, c_note = st.columns([3, 3, 4])
                with c_nome:
                    st.write(st_name)
                with c_stato:
                    stato = st.selectbox(f"Stato_{st_name}", STATI_VALIDI, key=f"st_{st_name}", label_visibility="collapsed")
                with c_note:
                    note = st.text_input(f"Note_{st_name}", key=f"nt_{st_name}", placeholder="Note opzionali...", label_visibility="collapsed")
                
                registrazioni.append({"studente": st_name, "stato": stato, "note": note})

            submitted = st.form_submit_button("💾 Salva Presenze")
            if submitted:
                registra_presenze(data_reg, plesso_sel, classe_sel, registrazioni)
                st.success("Presenze registrate con successo!")

    # --------------------------------------
    # MENU 2: GESTIONE ASSENZE
    # --------------------------------------
    elif menu == "Gestione Assenze":
        st.header("🏥 Gestione e Giustificazioni Assenze")
        st.info("Funzionalità di controllo e rilevazione assenze prolungate.")
        
        df_storico, _ = carica_statistiche()
        if not df_storico.empty:
            df_assenze = df_storico[df_storico["Stato"].isin(["Assente", "Non Giustificato"])]
            st.dataframe(df_assenze, use_container_width=True)
        else:
            st.warning("Nessun dato di assenza presente nel database.")

    # --------------------------------------
    # MENU 3: STATISTICHE & ARCHIVIAZIONE
    # --------------------------------------
    elif menu == "Statistiche & Archiviazione":
        st.header("📊 Analisi e Reportistica")
        
        df_storico, df_assenze = carica_statistiche()
        
        if not df_storico.empty:
            st.subheader("Riepilogo generale presenze")
            st.dataframe(df_storico, use_container_width=True)
        else:
            st.info("Nessun dato registrato nello storico.")

        st.markdown("---")
        
        # SEZIONE ARCHIVIAZIONE ANNO SCOLASTICO
        st.header("📦 Archivia anno scolastico")
        st.caption("L'archiviazione salverà i dati correnti in nuovi fogli separati e ripristinerà i fogli principali per il nuovo anno.")
        
        col_anno, col_btn = st.columns([3, 2])
        with col_anno:
            anno_scolastico = st.text_input(
                "Anno Scolastico da archiviare", 
                placeholder="es. 2023-2024",
                help="Inserire l'anno nel formato YYYY-YYYY o YYYY/YYYY"
            )
        
        with col_btn:
            st.write(" ") # Spaziatore verticale
            st.write(" ")
            if st.button("🚨 Archivia e Resetta", type="primary"):
                if not anno_scolastico.strip():
                    st.warning("Inserisci prima l'anno scolastico!")
                else:
                    with st.spinner("Archiviazione in corso..."):
                        if archivia_anno_scolastico(anno_scolastico.strip()):
                            st.success(f"Anno scolastico {anno_scolastico} archiviato con successo!")

if __name__ == "__main__":
    main()
