import streamlit as st
import pandas as pd
import os
import re
import gspread
import gspread_dataframe as gd
from oauth2client.service_account import ServiceAccountCredentials
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

# =========================
# CLIENT GOOGLE DRIVE
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

# =========================
# ...
# (codice invariato: funzioni carica/salva, gestione pivot ecc.)
# =========================

# --- GESTIONE ASSENZE ---
elif menu == "Gestione Assenze":
    st.header("🚨 Gestione Assenze")

    if orario_df.empty:
        st.warning("Non hai ancora caricato nessun orario.")
    else:
        # ... (codice invariato per docenti assenti, proposta sostituti, edited_df, tabella_df, testo_output)

        # --- Step 1: conferma tabella (non salva ancora) ---
        if st.button("✅ Conferma tabella (non salva ancora)"):
            # 🔧 FIX: definiamo edited_df_sorted prima del controllo conflitti
            ordine_ore = ["I", "II", "III", "IV", "V", "VI"]
            edited_df_sorted = edited_df.copy()
            edited_df_sorted["Ora"] = pd.Categorical(edited_df_sorted["Ora"], categories=ordine_ore, ordered=True)
            edited_df_sorted = edited_df_sorted.sort_values(["Ora", "Classe"]).reset_index(drop=True)

            # Controllo conflitti: stesso docente in più classi nella stessa ora
            conflitti = []
            for ora in edited_df_sorted["Ora"].unique():
                assegnazioni = edited_df_sorted[edited_df_sorted["Ora"] == ora]
                sostituti = [s for s in assegnazioni["Sostituto"] if s != "Nessuno"]
                duplicati = [s for s in sostituti if sostituti.count(s) > 1]
                if duplicati:
                    conflitti.append((ora, list(set(duplicati))))

            if conflitti:
                st.error("⚠️ Errore: lo stesso docente è stato assegnato a più classi nella stessa ora:")
                for ora, docs in conflitti:
                    st.write(f"- Ora {ora}: {', '.join(docs)}")
                st.stop()

            # Memorizza tutto in sessione per lo Step 2
            st.session_state["sostituzioni_confermate"] = edited_df_sorted.copy()
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
                    if salva_storico_assenze(data_tmp, giorno_tmp, sost_df, ore_assenti_session):
                        st.success("Assenze e sostituzioni salvate nello storico ✅")
                        for k in ["sostituzioni_confermate", "ore_assenti_confermate", "data_sostituzione_tmp", "giorno_assente_tmp"]:
                            st.session_state.pop(k, None)
                        try:
                            st.rerun()
                        except Exception:
                            pass
