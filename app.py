import streamlit as st
import pandas as pd
import re
import io
import json
import zipfile
import html as html_lib
import urllib.parse
import gspread
import gspread_dataframe as gd
from google.oauth2.service_account import Credentials
from datetime import datetime

# =========================
# VERSIONE APP & CONFIGURAZIONE
# =========================
APP_VERSION = "2.5"
REQUIRED_COLUMNS = ["Docente", "Giorno", "Ora", "Classe", "Tipo", "Escludi"]
ORARIO_SHEET     = "orario"
STORICO_SHEET    = "storico"
ASSENZE_SHEET    = "assenze"
GIORNI_SETTIMANA = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]
ORE_LEZIONE      = ["I", "II", "III", "IV", "V", "VI"]
TIPI_LEZIONE     = ["Lezione", "Sostegno", "Altro"]

try:
    SPREADSHEET_NAME = st.secrets["app"]["spreadsheet_name"]
    PLESSO_NAME      = st.secrets["app"]["plesso_name"]
except KeyError:
    st.error(
        "Configurazione mancante nei secrets. Aggiungi in Settings → Secrets:\n\n"
        "```toml\n[app]\nspreadsheet_name = \"OrarioSostituzioni_Centrale\"\n"
        "plesso_name      = \"Plesso Centrale\"\n```"
    )
    st.stop()

# =========================
# STILI CSS PERSONALIZZATI
# =========================
st.markdown("""
<style>
:root {
    --dc-marrone-scuro: #3A2E1F;
    --dc-marrone: #9C5F2C;
    --dc-arancio: #C97D3D;
    --dc-crema: #EFE6D3;
    --dc-crema-chiaro: #FBF4E6;
    --dc-bordo: #E3D9C2;
}

html, body, [class*="css"] {
    font-size: 17px;
}

.block-container {
    padding-left: 1rem;
    padding-right: 1rem;
    padding-top: calc(3.2rem + env(safe-area-inset-top, 0px));
}

header[data-testid="stHeader"] {
    background: #FFFFFF;
}

.stButton button {
    width: 100%;
    min-height: 3em;
    padding: 0.85em 1em;
    font-size: 1.1em;
    font-weight: 700;
    border-radius: 14px;
    border: none;
    background: linear-gradient(135deg, var(--dc-arancio), #B56A2E);
    color: white;
    box-shadow: 0 3px 8px rgba(58, 46, 31, 0.25);
    transition: transform 0.05s ease, box-shadow 0.15s ease;
}
.stButton button:hover {
    box-shadow: 0 5px 12px rgba(58, 46, 31, 0.3);
}
.stButton button:active {
    transform: scale(0.98);
}
.stButton button[kind="secondary"] {
    background: var(--dc-crema-chiaro);
    color: var(--dc-marrone-scuro);
    border: 1.5px solid var(--dc-bordo);
    box-shadow: none;
}

[data-testid="stSegmentedControl"] label {
    min-height: 3em;
    padding: 0.6em 0.8em !important;
    font-size: 1.08em !important;
    font-weight: 600;
    border-radius: 14px !important;
    border: 1.5px solid var(--dc-bordo) !important;
    background: var(--dc-crema-chiaro) !important;
    color: var(--dc-marrone-scuro) !important;
}
[data-testid="stSegmentedControl"] label[aria-checked="true"],
[data-testid="stSegmentedControl"] label[data-selected="true"] {
    background: var(--dc-arancio) !important;
    border-color: var(--dc-arancio) !important;
    color: white !important;
    box-shadow: 0 3px 8px rgba(58, 46, 31, 0.25);
}

h1, h2 {
    color: var(--dc-marrone-scuro);
    border-bottom: 2px solid var(--dc-bordo);
    padding-bottom: 0.35em;
    margin-top: 1.1em;
}
h1 { font-size: 1.7em; }
h2 { font-size: 1.4em; }
h3 { font-size: 1.2em; color: var(--dc-marrone); }

.btn-whatsapp {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 100%;
    padding: 0.8em;
    background-color: #25D366;
    color: white !important;
    font-weight: bold;
    text-decoration: none;
    border-radius: 14px;
    margin-top: 5px;
    box-shadow: 0 3px 8px rgba(0, 0, 0, 0.15);
}
</style>
""", unsafe_allow_html=True)

# =========================
# INTEGRAZIONE GOOGLE DRIVE / SHEETS
# =========================
SCOPE = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

@st.cache_resource(show_spinner=False)
def get_gdrive_client():
    creds = Credentials.from_service_account_info(st.secrets["gdrive"], scopes=SCOPE)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    client = get_gdrive_client()
    try:
        return client.open(SPREADSHEET_NAME)
    except gspread.SpreadsheetNotFound:
        return client.create(SPREADSHEET_NAME)

@st.cache_resource(show_spinner=False)
def get_worksheet(sheet_name: str):
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

def ensure_sheets_exist():
    for sheet in (ORARIO_SHEET, STORICO_SHEET, ASSENZE_SHEET):
        get_worksheet(sheet)

# =========================
# GESTIONE DATI (LOAD & SAVE)
# =========================
@st.cache_data(ttl=300, show_spinner=False)
def carica_orario():
    try:
        ws = get_worksheet(ORARIO_SHEET)
        df = gd.get_as_dataframe(ws, evaluate_formulas=True, header=0)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        for col in REQUIRED_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df.dropna(how='all').copy()
        df["Escludi"] = df["Escludi"].replace({pd.NA: False, "": False}).astype(bool)
        for col in ["Tipo", "Docente", "Giorno", "Ora", "Classe"]:
            df[col] = df[col].astype(str).str.strip().fillna("")
        return df[REQUIRED_COLUMNS]
    except Exception as e:
        st.error(f"Errore nel caricamento dell'orario: {e}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

def salva_orario(df):
    try:
        ws = get_worksheet(ORARIO_SHEET)
        df_to_save = df.copy()
        for col in REQUIRED_COLUMNS:
            if col not in df_to_save.columns:
                df_to_save[col] = ""
        gd.set_with_dataframe(ws, df_to_save[REQUIRED_COLUMNS], include_index=False, include_column_header=True)
        carica_orario.clear()
        return True
    except Exception as e:
        st.error(f"Errore durante il salvataggio dell'orario: {e}")
        return False

@st.cache_data(ttl=300, show_spinner=False)
def carica_statistiche():
    try:
        ws_storico = get_worksheet(STORICO_SHEET)
        ws_assenze = get_worksheet(ASSENZE_SHEET)
        df_storico = gd.get_as_dataframe(ws_storico, header=0).dropna(how='all')
        df_assenze = gd.get_as_dataframe(ws_assenze, header=0).dropna(how='all')
        
        if not df_storico.empty:
            df_storico["data"] = pd.to_datetime(df_storico["data"], errors="coerce").dt.strftime("%Y-%m-%d")
            df_storico["ore"] = pd.to_numeric(df_storico["ore"], errors="coerce").fillna(0).astype(int)
            for c in ["docente", "giorno"]:
                df_storico[c] = df_storico[c].astype(str).str.strip().str.lower()
                    
        if not df_assenze.empty:
            df_assenze["data"] = pd.to_datetime(df_assenze["data"], errors="coerce").dt.strftime("%Y-%m-%d")
            for c in ["docente", "giorno", "ora", "classe"]:
                df_assenze[c] = df_assenze[c].astype(str).str.strip()
                    
        return (
            df_storico if not df_storico.empty else pd.DataFrame(columns=["data", "giorno", "docente", "ore"]),
            df_assenze if not df_assenze.empty else pd.DataFrame(columns=["data", "giorno", "docente", "ora", "classe"])
        )
    except Exception as e:
        st.error(f"Errore nel caricamento delle statistiche: {e}")
        return pd.DataFrame(columns=["data", "giorno", "docente", "ore"]), pd.DataFrame(columns=["data", "giorno", "docente", "ora", "classe"])

def salva_storico_assenze(data_sostituzione, giorno_assente, sostituzioni_df, ore_assenti):
    try:
        ws_storico = get_worksheet(STORICO_SHEET)
        ws_assenze = get_worksheet(ASSENZE_SHEET)

        sostituzioni_effettive = sostituzioni_df[
            sostituzioni_df["Sostituto"].notna() &
            (sostituzioni_df["Sostituto"].str.strip() != "") &
            (sostituzioni_df["Sostituto"].str.strip().str.lower() != "nessuno")
        ].copy()

        storico_data = [
            [str(data_sostituzione), giorno_assente, row["Sostituto"], 1]
            for _, row in sostituzioni_effettive.iterrows()
        ]

        ore_effettivamente_assenti = ore_assenti[
            ore_assenti["Ora"].isin(sostituzioni_effettive["Ora"])
        ].copy()

        assenze_data = [
            [str(data_sostituzione), giorno_assente, row["Docente"], row["Ora"], row["Classe"]]
            for _, row in ore_effettivamente_assenti.iterrows()
        ]

        if storico_data:
            ws_storico.append_rows(storico_data, value_input_option="USER_ENTERED")
        if assenze_data:
            ws_assenze.append_rows(assenze_data, value_input_option="USER_ENTERED")

        carica_statistiche.clear()
        return True
    except Exception as e:
        st.error(f"Errore durante la registrazione: {e}")
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

def archivia_anno_scolastico(anno: str):
    try:
        sh = get_spreadsheet()
        suffisso = anno.replace("/", "-")
        nomi_archivio = {
            STORICO_SHEET: f"archivio_storico_{suffisso}",
            ASSENZE_SHEET: f"archivio_assenze_{suffisso}",
        }

        for sheet_src, nome_dest in nomi_archivio.items():
            try:
                sh.worksheet(nome_dest)
                st.error(f"Esiste già un archivio per l'anno {anno} ({nome_dest}).")
                return False
            except gspread.WorksheetNotFound:
                pass

            ws_src = get_worksheet(sheet_src)
            dati = ws_src.get_all_values()
            
            num_rows = max(len(dati) + 10, 50)
            num_cols = max([len(riga) for riga in dati], default=10) if dati else 10

            ws_dest = sh.add_worksheet(title=nome_dest, rows=num_rows, cols=num_cols)
            if dati:
                ws_dest.update(values=dati, value_input_option="USER_ENTERED")

        clear_sheet_content(STORICO_SHEET)
        clear_sheet_content(ASSENZE_SHEET)
        return True
    except Exception as e:
        st.error(f"Errore durante l'archiviazione: {e}")
        return False

# =========================
# UTILITIES EXPORT & STAMPA
# =========================
def create_backup():
    try:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for s_name in [ORARIO_SHEET, STORICO_SHEET, ASSENZE_SHEET]:
                ws = get_worksheet(s_name)
                df = gd.get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how='all')
                zip_file.writestr(f"{s_name}.csv", df.to_csv(index=False).encode('utf-8'))
        zip_buffer.seek(0)
        return zip_buffer
    except Exception as e:
        st.error(f"Errore creazione backup ZIP: {e}")
        return None

def create_excel_export():
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for label, s_name in [('Orario', ORARIO_SHEET), ('Storico Sostituzioni', STORICO_SHEET), ('Storico Assenze', ASSENZE_SHEET)]:
                ws = get_worksheet(s_name)
                df = gd.get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how='all')
                df.to_excel(writer, sheet_name=label, index=False)
                worksheet = writer.sheets[label]
                for idx, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max() if not df.empty else 0, len(str(col))) + 3
                    worksheet.set_column(idx, idx, max_len)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Errore generazione file Excel: {e}")
        return None

def build_docente_tipo_map(df):
    if df.empty:
        return {}
    tmp = df[df["Tipo"].astype(str).str.strip() != ""]
    return tmp.groupby("Docente")["Tipo"].first().to_dict() if not tmp.empty else {}

def trova_conflitti_orario(df):
    if df.empty:
        return []
    conteggi = df.groupby(["Docente", "Giorno", "Ora"]).size()
    return [idx for idx, n in conteggi.items() if n > 1 and idx[0].strip() != ""]

def _colore_tipo(label):
    if "[S]" in label and "[NP]" not in label:
        return "#6B8F71", "white", "🔵"
    elif "[C] [USCITA]" in label:
        return "#5E7A93", "white", "🟡"
    elif "[NP]" in label:
        return "#9C9C7A", "white", "🟢"
    elif "[C]" in label:
        return "#C97D3D", "white", "🔴"
    return "#E3D9C2", "#3A2E1F", "–"

def _badge_sostituto(label):
    if label == "Nessuno":
        return '<span style="color:#9C5F2C;font-style:italic;">— nessuno —</span>'
    bg, fg, _ = _colore_tipo(label)
    nome = re.sub(r'\[.*?\]|🔵|🟡|🟢|🔴', '', label).strip()
    return f'<span style="background:{bg};color:{fg};border-radius:8px;padding:3px 10px;font-weight:700;font-size:0.9em;">{nome}</span>'

def download_orario(df):
    if not df.empty:
        st.download_button("⬇️ Scarica orario (CSV)", data=df.to_csv(index=False), file_name="orario.csv", mime="text/csv")

def genera_html_stampa_sostituzioni(plesso_nome, data_sost, giorno_assente, tabella_df):
    titolo_giorno = f"{giorno_assente} {data_sost.strftime('%d/%m/%Y')}"
    righe_html = ""
    tabella_ordinata = tabella_df.copy()
    tabella_ordinata["Ora"] = pd.Categorical(tabella_ordinata["Ora"], categories=ORE_LEZIONE, ordered=True)
    tabella_ordinata = tabella_ordinata.sort_values(["Ora", "Classe"])

    for _, r in tabella_ordinata.iterrows():
        ora = html_lib.escape(str(r["Ora"]))
        classe = html_lib.escape(str(r["Classe"]))
        assente = html_lib.escape(str(r["Assente"]))
        sost_pulito = re.sub(r'\[.*?\]|🔵|🟡|🟢|🔴', '', str(r["Sostituzione"])).strip()
        if sost_pulito in ("Nessuno", "", "—"):
            sost_pulito = "— DA COPRIRE —"
        sostituto = html_lib.escape(sost_pulito)
        riga_scoperta = ' class="scoperta"' if sost_pulito == "— DA COPRIRE —" else ""
        righe_html += f"<tr{riga_scoperta}><td>{ora}</td><td>{classe}</td><td>{assente}</td><td>{sostituto}</td></tr>\n"

    return f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<title>Sostituzioni {html_lib.escape(titolo_giorno)} — {html_lib.escape(plesso_nome)}</title>
<style>
  @page {{ margin: 1.5cm; }}
  body {{ font-family: Georgia, serif; color: #1a1a1a; padding: 0 0.5cm; }}
  h1 {{ font-size: 1.5em; border-bottom: 3px solid #C97D3D; padding-bottom: 0.2em; }}
  table {{ width: 100%; border-collapse: collapse; margin-top: 1em; }}
  th, td {{ border: 1px solid #999; padding: 8px 10px; text-align: left; }}
  th {{ background: #EFE6D3; text-transform: uppercase; }}
  tr.scoperta td {{ font-weight: bold; }}
  @media print {{ .no-print {{ display: none; }} }}
</style>
</head>
<body>
  <h1>📚 Sostituzioni — {html_lib.escape(plesso_nome)}</h1>
  <p><strong>{html_lib.escape(titolo_giorno)}</strong></p>
  <table>
    <thead><tr><th>Ora</th><th>Classe</th><th>Assente</th><th>Sostituzione</th></tr></thead>
    <tbody>{righe_html if righe_html else '<tr><td colspan="4">Nessuna sostituzione.</td></tr>'}</tbody>
  </table>
</body>
</html>"""

def pulsante_stampa_sostituzioni(plesso_nome, data_sost, giorno_assente, tabella_df):
    pagina_json = json.dumps(genera_html_stampa_sostituzioni(plesso_nome, data_sost, giorno_assente, tabella_df))
    st.components.v1.html(f"""
<button id="stampa-btn" style="width:100%; padding:0.7em; font-size:1em; font-weight:bold; background:#3A2E1F; color:white; border:none; border-radius:10px; cursor:pointer;">🖨️ Stampa / PDF per la bacheca</button>
<script>
document.getElementById('stampa-btn').addEventListener('click', function() {{
    var win = window.open('', '_blank');
    win.document.write({pagina_json});
    win.document.close();
    win.focus();
    setTimeout(function() {{ win.print(); }}, 300);
}});
</script>
""", height=55)

def vista_pivot_docenti(df, mode="classi"):
    if df.empty:
        st.warning("Nessun orario disponibile.")
        return

    dfp = df.copy()
    if mode == "classi":
        dfp["Info"] = dfp["Docente"]
        pivot = dfp.pivot_table(
            index=["Giorno", "Ora"],
            columns="Classe",
            values="Info",
            aggfunc=lambda x: " / ".join(list(dict.fromkeys(x)))
        ).fillna("-")

        def sort_classi(c):
            m = re.match(r"(\d+)\s*([A-Za-zÀ-ÖØ-öø-ÿ]+)", str(c))
            return (int(m.group(1)), m.group(2).upper()) if m else (10**9, str(c))
        
        pivot = pivot.reindex(sorted(pivot.columns, key=sort_classi), axis=1)
        pivot = pivot.reindex(pd.MultiIndex.from_product([GIORNI_SETTIMANA, ORE_LEZIONE], names=["Giorno", "Ora"])).reset_index()

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

        st.dataframe(pivot.style.set_properties(**{"text-align": "center"}), use_container_width=True, hide_index=True)

# =========================
# INTERFACCIA E NAVIGAZIONE
# =========================
st.markdown(f"""
<div style="display:flex; align-items:center; justify-content:space-between; background:linear-gradient(135deg, #EFE6D3, #FBF4E6); border:1.5px solid #E3D9C2; border-radius:16px; padding:14px 18px; margin-bottom:0.8rem;">
  <div style="display:flex; align-items:center; gap:14px;">
    <div style="font-size:2.1em;">🏫</div>
    <div>
      <div style="font-size:1.35em; font-weight:800; color:#3A2E1F;">{PLESSO_NAME}</div>
      <div style="font-size:0.95em; color:#9C5F2C; font-weight:600;">Gestione sostituzioni docenti</div>
    </div>
  </div>
  <div style="display:flex; flex-direction:column; align-items:center;">
    <a href="?ricarica=1" target="_self" title="Ricarica dati" style="font-size:1.5em; text-decoration:none; color:#C97D3D;">🔄</a>
    <span style="font-size:0.7em; color:#B0A090; font-weight:600;">v{APP_VERSION}</span>
  </div>
</div>
""", unsafe_allow_html=True)

if st.query_params.get("ricarica") == "1":
    carica_orario.clear()
    carica_statistiche.clear()
    st.query_params.clear()
    st.rerun()

ensure_sheets_exist()
orario_df = carica_orario()

ETICHETTE_MENU = {
    "Inserisci/Modifica Orario": "📝 Orario",
    "Gestione Assenze":          "🚨 Assenze",
    "Visualizza Orario":         "📅 Vedi",
    "Statistiche":               "📊 Stats",
}

menu_scelto = st.segmented_control("Navigazione", list(ETICHETTE_MENU.values()), selection_mode="single", default="🚨 Assenze")
menu = {v: k for k, v in ETICHETTE_MENU.items()}.get(menu_scelto)

# --- TAB 1: MODIFICA ORARIO ---
if menu == "Inserisci/Modifica Orario":
    st.header("➕ Modifica orario")

    uploaded_file = st.file_uploader("Carica nuovo orario (CSV)", type="csv")
    if uploaded_file:
        df_tmp = pd.read_csv(uploaded_file)
        if not all(col in df_tmp.columns for col in REQUIRED_COLUMNS):
            st.info("💡 Mappa le colonne del file CSV con le voci di sistema:")
            col_map = {}
            cols_csv = ["-- Seleziona --"] + list(df_tmp.columns)
            
            for req_col in REQUIRED_COLUMNS:
                def_idx = next((i for i, c in enumerate(cols_csv) if req_col.lower() in c.lower()), 0)
                col_map[req_col] = st.selectbox(f"Colonna per {req_col}", cols_csv, index=def_idx, key=f"map_{req_col}")
            
            if st.button("Importa con questo mapping", type="primary"):
                df_nuovo = pd.DataFrame()
                for req_col in REQUIRED_COLUMNS:
                    df_nuovo[req_col] = df_tmp[col_map[req_col]] if col_map[req_col] != "-- Seleziona --" else ("" if req_col != "Escludi" else False)
                
                if salva_orario(df_nuovo):
                    st.success("Orario caricato con successo ✅")
                    st.rerun()
        else:
            if salva_orario(df_tmp[REQUIRED_COLUMNS]):
                st.success("Orario caricato con successo ✅")
                st.rerun()

    with st.expander("Aggiungi riga orario"):
        doc = st.selectbox("Docente", ["➕ Nuovo"] + (sorted(orario_df["Docente"].unique()) if not orario_df.empty else []))
        doc_val = st.text_input("Nome Docente") if doc == "➕ Nuovo" else doc
        gio_val = st.selectbox("Giorno", GIORNI_SETTIMANA)
        ora_val = st.selectbox("Ora", ORE_LEZIONE)
        cls = st.selectbox("Classe", ["➕ Nuova"] + (sorted(orario_df["Classe"].unique()) if not orario_df.empty else []))
        cls_val = st.text_input("Nome Classe") if cls == "➕ Nuova" else cls
        tip_val = st.selectbox("Tipo", TIPI_LEZIONE)
        esc_val = st.checkbox("Escludi da sostituzioni")

        if st.button("Aggiungi Lezione", type="primary"):
            if doc_val and cls_val:
                nuova_riga = pd.DataFrame([{"Docente": doc_val, "Giorno": gio_val, "Ora": ora_val, "Classe": cls_val, "Tipo": tip_val, "Escludi": esc_val}])
                orario_df = pd.concat([orario_df, nuova_riga], ignore_index=True)
                if salva_orario(orario_df):
                    st.success("Orario aggiornato ✅")
                    st.rerun()

    st.subheader("📝 Modifica diretta")
    if not orario_df.empty:
        edited_df = st.data_editor(
            orario_df.sort_values(by=["Docente"]),
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Giorno": st.column_config.SelectboxColumn("Giorno", options=GIORNI_SETTIMANA, required=True),
                "Ora": st.column_config.SelectboxColumn("Ora", options=ORE_LEZIONE, required=True),
                "Tipo": st.column_config.SelectboxColumn("Tipo", options=TIPI_LEZIONE, required=True),
                "Escludi": st.column_config.CheckboxColumn("Escludi"),
            }
        )
        if st.button("Salva modifiche", type="primary"):
            if salva_orario(edited_df):
                st.success("Orario salvato ✅")
                st.rerun()

    download_orario(orario_df)

# --- TAB 2: GESTIONE ASSENZE ---
elif menu == "Gestione Assenze":
    st.header("🚨 Gestione Assenze")

    if orario_df.empty:
        st.warning("Carica un orario per procedere.")
    else:
        data_sost = st.date_input("Data sostituzione")
        giorni_m = {"Monday": "Lunedì", "Tuesday": "Martedì", "Wednesday": "Mercoledì", "Thursday": "Giovedì", "Friday": "Venerdì"}
        giorno_assente = giorni_m.get(data_sost.strftime("%A"), data_sost.strftime("%A"))

        docenti_assenti = st.multiselect("Docenti assenti", sorted(orario_df["Docente"].unique()))
        
        classi_uscita_per_ora = {}
        with st.expander("🚌 Classi in uscita didattica"):
            classi_uscita = st.multiselect("Classi fuori sede", sorted(orario_df["Classe"].unique()))
            for c_u in classi_uscita:
                ore_c = [o for o in ORE_LEZIONE if not orario_df[(orario_df["Classe"] == c_u) & (orario_df["Giorno"] == giorno_assente) & (orario_df["Ora"] == o)].empty]
                ore_sel = st.multiselect(f"Ore uscita per {c_u}", ore_c, default=ore_c, key=f"u_{c_u}")
                for o_sel in ore_sel:
                    classi_uscita_per_ora.setdefault(o_sel, set()).add(c_u)

        if docenti_assenti:
            ore_assenti = orario_df[(orario_df["Docente"].isin(docenti_assenti)) & (orario_df["Giorno"] == giorno_assente)].copy()
            if not ore_assenti.empty:
                st.subheader("📌 Tabella di Algoritmo Sostituzioni")
                sostituzioni = []
                docente_tipo_map = build_docente_tipo_map(orario_df)
                tutti_docenti = sorted(orario_df["Docente"].unique())
                escludi_set = set(orario_df[orario_df["Escludi"]]["Docente"].unique())

                for _, row in ore_assenti.iterrows():
                    ora, classe, assente = row["Ora"], row["Classe"], row["Docente"]
                    
                    presenti_df = orario_df[(orario_df["Giorno"] == giorno_assente) & (orario_df["Ora"] == ora) & (~orario_df["Escludi"]) & (~orario_df["Docente"].isin(docenti_assenti))].copy()
                    
                    options, added = ["Nessuno"], set()
                    
                    # Logica priorità
                    same_sost = presenti_df[(presenti_df["Tipo"].str.lower() == "sostegno") & (presenti_df["Classe"] == classe)]["Docente"].tolist()
                    for d in sorted(same_sost):
                        options.append(f"🔵 [S] {d}"); added.add(d)

                    other_sost = presenti_df[(presenti_df["Tipo"].str.lower() == "sostegno")]["Docente"].tolist()
                    for d in sorted(other_sost):
                        if d not in added: options.append(f"🔵 [S] {d}"); added.add(d)

                    uscita_set = classi_uscita_per_ora.get(ora, set())
                    curr_liberi = presenti_df[(presenti_df["Tipo"].str.lower() != "sostegno") & (presenti_df["Classe"].isin(uscita_set))]["Docente"].tolist()
                    for d in sorted(curr_liberi):
                        if d not in added: options.append(f"🟡 [C] [USCITA] {d}"); added.add(d)

                    curr_occ = presenti_df[(presenti_df["Tipo"].str.lower() != "sostegno") & (~presenti_df["Classe"].isin(uscita_set))]["Docente"].tolist()
                    for d in sorted(curr_occ):
                        if d not in added: options.append(f"🔴 [C] {d}"); added.add(d)

                    np_cand = [d for d in tutti_docenti if d not in set(presenti_df["Docente"]) and d not in docenti_assenti and d not in escludi_set]
                    for d in sorted(np_cand):
                        if d not in added:
                            prefix = "🟢 [S] [NP]" if docente_tipo_map.get(d, "").lower() == "sostegno" else "🟢 [C] [NP]"
                            options.append(f"{prefix} {d}"); added.add(d)

                    scelta = st.selectbox(f"Ora {ora} - Cl. {classe} (Assente: {assente})", options, key=f"s_{assente}_{ora}_{classe}")
                    nome_pulito = re.sub(r'\[.*?\]|🔵|🟡|🟢|🔴', '', scelta).strip() if scelta != "Nessuno" else "Nessuno"
                    
                    sostituzioni.append({"Ora": ora, "Classe": classe, "Assente": assente, "Sostituto_display": scelta, "Sostituto": nome_pulito})

                sost_df = pd.DataFrame(sostituzioni)

                # Output Testuale e WhatsApp
                st.subheader("📝 Format Comunicazione Mobile")
                txt_out = "Buongiorno, supplenze.©\n\n"
                for o_c, grp in sost_df.groupby("Ora"):
                    txt_out += f"🕐 *{o_c} ORA*\n"
                    for _, r in grp.iterrows():
                        txt_out += f"Classe {r['Classe']}\n👩‍🏫 Assente: {r['Assente']}\n✅ Sostituzione: {r['Sostituto']}\n\n"

                st.text_area("Testo da copiare", value=txt_out.strip(), height=200)
                st.markdown(f'<a href="https://api.whatsapp.com/send?text={urllib.parse.quote(txt_out.strip())}" target="_blank" class="btn-whatsapp">📲 Invia su WhatsApp</a>', unsafe_allow_html=True)
                
                st.subheader("🖨️ Stampa per Bacheca")
                pulsante_stampa_sostituzioni(PLESSO_NAME, data_sost, giorno_assente, sost_df.rename(columns={"Sostituto_display": "Sostituzione"}))

                if st.button("💾 Salva Sostituzioni su Google Sheets", type="primary"):
                    if salva_storico_assenze(data_sost, giorno_assente, sost_df, ore_assenti):
                        st.success("Dati registrati nello Storico con successo! ✅")

# --- TAB 3: VISUALIZZA ORARIO ---
elif menu == "Visualizza Orario":
    st.header("📅 Vista Orario")
    if not orario_df.empty:
        c1, c2 = st.columns(2)
        f_doc = c1.multiselect("Filtra Docente", sorted(orario_df["Docente"].unique()))
        f_cls = c2.multiselect("Filtra Classe", sorted(orario_df["Classe"].unique()))

        df_v = orario_df.copy()
        if f_doc: df_v = df_v[df_v["Docente"].isin(f_doc)]
        if f_cls: df_v = df_v[df_v["Classe"].isin(f_cls)]

        vista_pivot_docenti(df_v, mode="classi")
        download_orario(orario_df)

# --- TAB 4: STATISTICHE ---
elif menu == "Statistiche":
    st.header("📊 Statistiche & Reporting")
    df_st, df_as = carica_statistiche()

    if not df_st.empty:
        st.subheader("🟢 Sostituzioni Eseguite")
        sum_st = df_st.groupby("docente")["ore"].sum().reset_index().sort_values("ore", ascending=False)
        st.dataframe(sum_st, use_container_width=True, hide_index=True)
        st.bar_chart(sum_st.set_index("docente"))

    if not df_as.empty:
        st.subheader("🟠 Conteggio Assenze")
        sum_as = df_as.groupby("docente")["ora"].count().reset_index().rename(columns={"ora": "Totale Ore Assenti"}).sort_values("Totale Ore Assenti", ascending=False)
        st.dataframe(sum_as, use_container_width=True, hide_index=True)

    st.subheader("📦 Download & Archiviazione Anno")
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        ex = create_excel_export()
        if ex: st.download_button("📊 Scarica Report Excel (.xlsx)", data=ex, file_name=f"Report_Orari_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_e2:
        bk = create_backup()
        if bk: st.download_button("⬇️ Scarica Backup ZIP", data=bk, file_name=f"Backup_Orari_{datetime.now().strftime('%Y%m%d')}.zip", mime="application/zip")

    st.markdown("---")
    anno_txt = st.text_input("Anno Scolastico da archiviare (es. 2024-25)")
    chk_arch = st.checkbox("Confermo la chiusura dell'anno corrente e l'azzeramento dello storico attivo")
    if st.button("📦 Archivia Anno Scolastico") and anno_txt and chk_arch:
        if archivia_anno_scolastico(anno_txt.strip()):
            st.success(f"Anno {anno_txt} archiviato su Google Sheets con successo ✅")
