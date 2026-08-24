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
# VERSIONE APP
# =========================
APP_VERSION = "2.4"

# =========================
# CONFIGURAZIONE FILE / SHEETS
# =========================
REQUIRED_COLUMNS    = ["Docente", "Giorno", "Ora", "Classe", "Tipo", "Escludi"]
ORARIO_SHEET        = "orario"
STORICO_SHEET       = "storico"
ASSENZE_SHEET       = "assenze"
GIORNI_SETTIMANA    = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]
ORE_LEZIONE         = ["I", "II", "III", "IV", "V", "VI"]
TIPI_LEZIONE        = ["Lezione", "Sostegno", "Altro"]

try:
    SPREADSHEET_NAME = st.secrets["app"]["spreadsheet_name"]
    PLESSO_NAME      = st.secrets["app"]["plesso_name"]
except KeyError:
    st.error(
        "Configurazione mancante nei secrets. Aggiungi in Settings → Secrets:\n\n"
        "```\n[app]\nspreadsheet_name = \"OrarioSostituzioni_Centrale\"\n"
        "plesso_name      = \"Plesso Centrale\"\n```"
    )
    st.stop()

# =========================
# STILI PERSONALIZZATI
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

.block-container > div:first-child img[alt] {
    width: 100% !important;
    height: auto !important;
    max-width: 100%;
    object-fit: contain;
    display: block;
    border-radius: 16px;
    box-shadow: 0 4px 14px rgba(58, 46, 31, 0.15);
    margin-bottom: 0.6rem;
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

[data-testid="stSegmentedControl"] {
    gap: 0.4rem;
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

.stDataFrame, .stDataEditor {
    font-size: 0.95em !important;
    border-radius: 12px !important;
}

.stSelectbox, .stTextInput, .stDateInput, .stMultiSelect {
    width: 100% !important;
}
.stSelectbox > div, .stTextInput > div, .stDateInput > div, .stMultiSelect > div {
    min-height: 2.6em;
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
.btn-telegram {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 100%;
    padding: 0.8em;
    background-color: #0088cc;
    color: white !important;
    font-weight: bold;
    text-decoration: none;
    border-radius: 14px;
    margin-top: 5px;
    box-shadow: 0 3px 8px rgba(0, 0, 0, 0.15);
}

@media (max-width: 480px) {
    html, body, [class*="css"] { font-size: 18px; }
    .stButton button { font-size: 1.15em; padding: 1em; }
    [data-testid="stSegmentedControl"] label { font-size: 1.1em !important; }
}
</style>
""", unsafe_allow_html=True)

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
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        for col in REQUIRED_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df.dropna(how='all').copy()
        if "Escludi" in df.columns:
            df["Escludi"] = df["Escludi"].replace({pd.NA: False, "": False}).astype(bool)
        else:
            df["Escludi"] = False
        for col in ["Tipo", "Docente", "Giorno", "Ora", "Classe"]:
            df[col] = df[col].astype(str).str.strip().fillna("")
        df = df.loc[:, REQUIRED_COLUMNS]
        return df
    except Exception as e:
        st.error(f"Errore nel caricamento dell'orario da Google Sheets: {e}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

def salva_orario(df):
    try:
        ws = get_worksheet(ORARIO_SHEET)
        df_to_save = df.copy()
        for col in REQUIRED_COLUMNS:
            if col not in df_to_save.columns:
                df_to_save[col] = ""
        df_to_save = df_to_save[REQUIRED_COLUMNS]
        gd.set_with_dataframe(ws, df_to_save, include_index=False, include_column_header=True)
        carica_orario.clear()
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio dell'orario su Google Sheets: {e}")
        return False

# =========================
# CARICAMENTO / SALVATAGGIO STATISTICHE
# =========================
@st.cache_data(ttl=300, show_spinner=False)
def carica_statistiche():
    try:
        ws_storico = get_worksheet(STORICO_SHEET)
        ws_assenze = get_worksheet(ASSENZE_SHEET)
        df_storico = gd.get_as_dataframe(ws_storico, header=0).dropna(how='all')
        df_assenze = gd.get_as_dataframe(ws_assenze, header=0).dropna(how='all')
        
        if not df_storico.empty:
            if "data" in df_storico.columns:
                df_storico["data"] = pd.to_datetime(df_storico["data"], errors="coerce").dt.strftime("%Y-%m-%d")
            if "ore" in df_storico.columns:
                df_storico["ore"] = pd.to_numeric(df_storico["ore"], errors="coerce").fillna(0).astype(int)
            for c in ["docente", "giorno"]:
                if c in df_storico.columns:
                    df_storico[c] = df_storico[c].astype(str).str.strip().str.lower()
                    
        if not df_assenze.empty:
            if "data" in df_assenze.columns:
                df_assenze["data"] = pd.to_datetime(df_assenze["data"], errors="coerce").dt.strftime("%Y-%m-%d")
            for c in ["docente", "giorno", "ora", "classe"]:
                if c in df_assenze.columns:
                    df_assenze[c] = df_assenze[c].astype(str).str.strip()
                    
        if df_storico.empty:
            df_storico = pd.DataFrame(columns=["data", "giorno", "docente", "ore"])
        if df_assenze.empty:
            df_assenze = pd.DataFrame(columns=["data", "giorno", "docente", "ora", "classe"])
        return df_storico, df_assenze
    except Exception as e:
        st.error(f"Errore nel caricamento delle statistiche da Google Sheets: {e}")
        return pd.DataFrame(columns=["data", "giorno", "docente", "ore"]), pd.DataFrame(columns=["data", "giorno", "docente", "ora", "classe"])

def salva_storico_assenze(data_sostituzione, giorno_assente, sostituzioni_df, ore_assenti):
    """Salvataggio coordinato ed esplicito su Google Sheets."""
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
        st.error(f"Errore durante il salvataggio dei dati su Google Sheets: {e}")
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
                st.error(f"Esiste già un archivio per l'anno {anno} ({nome_dest}). Scegli un anno diverso.")
                return False
            except gspread.WorksheetNotFound:
                pass

            ws_src = get_worksheet(sheet_src)
            dati = ws_src.get_all_values()
            ws_dest = sh.add_worksheet(title=nome_dest, rows=max(len(dati) + 10, 50), cols=10)
            if dati:
                ws_dest.update(values=dati, value_input_option="USER_ENTERED")

        clear_sheet_content(STORICO_SHEET)
        clear_sheet_content(ASSENZE_SHEET)
        carica_statistiche.clear()
        return True
    except Exception as e:
        st.error(f"Errore durante l'archiviazione: {e}")
        return False

# =========================
# BACKUP ZIP ED EXPORT EXCEL (.xlsx)
# =========================
def create_backup():
    try:
        ws_orario = get_worksheet(ORARIO_SHEET)
        df_orario = gd.get_as_dataframe(ws_orario, evaluate_formulas=True, header=0).dropna(how='all')

        ws_storico = get_worksheet(STORICO_SHEET)
        df_storico = gd.get_as_dataframe(ws_storico, evaluate_formulas=True, header=0).dropna(how='all')

        ws_assenze = get_worksheet(ASSENZE_SHEET)
        df_assenze = gd.get_as_dataframe(ws_assenze, evaluate_formulas=True, header=0).dropna(how='all')

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            zip_file.writestr("orario.csv", df_orario.to_csv(index=False).encode('utf-8'))
            zip_file.writestr("storico.csv", df_storico.to_csv(index=False).encode('utf-8'))
            zip_file.writestr("assenze.csv", df_assenze.to_csv(index=False).encode('utf-8'))

        zip_buffer.seek(0)
        return zip_buffer
    except Exception as e:
        st.error(f"Errore durante la creazione del backup: {e}")
        return None

def create_excel_export():
    try:
        ws_orario = get_worksheet(ORARIO_SHEET)
        df_orario = gd.get_as_dataframe(ws_orario, evaluate_formulas=True, header=0).dropna(how='all')

        ws_storico = get_worksheet(STORICO_SHEET)
        df_storico = gd.get_as_dataframe(ws_storico, evaluate_formulas=True, header=0).dropna(how='all')

        ws_assenze = get_worksheet(ASSENZE_SHEET)
        df_assenze = gd.get_as_dataframe(ws_assenze, evaluate_formulas=True, header=0).dropna(how='all')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_orario.to_excel(writer, sheet_name='Orario', index=False)
            df_storico.to_excel(writer, sheet_name='Storico Sostituzioni', index=False)
            df_assenze.to_excel(writer, sheet_name='Storico Assenze', index=False)
            
            for sheet_name, df_sheet in [('Orario', df_orario), ('Storico Sostituzioni', df_storico), ('Storico Assenze', df_assenze)]:
                worksheet = writer.sheets[sheet_name]
                for idx, col in enumerate(df_sheet.columns):
                    max_len = max(df_sheet[col].astype(str).map(len).max() if not df_sheet.empty else 0, len(str(col))) + 3
                    worksheet.set_column(idx, idx, max_len)
                    
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Errore durante la generazione dell'Excel: {e}")
        return None

# =========================
# UTILITA' PER DOWNLOAD E PIVOT
# =========================
def build_docente_tipo_map(df):
    if df.empty:
        return {}
    tmp = df[df["Tipo"].astype(str).str.strip() != ""]
    if tmp.empty:
        return {}
    return tmp.groupby("Docente")["Tipo"].first().to_dict()

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
    nome = (label.replace("[S] [NP] ", "").replace("[C] [NP] ", "")
                 .replace("[C] [USCITA] ", "").replace("[S] ", "")
                 .replace("[C] ", "").replace("🔵 ", "").replace("🟡 ", "")
                 .replace("🟢 ", "").replace("🔴 ", "").strip())
    return (f'<span style="background:{bg};color:{fg};border-radius:8px;'
            f'padding:3px 10px;font-weight:700;font-size:0.9em;">{nome}</span>')

def download_orario(df):
    if not df.empty:
        st.download_button(
            "⬇️ Scarica orario in CSV",
            data=df.to_csv(index=False),
            file_name="orario.csv",
            mime="text/csv"
        )

def genera_html_stampa_sostituzioni(plesso_nome, data_sost, giorno_assente, tabella_df):
    data_leggibile = data_sost.strftime("%d/%m/%Y")
    titolo_giorno = f"{giorno_assente} {data_leggibile}"

    righe_html = ""
    ordine_ore = ORE_LEZIONE
    tabella_ordinata = tabella_df.copy()
    tabella_ordinata["Ora"] = pd.Categorical(tabella_ordinata["Ora"], categories=ordine_ore, ordered=True)
    tabella_ordinata = tabella_ordinata.sort_values(["Ora", "Classe"])

    for _, r in tabella_ordinata.iterrows():
        ora = html_lib.escape(str(r["Ora"]))
        classe = html_lib.escape(str(r["Classe"]))
        assente = html_lib.escape(str(r["Assente"]))
        sost_raw = str(r["Sostituzione"])
        sost_pulito = (sost_raw.replace("[S] [NP] ", "").replace("[C] [NP] ", "")
                                .replace("[C] [USCITA] ", "").replace("[S] ", "")
                                .replace("[C] ", "").replace("🔵 ", "").replace("🟡 ", "")
                                .replace("🟢 ", "").replace("🔴 ", "").strip())
        if sost_pulito in ("Nessuno", "", "—"):
            sost_pulito = "— DA COPRIRE —"
        sostituto = html_lib.escape(sost_pulito)
        riga_scoperta = ' class="scoperta"' if sost_pulito == "— DA COPRIRE —" else ""
        righe_html += (
            f"<tr{riga_scoperta}><td>{ora}</td><td>{classe}</td>"
            f"<td>{assente}</td><td>{sostituto}</td></tr>\n"
        )

    return f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<title>Sostituzioni {html_lib.escape(titolo_giorno)} — {html_lib.escape(plesso_nome)}</title>
<style>
  @page {{ margin: 1.5cm; }}
  body {{ font-family: Georgia, 'Times New Roman', serif; color: #1a1a1a; margin: 0; padding: 0 0.5cm; }}
  h1 {{ font-size: 1.5em; margin: 0 0 0.1em 0; border-bottom: 3px solid #C97D3D; padding-bottom: 0.2em; }}
  .sottotitolo {{ font-size: 1.15em; color: #444; margin: 0 0 0.9em 0; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 1.05em; }}
  th, td {{ border: 1px solid #999; padding: 8px 10px; text-align: left; }}
  th {{ background: #EFE6D3; font-size: 0.95em; text-transform: uppercase; letter-spacing: 0.03em; }}
  tr.scoperta td {{ font-weight: bold; }}
  .piepagina {{ margin-top: 1.2em; font-size: 0.8em; color: #777; }}
  @media print {{ .no-print {{ display: none; }} }}
</style>
</head>
<body>
  <h1>📚 Sostituzioni — {html_lib.escape(plesso_nome)}</h1>
  <p class="sottotitolo">{html_lib.escape(titolo_giorno)}</p>
  <table>
    <thead>
      <tr><th>Ora</th><th>Classe</th><th>Assente</th><th>Sostituzione</th></tr>
    </thead>
    <tbody>
      {righe_html if righe_html else '<tr><td colspan="4">Nessuna sostituzione da mostrare.</td></tr>'}
    </tbody>
  </table>
  <p class="piepagina">Generato automaticamente il {datetime.now().strftime('%d/%m/%Y alle %H:%M')}.</p>
</body>
</html>"""

def pulsante_stampa_sostituzioni(plesso_nome, data_sost, giorno_assente, tabella_df):
    pagina_html = genera_html_stampa_sostituzioni(plesso_nome, data_sost, giorno_assente, tabella_df)
    pagina_json = json.dumps(pagina_html)
    st.components.v1.html(f"""
<button id="stampa-sostituzioni-btn" style="
    width:100%; padding:0.7em; font-size:1em; font-weight:bold;
    background:#3A2E1F; color:white; border:none; border-radius:10px; cursor:pointer;
">🖨️ Stampa / PDF per la bacheca</button>
<script>
document.getElementById('stampa-sostituzioni-btn').addEventListener('click', function() {{
    var finestra = window.open('', '_blank');
    finestra.document.write({pagina_json});
    finestra.document.close();
    finestra.focus();
    setTimeout(function() {{ finestra.print(); }}, 300);
}});
</script>
""", height=55)

def vista_pivot_docenti(df, mode="docenti"):
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
                return "color: #3B6D11; font-weight: bold;"
            elif text.strip() != "":
                return "color: #9C5F2C;"
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
def mostra_intestazione():
    st.markdown(
        f"""
<div style="
    display:flex; align-items:center; justify-content:space-between; gap:14px;
    background:linear-gradient(135deg, #EFE6D3, #FBF4E6);
    border:1.5px solid #E3D9C2; border-radius:16px;
    padding:14px 18px; margin-bottom:0.8rem;
    box-shadow:0 3px 10px rgba(58,46,31,0.12);
">
  <div style="display:flex; align-items:center; gap:14px;">
    <div style="font-size:2.1em; line-height:1;">🏫</div>
    <div>
      <div style="font-size:1.35em; font-weight:800; color:#3A2E1F; line-height:1.15;">
        {PLESSO_NAME}
      </div>
      <div style="font-size:0.95em; color:#9C5F2C; font-weight:600;">
        Gestione sostituzioni docenti
      </div>
    </div>
  </div>
  <div style="display:flex; flex-direction:column; align-items:center; gap:2px; flex-shrink:0;">
    <a href="?ricarica=1" target="_self" title="Ricarica dati da Google Sheets" style="
      font-size:1.5em; line-height:1; text-decoration:none;
      color:#C97D3D;
    ">🔄</a>
    <span style="font-size:0.7em; color:#B0A090; font-weight:600;">v{APP_VERSION}</span>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

mostra_intestazione()

if st.query_params.get("ricarica") == "1":
    carica_orario.clear()
    carica_statistiche.clear()
    st.query_params.clear()
    st.rerun()

try:
    with st.spinner('Caricamento dati...'):
        ensure_sheets_exist()
except Exception as e:
    st.error(f"Impossibile inizializzare i fogli Google: {e}")

with st.spinner('Caricamento orario...'):
    orario_df = carica_orario()

# =========================
# MENU PRINCIPALE
# =========================
ETICHETTE_MENU = {
    "Inserisci/Modifica Orario": "📝 Orario",
    "Gestione Assenze":          "🚨 Assenze",
    "Visualizza Orario":         "📅 Vedi",
    "Statistiche":               "📊 Stats",
}

menu_scelto_label = st.segmented_control(
    "Navigazione",
    list(ETICHETTE_MENU.values()),
    selection_mode="single",
    default="🚨 Assenze",
)
etichette_inverse = {v: k for k, v in ETICHETTE_MENU.items()}
menu = etichette_inverse.get(menu_scelto_label)

# --- INSERIMENTO/MODIFICA ORARIO ---
if menu == "Inserisci/Modifica Orario":
    st.header("➕ Modifica orario")

    uploaded_file = st.file_uploader("Carica un nuovo orario (CSV)", type="csv")
    if uploaded_file:
        df_tmp = pd.read_csv(uploaded_file)
        
        # MAPPING GUIDATO DELLE COLONNE
        if not all(col in df_tmp.columns for col in REQUIRED_COLUMNS):
            st.info("💡 Mappa le colonne del tuo file CSV con le voci richieste dal sistema:")
            col_map = {}
            cols_csv = ["-- Seleziona --"] + list(df_tmp.columns)
            
            for req_col in REQUIRED_COLUMNS:
                default_idx = 0
                for idx, c in enumerate(cols_csv):
                    if req_col.lower() in c.lower():
                        default_idx = idx
                        break
                col_map[req_col] = st.selectbox(f"Colonna per **{req_col}**", cols_csv, index=default_idx, key=f"map_{req_col}")
            
            if st.button("Importa con questo mapping", type="primary"):
                if any(v == "-- Seleziona --" for k, v in col_map.items() if k != "Escludi"):
                    st.error("Associa tutte le colonne obbligatorie prima di procedere.")
                else:
                    df_nuovo = pd.DataFrame()
                    for req_col in REQUIRED_COLUMNS:
                        if col_map[req_col] != "-- Seleziona --":
                            df_nuovo[req_col] = df_tmp[col_map[req_col]]
                        else:
                            df_nuovo[req_col] = False if req_col == "Escludi" else ""
                    
                    for col in ["Docente", "Giorno", "Ora", "Classe", "Tipo"]:
                        df_nuovo[col] = df_nuovo[col].astype(str).str.strip()
                    if "Escludi" in df_nuovo.columns:
                        df_nuovo["Escludi"] = df_nuovo["Escludi"].astype(bool)

                    valori_non_validi = df_nuovo[
                        ~df_nuovo["Giorno"].isin(GIORNI_SETTIMANA) | ~df_nuovo["Ora"].isin(ORE_LEZIONE)
                    ]
                    conflitti_csv = trova_conflitti_orario(df_nuovo)
                    if not valori_non_validi.empty:
                        st.error(
                            f"Il CSV contiene valori di Giorno/Ora non validi (attesi {GIORNI_SETTIMANA} e {ORE_LEZIONE})."
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
        else:
            df_nuovo = df_tmp[REQUIRED_COLUMNS].copy()
            for col in ["Docente", "Giorno", "Ora", "Classe", "Tipo"]:
                df_nuovo[col] = df_nuovo[col].astype(str).str.strip()
            valori_non_validi = df_nuovo[
                ~df_nuovo["Giorno"].isin(GIORNI_SETTIMANA) | ~df_nuovo["Ora"].isin(ORE_LEZIONE)
            ]
            conflitti_csv = trova_conflitti_orario(df_nuovo)
            if not valori_non_validi.empty:
                st.error("Giorno o Ora non validi nel CSV.")
                st.dataframe(valori_non_validi, use_container_width=True, hide_index=True)
            elif conflitti_csv:
                st.error("Conflitti trovati nel CSV:")
                for docente, giorno, ora in conflitti_csv:
                    st.write(f"- {docente}: {giorno} ora {ora}")
            else:
                orario_df = df_nuovo
                if salva_orario(orario_df):
                    st.success("Orario caricato con successo ✅")
                    st.rerun()

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
        if st.button("Aggiungi", key="add_lesson", type="primary"):
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
        if st.button("Salva modifiche", type="primary"):
            df_da_salvare = edited_df.copy()
            for col in ["Docente", "Giorno", "Ora", "Classe", "Tipo"]:
                df_da_salvare[col] = df_da_salvare[col].astype(str).str.strip()
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

# --- GESTIONE ASSENZE ---
elif menu == "Gestione Assenze":
    st.header("🚨 Assenze")

    if orario_df.empty:
        st.warning("Non hai ancora caricato nessun orario.")
    else:
        data_sostituzione = st.date_input("Data della sostituzione")

        giorno_assente = data_sostituzione.strftime("%A")
        traduzione_giorni = {
            "Monday": "Lunedì", "Tuesday": "Martedì", "Wednesday": "Mercoledì",
            "Thursday": "Giovedì", "Friday": "Venerdì", "Saturday": "Sabato", "Sunday": "Domenica"
        }
        giorno_assente = traduzione_giorni.get(giorno_assente, giorno_assente)

        if giorno_assente not in GIORNI_SETTIMANA:
            st.warning(f"Hai selezionato {giorno_assente}, un giorno non presente nell'orario scolastico (Lun-Ven).")

        docenti_assenti = st.multiselect("Seleziona docenti assenti", sorted(orario_df["Docente"].unique()))

        classi_uscita_per_ora = {}
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
            ore_assenti = orario_df[
                (orario_df["Docente"].isin(docenti_assenti)) &
                (orario_df["Giorno"] == giorno_assente)
            ].copy()

            if ore_assenti.empty:
                st.info("I docenti selezionati non hanno lezioni in quel giorno.")
            else:
                st.subheader("📌 Ore scoperte")

                ore_assenti_display = ore_assenti.copy()
                sostegni_presenti = []

                for _, r in ore_assenti.iterrows():
                    ora = r["Ora"]
                    classe = r["Classe"]

                    sost_df = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == ora) &
                        (orario_df["Classe"] == classe) &
                        (orario_df["Tipo"].str.lower() == "sostegno")
                    ]

                    if sost_df.empty:
                        sostegni_presenti.append("—")
                    else:
                        lista = ", ".join(sorted(sost_df["Docente"].unique()))
                        sostegni_presenti.append(lista)

                ore_assenti_display["Sostegni in servizio"] = sostegni_presenti
                ore_assenti_display = ore_assenti_display[["Docente", "Ora", "Classe", "Tipo", "Sostegni in servizio"]]
                st.dataframe(ore_assenti_display, use_container_width=True, hide_index=True)

                st.subheader("🔄 Possibili sostituti")
                sostituzioni = []

                tutti_docenti = sorted(orario_df["Docente"].unique())
                escludi_docenti = set(orario_df[orario_df["Escludi"]]["Docente"].unique())
                docenti_assenti_set = set(docenti_assenti)

                docente_tipo_map = build_docente_tipo_map(orario_df)
                def tipo_docente(d):
                    return docente_tipo_map.get(d, "")

                ore_assenti["Ora"] = pd.Categorical(
                    ore_assenti["Ora"], categories=ORE_LEZIONE, ordered=True
                )
                ore_assenti = ore_assenti.sort_values(["Ora", "Docente"]).reset_index(drop=True)

                ora_corrente = None

                for _, row in ore_assenti.iterrows():
                    ora = row["Ora"]
                    classe = row["Classe"]
                    assente = row["Docente"]

                    if ora != ora_corrente:
                        if ora_corrente is not None:
                            st.markdown("<hr style='border:none;border-top:2px solid #E3D9C2;margin:16px 0 12px;'>", unsafe_allow_html=True)
                        st.markdown(
                            f'<div style="background:#EFE6D3;border-radius:10px;padding:8px 14px;'
                            f'font-weight:800;font-size:1.05em;color:#3A2E1F;margin-bottom:10px;">'
                            f'🕐 {ora} ora</div>',
                            unsafe_allow_html=True
                        )
                        ora_corrente = ora

                    presenti_ora_df = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == ora) &
                        (~orario_df["Escludi"]) &
                        (~orario_df["Docente"].isin(docenti_assenti_set))
                    ].copy()

                    presenti_ora = list(dict.fromkeys(presenti_ora_df["Docente"].tolist()))

                    added = set()
                    options = ["Nessuno"]

                    # 1. VISUAL BADGE INTEGRATED INTO SELECTBOX OPTIONS
                    same_class_sost = presenti_ora_df[
                        (presenti_ora_df["Tipo"].str.lower() == "sostegno") &
                        (presenti_ora_df["Classe"] == classe) &
                        (presenti_ora_df["Docente"] != assente)
                    ]["Docente"].unique().tolist()
                    for d in sorted(same_class_sost):
                        label = f"🔵 [S] {d}"
                        if d not in added:
                            options.append(label); added.add(d)

                    other_sost = presenti_ora_df[
                        (presenti_ora_df["Tipo"].str.lower() == "sostegno") &
                        (presenti_ora_df["Docente"] != assente)
                    ]["Docente"].unique().tolist()
                    for d in sorted(other_sost):
                        if d in added: 
                            continue
                        label = f"🔵 [S] {d}"
                        options.append(label); added.add(d)

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

                    for d in sorted(curricolari_liberi_uscita):
                        if d in added:
                            continue
                        label = f"🟡 [C] [USCITA] {d}"
                        options.append(label); added.add(d)

                    for d in sorted(curricolari_occupati):
                        if d in added:
                            continue
                        label = f"🔴 [C] {d}"
                        options.append(label); added.add(d)

                    presenti_ora_set = set(presenti_ora)
                    np_candidates = [
                        d for d in tutti_docenti
                        if d not in presenti_ora_set
                        and d not in docenti_assenti_set
                        and d not in escludi_docenti
                    ]

                    np_sost = []
                    np_curr = []
                    for d in np_candidates:
                        t = tipo_docente(d).lower()
                        if t == "sostegno":
                            np_sost.append(d)
                        else:
                            np_curr.append(d)

                    for d in sorted(np_sost):
                        if d in added:
                            continue
                        label = f"🟢 [S] [NP] {d}"
                        options.append(label); added.add(d)

                    for d in sorted(np_curr):
                        if d in added:
                            continue
                        label = f"🟢 [C] [NP] {d}"
                        options.append(label); added.add(d)

                    proposto_display = "Nessuno"
                    if len(same_class_sost) > 0:
                        proposto_display = f"🔵 [S] {sorted(same_class_sost)[0]}"
                    elif len(other_sost) > 0:
                        proposto_display = f"🔵 [S] {sorted(other_sost)[0]}"
                    elif len(curricolari_liberi_uscita) > 0:
                        proposto_display = f"🟡 [C] [USCITA] {sorted(curricolari_liberi_uscita)[0]}"
                    elif len(curricolari_occupati) > 0:
                        proposto_display = f"🔴 [C] {sorted(curricolari_occupati)[0]}"
                    elif len(np_sost) > 0:
                        proposto_display = f"🟢 [S] [NP] {sorted(np_sost)[0]}"
                    elif len(np_curr) > 0:
                        proposto_display = f"🟢 [C] [NP] {sorted(np_curr)[0]}"

                    default_index = options.index(proposto_display) if proposto_display in options else 0

                    col_sx, col_dx = st.columns([3, 1])
                    with col_sx:
                        st.markdown(
                            f"Classe **{classe}** · Assente: *{assente}*",
                        )
                    bg, fg, ico = _colore_tipo(proposto_display)
                    with col_dx:
                        st.markdown(
                            f'<div style="background:{bg};color:{fg};border-radius:8px;'
                            f'padding:4px 8px;font-size:0.78em;font-weight:700;text-align:center;">'
                            f'{ico} proposto</div>',
                            unsafe_allow_html=True
                        )

                    scelta = st.selectbox(
                        f"Sostituto",
                        options,
                        index=default_index,
                        key=f"sost_{assente}_{ora}_{classe}",
                        label_visibility="collapsed",
                    )

                    bg2, fg2, ico2 = _colore_tipo(scelta)
                    tipo_label = (
                        "Sostegno" if "[S]" in scelta and "[NP]" not in scelta
                        else "Uscita" if "[USCITA]" in scelta
                        else "Non in orario (Libero)" if "[NP]" in scelta
                        else "Curricolare (Occupato)" if "[C]" in scelta
                        else "—"
                    )
                    if scelta != "Nessuno":
                        st.markdown(
                            f'<div style="background:{bg2};color:{fg2};border-radius:10px;'
                            f'padding:6px 12px;font-size:0.85em;font-weight:600;'
                            f'margin-bottom:8px;display:inline-block;">'
                            f'{ico2} {tipo_label}</div>',
                            unsafe_allow_html=True
                        )
                    st.markdown("---")

                    if scelta == "Nessuno":
                        nome_pulito = "Nessuno"
                    else:
                        nome_pulito = (
                            scelta.replace("[S] [NP] ", "")
                                  .replace("[C] [NP] ", "")
                                  .replace("[C] [USCITA] ", "")
                                  .replace("[S] ", "")
                                  .replace("[C] ", "")
                                  .replace("🔵 ", "").replace("🟡 ", "")
                                  .replace("🟢 ", "").replace("🔴 ", "")
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

                ordine_ore = ORE_LEZIONE
                if not sostituzioni_df.empty:
                    sostituzioni_df["Ora"] = pd.Categorical(sostituzioni_df["Ora"], categories=ordine_ore, ordered=True)
                    sostituzioni_df = sostituzioni_df.sort_values("Ora").reset_index(drop=True)

                tabella_df = sostituzioni_df[["Ora", "Classe", "Assente", "Sostituto_display"]].copy()
                tabella_df = tabella_df.rename(columns={"Sostituto_display": "Sostituzione"})
                tabella_df["Ora"] = pd.Categorical(tabella_df["Ora"], categories=ordine_ore, ordered=True)
                tabella_df = tabella_df.sort_values(["Ora", "Classe"]).reset_index(drop=True)
                st.subheader("📋 Riepilogo sostituzioni")

                cards_html = ""
                for ora_c, grp in tabella_df.groupby("Ora", sort=False):
                    righe_html = ""
                    for _, r in grp.iterrows():
                        badge = _badge_sostituto(r["Sostituzione"])
                        righe_html += (
                            f'<div style="display:flex;justify-content:space-between;'
                            f'align-items:center;padding:8px 0;border-bottom:1px solid #EFE6D3;">'
                            f'<div><span style="font-weight:700;color:#3A2E1F;">Cl. {r["Classe"]}</span>'
                            f'<span style="color:#9C5F2C;font-size:0.85em;margin-left:6px;">ass. {r["Assente"]}</span></div>'
                            f'<div>{badge}</div></div>'
                        )
                    cards_html += (
                        f'<div style="background:#FBF4E6;border:1.5px solid #E3D9C2;border-radius:14px;'
                        f'padding:12px 14px;margin-bottom:10px;">'
                        f'<div style="font-size:1em;font-weight:800;color:#C97D3D;margin-bottom:4px;">'
                        f'🕐 {ora_c} ora</div>{righe_html}</div>'
                    )
                st.markdown(cards_html, unsafe_allow_html=True)

                st.subheader("📝 Sostituzioni in formato testo (mobile/copincolla)")
                testo_output = "Buongiorno, supplenze:\n\n"

                for ora, gruppo in sostituzioni_df.groupby("Ora"):
                    if not gruppo.empty:
                        testo_output += f"🕐 *{ora} ORA*\n"
                        for _, r in gruppo.iterrows():
                            sost_pulito = r['Sostituto'] if r['Sostituto'] not in ["Nessuno", "", "—"] else "—"
                            testo_output += f"Classe {r['Classe']}\n"
                            testo_output += f"👩‍🏫 Assente: {r['Assente']}\n"
                            testo_output += f"✅ Sostituzione: {sost_pulito}\n\n"

                testo_strip = testo_output.strip()
                st.text_area("Testo pronto da copiare", value=testo_strip, height=260)
                
                # 2. BOTTONI RAPIDI WHATSAPP / TELEGRAM E COPIA
                encoded_text = urllib.parse.quote(testo_strip)
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    st.markdown(
                        f'<a href="https://api.whatsapp.com/send?text={encoded_text}" target="_blank" class="btn-whatsapp">📲 Invia su WhatsApp</a>',
                        unsafe_allow_html=True
                    )
                with col_btn2:
                    st.markdown(
                        f'<a href="https://t.me/share/url?url=&text={encoded_text}" target="_blank" class="btn-telegram">✈️ Invia su Telegram</a>',
                        unsafe_allow_html=True
                    )
                
                testo_json = json.dumps(testo_strip)
                st.components.v1.html(f"""
<button id="copia-sostituzioni-btn" style="
    width:100%; padding:0.7em; font-size:1em; font-weight:bold;
    background:#C97D3D; color:white; border:none; border-radius:10px; cursor:pointer; margin-top:10px;
">📋 Copia negli appunti</button>
<script>
document.getElementById('copia-sostituzioni-btn').addEventListener('click', function() {{
    navigator.clipboard.writeText({testo_json}).then(() => {{
        this.innerText = '✅ Copiato!';
        setTimeout(() => {{ this.innerText = '📋 Copia negli appunti'; }}, 2000);
    }});
}});
</script>
""", height=65)

                st.subheader("🖨️ Riepilogo per la bacheca")
                st.caption(
                    "Apre una finestra pronta per la stampa. Dal dialogo di stampa del "
                    "browser puoi scegliere una stampante fisica oppure \"Salva come PDF\"."
                )
                pulsante_stampa_sostituzioni(PLESSO_NAME, data_sostituzione, giorno_assente, tabella_df)

                if st.button("✅ Conferma tabella (non salva ancora)", type="primary"):
                    check_df = sostituzioni_df.copy()
                    conflitti = []
                    conflitti_orario = []

                    for ora_val in check_df["Ora"].unique():
                        assegnazioni = check_df[check_df["Ora"] == ora_val]
                        sostituti = [s for s in assegnazioni["Sostituto"] if s not in ["Nessuno", "", "—"]]

                        duplicati = [s for s in sostituti if sostituti.count(s) > 1]
                        if duplicati:
                            conflitti.append((ora_val, list(set(duplicati))))

                        for s in sostituti:
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

                    st.session_state["sostituzioni_confermate"] = sostituzioni_df.copy()
                    st.session_state["ore_assenti_confermate"] = ore_assenti.copy()
                    st.session_state["data_sostituzione_tmp"] = data_sostituzione
                    st.session_state["giorno_assente_tmp"] = giorno_assente

                    st.success("Tabella confermata ✅ Ora puoi salvarla nello storico.")

                if st.session_state.get("sostituzioni_confermate") is not None:
                    if st.button("💾 Salva nello storico", key="save_storico_main", type="primary"):
                        sost_df = st.session_state.get("sostituzioni_confermate")
                        ore_assenti_session = st.session_state.get("ore_assenti_confermate")
                        data_tmp = st.session_state.get("data_sostituzione_tmp")
                        giorno_tmp = st.session_state.get("giorno_assente_tmp")

                        if sost_df is not None and ore_assenti_session is not None:
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
        # 3. FILTRO PER CLASSE E DOCENTE
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            docenti_selezionati = st.multiselect(
                "🔍 Filtra per Docente",
                sorted(orario_df["Docente"].unique())
            )
        with col_dx if 'col_dx' in locals() else col_f2:
            classi_selezionate = st.multiselect(
                "🏫 Filtra per Classe",
                sorted(orario_df["Classe"].unique())
            )

        df_filtrato = orario_df.copy()

        if docenti_selezionati:
            df_base = df_filtrato[df_filtrato["Docente"].isin(docenti_selezionati)]
            chiavi = df_base[["Classe", "Giorno", "Ora"]].drop_duplicates()
            df_compresenze = df_filtrato[
                df_filtrato["Tipo"].str.lower() == "sostegno"
            ].merge(chiavi, on=["Classe", "Giorno", "Ora"], how="inner")
            df_filtrato = pd.concat([df_base, df_compresenze]).drop_duplicates()

        if classi_selezionate:
            df_filtrato = df_filtrato[df_filtrato["Classe"].isin(classi_selezionate)]

        if df_filtrato.empty:
            st.warning("Nessun risultato corrisponde ai filtri selezionati.")
        else:
            vista_pivot_docenti(df_filtrato, mode="classi")

        download_orario(orario_df)

# --- STATISTICHE ---
elif menu == "Statistiche":
    st.header("📊 Statistiche")
    with st.spinner('Caricamento statistiche...'):
        df_storico, df_assenze = carica_statistiche()

    date_storico = pd.to_datetime(df_storico["data"], errors="coerce") if "data" in df_storico.columns else pd.Series([], dtype="datetime64[ns]")
    date_assenze = pd.to_datetime(df_assenze["data"], errors="coerce") if "data" in df_assenze.columns else pd.Series([], dtype="datetime64[ns]")
    tutte_le_date = pd.concat([date_storico, date_assenze]).dropna()

    if tutte_le_date.empty:
        data_min_default = data_max_default = datetime.now().date()
    else:
        data_min_default = tutte_le_date.min().date()
        data_max_default = tutte_le_date.max().date()

    intervallo = st.date_input(
        "📅 Filtra per intervallo di date",
        value=(data_min_default, data_max_default),
        min_value=data_min_default,
        max_value=data_max_default,
        key="statistiche_intervallo_date",
        help="Filtra le statistiche qui sotto per periodo."
    )
    if isinstance(intervallo, tuple) and len(intervallo) == 2:
        data_inizio, data_fine = intervallo
    elif isinstance(intervallo, tuple) and len(intervallo) == 1:
        data_inizio = data_fine = intervallo[0]
    else:
        data_inizio = data_fine = intervallo

    if not df_storico.empty:
        df_storico = df_storico[
            pd.to_datetime(df_storico["data"], errors="coerce").between(
                pd.Timestamp(data_inizio), pd.Timestamp(data_fine)
            )
        ].copy()
    if not df_assenze.empty:
        df_assenze = df_assenze[
            pd.to_datetime(df_assenze["data"], errors="coerce").between(
                pd.Timestamp(data_inizio), pd.Timestamp(data_fine)
            )
        ].copy()

    def render_cards(titolo, righe, colore):
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
        if tutte_le_date.empty:
            st.info("Nessuna statistica disponibile. Registra prima delle sostituzioni.")
        else:
            st.info("Nessuna sostituzione registrata nell'intervallo di date selezionato.")
    else:
        df_sum = df_storico.groupby("docente")["ore"].sum().reset_index()
        df_sum = df_sum.rename(columns={"ore": "Totale Ore Sostituite"})
        df_sorted = df_sum.sort_values("Totale Ore Sostituite", ascending=False).reset_index(drop=True)

        top3 = [(r["docente"], int(r["Totale Ore Sostituite"])) for _, r in df_sorted.head(3).iterrows()]

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
            render_cards("🟢 Più sostituzioni", top3, "#6B8F71")
        with col_bot:
            render_cards("🟠 Meno sostituzioni [S]", bot3, "#C9933D")

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
        if tutte_le_date.empty:
            st.info("Nessuna assenza registrata.")
        else:
            st.info("Nessuna assenza registrata nell'intervallo di date selezionato.")
    else:
        df_ore = df_assenze.groupby("docente")["ora"].count().reset_index().rename(columns={"ora": "Totale Ore Assenti"})
        df_giorni = df_assenze.groupby("docente")["data"].nunique().reset_index().rename(columns={"data": "Giorni Assenti"})
        df_assenze_agg = df_ore.merge(df_giorni, on="docente")
        df_assenze_agg = df_assenze_agg.sort_values("Totale Ore Assenti", ascending=False).reset_index(drop=True)

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
<div style="background:#C9933D;border-radius:14px;padding:16px 20px;color:white;margin-bottom:16px;">
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

    st.subheader("💾 Esportazioni & Cloud Backup")
    col_exp1, col_exp2 = st.columns(2)
    
    with col_exp1:
        st.info("Scarica un unico file Excel (.xlsx) con tutti i dati formattati su fogli distinti.")
        excel_file = create_excel_export()
        if excel_file:
            st.download_button(
                label="📊 Scarica Report Excel (.xlsx)",
                data=excel_file,
                file_name=f"{SPREADSHEET_NAME}_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    with col_exp2:
        st.info("Scarica un backup compresso dei dati grezzi dei fogli Orario, Storico e Assenze.")
        backup_file = create_backup()
        if backup_file:
            st.download_button(
                label="⬇️ Scarica Backup (ZIP)",
                data=backup_file,
                file_name=f"{SPREADSHEET_NAME}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )

    st.header("📦 Archivia anno scolastico")
    st.info(
        "Copia storico e assenze dell'anno corrente in fogli separati nel documento Google "
        "(es. archivio_storico_2024-25), poi svuota i fogli attivi per il nuovo anno. "
        "L'orario non viene toccato."
    )
    anno_input = st.text_input(
        "Anno scolastico da archiviare",
        placeholder="es. 2024-25",
        max_chars=10
    )
    conferma_archivio = st.checkbox(
        "Confermo: voglio archiviare l'anno e azzerare storico e assenze attivi",
        key="conf_archivio"
    )
    if st.button("📦 Archivia anno scolastico", key="btn_archivio"):
        if not anno_input.strip():
            st.warning("Inserisci l'anno scolastico (es. 2024-25).")
        elif not conferma_archivio:
            st.warning("Spunta la casella di conferma prima di procedere.")
        else:
            with st.spinner("Archiviazione in corso..."):
                if archivia_anno_scolastico(anno_input.strip()):
                    st.success(
                        f"Anno {anno_input.strip()} archiviato ✅ "
                        f"I fogli archivio_storico_{anno_input.strip().replace('/', '-')} e "
                        f"archivio_assenze_{anno_input.strip().replace('/', '-')} "
                        f"sono ora disponibili nel documento Google."
                    )
