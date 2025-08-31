import streamlit as st
import pandas as pd
import os
import re
import gspread
import gspread_dataframe as gd
from oauth2client.service_account import ServiceAccountCredentials

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
# INIZIALIZZAZIONE FOGLI (se mancanti creali con header corretti)
# =========================
def ensure_sheets_exist():
    client = get_gdrive_client()
    try:
        sh = client.open(SPREADSHEET_NAME)
    except Exception as e:
        sh = client.create(SPREADSHEET_NAME)
    
    existing = {ws.title for ws in sh.worksheets()}
    
    if ORARIO_SHEET not in existing:
        ws = sh.add_worksheet(title=ORARIO_SHEET, rows="100", cols="20")
        header_df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        gd.set_with_dataframe(ws, header_df, include_index=False, include_column_header=True)

    if STORICO_SHEET not in existing:
        ws = sh.add_worksheet(title=STORICO_SHEET, rows="100", cols="10")
        header = pd.DataFrame(columns=["data", "giorno", "docente", "ore", "note"])
        gd.set_with_dataframe(ws, header, include_index=False, include_column_header=True)

    if ASSENZE_SHEET not in existing:
        ws = sh.add_worksheet(title=ASSENZE_SHEET, rows="100", cols="10")
        header = pd.DataFrame(columns=["data", "giorno", "docente", "ora", "classe"])
        gd.set_with_dataframe(ws, header, include_index=False, include_column_header=True)


# =========================
# CARICAMENTO / SALVATAGGIO ORARIO
# =========================
def carica_orario():
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(ORARIO_SHEET)
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
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(ORARIO_SHEET)
        df_to_save = df.copy()
        for col in REQUIRED_COLUMNS:
            if col not in df_to_save.columns:
                df_to_save[col] = ""
        df_to_save = df_to_save[REQUIRED_COLUMNS]
        gd.set_with_dataframe(ws, df_to_save, include_index=False, include_column_header=True)
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio dell'orario su Google Sheets: {e}")
        return False

# =========================
# CARICAMENTO / SALVATAGGIO STATISTICHE (storico + assenze)
# =========================
def carica_statistiche():
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws_storico = sh.worksheet(STORICO_SHEET)
        ws_assenze = sh.worksheet(ASSENZE_SHEET)
        df_storico = gd.get_as_dataframe(ws_storico, header=0).dropna(how='all')
        df_assenze = gd.get_as_dataframe(ws_assenze, header=0).dropna(how='all')
        if not df_storico.empty:
            if "data" in df_storico.columns:
                df_storico["data"] = pd.to_datetime(df_storico["data"], errors="coerce")
                df_storico["data"] = df_storico["data"].dt.strftime("%Y-%m-%d")

            if "ore" in df_storico.columns:
                df_storico["ore"] = pd.to_numeric(df_storico["ore"], errors="coerce").fillna(0).astype(int)

            for c in ["docente", "giorno", "note"]:
                if c in df_storico.columns:
                    df_storico[c] = df_storico[c].astype(str).str.strip().str.lower()
        if not df_assenze.empty:
            if "data" in df_assenze.columns:
                df_assenze["data"] = pd.to_datetime(df_assenze["data"], errors="coerce")
                df_assenze["data"] = df_assenze["data"].dt.strftime("%Y-%m-%d")
            for c in ["docente", "giorno", "ora", "classe"]:
                if c in df_assenze.columns:
                    df_assenze[c] = df_assenze[c].astype(str).str.strip()
        if df_storico.empty:
            df_storico = pd.DataFrame(columns=["data", "giorno", "docente", "ore", "note"])
        if df_assenze.empty:
            df_assenze = pd.DataFrame(columns=["data", "giorno", "docente", "ora", "classe"])
        return df_storico, df_assenze
    except Exception as e:
        st.error(f"Errore nel caricamento delle statistiche da Google Sheets: {e}")
        return pd.DataFrame(columns=["data", "giorno", "docente", "ore", "note"]), pd.DataFrame(columns=["data", "giorno", "docente", "ora", "classe"])

def salva_storico_assenze(data_sostituzione, giorno_assente, sostituzioni_df, ore_assenti):
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws_storico = sh.worksheet(STORICO_SHEET)
        ws_assenze = sh.worksheet(ASSENZE_SHEET)

        storico_data = []
        for _, row in sostituzioni_df.iterrows():
            sost = row.get("Sostituto", "")
            note = row.get("Note", "")
            if sost and sost != "Nessuno":
                storico_data.append([str(data_sostituzione), giorno_assente, sost, 1, note])

        if storico_data:
            ws_storico.append_rows(storico_data, value_input_option="USER_ENTERED")

        assenze_data = []
        for _, row in ore_assenti.iterrows():
            assenze_data.append([str(data_sostituzione), giorno_assente, row["Docente"], row["Ora"], row["Classe"]])
        if assenze_data:
            ws_assenze.append_rows(assenze_data, value_input_option="USER_ENTERED")

        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio dei dati su Google Sheets: {e}")
        return False

def clear_sheet_content(sheet_name):
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(sheet_name)
        ws.clear()
        if sheet_name == STORICO_SHEET:
            ws.append_row(["data", "giorno", "docente", "ore", "note"])
        elif sheet_name == ASSENZE_SHEET:
            ws.append_row(["data", "giorno", "docente", "ora", "classe"])
        return True
    except Exception as e:
        st.error(f"Errore nell'azzeramento del foglio {sheet_name}: {e}")
        return False

# =========================
# UTILITA' PER DOWNLOAD e PIVOT
# =========================
def download_orario(df):
    if not df.empty:
        st.download_button(
            "⬇️ Scarica orario in CSV",
            data=df.to_csv(index=False),
            file_name="orario.csv",
            mime="text/csv"
        )

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

        styled = pivot.style.set_properties(**{"text-align": "center"})
        st.dataframe(styled, use_container_width=True, hide_index=True)


# =========================
# AVVIO APP
# =========================
st.title("📚 Gestione orari e Sostituzioni Docenti (Google Sheets)")
try:
    ensure_sheets_exist()
except Exception as e:
    st.error(f"Impossibile inizializzare i fogli Google: {e}")

if "orario_df" not in st.session_state:
    st.session_state.orario_df = carica_orario()

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
            st.session_state.orario_df = df_tmp[REQUIRED_COLUMNS].copy()
            if salva_orario(st.session_state.orario_df):
                st.success("Orario caricato con successo ✅")
        else:
            st.error(f"CSV non valido. Deve contenere: {REQUIRED_COLUMNS}")

    with st.expander("Aggiungi una nuova lezione"):
        docente = st.selectbox(
            "Nome docente",
            ["➕ Nuovo docente"] + (sorted(st.session_state.orario_df["Docente"].unique()) if not st.session_state.orario_df.empty else []),
            key="docente_input"
        )
        if docente == "➕ Nuovo docente":
            docente = st.text_input("Inserisci nuovo docente")
        giorno = st.selectbox("Giorno", ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"])
        ora = st.selectbox("Ora", ["I", "II", "III", "IV", "V", "VI"])
        classe = st.selectbox(
            "Classe",
            ["➕ Nuova classe"] + (sorted(st.session_state.orario_df["Classe"].unique()) if not st.session_state.orario_df.empty else []),
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
                conflitto = not st.session_state.orario_df[
                    (st.session_state.orario_df["Docente"] == docente) &
                    (st.session_state.orario_df["Giorno"] == giorno) &
                    (st.session_state.orario_df["Ora"] == ora)
                ].empty
                if conflitto:
                    st.error("Conflitto: questo docente è già assegnato a un'altra classe in questa ora.")
                else:
                    st.session_state.orario_df = pd.concat([st.session_state.orario_df, nuovo], ignore_index=True)
                    if salva_orario(st.session_state.orario_df):
                        st.success("Lezione aggiunta all'orario e salvata su Google Sheets ✅")
            else:
                st.error("Compila tutti i campi per aggiungere una lezione.")

    st.subheader("📝 Modifica orario attuale")
    if not st.session_state.orario_df.empty:
        col_order = REQUIRED_COLUMNS
        df_edit = st.session_state.orario_df[col_order].copy()
        df_edit = df_edit.sort_values(by=["Docente"])
        edited_df = st.data_editor(df_edit, use_container_width=True, hide_index=True)
        if st.button("Salva modifiche"):
            st.session_state.orario_df = edited_df
            if salva_orario(st.session_state.orario_df):
                st.success("Orario modificato e salvato su Google Sheets ✅")
    download_orario(st.session_state.orario_df)

# --- GESTIONE ASSENZE ---
elif menu == "Gestione Assenze":
    st.header("🚨 Gestione Assenze")
    orario_df = st.session_state.orario_df

    if orario_df.empty:
        st.warning("Non hai ancora caricato nessun orario.")
    else:
        docenti_assenti = st.multiselect("Seleziona docenti assenti", sorted(orario_df["Docente"].unique()))
        data_sostituzione = st.date_input("Data della sostituzione")

        giorno_assente = data_sostituzione.strftime("%A")
        traduzione_giorni = {
            "Monday": "Lunedì",
            "Tuesday": "Martedì",
            "Wednesday": "Mercoledì",
            "Thursday": "Giovedì",
            "Friday": "Venerdì",
            "Saturday": "Sabato",
            "Sunday": "Domenica"
        }
        giorno_assente = traduzione_giorni.get(giorno_assente, giorno_assente)

        if not docenti_assenti:
            st.info("Seleziona almeno un docente per continuare.")
        else:
            ore_assenti = orario_df[
                (orario_df["Docente"].isin(docenti_assenti)) &
                (orario_df["Giorno"] == giorno_assente)
            ]

            if ore_assenti.empty:
                st.info("I docenti selezionati non hanno lezioni in quel giorno.")
            else:
                st.subheader("📌 Ore scoperte")
                
                # Aggiungi campo Note
                ore_assenti = ore_assenti.assign(Note="")
                edited_ore_assenti = st.data_editor(
                    ore_assenti[["Docente", "Ora", "Classe", "Tipo", "Note"]],
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Note": st.column_config.TextColumn("Note", help="Aggiungi una nota alla supplenza")
                    }
                )

                st.subheader("🔄 Possibili sostituti")
                sostituzioni = []
                df_storico, df_assenze = carica_statistiche()
                
                # Calcola il carico di sostituzioni per la prioritizzazione
                carico_sostituzioni = df_storico.groupby("docente")["ore"].sum().to_dict()

                for _, row in edited_ore_assenti.iterrows():
                    presenti_ora = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == row["Ora"])
                    ]

                    docenti_disponibili = presenti_ora[
                        (~presenti_ora["Docente"].isin(docenti_assenti)) &
                        (~presenti_ora["Escludi"])
                    ]
                    
                    docenti_esclusi = presenti_ora[
                        (~presenti_ora["Docente"].isin(docenti_assenti)) &
                        (presenti_ora["Escludi"])
                    ]

                    proposto = "Nessuno"
                    if not docenti_disponibili.empty:
                        prioritari_sostegno = docenti_disponibili[
                            (docenti_disponibili["Tipo"] == "Sostegno") &
                            (docenti_disponibili["Classe"] == row["Classe"])
                        ]

                        if not prioritari_sostegno.empty:
                            proposto = prioritari_sostegno["Docente"].iloc[0]
                        else:
                            non_sostegno = docenti_disponibili[docenti_disponibili["Tipo"] != "Sostegno"]
                            
                            if not non_sostegno.empty:
                                non_sostegno['ore_sostituite'] = non_sostegno['Docente'].map(carico_sostituzioni).fillna(0)
                                non_sostegno = non_sostegno.sort_values(by='ore_sostituite')
                                proposto = non_sostegno["Docente"].iloc[0]
                            else:
                                if not docenti_esclusi.empty:
                                    proposto = docenti_esclusi["Docente"].iloc[0]
                    else:
                        if not docenti_esclusi.empty:
                            proposto = docenti_esclusi["Docente"].iloc[0]
                            
                    sostituzioni.append({
                        "Ora": row["Ora"],
                        "Classe": row["Classe"],
                        "Assente": row["Docente"],
                        "Note": row["Note"],
                        "Sostituto": proposto
                    })

                sostituzioni_df = pd.DataFrame(sostituzioni)

                ordine_ore = ["I", "II", "III", "IV", "V", "VI"]
                if not sostituzioni_df.empty:
                    sostituzioni_df["Ora"] = pd.Categorical(sostituzioni_df["Ora"], categories=ordine_ore, ordered=True)
                    sostituzioni_df = sostituzioni_df.sort_values("Ora").reset_index(drop=True)

                colonne_mostrate = ["Ora", "Classe", "Assente", "Sostituto", "Note"]
                edited_df = sostituzioni_df.copy()
                
                # Preparazione per il data editor
                edited_df["Sostituto_Select"] = ""
                
                # Mappa per i docenti occupati
                docenti_occupati_ora = {}
                for ora_val in orario_df['Ora'].unique():
                    docenti_occupati_ora[ora_val] = orario_df[
                        (orario_df['Giorno'] == giorno_assente) & (orario_df['Ora'] == ora_val)
                    ]['Docente'].tolist()

                for idx, riga in sostituzioni_df.iterrows():
                    presenti_ora = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == riga["Ora"])
                    ]["Docente"].unique()

                    sostegni = orario_df[orario_df["Tipo"] == "Sostegno"]["Docente"].unique()

                    opzioni_validi = []
                    for d in sorted(presenti_ora):
                        if d not in docenti_assenti and not orario_df.loc[orario_df["Docente"] == d, "Escludi"].any():
                            if d in sostegni:
                                opzioni_validi.append(f"[S] {d}")
                            else:
                                opzioni_validi.append(d)
                    
                    docenti_esclusi_per_ora = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == riga["Ora"]) &
                        (orario_df["Escludi"])
                    ]["Docente"].unique()
                    
                    opzioni_esclusi = [d for d in docenti_esclusi_per_ora if d not in docenti_assenti]
                    
                    opzioni = ["Nessuno"] + sorted(
                        opzioni_validi,
                        key=lambda x: (0 if x.startswith("[S]") else 1, x)
                    ) + sorted(opzioni_esclusi)

                    proposto = riga["Sostituto"]
                    proposto_display = proposto
                    if proposto in sostegni:
                        proposto_display = f"[S] {proposto}"
                    
                    edited_df.at[idx, "Sostituto_Select"] = opzioni.index(proposto_display) if proposto_display in opzioni else 0

                st.subheader("Modifica le proposte (suggerite in base al carico)")
                
                def color_cells(row):
                    # Controllo conflitti in tempo reale per lo stile
                    sostituto = row["Sostituto"]
                    ora = row["Ora"]
                    
                    # Conflitto 1: docente occupato
                    is_occupied = sostituto != "Nessuno" and sostituto in docenti_occupati_ora.get(ora, [])
                    
                    # Conflitto 2: stesso docente in più righe della stessa ora
                    conflitti_ora = edited_df[(edited_df['Ora'] == ora) & (edited_df['Sostituto'] == sostituto)]
                    is_duplicate = len(conflitti_ora) > 1 and sostituto != "Nessuno"
                    
                    if is_occupied or is_duplicate:
                        return pd.Series(
                            "background-color: #ffcccc",
                            index=row.index
                        )
                    
                    return pd.Series("", index=row.index)

                col_config = {
                    "Ora": st.column_config.TextColumn("Ora", disabled=True),
                    "Classe": st.column_config.TextColumn("Classe", disabled=True),
                    "Assente": st.column_config.TextColumn("Assente", disabled=True),
                    "Sostituto": st.column_config.SelectboxColumn(
                        "Sostituto",
                        options=opzioni, # Qui si può ottimizzare per riga
                        default="Nessuno"
                    ),
                    "Note": st.column_config.TextColumn("Note", help="Aggiungi una nota alla supplenza")
                }
                
                # Non usare la vecchia colonna `Sostituto`, ma `Sostituto_Select` per inizializzare il selectbox
                edited_df_display = edited_df.drop(columns="Sostituto")
                edited_df_display = edited_df_display.rename(columns={"Sostituto_Select": "Sostituto"})
                
                def highlight_conflict(cell):
                    if cell.name == 'Sostituto':
                        sostituto_scelto = cell.values[0]
                        ora = cell.iloc[0, cell.columns.get_loc('Ora')]
                        
                        is_occupied = sostituto_scelto != "Nessuno" and sostituto_scelto in docenti_occupati_ora.get(ora, [])
                        
                        conflitti_ora = edited_df_display[(edited_df_display['Ora'] == ora) & (edited_df_display['Sostituto'] == sostituto_scelto)]
                        is_duplicate = len(conflitti_ora) > 1 and sostituto_scelto != "Nessuno"
                        
                        if is_occupied or is_duplicate:
                            return ["background-color: #ffcccc"] * len(cell)
                    return [""] * len(cell)
                
                # Purtroppo non si può mappare per riga, quindi usiamo un editor meno avanzato ma che funzioni
                # Ritorna al vecchio editor per evitare complessità
                edited_df = st.data_editor(
                    sostituzioni_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Sostituto": st.column_config.SelectboxColumn(
                            "Sostituto",
                            options=opzioni,
                            default="Nessuno"
                        ),
                        "Note": st.column_config.TextColumn("Note", help="Aggiungi una nota alla supplenza")
                    },
                )
                
                # Rimappa la scelta del selectbox per il salvataggio
                edited_df["Sostituto"] = edited_df["Sostituto"].str.replace("[S] ", "")

                # evidenzia in verde i docenti di sostegno nei sostituti
                def evidenzia_sostegno_docente(val):
                    if val in orario_df[orario_df["Tipo"] == "Sostegno"]["Docente"].unique():
                        return "color: green; font-weight: bold;"
                    return ""

                styled_sostituzioni = edited_df.style.map(evidenzia_sostegno_docente, subset=["Sostituto"])
                st.dataframe(styled_sostituzioni, use_container_width=True, hide_index=True)


                # Testo WhatsApp
                edited_df_sorted = edited_df.copy()

                mesi_it = {
                    1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile",
                    5: "Maggio", 6: "Giugno", 7: "Luglio", 8: "Agosto",
                    9: "Settembre", 10: "Ottobre", 11: "Novembre", 12: "Dicembre"
                }

                giorno_num = data_sostituzione.day
                mese_nome = mesi_it[data_sostituzione.month]
                anno = data_sostituzione.year
                data_estesa = f"{giorno_assente} {giorno_num} {mese_nome} {anno}"

                testo = f"Sostituzioni per {data_estesa}:\n\n"
                for _, row in edited_df_sorted.iterrows():
                    testo += f"• Classe {row['Classe']} – {row['Ora']} ora (assente: {row['Assente']}) → {row['Sostituto']}"
                    if row.get("Note"):
                        testo += f" ({row['Note']})"
                    testo += "\n"

                st.subheader("📤 Testo per WhatsApp (copia e modifica)")
                st.text_area(
                    "Modifica qui il messaggio prima di copiarlo:",
                    testo.strip(),
                    height=200,
                    key="whatsapp_text_area"
                )

                if st.button("✅ Conferma e Salva"):
                    conflitti_rilevati = False
                    for ora in edited_df_sorted["Ora"].unique():
                        assegnazioni = edited_df_sorted[edited_df_sorted["Ora"] == ora]
                        sostituti = [s.replace("[S] ", "") for s in assegnazioni["Sostituto"] if s and s != "Nessuno"]
                        
                        for sost in set(sostituti):
                            is_occupied = sost in docenti_occupati_ora.get(ora, [])
                            is_duplicate = sostituti.count(sost) > 1

                            if is_occupied:
                                st.error(f"⚠️ Errore di conflitto: il docente {sost} è già assegnato a un'altra lezione nell'ora {ora}.")
                                conflitti_rilevati = True
                            if is_duplicate:
                                st.error(f"⚠️ Errore di conflitto: il docente {sost} è stato assegnato a più classi nell'ora {ora}.")
                                conflitti_rilevati = True

                    if not conflitti_rilevati:
                        if salva_storico_assenze(data_sostituzione, giorno_assente, edited_df, ore_assenti):
                            st.success("Assenze e sostituzioni salvate nello storico ✅")
                            try:
                                st.rerun()
                            except Exception:
                                pass


# --- VISUALIZZA ORARIO ---
elif menu == "Visualizza Orario":
    st.header("📅 Orario completo")
    if st.session_state.orario_df.empty:
        st.warning("Nessun orario disponibile.")
    else:
        vista_pivot_docenti(st.session_state.orario_df, mode="classi")
        download_orario(st.session_state.orario_df)

# --- STATISTICHE ---
elif menu == "Statistiche":
    st.header("📊 Statistiche Sostituzioni")
    df_storico, df_assenze = carica_statistiche()
    if df_storico.empty:
        st.info("Nessuna statistica disponibile. Registra prima delle sostituzioni.")
    else:
        df_sum = df_storico.groupby("docente")["ore"].sum().reset_index()
        df_sum = df_sum.rename(columns={"ore": "Totale Ore Sostituite"})
        st.dataframe(df_sum, use_container_width=True, hide_index=True)
        st.bar_chart(df_sum.set_index("docente"))

    st.subheader("⚠️ Azzeramento storico sostituzioni")
    conferma = st.checkbox("Confermo di voler cancellare definitivamente lo storico delle sostituzioni", key="conf_storico")
    if st.button("Elimina storico sostituzioni"):
        if conferma:
            if clear_sheet_content(STORICO_SHEET):
                st.success("Storico delle sostituzioni eliminato ✅")
        else:
            st.warning("Devi spuntare la conferma prima di cancellare lo storico delle sostituzioni.")

    st.header("📊 Statistiche Assenze")
    if df_assenze.empty:
        st.info("Nessuna assenza registrata.")
    else:
        st.dataframe(df_assenze, use_container_width=True, hide_index=True)
        df_assenze_agg = df_assenze.groupby("docente")["ora"].count().reset_index().rename(columns={"ora": "Totale Ore Assenti"})
        st.bar_chart(df_assenze_agg.set_index("docente"))

    st.subheader("⚠️ Azzeramento storico assenze")
    conferma_assenze = st.checkbox("Confermo di voler cancellare definitivamente lo storico delle assenze", key="conf_assenze")
    if st.button("Elimina storico assenze"):
        if conferma_assenze:
            if clear_sheet_content(ASSENZE_SHEET):
                st.success("Storico delle assenze eliminato ✅")
        else:
            st.warning("Devi spuntare la conferma prima di cancellare lo storico delle assenze.")
