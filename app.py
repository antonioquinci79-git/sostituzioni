# app_googlesheets.py
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
        # Proviamo a crearne uno nuovo se non esiste
        sh = client.create(SPREADSHEET_NAME)
        # la creazione potrebbe non impostare permessi: l'utente deve condividere il foglio con il service account se necessario
    # Verifica e crea fogli con intestazioni se mancanti
    existing = {ws.title for ws in sh.worksheets()}
    if ORARIO_SHEET not in existing:
        ws = sh.add_worksheet(title=ORARIO_SHEET, rows="100", cols="20")
        header_df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        gd.set_with_dataframe(ws, header_df, include_index=False, include_column_header=True)
    if STORICO_SHEET not in existing:
        ws = sh.add_worksheet(title=STORICO_SHEET, rows="100", cols="10")
        header = pd.DataFrame(columns=["data", "giorno", "docente", "ore"])
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
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(ORARIO_SHEET)
        # assicurati che le colonne siano quelle giuste e in ordine
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
        # Normalizza nomi e tipi
        if not df_storico.empty:
            if "data" in df_storico.columns:
                df_storico["data"] = pd.to_datetime(df_storico["data"], errors="coerce")
                df_storico["data"] = df_storico["data"].dt.strftime("%Y-%m-%d")
            for c in ["docente", "giorno", "ore"]:
                if c in df_storico.columns:
                    df_storico[c] = df_storico[c].astype(str).str.strip()
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
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws_storico = sh.worksheet(STORICO_SHEET)
        ws_assenze = sh.worksheet(ASSENZE_SHEET)

        # Prepara dati per append
        storico_data = []
        for _, row in sostituzioni_df.iterrows():
            sost = row.get("Sostituto", "")
            if sost and sost != "Nessuno":
                storico_data.append([str(data_sostituzione), giorno_assente, sost, 1])

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
            ws.append_row(["data", "giorno", "docente", "ore"])
        elif sheet_name == ASSENZE_SHEET:
            ws.append_row(["data", "giorno", "docente", "ora", "classe"])
        return True
    except Exception as e:
        st.error(f"Errore nell'azzeramento del foglio {sheet_name}: {e}")
        return False

def rimuovi_righe_per_data(sheet_name, data):
    """Rimuove tutte le righe di uno sheet che corrispondono alla data passata"""
    try:
        client = get_gdrive_client()
        sh = client.open(SPREADSHEET_NAME)
        ws = sh.worksheet(sheet_name)

        # Scarica tutti i dati
        df = gd.get_as_dataframe(ws, header=0).dropna(how="all")

        if "data" not in df.columns:
            return False  # se il foglio non ha colonna data, non facciamo nulla

        # Filtra le righe diverse dalla data indicata
        df_filtrato = df[df["data"] != str(data)]

        # Riscrivi il foglio con i dati filtrati
        gd.set_with_dataframe(ws, df_filtrato, include_index=False, include_column_header=True)
        return True
    except Exception as e:
        st.error(f"Errore nella rimozione delle righe del foglio {sheet_name} per la data {data}: {e}")
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
# assicurati che i fogli esistano con le intestazioni
try:
    ensure_sheets_exist()
except Exception as e:
    st.error(f"Impossibile inizializzare i fogli Google: {e}")

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
            orario_df = df_tmp[REQUIRED_COLUMNS].copy()
            if salva_orario(orario_df):
                st.success("Orario caricato con successo ✅")
        else:
            st.error(f"CSV non valido. Deve contenere: {REQUIRED_COLUMNS}")

    # Inserimento nuova lezione
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

    # Modifica orario esistente
    st.subheader("📝 Modifica orario attuale")
    if not orario_df.empty:
        col_order = REQUIRED_COLUMNS
        df_edit = orario_df[col_order].copy()
        df_edit = df_edit.sort_values(by=["Docente"])
        edited_df = st.data_editor(df_edit, use_container_width=True, hide_index=True)
        if st.button("Salva modifiche"):
            # backup: salviamo la versione corrente come foglio temporaneo? Qui semplicemente sovrascriviamo sul foglio
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

        # giorno calcolato automaticamente dalla data (in italiano)
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
            # Ore scoperte (rispettando Escludi come in app.py)
            ore_assenti = orario_df[
                (orario_df["Docente"].isin(docenti_assenti)) &
                (orario_df["Giorno"] == giorno_assente)
            ]

            if ore_assenti.empty:
                st.info("I docenti selezionati non hanno lezioni in quel giorno.")
            else:
                st.subheader("📌 Ore scoperte")

                def evidenzia_sostegno(val, tipo):
                    if tipo == "Sostegno":
                        return "color: green; font-weight: bold;"
                    return ""

                styled_ore_assenti = ore_assenti[["Docente", "Ora", "Classe", "Tipo"]].style.apply(
                    lambda col: [evidenzia_sostegno(v, t) for v, t in zip(ore_assenti["Docente"], ore_assenti["Tipo"])],
                    axis=0
                )
                st.dataframe(styled_ore_assenti, use_container_width=True, hide_index=True)

                st.subheader("🔄 Possibili sostituti")
                sostituzioni = []
                # Carica storico per prioritizzazione del carico (come in app.py)
                df_storico, df_assenze = carica_statistiche()
                storici = pd.DataFrame()
                if not df_storico.empty:
                    storici = df_storico.groupby("docente")["ore"].sum().reset_index().rename(columns={"ore": "total"})
                # Itera le ore scoperte e propone sostituti seguendo la stessa logica di app.py
                for _, row in ore_assenti.iterrows():
                    presenti_ora = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == row["Ora"])
                    ]["Docente"].unique()

                    occupati_curr = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == row["Ora"]) &
                        (orario_df["Tipo"] != "Sostegno")
                    ]["Docente"].unique()

                    # liberi: presenti in orario per quell'ora ma non occupati_curr, non assenti e non con *
                    liberi = [
                        d for d in presenti_ora
                        if d not in occupati_curr
                        and d not in docenti_assenti
                        and "*" not in d
                    ]

                    # Priorità: stesso sostegno di quella classe e ora
                    prioritari = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == row["Ora"]) &
                        (orario_df["Tipo"] == "Sostegno") &
                        (orario_df["Classe"] == row["Classe"]) &
                        (~orario_df["Docente"].str.contains('\*', regex=True, na=False))
                    ]
                    if not prioritari.empty:
                        proposto = prioritari["Docente"].iloc[0]
                    else:
                        sostegni_disponibili = orario_df[
                            (orario_df["Giorno"] == giorno_assente) &
                            (orario_df["Ora"] == row["Ora"]) &
                            (orario_df["Tipo"] == "Sostegno") &
                            (~orario_df["Docente"].str.contains('\*', regex=True, na=False))
                        ]["Docente"].unique()

                        candidati = [d for d in sostegni_disponibili if d in liberi]

                        if candidati:
                            proposto = candidati[0]
                        else:
                            proposto = "Nessuno"

                    sostituzioni.append({
                        "Ora": row["Ora"],
                        "Classe": row["Classe"],
                        "Assente": row["Docente"],
                        "Sostituto": proposto
                    })

                sostituzioni_df = pd.DataFrame(sostituzioni)

                # Ordina per ora
                ordine_ore = ["I", "II", "III", "IV", "V", "VI"]
                if not sostituzioni_df.empty:
                    sostituzioni_df["Ora"] = pd.Categorical(sostituzioni_df["Ora"], categories=ordine_ore, ordered=True)
                    sostituzioni_df = sostituzioni_df.sort_values("Ora").reset_index(drop=True)

                # Editor: per ogni riga, mostra selectbox con opzioni coerenti (inclusi [S])
                edited_df = sostituzioni_df.copy()
                for idx, riga in sostituzioni_df.iterrows():
                    presenti_ora = orario_df[
                        (orario_df["Giorno"] == giorno_assente) &
                        (orario_df["Ora"] == riga["Ora"])
                    ]["Docente"].unique()

                    # Lista sostegni
                    sostegni = orario_df[orario_df["Tipo"] == "Sostegno"]["Docente"].unique()

                    # Costruisci opzioni con prefisso [S] per i sostegni
                    opzioni_validi = []
                    for d in presenti_ora:
                        if d not in docenti_assenti and "*" not in d:
                            if d in sostegni:
                                opzioni_validi.append(f"[S] {d}")
                            else:
                                opzioni_validi.append(d)

                    # Metti i sostegni in cima
                    opzioni = ["Nessuno"] + sorted(
                        opzioni_validi,
                        key=lambda x: (0 if x.startswith("[S]") else 1, x)
                    )

                    # Mantieni coerenza con la proposta iniziale (aggiungendo [S] se serve)
                    proposto = riga["Sostituto"]
                    if proposto in sostegni:
                        proposto_display = f"[S] {proposto}"
                    else:
                        proposto_display = proposto

                    scelta = st.selectbox(
                        f"Sostituto per {riga['Ora']} ora - Classe {riga['Classe']} (assente {riga['Assente']})",
                        opzioni,
                        index=opzioni.index(proposto_display) if proposto_display in opzioni else 0,
                        key=f"select_{idx}"
                    )

                    # Rimuovi eventuale [S] quando salviamo la scelta
                    edited_df.at[idx, "Sostituto"] = scelta.replace("[S] ", "")

                # evidenzia in verde i docenti di sostegno nei sostituti
                def evidenzia_sostegno_docente(val):
                    if val in orario_df[orario_df["Tipo"] == "Sostegno"]["Docente"].unique():
                        return "color: green; font-weight: bold;"
                    return ""

                edited_df = edited_df.sort_values("Ora").reset_index(drop=True)
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
                    testo += f"• Classe {row['Classe']} – {row['Ora']} ora (assente: {row['Assente']}) → {row['Sostituto']}\n"

                st.subheader("📤 Testo per WhatsApp (copia e modifica)")
                st.text_area(
                    "Modifica qui il messaggio prima di copiarlo:",
                    testo.strip(),
                    height=200,
                    key="whatsapp_text_area"
                )

# --- Step 1: conferma tabella (non salva ancora) ---
if st.button("✅ Conferma tabella (non salva ancora)"):
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

# --- Step 2: Salva nello storico (con controllo se già esistono dati per quella data) ---
if st.session_state.get("sostituzioni_confermate") is not None:
    sost_df = st.session_state.get("sostituzioni_confermate")
    ore_assenti_session = st.session_state.get("ore_assenti_confermate")
    data_tmp = st.session_state.get("data_sostituzione_tmp")
    giorno_tmp = st.session_state.get("giorno_assente_tmp")

    # Carica lo storico attuale
    df_storico, df_assenze = carica_statistiche()
    gia_presenti = False
    if not df_storico.empty and str(data_tmp) in df_storico["data"].values:
        gia_presenti = True

    if gia_presenti:
        st.warning(f"⚠️ Ci sono già sostituzioni salvate per la data {data_tmp}. Vuoi aggiungere oppure sostituire quelle esistenti?")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("➕ Aggiungi allo storico"):
                if salva_storico_assenze(data_tmp, giorno_tmp, sost_df, ore_assenti_session):
                    st.success("Nuove assenze e sostituzioni aggiunte allo storico ✅")
                    for k in ["sostituzioni_confermate", "ore_assenti_confermate", "data_sostituzione_tmp", "giorno_assente_tmp"]:
                        st.session_state.pop(k, None)
                    st.rerun()

        with col2:
            if st.button("♻️ Sostituisci quelle del giorno"):
                # rimuovi le righe di quella data sia da storico che da assenze
                ok1 = rimuovi_righe_per_data(STORICO_SHEET, data_tmp)
                ok2 = rimuovi_righe_per_data(ASSENZE_SHEET, data_tmp)
                if ok1 and ok2:
                    if salva_storico_assenze(data_tmp, giorno_tmp, sost_df, ore_assenti_session):
                        st.success("Assenze e sostituzioni aggiornate nello storico ✅")
                        for k in ["sostituzioni_confermate", "ore_assenti_confermate", "data_sostituzione_tmp", "giorno_assente_tmp"]:
                            st.session_state.pop(k, None)
                        st.rerun()
    else:
        # Se non ci sono già dati per quel giorno → salvataggio normale
        if st.button("💾 Salva nello storico", key="save_storico_main"):
            if sost_df is not None and ore_assenti_session is not None:
                if salva_storico_assenze(data_tmp, giorno_tmp, sost_df, ore_assenti_session):
                    st.success("Assenze e sostituzioni salvate nello storico ✅")
                    for k in ["sostituzioni_confermate", "ore_assenti_confermate", "data_sostituzione_tmp", "giorno_assente_tmp"]:
                        st.session_state.pop(k, None)
                    st.rerun()




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
    df_storico, df_assenze = carica_statistiche()
    if df_storico.empty:
        st.info("Nessuna statistica disponibile. Registra prima delle sostituzioni.")
    else:
        st.dataframe(df_storico, use_container_width=True, hide_index=True)
        # barra: somma ore per docente
        try:
            df_sum = df_storico.groupby("docente")["ore"].sum().reset_index()
            df_sum = df_sum.rename(columns={"ore": "Totale Ore Sostituite"})
            st.bar_chart(df_sum.set_index("docente"))
        except Exception:
            # se 'ore' non numerico, proviamo a convertire
            df_storico["ore"] = pd.to_numeric(df_storico["ore"], errors="coerce").fillna(0)
            df_sum = df_storico.groupby("docente")["ore"].sum().reset_index().rename(columns={"ore": "Totale Ore Sostituite"})
            st.bar_chart(df_sum.set_index("docente"))

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
        st.dataframe(df_assenze, use_container_width=True, hide_index=True)
        try:
            df_assenze_agg = df_assenze.groupby("docente")["ora"].count().reset_index().rename(columns={"ora": "Totale Ore Assenti"})
            st.bar_chart(df_assenze_agg.set_index("docente"))
        except Exception:
            st.info("Impossibile calcolare le statistiche delle assenze (controlla intestazioni nel foglio).")

    st.subheader("⚠️ Azzeramento storico assenze")
    conferma_assenze = st.checkbox("Confermo di voler cancellare definitivamente lo storico delle assenze", key="conf_assenze")
    if st.button("Elimina storico assenze"):
        if conferma_assenze:
            if clear_sheet_content(ASSENZE_SHEET):
                st.success("Storico delle assenze eliminato ✅")
        else:
            st.warning("Devi spuntare la conferma prima di cancellare lo storico delle assenze.")
