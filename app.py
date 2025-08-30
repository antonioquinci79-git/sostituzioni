import streamlit as st
import pandas as pd
import sqlite3
import os
# =========================
# CONFIGURAZIONE FILE
# =========================
FILE_ORARIO = "orario.csv"
DB_FILE = "sostituzioni.db"

# =========================
# FUNZIONI DI SUPPORTO
# =========================
def get_connection():
    return sqlite3.connect(DB_FILE)

def init_db():
    with get_connection() as conn:
        c = conn.cursor()
        # 🔹 tabella storico (sostituzioni)
        c.execute('''CREATE TABLE IF NOT EXISTS storico 
                     (data TEXT, giorno TEXT, docente TEXT, ore INTEGER)''')
        # 🔹 nuova tabella assenze
        c.execute('''CREATE TABLE IF NOT EXISTS assenze
                     (data TEXT, giorno TEXT, docente TEXT, ora TEXT, classe TEXT)''')
        conn.commit()

def carica_orario():
    if os.path.exists(FILE_ORARIO):
        df = pd.read_csv(FILE_ORARIO)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    else:
        df = pd.DataFrame(columns=["Docente", "Giorno", "Ora", "Classe", "Tipo"])
    return df

def salva_orario(df):
    df.to_csv(FILE_ORARIO, index=False)

def backup_orario(df):
    backup_file = f"orario_backup_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv"
    df.to_csv(backup_file, index=False)
    return backup_file

def download_orario(df):
    if not df.empty:
        st.download_button(
            "⬇️ Scarica orario in CSV",
            data=df.to_csv(index=False),
            file_name="orario.csv",
            mime="text/csv"
        )

def vista_pivot_docenti(df):
    if df.empty:
        st.warning("Nessun orario disponibile.")
        return

    dfp = df.copy()

    # "Docente (Classe)" e prefisso [S] per sostegno
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

    # colorazione: verde se c'è almeno un sostegno, blu altrimenti (celle non vuote)
    def color_cells(val):
        text = str(val)
        if "[S]" in text:
            return "color: green; font-weight: bold;"
        elif text.strip() != "":
            return "color: blue;"
        return ""

    styled = pivot.style.map(color_cells)
    st.dataframe(styled, use_container_width=True)

    
    

# =========================
# AVVIO APP
# =========================
st.title("📚 Gestione orari e Sostituzioni Docenti")
init_db()
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
        orario_df = pd.read_csv(uploaded_file)
        orario_df = orario_df.loc[:, ~orario_df.columns.str.contains('^Unnamed')]
        salva_orario(orario_df)
        st.success("Orario caricato con successo ✅")

    # 🔹 Inserimento nuova lezione
    with st.expander("Aggiungi una nuova lezione"):
        docente = st.text_input("Nome docente")
        giorno = st.selectbox("Giorno", ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"])
        ora = st.selectbox("Ora", ["I", "II", "III", "IV", "V", "VI"])
        classe = st.text_input("Classe (es. 2A, 1B...)")
        tipo = st.selectbox("Tipo", ["Lezione", "Sostegno", "Altro"])
        asterisco = st.checkbox("Escludi da sostituzioni (asterisco)")

        if st.button("Aggiungi", key="add_lesson"):
            if docente and giorno and ora and classe and tipo:
                tipo_val = tipo if not asterisco else f"{tipo}*"
                row = {"Docente": docente, "Giorno": giorno, "Ora": ora, "Classe": classe, "Tipo": tipo_val}
                if "Tipo_no_asterisco" in orario_df.columns:
                    row["Tipo_no_asterisco"] = tipo
                nuovo = pd.DataFrame([row])

                if not orario_df.empty:
                    conflitto = orario_df[
                        (orario_df["Docente"] == docente) &
                        (orario_df["Giorno"] == giorno) &
                        (orario_df["Ora"] == ora)
                    ].empty
                    if not conflitto:
                        st.error("Conflitto: questo docente è già assegnato a un'altra classe in questa ora.")
                    else:
                        orario_df = pd.concat([orario_df, nuovo], ignore_index=True)
                        salva_orario(orario_df)
                        st.success("Lezione aggiunta all'orario ✅")
                else:
                    orario_df = pd.concat([orario_df, nuovo], ignore_index=True)
                    salva_orario(orario_df)
                    st.success("Lezione aggiunta all'orario ✅")
            else:
                st.error("Compila tutti i campi per aggiungere una lezione.")

    # 🔹 Modifica orario esistente
    st.subheader("📝 Modifica orario attuale")
    if not orario_df.empty:
        # Reimposta l'ordine delle colonne
        col_order = ["Giorno", "Ora", "Docente", "Classe", "Tipo"]
        df_edit = orario_df.copy()
        df_edit = df_edit[col_order]

        # Ordina per Docente
        df_edit = df_edit.sort_values(by=["Docente"])

        # Mostra l’editor con colonne e ordine desiderati
        edited_df = st.data_editor(
            df_edit,
            use_container_width=True, hide_index=True
        )

        if st.button("Salva modifiche"):
            backup_file = backup_orario(orario_df)
            edited_df["Tipo_no_asterisco"] = edited_df["Tipo"].str.replace('*', '', regex=False)
            orario_df = edited_df
            salva_orario(orario_df)
            st.success(f"Orario modificato e salvato. Backup creato: {backup_file} ✅")



    # 🔹 Eliminazione riga specifica
    st.subheader("🗑️ Elimina una riga dall'orario")
    if not orario_df.empty:
        row_options = [
            f"{row['Docente']} - {row['Giorno']} - {row['Ora']} - {row['Classe']} - {row['Tipo']}"
            for _, row in orario_df.iterrows()
        ]
        row_to_delete = st.selectbox("Seleziona la riga da eliminare", ["Seleziona una riga"] + row_options)

        if row_to_delete != "Seleziona una riga" and st.button("Elimina riga"):
            backup_file = backup_orario(orario_df)
            row_index = [i for i, r in enumerate(row_options) if r == row_to_delete][0]
            orario_df = orario_df.drop(index=row_index).reset_index(drop=True)
            salva_orario(orario_df)
            st.success(f"Riga eliminata: {row_to_delete}. Backup creato: {backup_file} ✅")
    else:
        st.info("Nessun orario da eliminare.")

    # 🔹 Download orario
    download_orario(orario_df)
    
    
    # 🔹 Cancella tutto l'orario
    st.subheader("⚠️ Cancella tutto l'orario")
    conferma = st.checkbox("Confermo di voler cancellare definitivamente l'orario")

    if st.button("Elimina completamente l'orario"):
        if conferma:
            backup_file = backup_orario(orario_df)
            orario_df = pd.DataFrame(columns=["Docente", "Giorno", "Ora", "Classe", "Tipo"])
            salva_orario(orario_df)
            st.success(f"Tutto l'orario è stato eliminato. Backup creato: {backup_file} ✅")
        else:
            st.warning("Devi spuntare la conferma prima di cancellare l'orario.")



# --- GESTIONE ASSENZE ---
elif menu == "Gestione Assenze":
    st.header("🚨 Gestione Assenze")

    if orario_df.empty:
        st.warning("Non hai ancora caricato nessun orario.")
    else:
        docenti_assenti = st.multiselect("Seleziona docenti assenti", sorted(orario_df["Docente"].unique()))

        # 🔹 data con calendario
        data_sostituzione = st.date_input("Data della sostituzione")

        # 🔹 giorno calcolato automaticamente dalla data
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
        giorno_assente = traduzione_giorni[giorno_assente]

        if not docenti_assenti:
            st.info("Seleziona almeno un docente per continuare.")
        else:
            # Ore scoperte
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
                with get_connection() as conn:
                    storici = pd.read_sql("SELECT docente, SUM(ore) as total FROM storico GROUP BY docente", conn)
                storici = storici.fillna(0)

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

                    liberi = [
                        d for d in presenti_ora
                        if d not in occupati_curr
                        and d not in docenti_assenti
                        and "*" not in d
                    ]

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

                # --- Ordina subito le proposte per Ora ---
                ordine_ore = ["I", "II", "III", "IV", "V", "VI"]
                sostituzioni_df["Ora"] = pd.Categorical(sostituzioni_df["Ora"], categories=ordine_ore, ordered=True)
                sostituzioni_df = sostituzioni_df.sort_values("Ora").reset_index(drop=True)

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
                        proposto = f"[S] {proposto}"

                    scelta = st.selectbox(
                        f"Sostituto per {riga['Ora']} ora - Classe {riga['Classe']} (assente {riga['Assente']})",
                        opzioni,
                        index=opzioni.index(proposto) if proposto in opzioni else 0,
                        key=f"select_{idx}"
                    )

                    # Rimuovi eventuale [S] quando salviamo la scelta
                    edited_df.at[idx, "Sostituto"] = scelta.replace("[S] ", "")

                # --- evidenzia in verde i docenti di sostegno nei sostituti ---
                def evidenzia_sostegno_docente(val):
                    if val in orario_df[orario_df["Tipo"] == "Sostegno"]["Docente"].unique():
                        return "color: green; font-weight: bold;"
                    return ""

                # --- Tabella riassuntiva ordinata per Ora ---
                edited_df = edited_df.sort_values("Ora").reset_index(drop=True)
                styled_sostituzioni = edited_df.style.map(evidenzia_sostegno_docente, subset=["Sostituto"])
                st.dataframe(styled_sostituzioni, use_container_width=True, hide_index=True)

                # --- ORDINA testo WhatsApp per ora ---
                edited_df_sorted = edited_df.copy()
                testo = "Sostituzioni per " + giorno_assente + ":\n"
                for _, row in edited_df_sorted.iterrows():
                    testo += f"- {row['Ora']} ora - Classe {row['Classe']} (assente {row['Assente']}) → {row['Sostituto']}\n"

                st.subheader("📤 Testo per WhatsApp (copia e modifica)")
                st.text_area("Modifica qui il messaggio prima di copiarlo:", testo, height=200, key="whatsapp_text")

                # --- Step 1: Conferma tabella (senza salvare nello storico) ---
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

                    # Se nessun conflitto, salviamo la tabella in sessione
                    st.session_state["sostituzioni_confermate"] = edited_df_sorted.copy()
                    st.success("Tabella confermata ✅ Ora puoi salvarla nello storico.")

# --- Step 2: Salva nello storico (solo se confermata) ---
if "sostituzioni_confermate" in st.session_state:
    if st.button("💾 Salva nello storico"):
        with get_connection() as conn:
            c = conn.cursor()

            # 🔹 Salva sostituzioni
            for _, row in st.session_state["sostituzioni_confermate"].iterrows():
                if row["Sostituto"] != "Nessuno":
                    c.execute(
                        "INSERT INTO storico VALUES (?, ?, ?, ?)",
                        (str(data_sostituzione), giorno_assente, row["Sostituto"], 1)
                    )

            # 🔹 Salva assenze (docente, giorno, ora, classe)
            for _, row in ore_assenti.iterrows():
                c.execute(
                    "INSERT INTO assenze VALUES (?, ?, ?, ?, ?)",
                    (str(data_sostituzione), giorno_assente, row["Docente"], row["Ora"], row["Classe"])
                )

            conn.commit()
        st.success("Assenze e sostituzioni salvate nello storico ✅")
        del st.session_state["sostituzioni_confermate"]




# --- VISUALIZZA ORARIO ---
elif menu == "Visualizza Orario":
    st.header("📅 Orario completo")
    if orario_df.empty:
        st.warning("Nessun orario disponibile.")
    else:
        import re

        def vista_pivot_docenti(df):
            if df.empty:
                st.warning("Nessun orario disponibile.")
                return

            dfp = df.copy()
            dfp["Info"] = dfp["Docente"]

            # Pivot: righe = Giorno + Ora, colonne = Classe
            pivot = dfp.pivot_table(
                index=["Giorno", "Ora"],
                columns="Classe",
                values="Info",
                aggfunc=lambda x: " / ".join(list(dict.fromkeys(x)))
            ).fillna("-")

            # Ordina colonne classe tipo 1A, 2A, 3A, ...
            def sort_classi(classe):
                m = re.match(r"(\d+)\s*([A-Za-zÀ-ÖØ-öø-ÿ]+)", str(classe))
                if m:
                    return (int(m.group(1)), m.group(2).upper())
                return (10**9, str(classe))
            pivot = pivot.reindex(sorted(pivot.columns, key=sort_classi), axis=1)

            # Ordine settimanale (lun-ven) e ore
            ordine_giorni = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]
            ordine_ore = ["I", "II", "III", "IV", "V", "VI"]
            pivot = pivot.reindex(pd.MultiIndex.from_product([ordine_giorni, ordine_ore], names=["Giorno", "Ora"]))

            # Trasforma l’indice in colonne per mostrare Giorno e Ora
            pivot = pivot.reset_index()

            # 🔹 Inserisci una riga separatrice tra i giorni
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

            # 🔹 Mostra il giorno solo nella prima riga del blocco
            last_giorno = None
            for i in range(len(pivot)):
                if pivot.loc[i, "Giorno"] == last_giorno and pivot.loc[i, "Giorno"] not in ["", "──"]:
                    pivot.loc[i, "Giorno"] = ""
                else:
                    last_giorno = pivot.loc[i, "Giorno"]

            # Centra testo celle e rimuovi numeri progressivi
            styled = pivot.style.set_properties(**{"text-align": "center"})
            st.dataframe(styled, use_container_width=True, hide_index=True)

        vista_pivot_docenti(orario_df)

        # Download (invariato)
        download_orario(orario_df)



# --- STATISTICHE ---
elif menu == "Statistiche":
    st.header("📊 Statistiche Sostituzioni")
    with get_connection() as conn:
        # 🔹 Sostituzioni (con data e giorno)
        df_storico = pd.read_sql(
            "SELECT docente, data, giorno, SUM(ore) as total FROM storico GROUP BY docente, data, giorno",
            conn
        )
    df_storico = df_storico.fillna(0)

    if df_storico.empty:
        st.info("Nessuna statistica disponibile. Registra prima delle sostituzioni.")
    else:
        st.dataframe(df_storico, use_container_width=True, hide_index=True)
        st.bar_chart(df_storico.groupby("docente")["total"].sum())

    st.subheader("⚠️ Azzeramento storico sostituzioni")
    conferma = st.checkbox("Confermo di voler cancellare definitivamente lo storico delle sostituzioni")

    if st.button("Elimina storico sostituzioni"):
        if conferma:
            with get_connection() as conn:
                c = conn.cursor()
                c.execute("DELETE FROM storico")
                conn.commit()
            st.success("Storico delle sostituzioni eliminato ✅")
        else:
            st.warning("Devi spuntare la conferma prima di cancellare lo storico delle sostituzioni.")

    # --- STATISTICHE ASSENZE ---
    st.header("📊 Statistiche Assenze")
    with get_connection() as conn:
        df_assenze = pd.read_sql(
            "SELECT docente, data, giorno, COUNT(*) as totale_ore_assenti FROM assenze GROUP BY docente, data, giorno",
            conn
        )
    df_assenze = df_assenze.fillna(0)

    if df_assenze.empty:
        st.info("Nessuna assenza registrata.")
    else:
        st.dataframe(df_assenze, use_container_width=True, hide_index=True)
        st.bar_chart(df_assenze.groupby("docente")["totale_ore_assenti"].sum())

    st.subheader("⚠️ Azzeramento storico assenze")
    conferma_assenze = st.checkbox("Confermo di voler cancellare definitivamente lo storico delle assenze")

    if st.button("Elimina storico assenze"):
        if conferma_assenze:
            with get_connection() as conn:
                c = conn.cursor()
                c.execute("DELETE FROM assenze")
                conn.commit()
            st.success("Storico delle assenze eliminato ✅")
        else:
            st.warning("Devi spuntare la conferma prima di cancellare lo storico delle assenze.")


