import streamlit as st
import streamlit.components.v1 as components

def copy_to_clipboard_button(label, text_to_copy, key=None):
    """
    Crea un pulsante Streamlit che copia il testo negli appunti quando viene cliccato.
    
    Args:
        label (str): Il testo da mostrare sul pulsante.
        text_to_copy (str): Il testo che verrà copiato negli appunti.
        key (str, optional): Una chiave unica per il componente.
    """
    button_js = f"""
    <button id="copy-button-{key}" 
            style="
                background-color: #007bff; 
                color: white; 
                padding: 0.5rem 1rem;
                border-radius: 0.25rem;
                border: none;
                cursor: pointer;
            "
            onclick="copyToClipboard(`{text_to_copy.replace('`', '\\`')}`)">
        {label}
    </button>
    <script>
        function copyToClipboard(text) {{
            navigator.clipboard.writeText(text).then(function() {{
                const button = document.getElementById('copy-button-{key}');
                button.innerText = 'Copiato!';
                setTimeout(() => {{
                    button.innerText = '{label}';
                }}, 2000);
            }}, function(err) {{
                console.error('Errore nella copia: ', err);
            }});
        }}
    </script>
    """
    components.html(button_js, height=50)

if __name__ == '__main__':
    st.title("Esempio di pulsante personalizzato")
    
    testo_da_copiare = st.text_area("Scrivi il testo da copiare qui:", "Questo è un testo di prova.")
    
    # Usa il componente personalizzato
    copy_to_clipboard_button("Copia Testo", testo_da_copiare, key="example_key")
