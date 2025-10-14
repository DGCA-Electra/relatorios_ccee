import streamlit as st
import os
import logging
from src.view.main_page import show_main_page
from src.view.config_page import show_config_page

LOG_DIR = 'logs'
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, 'app.log')
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s', encoding='utf-8')

st.set_page_config(
    page_title="Envio de Relatórios CCEE",
    page_icon="static/icon.png",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={
        'About': "Aplicação para automação de envio de e-mails DGCA."
    }
)

def main() -> None:
    st.image("static/logo.png", width=250)

    st.sidebar.title("🧭 Navegação")
    page_options = ["Envio de Relatórios", "Configurações"]
    page = st.sidebar.radio("Escolha a página:", page_options, label_visibility="hidden", key="sidebar_radio")
    
    if page == "Envio de Relatórios":
        show_main_page()
    else:
        show_config_page()

if __name__ == "__main__":
    main()
    st.sidebar.info("Aplicação desenvolvida para automação de envio de e-mails DGCA.")
    st.sidebar.warning("Nota: Ao processar, janelas do Outlook podem abrir para sua revisão. Isso é esperado.")
    st.sidebar.markdown("---")
    st.sidebar.markdown("© 2025 Desenvolvido por Malik Ribeiro")