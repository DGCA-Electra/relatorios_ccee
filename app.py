import streamlit as st
import os
import logging
from dotenv import load_dotenv

load_dotenv()

from src.view.main_page import show_main_page
from src.view.config_page import show_config_page
from src.view.login_page import show_login_page

LOG_DIR = 'logs'
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, 'app.log')
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s', encoding='utf-8')

st.set_page_config(
    page_title="Envio de Relatórios CCEE",
    page_icon="static/icon.png",
    layout="wide",
    initial_sidebar_state="auto",
    menu_items={
        'About': "Aplicação para automação de envio de e-mails DGCA."
    }
)

def logout():
    """Limpa as informações de sessão relacionadas ao login."""
    keys_to_remove = ["ms_token", "user_info"]
    for key in keys_to_remove:
        if key in st.session_state:
            del st.session_state[key]
    logging.info("Usuário fez logout.")
    st.query_params.clear()
    st.rerun()

def main() -> None:
    if "ms_token" not in st.session_state:
        show_login_page()
    else:
        st.image("static/logo.png", width=250)

        user_info = st.session_state.get("user_info", {})
        user_name = user_info.get("displayName", "Usuário")
        user_email = user_info.get("userPrincipalName", "")
        st.sidebar.success(f"Logado como: {user_name}")
        if user_email:
             st.sidebar.caption(user_email)
        if st.sidebar.button("Logout"):
            logout()

        st.sidebar.title("🧭 Navegação")
        page_options = ["Envio de Relatórios", "Configurações"]
        page = st.sidebar.radio("Escolha a página:", page_options, label_visibility="collapsed", key="main_sidebar_radio")

        if page == "Envio de Relatórios":
            show_main_page()
        elif page == "Configurações":
            show_config_page()

        st.sidebar.info("Aplicação desenvolvida para automação de envio de e-mails DGCA.")
        st.sidebar.warning("Nota: Os e-mails serão criados como rascunhos na sua caixa de entrada.")
        st.sidebar.markdown("---")
        st.sidebar.markdown("© 2025 Desenvolvido por Malik Ribeiro")


if __name__ == "__main__":
    main()