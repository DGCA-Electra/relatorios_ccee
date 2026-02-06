import streamlit as st
import os
import logging
from dotenv import load_dotenv

load_dotenv()

from apps.relatorios_ccee.view.ui_relatorios import exibir_pagina_principal
from apps.relatorios_ccee.view.configuracao import show_config_page
from apps.relatorios_ccee.view.login import show_login_page
from apps.relatorios_ccee.controller import auth_controller
from apps.relatorios_ccee.model.arquivos import obtem_asset_path

LOG_DIR = 'logs'
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, 'app.log')
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s', encoding='utf-8')

def logout():
    """Delegates logout to the AuthController (clears auth session)."""
    auth_controller.logout()


def main() -> None:
    st.set_page_config(page_title="Envio Relatórios CCEE", layout="wide")

    if "ms_token" not in st.session_state:
        show_login_page()
        return

    logo_path = obtem_asset_path("logo.png")
    if logo_path:
        st.image(logo_path, width=250)

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
        exibir_pagina_principal()
    elif page == "Configurações":
        show_config_page()

    st.sidebar.warning("Nota: Os e-mails serão criados como rascunhos na sua caixa de entrada.")
    st.sidebar.markdown("---")
    st.sidebar.markdown("© 2025 Desenvolvido por Malik Ribeiro")


if __name__ == "__main__":
    main()