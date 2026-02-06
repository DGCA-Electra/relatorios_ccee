import streamlit as st
import os
import logging
from dotenv import load_dotenv
import msal
import requests
from apps.relatorios_ccee.model.arquivos import obtem_asset_path
from apps.relatorios_ccee.controller import auth_controller

# --- CARREGAMENTO DO AMBIENTE ---
# Garante que o .env seja lido da pasta atual do aplicativo
current_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(current_dir, ".env")
load_dotenv(env_path)

# --- CONFIGURAÇÕES ---
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
TENANT_ID = os.environ.get("AZURE_TENANT_ID")
REDIRECT_URI = os.environ.get("AZURE_REDIRECT_URI")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Mail.Send"]

# --- DIAGNÓSTICO VISUAL (Para encontrarmos o erro) ---
# Se algum valor estiver estranho, você verá aqui na tela
if not CLIENT_ID or len(CLIENT_ID) < 5:
    st.error(f"❌ CLIENT_ID parece inválido ou vazio: '{CLIENT_ID}'")
if not TENANT_ID or len(TENANT_ID) < 5:
    st.error(f"❌ TENANT_ID parece inválido ou vazio: '{TENANT_ID}'")
if not REDIRECT_URI or not REDIRECT_URI.startswith("http"):
    st.error(f"❌ REDIRECT_URI inválida: '{REDIRECT_URI}'")

# Validação crítica
if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID, REDIRECT_URI]):
    st.error("Erro Crítico: Variáveis do .env não carregaram.")
    st.info(f"Tentando ler de: {env_path}")
    st.stop()

# Instância MSAL
try:
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
except Exception as e_init:
    st.error(f"❌ Erro ao criar app MSAL (Verifique Tenant/Client ID): {e_init}")
    st.stop()

# Validação crítica
if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID, REDIRECT_URI]):
    st.error("Erro Crítico: Variáveis de configuração do Azure AD não encontradas.")
    st.info(f"O sistema buscou o arquivo .env em: {env_path}")
    st.stop()

# Instância MSAL
msal_app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)

def obter_url_autenticacao():
    try:
        # Tenta gerar a URL
        url = msal_app.get_authorization_request_url(
            SCOPES,
            redirect_uri=REDIRECT_URI,
            response_type='code'
        )
        return url
    except Exception as e:
        # AQUI ESTÁ O SEGREDO: Mostra o erro real na tela
        st.error(f"❌ Erro técnico ao gerar link: {str(e)}")
        logging.error(f"Erro detalhado: {e}")
        return None

def obter_token_do_codigo(codigo_autorizacao):
    try:
        resultado = msal_app.acquire_token_by_authorization_code(
            codigo_autorizacao,
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI
        )
        if "error" in resultado:
            st.error(f"Erro no Token: {resultado.get('error_description')}")
            return None
        return resultado
    except Exception as e:
        logging.error(f"Exceção Token: {e}")
        st.error("Falha na comunicação com autenticação.")
        return None

def obter_info_usuario(token_acesso):
    headers = {'Authorization': 'Bearer ' + token_acesso}
    try:
        resposta = requests.get('https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName', headers=headers)
        resposta.raise_for_status()
        return resposta.json()
    except Exception as e:
        logging.error(f"Erro Graph API: {e}")
        return None

def show_login_page():
    st.image(obtem_asset_path("logo.png"), width=250)
    st.title("Login - Envio de Relatórios CCEE")
    st.write("Por favor, autentique-se com sua conta Microsoft para continuar.")

    params = st.query_params
    codigo = params.get("code")

    if codigo:
        try:
            auth_controller.processar_callback(codigo)
        except Exception as e:
            st.error(f"Falha ao validar login: {e}")
            logging.exception("Erro durante callback de autenticação:")
            url_auth = auth_controller.obter_url_autenticacao()
            if url_auth:
                st.markdown(f'<a href="{url_auth}" target="_self" class="button">Tentar Novamente</a>', unsafe_allow_html=True)
    else:
        url_auth = auth_controller.obter_url_autenticacao()
        if url_auth:
            st.markdown("""
            <style>
            .button {
                display: inline-block;
                padding: 0.5em 1em;
                background-color: #0078D4;
                color: white !important;
                border-radius: 4px;
                text-decoration: none;
                font-size: 16px;
                cursor: pointer;
            }
            .button:hover { background-color: #005A9E; }
            </style>
            """, unsafe_allow_html=True)
            st.markdown(f'<a href="{url_auth}" target="_self" class="button">Entrar com Microsoft</a>', unsafe_allow_html=True)
        else:
            st.error("Erro ao gerar link de login.")
    st.stop()