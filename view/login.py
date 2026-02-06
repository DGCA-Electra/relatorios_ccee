import streamlit as st
import msal
import requests
import os
import logging
from dotenv import load_dotenv
from model.arquivos import obtem_asset_path

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
    st.image(obtem_asset_path("logo.png"))
    
    st.title("Login - Envio de Relatórios CCEE")
    st.write("Por favor, autentique-se com sua conta Microsoft para continuar.")

    # Verifica retorno do Login (Callback)
    parametros = st.query_params
    codigo = parametros.get("code")

    if codigo:
        # Troca código por token
        token_resp = obter_token_do_codigo(codigo)
        
        if token_resp:
            st.session_state["ms_token"] = token_resp
            user_data = obter_info_usuario(token_resp['access_token'])
            
            if user_data:
                st.session_state["user_info"] = {
                    "displayName": user_data.get("displayName", "Usuário"),
                    "userPrincipalName": user_data.get("userPrincipalName", "")
                }
            
            st.rerun() # Recarrega para entrar no app principal
        else:
            st.error("Falha ao validar login. Tente novamente.")
            # Gera novo link para tentar de novo
            url_auth = obter_url_autenticacao()
            if url_auth:
                st.markdown(f'<a href="{url_auth}" target="_self" class="button">Tentar Novamente</a>', unsafe_allow_html=True)

    else:
        # Exibe botão de login inicial
        url_auth = obter_url_autenticacao()
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

    st.stop() # Para a execução aqui até logar