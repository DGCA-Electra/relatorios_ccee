import streamlit as st
import msal
import requests
import os
from urllib.parse import urlparse, parse_qs
import logging

CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
TENANT_ID = os.environ.get("AZURE_TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = os.environ.get("AZURE_REDIRECT_URI", "http://localhost:8501") 
SCOPES = ["User.Read", "Mail.ReadWrite"]

if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID, REDIRECT_URI]):
    st.error("Erro Crítico: Variáveis de configuração do Azure AD não encontradas no ambiente. Verifique o arquivo .env ou as configurações do servidor.")
    st.stop()

msal_app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)

def get_auth_url():
    try:
        auth_url = msal_app.get_authorization_request_url(
            SCOPES,
            redirect_uri=REDIRECT_URI,
            response_type='code'
        )
        return auth_url
    except Exception as e:
        logging.error(f"Erro ao gerar URL de autenticação: {e}")
        st.error("Ocorreu um erro ao preparar o login. Verifique as configurações.")
        return None

def get_token_from_code(auth_code):
    """Troca o código de autorização por um token de acesso."""
    result = None
    try:
        result = msal_app.acquire_token_by_authorization_code(
            auth_code,
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI
        )
        if "error" in result:
            logging.error(f"Erro MSAL ao obter token: {result.get('error_description')}")
            st.error(f"Erro ao obter token: {result.get('error_description')}")
            return None
        if "access_token" not in result:
             logging.error(f"Resposta inesperada do MSAL (sem access_token): {result}")
             st.error("Falha ao obter o token de acesso da Microsoft.")
             return None
        return result
    except Exception as e:
        logging.error(f"Exceção ao adquirir token: {e}")
        st.error(f"Falha na comunicação com o serviço de autenticação: {e}")
        return None

def get_user_info(access_token):
    """Busca informações básicas do usuário logado usando o token."""
    headers = {'Authorization': 'Bearer ' + access_token}
    try:
        response = requests.get('https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName', headers=headers)
        response.raise_for_status()
        user_data = response.json()
        return user_data
    except requests.exceptions.RequestException as e:
        logging.error(f"Erro ao buscar informações do usuário na API Graph: {e}")
        st.error("Não foi possível buscar as informações do usuário.")
        return None

def show_login_page():
    """Renderiza a página de login e processa o callback."""

    # st.image("static/logo.png", width=250)
    st.title("Login - Envio de Relatórios CCEE")
    st.write("Por favor, autentique-se com sua conta Microsoft para continuar.")

    query_params = st.query_params
    auth_code = query_params.get("code")

    if auth_code:
        st.query_params["code"] = ""

        token_response = get_token_from_code(auth_code)

        if token_response:
            st.session_state["ms_token"] = token_response

            user_data = get_user_info(token_response['access_token'])
            if user_data:
                st.session_state["user_info"] = {
                    "displayName": user_data.get("displayName", "Usuário Desconhecido"),
                    "userPrincipalName": user_data.get("userPrincipalName", "")
                }
            else:
                 st.session_state["user_info"] = {"displayName": "Usuário (Erro Info)", "userPrincipalName": ""}

            logging.info(f"Login bem-sucedido para: {st.session_state['user_info'].get('userPrincipalName')}")
            st.rerun()
        else:
            auth_url = get_auth_url()
            if auth_url:
                st.markdown(f'<a href="{auth_url}" target="_self" class="button">Entrar com Microsoft</a>', unsafe_allow_html=True)
            st.error("Falha ao processar o login. Tente novamente.")
    else:
        auth_url = get_auth_url()
        if auth_url:
            st.markdown("""
        <style>
        .button { /* Estilos gerais do botão (fundo, padding, etc.) */
            display: inline-block;
            padding: 0.5em 1em;
            /* A cor do texto será definida na regra a.button abaixo */
            background-color: #0078D4; /* Cor de fundo azul */
            border: none;
            border-radius: 4px;
            text-align: center;
            text-decoration: none; /* Remove sublinhado do link */
            font-size: 16px;
            cursor: pointer;
            line-height: normal; /* Garante alinhamento vertical */
        }
        a.button { /* Estilo específico para o link QUE TEM a classe button */
             color: white !important; /* COR PADRÃO DO TEXTO - altere para 'red' se quiser */
             text-decoration: none !important; /* Garante que não haja sublinhado */
        }
        /* Não precisa mais de display: block; aqui pois o 'a' já é o botão */

        a.button:hover { /* Estilo do link/botão ao passar o mouse */
             background-color: #005A9E; /* Fundo mais escuro */
             color: white !important; /* Cor do texto no hover (pode manter branco) */
             text-decoration: none !important;
        }
        /* Removidas as regras separadas .button:hover e .button:hover a */
        </style>
        """, unsafe_allow_html=True)
            st.markdown(f'<a href="{auth_url}" target="_self" class="button">Entrar com Microsoft</a>', unsafe_allow_html=True)
        else:
            st.error("Não foi possível gerar o link de login. Verifique as configurações da aplicação.")

    st.stop()