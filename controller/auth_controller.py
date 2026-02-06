import os
import logging
import msal
import requests
import streamlit as st
from dotenv import load_dotenv
from apps.relatorios_ccee.model.arquivos import obtem_asset_path

# Load .env from same folder as this file
current_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(current_dir, ".env")
load_dotenv(env_path)

CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
TENANT_ID = os.environ.get("AZURE_TENANT_ID")
REDIRECT_URI = os.environ.get("AZURE_REDIRECT_URI")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Mail.Send"]

# Lazily create MSAL app
_msal_app = None

def _get_msal_app():
    global _msal_app
    if _msal_app is None:
        _msal_app = msal.ConfidentialClientApplication(
            CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
        )
    return _msal_app


def obter_url_autenticacao() -> str:
    """Retorna a URL de autenticação para redirecionar o usuário."""
    try:
        app = _get_msal_app()
        url = app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI, response_type='code')
        return url
    except Exception as e:
        logging.error(f"Erro ao gerar URL de autenticação: {e}")
        return ""


def obter_token_do_codigo(codigo_autorizacao: str) -> dict:
    """Troca o código de autorização por um token MSAL.

    Returns:
        dict: objeto de token MSAL ou levanta exceção em caso de falha.
    """
    try:
        app = _get_msal_app()
        resultado = app.acquire_token_by_authorization_code(
            codigo_autorizacao,
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI
        )
        if "error" in resultado:
            raise Exception(resultado.get("error_description") or resultado.get("error"))
        return resultado
    except Exception as e:
        logging.error(f"Falha ao adquirir token por código: {e}")
        raise


def obter_info_usuario(token_acesso: str) -> dict:
    """Obtém informações do usuário via Graph API.

    Args:
        token_acesso: Token de acesso (string).

    Returns:
        dict: Informações do usuário obtidas da API Graph.
    """
    headers = {'Authorization': 'Bearer ' + token_acesso}
    try:
        resposta = requests.get('https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName', headers=headers)
        resposta.raise_for_status()
        return resposta.json()
    except Exception as e:
        logging.error(f"Erro ao obter informações do usuário via Graph: {e}")
        raise


def processar_callback(codigo: str) -> None:
    """Processa o callback de autenticação: obtém token e popula `st.session_state`.

    Levanta `Exception` em caso de falha.
    """
    token_resp = obter_token_do_codigo(codigo)
    st.session_state["ms_token"] = token_resp
    user_data = obter_info_usuario(token_resp['access_token'])
    st.session_state["user_info"] = {
        "displayName": user_data.get("displayName", "Usuário"),
        "userPrincipalName": user_data.get("userPrincipalName", "")
    }
    # Força recarregamento da página
    st.rerun()


def logout():
    """Limpa a sessão do Streamlit referente à autenticação."""
    for k in ("ms_token", "user_info"):
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()
