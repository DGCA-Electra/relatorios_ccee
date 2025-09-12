import os
from dotenv import load_dotenv
load_dotenv()
"""
mail/email_utils.py
Funções utilitárias para montagem e envio de e-mails.
"""
from typing import List, Dict, Any
import services

def montar_corpo_email(dados_empresa: Dict[str, Any], template: str) -> str:
    """
    Monta o corpo do e-mail usando os dados da empresa e um template HTML.
    """
    # Carregar template HTML
    from jinja2 import Template as JinjaTemplate
    campos = {
        'analista': dados_empresa.get('Analista', ''),
        'empresa': dados_empresa.get('Empresa', ''),
        'tipo_agente': dados_empresa.get('TipoAgente', ''),
        'valor': dados_empresa.get('Valor', ''),
        'data': dados_empresa.get('Data', ''),
        'situacao': dados_empresa.get('Situacao', ''),
        'email': dados_empresa.get('Email', '')
    }
    corpo = JinjaTemplate(template).render(**campos)
    return corpo

def enviar_email(destinatario: str, assunto: str, corpo: str, anexos: List[str]) -> None:
    """
    Envia o e-mail usando a função do services (Outlook).
    """
    # Exemplo de uso de credenciais (caso precise para SMTP ou API)
    email_user = os.getenv('EMAIL_USER')
    email_password = os.getenv('EMAIL_PASSWORD')
    # Aqui você pode usar email_user/email_password para autenticação segura
    services._create_outlook_draft(destinatario, assunto, corpo, anexos)
