import os
from dotenv import load_dotenv
load_dotenv()
from typing import List, Dict, Any
import services

def montar_corpo_email(dados_empresa: Dict[str, Any], template: str) -> str:
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
    email_user = os.getenv('EMAIL_USER')
    email_password = os.getenv('EMAIL_PASSWORD')
    services._create_outlook_draft(destinatario, assunto, corpo, anexos)
