import pandas as pd
from pathlib import Path
import sys
import pythoncom
from datetime import timedelta, datetime
import re
import streamlit as st
import os
from typing import Dict, List, Any, Optional, Tuple
import json
from jinja2 import Environment, BaseLoader, StrictUndefined, meta
from jinja2.sandbox import SandboxedEnvironment
import logging

try:
    import win32com.client as win32
    WIN32_AVAILABLE = sys.platform == "win32"
except ImportError:
    WIN32_AVAILABLE = False

from config.config import load_configs, MESES, build_report_paths
from utils.security_utils import sanitize_html, sanitize_subject, is_safe_path, within_size_limit

class ReportProcessingError(Exception):
    """Exceção customizada para erros de processamento."""
    pass

# ==============================================================================
# FUNÇÕES AUXILIARES (sem alterações)
# ==============================================================================
def _create_outlook_draft(recipient: str, subject: str, body: str, attachments: List[Path]) -> None:
    if not WIN32_AVAILABLE:
        print("--- MODO DE SIMULAÇÃO ---")
        print(f"PARA: {recipient}")
        print(f"ASSUNTO: {subject}")
        print(f"ANEXOS: {[p.name for p in attachments if p and p.exists()]}")
        print("-------------------------")
        return
    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.HTMLBody = body
        for attachment_path in attachments:
            if attachment_path and attachment_path.exists():
                mail.Attachments.Add(str(attachment_path.resolve()))
            else:
                print(f"AVISO: Anexo não encontrado e não será adicionado: {attachment_path}")
        mail.Display(True)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao interagir com o Outlook: {e}")
    finally:
        pythoncom.CoUninitialize()

def _build_filename(company: str, report_type: str, month: str, year: str) -> str:
    company_clean = str(company).strip()
    company_part = re.sub(r'[\s_-]+', '_', company_clean).upper()
    report_part = str(report_type).upper()
    month_part = str(month).lower()[:3]
    year_part = str(year)[-2:]
    return f"{company_part}_{report_part}_{month_part}_{year_part}.pdf"

def _format_currency(value: Any) -> str:
    try:
        val = float(value)
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "R$ 0,00"

def _format_date(date_value: Any) -> str:
    try:
        if date_value is None or pd.isna(date_value):
            return "Data não informada"
        return pd.to_datetime(date_value).strftime('%d/%m/%Y')
    except (ValueError, TypeError):
        return "Data Inválida"

def _load_excel_data(excel_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    if not Path(excel_path).exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {excel_path}")
    return pd.read_excel(Path(excel_path), sheet_name=sheet_name, header=header_row)

def _find_attachment(pdf_dir: str, filename: str) -> Tuple[Optional[Path], List[str]]:
    warnings = []
    attachment_path = Path(pdf_dir) / filename
    if attachment_path.exists():
        return attachment_path, warnings
    warnings.append(f"Anexo não encontrado no caminho principal: {attachment_path}")
    return None, warnings

# ==============================================================================
# ENGINE DE TEMPLATES (Versão Refatorada)
# ==============================================================================
TEMPLATES_JSON_PATH = "config/email_templates.json"

def load_email_templates() -> Dict[str, Any]:
    try:
        with open(TEMPLATES_JSON_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao carregar {TEMPLATES_JSON_PATH}: {e}")

def render_email_from_template(report_type: str, row: Dict[str, Any], common: Dict[str, Any], cfg: Dict[str, Any], auto_send: bool = False) -> Optional[Dict[str, Any]]:
    templates = load_email_templates()
    if report_type not in templates:
        return None # Retorna None se não houver template; a lógica de fallback será usada

    report_cfg = templates[report_type]
    context = {**row, **common, **cfg}

    # Normaliza variáveis para garantir compatibilidade com os templates
    context.setdefault('empresa', row.get('Empresa'))
    context.setdefault('mes', common.get('month_num'))
    context.setdefault('mesext', common.get('month_long'))
    context.setdefault('ano', common.get('year'))
    context.setdefault('valor', row.get('Valor'))
    context.setdefault('situacao', row.get('Situacao'))
    context.setdefault('TipoAgente', row.get('TipoAgente'))
    
    # Lógica de negócio específica para relatórios
    if report_type == 'SUM001':
        try:
            df_raw = _load_excel_data(cfg['excel_dados'], cfg['sheet_dados'], -1)
            data_debito = df_raw.iloc[23, 0]
            data_credito = df_raw.iloc[23, 1]
        except Exception:
            data_debito, data_credito = None, None
        
        situacao = str(row.get('Situacao', '')).strip()
        
        if situacao == 'Crédito':
            context['texto1'] = "crédito"
            context['data_liquidacao'] = _format_date(data_credito)
            context['texto2'] = "ressaltamos que esse crédito está sujeito ao rateio da eventual inadimplência..."
        elif situacao == 'Débito':
            context['texto1'] = "débito"
            context['data_liquidacao'] = _format_date(data_debito)
            context['texto2'] = "teoricamente a conta possui o saldo necessário..."
        else:
            context['texto1'], context['texto2'] = "transação", "verifique os dados na planilha."
        
        context['valor'] = abs(row.get('Valor', 0))

    # Formatação de valores ANTES de renderizar
    for key in ['ValorLiquidacao', 'ValorLiquidado', 'ValorInadimplencia', 'valor']:
        if key in context and isinstance(context[key], (int, float)):
            context[key] = _format_currency(context[key])

    # Lógica de anexo
    final_attachments, attach_warnings = [], []
    filename = _build_filename(str(row.get('Empresa','')), report_type, str(common.get('month_long','')).upper(), str(common.get('year','')))
    if cfg.get('pdfs_dir'):
        path, warnings = _find_attachment(cfg['pdfs_dir'], filename)
        if path: final_attachments.append(path)
        attach_warnings.extend(warnings)
    
    # Renderização
    body_tpl = ''
    if report_type == 'LFN001': # Lógica para escolher template de crédito/débito
        body_tpl = report_cfg.get('body_html_credit') if str(row.get('Situacao','')).strip() == 'Crédito' else report_cfg.get('body_html_debit','')
    else:
        body_tpl = report_cfg.get('body_html', '')

    subject_tpl = report_cfg.get('subject_template', '')

    def normalize_placeholders(s: str) -> str:
        return re.sub(r"\{(\w+)\}", r"{{ \1 }}", s) if isinstance(s, str) else s

    for key, val in context.items():
        if isinstance(val, str):
            context[key] = val.replace('{', '{{').replace('}', '}}')

    env = Environment(loader=BaseLoader())
    subject = env.from_string(normalize_placeholders(subject_tpl)).render(context)
    body = env.from_string(normalize_placeholders(body_tpl)).render(context)
    
    result = {
        'subject': sanitize_subject(subject),
        'body': sanitize_html(body),
        'attachments': final_attachments
    }

    if auto_send:
        # A assinatura é adicionada aqui para o envio final
        assinatura = f"<br><br><p>Atenciosamente,</p><p><strong>{common['analyst']}</strong></p>"
        result['body'] += assinatura
        _create_outlook_draft(row.get('Email', ''), result['subject'], result['body'], result['attachments'])
    
    return result


# ==============================================================================
# HANDLERS DE E-MAIL (Legado - Mantidos para relatórios não migrados)
# ==============================================================================

def handle_gfn001(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    # (Esta função e as outras handle_... permanecem exatamente como estavam no seu código original)
    warnings = []
    # ... (código completo do seu handle_gfn001 original)
    return {'subject': '...', 'body': '...', 'attachments': []}

def handle_lfres(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    # (código completo do seu handle_lfres original)
    return {'subject': '...', 'body': '...', 'attachments': []}

def handle_lembrete(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    # (código completo do seu handle_lembrete original)
    return {'subject': '...', 'body': '...', 'attachments': []}

def handle_lfrcap(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    # (código completo do seu handle_lfrcap original)
    return {'subject': '...', 'body': '...', 'attachments': []}

def handle_rcap(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    # (código completo do seu handle_rcap original)
    return {'subject': '...', 'body': '...', 'attachments': []}

REPORT_HANDLERS = {
    'GFN001': handle_gfn001,
    'LFRES001': handle_lfres,
    'GFN - LEMBRETE': handle_lembrete,
    'LFRCAP001': handle_lfrcap,
    'RCAP002': handle_rcap,
}

# ==============================================================================
# FUNÇÃO PRINCIPAL DE PROCESSAMENTO
# ==============================================================================
def _load_and_process_data(cfg: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # (Esta função permanece como estava)
    header = int(cfg.get('header_row', 0))
    df_dados = _load_excel_data(cfg['excel_dados'], cfg['sheet_dados'], header)
    df_contatos = _load_excel_data(cfg['excel_contatos'], cfg['sheet_contatos'], 0)
    
    column_mapping = dict(item.split(':') for item in cfg['data_columns'].split(','))
    df_dados.rename(columns=column_mapping, inplace=True)
    df_contatos.rename(columns={'AGENTE': 'Empresa', 'ANALISTA': 'Analista', 'E-MAILS RELATÓRIOS CCEE': 'Email'}, inplace=True)
    
    return df_dados, df_contatos

def process_reports(report_type: str, analyst: str, month: str, year: str) -> List[Dict[str, Any]]:
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    report_paths = build_report_paths(report_type, year, month)
    cfg.update(report_paths)
    
    df_dados, df_contatos = _load_and_process_data(cfg)
    
    df_merged = pd.merge(df_dados, df_contatos, on='Empresa', how='left')
    df_filtered = df_merged[df_merged['Analista'] == analyst].copy()
    
    if df_filtered.empty:
        return []

    df_filtered['Email'] = df_filtered['Email'].fillna('EMAIL_NAO_ENCONTRADO')

    common_data = {
        'analyst': analyst,
        'month_long': month.title(),
        'month_num': {m: f"{i+1:02d}" for i, m in enumerate(MESES)}.get(month.upper()),
        'month': month,
        'year': year
    }

    results, created_count = [], 0
    for _, row in df_filtered.iterrows():
        try:
            row_dict = row.to_dict()
            # Tenta usar o novo sistema de templates primeiro
            email_data = render_email_from_template(report_type, row_dict, common_data, cfg, auto_send=True)

            # Se não houver template, usa o handler legado como fallback
            if not email_data and report_type in REPORT_HANDLERS:
                handler = REPORT_HANDLERS[report_type]
                email_data = handler(row, cfg, common_data)
                if email_data:
                    assinatura = f"<br><br><p>Atenciosamente,</p><p><strong>{analyst}</strong></p>"
                    email_data['body'] += assinatura
                    _create_outlook_draft(row['Email'], **email_data)

            if email_data:
                created_count += 1
                results.append({
                    'empresa': row['Empresa'],
                    'valor': _format_currency(row.get('Valor')),
                    'email': row['Email'], 
                    'anexos_count': len(email_data.get('attachments', [])), 
                    'created_count': created_count
                })
        except Exception as e:
            print(f"Erro ao processar linha para {row.get('Empresa', 'Empresa desconhecida')}: {e}")
            continue
            
    return results

@st.cache_data(show_spinner=True)
def preview_dados(report_type: str, analyst: str, month: str, year: str) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    report_paths = build_report_paths(report_type, year, month)
    cfg.update(report_paths)
    
    df_dados, df_contatos = _load_and_process_data(cfg)
    
    df_merged = pd.merge(df_dados, df_contatos, on='Empresa', how='left')
    df_filtered = df_merged[df_merged['Analista'] == analyst].copy()
    if df_filtered.empty:
        raise ReportProcessingError(f"Nenhum registro encontrado para o analista '{analyst}'.")

    df_filtered['Email'] = df_filtered['Email'].fillna('EMAIL_NAO_ENCONTRADO')
    
    return df_filtered, df_filtered.head(20), cfg