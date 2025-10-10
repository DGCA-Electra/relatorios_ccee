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
# FUNÇÕES AUXILIARES
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
# MOTOR DE TEMPLATES E LÓGICA DE E-MAIL
# ==============================================================================
TEMPLATES_JSON_PATH = "config/email_templates.json"

def load_email_templates() -> Dict[str, Any]:
    try:
        with open(TEMPLATES_JSON_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao carregar {TEMPLATES_JSON_PATH}: {e}")

def render_email_from_template(report_type: str, row: Dict[str, Any], common: Dict[str, Any], cfg: Dict[str, Any], auto_send: bool = False) -> Dict[str, Any]:
    templates = load_email_templates()
    if report_type not in templates:
        raise ReportProcessingError(f"Template para '{report_type}' não encontrado em {TEMPLATES_JSON_PATH}")

    report_cfg = templates[report_type]
    context = {**row, **common, **cfg}

    # Normaliza variáveis para garantir compatibilidade com os templates
    context.setdefault('empresa', row.get('Empresa'))
    context.setdefault('mesext', common.get('month_long'))
    context.setdefault('mes', common.get('month_num'))
    context.setdefault('ano', common.get('year'))

    # Lógica de negócio específica para relatórios
    if report_type == 'SUM001':
        try:
            df_raw = _load_excel_data(cfg['excel_dados'], cfg['sheet_dados'], -1)
            data_debito, data_credito = df_raw.iloc[23, 0], df_raw.iloc[23, 1]
        except Exception:
            data_debito, data_credito = None, None
        
        situacao = str(row.get('Situacao', '')).strip()
        data_liquidacao = datetime.now().strftime('%d/%m/%Y')

        if situacao == 'Crédito':
            context['texto1'], data_liquidacao = "crédito", _format_date(data_credito)
            context['texto2'] = "ressaltamos que esse crédito está sujeito ao rateio..."
        elif situacao == 'Débito':
            context['texto1'], data_liquidacao = "débito", _format_date(data_debito)
            context['texto2'] = "teoricamente a conta possui o saldo necessário..."
        else:
            context['texto1'], context['texto2'] = "transação", "verifique os dados na planilha."
        
        context['data_liquidacao'] = data_liquidacao
        context['valor'] = abs(row.get('Valor', 0))

    # Formatação de valores ANTES de renderizar
    for key in ['ValorLiquidacao', 'ValorLiquidado', 'ValorInadimplencia', 'valor']:
        if key in context and isinstance(context[key], (int, float)):
            context[key] = _format_currency(context[key])

    # --- LÓGICA DE ANEXOS ---
    final_attachments, attach_warnings = [], []
    filename = _build_filename(str(row.get('Empresa','')), report_type, common['month_long'].upper(), str(common.get('year','')))
    if cfg.get('pdfs_dir'):
        path, warnings = _find_attachment(cfg['pdfs_dir'], filename)
        if path: final_attachments.append(path)
        attach_warnings.extend(warnings)
    
    # LÓGICA RESTAURADA: Anexo secundário para GFN001
    if report_type == 'GFN001':
        filename_sum = _build_filename(str(row.get('Empresa','')), 'SUM001', common['month_long'].upper(), str(common.get('year','')))
        base_dir = Path(cfg.get('pdfs_dir', ''))
        memoria_calc_dir = base_dir.parent.parent / 'Sumário' / 'SUM001 - Memória_de_Cálculo'
        sum_path, sum_warnings = _find_attachment(str(memoria_calc_dir), filename_sum)
        if sum_path: final_attachments.append(sum_path)
        attach_warnings.extend(sum_warnings)
    
    # --- RENDERIZAÇÃO DO TEMPLATE ---
    selected, variant_name = resolve_variant(report_type, report_cfg, context)
    if variant_name == 'SKIP':
        return None # Retorna None para pular a criação do e-mail

    subject_tpl = selected.get('subject_template', '')
    body_tpl = selected.get('body_html', '')

    # Lógica para LFN001 continua necessária aqui
    if report_type == 'LFN001' and str(row.get('Situacao','')).strip() == 'Crédito':
        body_tpl = selected.get('body_html_credit', body_tpl)
    elif report_type == 'LFN001':
        body_tpl = selected.get('body_html_debit', body_tpl)

    def normalize(s: str): return re.sub(r"\{(\w+)\}", r"{{ \1 }}", s) if isinstance(s, str) else s
    for k in meta.find_undeclared_variables(Environment().parse(normalize(body_tpl))): context.setdefault(k, f'[{k} N/D]')

    env = Environment(loader=BaseLoader())
    subject = env.from_string(normalize(subject_tpl)).render(context)
    body = env.from_string(normalize(body_tpl)).render(context)
    
    result = {'subject': sanitize_subject(subject), 'body': sanitize_html(body), 'attachments': final_attachments}

    if auto_send:
        result['body'] += f"<br><br><p>Atenciosamente,</p><p><strong>{common['analyst']}</strong></p>"
        _create_outlook_draft(row.get('Email', ''), result['subject'], result['body'], result['attachments'])
    
    return result

# ==============================================================================
# FUNÇÃO PRINCIPAL DE PROCESSAMENTO
# ==============================================================================
def _load_and_process_data(cfg: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
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
    cfg.update(build_report_paths(report_type, year, month))
    
    df_dados, df_contatos = _load_and_process_data(cfg)
    
    df_merged = pd.merge(df_dados, df_contatos, on='Empresa', how='left')
    df_filtered = df_merged[df_merged['Analista'] == analyst].copy()
    
    if df_filtered.empty: return []

    df_filtered['Email'] = df_filtered['Email'].fillna('EMAIL_NAO_ENCONTRADO')

    common_data = {
        'analyst': analyst,
        'month_long': month.title(),
        'month_num': {m: f"{i+1:02d}" for i, m in enumerate(MESES)}.get(month.upper()),
        'month': month, 'year': year
    }

    results, created_count = [], 0
    for _, row in df_filtered.iterrows():
        try:
            email_data = render_email_from_template(report_type, row.to_dict(), common_data, cfg, auto_send=True)
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

def resolve_variant(report_type: str, report_config: Dict[str, Any], context: Dict[str, Any]) -> Tuple[Dict[str, Any], str]:
    """Seleciona a variante correta baseada nas regras de negócio."""

    if 'variants' not in report_config:
        return report_config, 'default'

    variants = report_config['variants']

    # Lógica específica para LFRES
    if report_type == 'LFRES':
        # Converte 'valor' de string formatada para float para a lógica
        try:
            # Remove o R$, o ponto de milhar e troca a vírgula do decimal por ponto
            valor_str = str(context.get('valor', '0')).replace('R$', '').replace('.', '').replace(',', '.').strip()
            valor = float(valor_str)
        except (ValueError, TypeError):
            valor = 0.0

        tipo_agente = str(context.get('TipoAgente', '')).strip()

        # Cenário 3: Valor é zero
        if valor == 0:
            # Se for Gerador-EER com valor zero, não envia e-mail
            if tipo_agente == 'Gerador-EER':
                return {}, 'SKIP' # Retorna um sinal para pular o e-mail
            return variants.get('ZERO_VALOR', {}), 'ZERO_VALOR'

        # Cenário 1: Valor > 0 e é Gerador-EER
        if tipo_agente == 'Gerador-EER':
            return variants.get('COM_VALOR_GERADOR', {}), 'COM_VALOR_GERADOR'

        # Cenário 2: Valor > 0 e não é Gerador-EER
        else:
            return variants.get('COM_VALOR_OUTROS', {}), 'COM_VALOR_OUTROS'

    # Fallback para outros relatórios com variantes (se houver no futuro)
    first_key = next(iter(variants))
    return variants[first_key], first_key

@st.cache_data(show_spinner=True)
def preview_dados(report_type: str, analyst: str, month: str, year: str) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    if not cfg:
        raise ReportProcessingError(f"'{report_type}' não encontrado nas configurações.")

    report_paths = build_report_paths(report_type, year, month)
    cfg.update(report_paths)
    
    df_dados, df_contatos = _load_and_process_data(cfg)
    
    df_merged = pd.merge(df_dados, df_contatos, on='Empresa', how='left')
    df_filtered = df_merged[df_merged['Analista'] == analyst].copy()
    if df_filtered.empty:
        raise ReportProcessingError(f"Nenhum registro encontrado para o analista '{analyst}'.")

    df_filtered['Email'] = df_filtered['Email'].fillna('EMAIL_NAO_ENCONTRADO')
    
    return df_filtered, cfg