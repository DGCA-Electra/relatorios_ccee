

import pandas as pd
from pathlib import Path
import sys
import pythoncom
from datetime import datetime
import re
import streamlit as st
import os
from typing import Dict, List, Any, Optional, Tuple
import json
from jinja2 import Environment, BaseLoader, meta
import logging

try:
    import win32com.client as win32
    WIN32_AVAILABLE = sys.platform == "win32"
except ImportError:
    WIN32_AVAILABLE = False

from config.config import load_configs, MESES, build_report_paths
from utils.security_utils import sanitize_html, sanitize_subject

class ReportProcessingError(Exception):
    """Exceção customizada para erros de processamento."""
    pass

# --- Funções Auxiliares --- #

def _create_outlook_draft(recipient: str, subject: str, body: str, attachments: List[Path]) -> None:
    if not WIN32_AVAILABLE:
        print("-- MODO DE SIMULAÇÃO ---")
        print(f"PARA: {recipient}")
        print(f"ASSUNTO: {subject}")
        print(f"ANEXOS: {[p.name for p in attachments if p and p.exists()]}")
        print("-------------------------")
        return
    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.HTMLBody = body
        for attachment_path in attachments:
            if attachment_path and attachment_path.exists():
                mail.Attachments.Add(str(attachment_path.resolve()))
            else:
                logging.warning(f"Anexo não encontrado: {attachment_path}")
        mail.Display(True)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao interagir com o Outlook: {e}")
    finally:
        pythoncom.CoUninitialize()

def _build_filename(company: str, report_type: str, month: str, year: str) -> str:
    company_clean = str(company).strip()
    company_part = re.sub(r"[\s_-]+", "_", company_clean).upper()
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
        # Converte para datetime e depois formata
        return pd.to_datetime(date_value).strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return "Data Inválida"

def _load_excel_data(excel_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    if not Path(excel_path).exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {excel_path}")
    # Se header_row for -1, significa que não há cabeçalho e a primeira linha é de dados
    if header_row == -1:
        return pd.read_excel(Path(excel_path), sheet_name=sheet_name, header=None)
    return pd.read_excel(Path(excel_path), sheet_name=sheet_name, header=header_row)

def _find_attachment(pdf_dir: str, filename: str) -> Optional[Path]:
    attachment_path = Path(pdf_dir) / filename
    if attachment_path.exists():
        return attachment_path
    logging.warning(f"Anexo não encontrado no caminho principal: {attachment_path}")
    return None

TEMPLATES_JSON_PATH = "config/email_templates.json"

def load_email_templates() -> Dict[str, Any]:
    try:
        with open(TEMPLATES_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao carregar {TEMPLATES_JSON_PATH}: {e}")

def save_email_templates(data: Dict[str, Any]) -> None:
    try:
        with open(TEMPLATES_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao salvar {TEMPLATES_JSON_PATH}: {e}")

def resolve_variant(report_type: str, report_config: Dict[str, Any], context: Dict[str, Any]) -> Tuple[Dict[str, Any], str]:
    if "variants" not in report_config:
        return report_config, "default"

    variants = report_config["variants"]

    if report_type.startswith("LFRES"):
        try:
            valor = float(context.get("valor", 0))
        except (ValueError, TypeError):
            valor = 0.0
        tipo_agente = str(context.get("TipoAgente", "")).strip()

        if valor == 0:
            if tipo_agente == "Gerador-EER":
                return {}, "SKIP"
            return variants.get("ZERO_VALOR", {}), "ZERO_VALOR"
        
        if tipo_agente == "Gerador-EER":
            return variants.get("COM_VALOR_GERADOR", {}), "COM_VALOR_GERADOR"
        else:
            return variants.get("COM_VALOR_OUTROS", {}), "COM_VALOR_OUTROS"

    first_key = next(iter(variants), "default")
    return variants.get(first_key, report_config), first_key

def render_email_from_template(report_type: str, row: Dict[str, Any], common: Dict[str, Any], cfg: Dict[str, Any], auto_send: bool = False) -> Optional[Dict[str, Any]]:
    templates = load_email_templates()
    
    # 1. Identificação do Template
    template_key = "LFRES" if report_type.startswith("LFRES") else report_type
    if template_key not in templates:
        raise ReportProcessingError(f"Template para \'{template_key}\' não encontrado.")
    report_cfg = templates[template_key]

    # 2. Construção do Contexto Inicial
    context = {**row, **common, **cfg}
    
    # 3. Normalize as chaves
    context["empresa"] = row.get("Empresa")
    context["mesext"] = common.get("month_long")
    context["mes"] = common.get("month_num")
    context["ano"] = common.get("year")
    # Garante que 'data' e 'valor' estejam no contexto inicial, mesmo que None
    context["data"] = row.get("Data")
    context["valor"] = row.get("Valor") # Valor bruto

    # 4. Execução de Lógicas Específicas (ANTES da formatação)
    if template_key.startswith("LFRES"):
        try:
            # Carrega a planilha sem cabeçalho para acessar posições fixas
            df_raw_date = _load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], -1)
            # As células A27 e B27 correspondem aos índices [26, 0] e [26, 1] em pandas (base 0)
            data_debito = df_raw_date.iloc[26, 0]
            data_credito = df_raw_date.iloc[26, 1]
            
            situacao = str(row.get("Situacao", "")).strip()
            if situacao == "Crédito":
                context["data"] = data_credito
            else: # Débito ou outro
                context["data"] = data_debito

            # Adiciona o valor de liquidação ao contexto, se disponível
            # O VBA usa Cells(i, 4).Value para o valor, que é o 'Valor' da row
            context["ValorLiquidacao"] = row.get("Valor") # Usar o valor bruto da row

        except Exception as e:
            logging.warning(f"LFRES: Não foi possível extrair a data ou valor do cabeçalho. Erro: {e}")
            context["data"] = None
            context["ValorLiquidacao"] = None

    # Adiciona os valores de liquidação ao contexto para LFN001
    if report_type == "LFN001":
        context["ValorLiquidacao"] = row.get("ValorLiquidacao")
        context["ValorLiquidado"] = row.get("ValorLiquidado")
        context["ValorInadimplencia"] = row.get("ValorInadimplencia")

    if report_type == "SUM001":
        try:
            df_raw = _load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], -1)
            data_debito, data_credito = df_raw.iloc[23, 0], df_raw.iloc[23, 1]
        except Exception:
            data_debito, data_credito = None, None
        
        situacao = str(row.get("Situacao", "")).strip()
        if situacao == "Crédito":
            context["texto1"] = "crédito"
            context["texto2"] = "ressaltamos que esse crédito está sujeito ao rateio..."
            context["data"] = data_credito
        elif situacao == "Débito":
            context["texto1"] = "débito"
            context["texto2"] = "teoricamente a conta possui o saldo necessário..."
            context["data"] = data_debito
        else:
            context["texto1"], context["texto2"] = "transação", "verifique os dados na planilha."
        context["valor"] = abs(row.get("Valor", 0))

    # 5. Seleção da Variante do Template
    selected_template, variant_name = resolve_variant(template_key, report_cfg, context)
    if variant_name == "SKIP":
        return None

    # 6. Formatação dos Dados para Exibição (APÓS a lógica)
    # Garante que 'valor' e 'ValorLiquidacao' sejam formatados se existirem
    for key in ["valor", "ValorLiquidacao", "ValorLiquidado", "ValorInadimplencia"]:
        if key in context and context[key] is not None:
            context[key] = _format_currency(context[key])
    if "data" in context and context["data"] is not None:
        context["data"] = _format_date(context["data"])

    # 7. Construção dos Anexos
    attachments = []
    filename = _build_filename(str(row.get("Empresa","")), report_type, common["month_long"].upper(), str(common.get("year","")))
    if cfg.get("pdfs_dir"):
        path = _find_attachment(cfg["pdfs_dir"], filename)
        if path: attachments.append(path)

    if report_type == "GFN001":
        filename_sum = _build_filename(str(row.get("Empresa","")), "SUM001", common["month_long"].upper(), str(common.get("year","")))
        base_dir = Path(cfg.get("pdfs_dir", ""))
        memoria_calc_dir = base_dir.parent.parent / "Sumário" / "SUM001 - Memória_de_Cálculo"
        sum_path = _find_attachment(str(memoria_calc_dir), filename_sum)
        if sum_path: attachments.append(sum_path)

    # 8. Renderização Final
    subject_tpl = selected_template.get("subject_template", "")
    body_tpl = selected_template.get("body_html", "")

    if report_type == "LFN001":
        body_tpl = selected_template.get("body_html_credit") if str(row.get("Situacao","")).strip() == "Crédito" else selected_template.get("body_html_debit", "")

    env = Environment(loader=BaseLoader())
    def normalize(s: str): return re.sub(r"\{(\w+)\}", r"{{ \1 }}", s) if isinstance(s, str) else s
    
    # Preenche placeholders não encontrados para evitar erros
    for k in meta.find_undeclared_variables(env.parse(normalize(body_tpl))): context.setdefault(k, f"[{k} N/D]")

    subject = env.from_string(normalize(subject_tpl)).render(context)
    body = env.from_string(normalize(body_tpl)).render(context)
    
    result = {"subject": sanitize_subject(subject), "body": sanitize_html(body), "attachments": attachments}

    if auto_send:
        result["body"] += f"  <p>Atenciosamente,</p><p><strong>{common["analyst"]}</strong></p>"
        _create_outlook_draft(row.get("Email", ""), result["subject"], result["body"], result["attachments"])
    
    return result

# --- Funções de Processamento de Relatórios ---

def _load_and_process_data(cfg: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    header = int(cfg.get("header_row", 0))
    df_dados = _load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], header)
    df_contatos = _load_excel_data(cfg["excel_contatos"], cfg["sheet_contatos"], 0)
    
    column_mapping = dict(item.split(":") for item in cfg["data_columns"].split(","))
    df_dados.rename(columns=column_mapping, inplace=True)
    df_contatos.rename(columns={
        "AGENTE": "Empresa", 
        "ANALISTA": "Analista", 
        "E-MAILS RELATÓRIOS CCEE": "Email"
    }, inplace=True)
    return df_dados, df_contatos

def process_reports(report_type: str, analyst: str, month: str, year: str) -> List[Dict[str, Any]]:
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    cfg.update(build_report_paths(report_type, year, month))
    
    df_dados, df_contatos = _load_and_process_data(cfg)
    
    df_merged = pd.merge(df_dados, df_contatos, on="Empresa", how="left")
    df_filtered = df_merged[df_merged["Analista"] == analyst].copy()
    
    if df_filtered.empty: return []

    df_filtered["Email"] = df_filtered["Email"].fillna("EMAIL_NAO_ENCONTRADO")

    common_data = {
        "analyst": analyst,
        "month_long": month.title(),
        "month_num": {m: f"{i+1:02d}" for i, m in enumerate(MESES)}.get(month.upper()),
        "year": year
    }

    results, created_count = [], 0
    for _, row in df_filtered.iterrows():
        try:
            email_data = render_email_from_template(report_type, row.to_dict(), common_data, cfg, auto_send=True)
            if email_data:
                created_count += 1
                results.append({
                    "empresa": row["Empresa"],
                    "valor": _format_currency(row.get("Valor")),
                    "email": row["Email"], 
                    "anexos_count": len(email_data.get("attachments", [])), 
                    "created_count": created_count
                })
        except Exception as e:
            logging.error(f"Erro ao processar linha para {row.get("Empresa", "Empresa desconhecida")}: {e}")
            continue
            
    return results

@st.cache_data(show_spinner=False)
def preview_dados(report_type: str, analyst: str, month: str, year: str) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    if not cfg:
        raise ReportProcessingError(f"\'{report_type}\' não encontrado nas configurações.")

    report_paths = build_report_paths(report_type, year, month)
    cfg.update(report_paths)
    
    df_dados, df_contatos = _load_and_process_data(cfg)
    
    df_merged = pd.merge(df_dados, df_contatos, on="Empresa", how="left")
    df_filtered = df_merged[df_merged["Analista"] == analyst].copy()
    if df_filtered.empty:
        raise ReportProcessingError(f"Nenhum registro encontrado para o analista \'{analyst}\'")

    df_filtered["Email"] = df_filtered["Email"].fillna("EMAIL_NAO_ENCONTRADO")
    
    return df_filtered, cfg