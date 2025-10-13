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

def _parse_brazilian_number(val: Any) -> float:
    """Converte 'R$ 1.234,56' ou '(1.234,56)' ou 1234.56 para float. Retorna 0.0 em erro."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s == "":
        return 0.0
    is_neg = False
    if s.startswith("(") and s.endswith(")"):
        is_neg = True
        s = s[1:-1]
    s = s.replace("R$", "").replace("r$", "").replace("\xa0", "").replace(" ", "")
    s = s.replace(".", "")
    s = s.replace(",", ".")
    s = re.sub(r"[^0-9\.-]", "", s)
    try:
        n = float(s) if s not in ("", "-", ".") else 0.0
        return -n if is_neg else n
    except Exception:
        return 0.0


def _create_outlook_draft(recipient: str, subject: str, body: str, attachments: List[Path]) -> None:
    """Cria um rascunho de e-mail no Outlook."""
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
    """Constrói o nome do arquivo PDF baseado nos dados da empresa."""
    company_clean = str(company).strip()
    company_part = re.sub(r"[\s_-]+", "_", company_clean).upper()
    report_part = str(report_type).upper()
    month_part = str(month).lower()[:3]
    year_part = str(year)[-2:]
    return f"{company_part}_{report_part}_{month_part}_{year_part}.pdf"


def _format_currency(value: Any) -> str:
    """Formata um valor numérico para formato de moeda brasileira."""
    try:
        val = float(value)
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "R$ 0,00"


def _format_date(date_value: Any) -> str:
    """Formata uma data para o formato brasileiro (dd/mm/aaaa)."""
    try:
        if date_value is None or pd.isna(date_value):
            return "Data não informada"
        return pd.to_datetime(date_value).strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return "Data Inválida"


def _load_excel_data(excel_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    """Carrega dados de uma planilha Excel."""
    if not Path(excel_path).exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {excel_path}")
    if header_row == -1:
        return pd.read_excel(Path(excel_path), sheet_name=sheet_name, header=None)
    return pd.read_excel(Path(excel_path), sheet_name=sheet_name, header=header_row)


def _find_attachment(pdf_dir: str, filename: str) -> Optional[Path]:
    """Procura por um arquivo PDF no diretório especificado."""
    attachment_path = Path(pdf_dir) / filename
    if attachment_path.exists():
        return attachment_path
    logging.warning(f"Anexo não encontrado no caminho principal: {attachment_path}")
    return None


TEMPLATES_JSON_PATH = "config/email_templates.json"


def load_email_templates() -> Dict[str, Any]:
    """Carrega os templates de e-mail do arquivo JSON."""
    try:
        with open(TEMPLATES_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao carregar {TEMPLATES_JSON_PATH}: {e}")


def save_email_templates(data: Dict[str, Any]) -> None:
    """Salva os templates de e-mail no arquivo JSON."""
    try:
        with open(TEMPLATES_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao salvar {TEMPLATES_JSON_PATH}: {e}")


def resolve_variant(report_type: str, report_config: Dict[str, Any], context: Dict[str, Any]) -> Tuple[Dict[str, Any], str]:
    """
    Resolve qual variante do template usar baseado no tipo de relatório e contexto.
    Retorna (template_selecionado, nome_da_variante).
    """
    if "variants" not in report_config:
        return report_config, "default"

    variants = report_config["variants"]

    if report_type.startswith("LFRES"):
        raw_val = context.get("valor", 0.0)
        try:
            valor = float(raw_val)
        except (ValueError, TypeError):
            try:
                valor = _parse_brazilian_number(raw_val)
            except Exception:
                valor = 0.0

        valor_abs = abs(valor)
        tipo_agente = str(context.get("TipoAgente", "")).strip()

        logging.info(f"Resolve_variant LFRES - empresa={context.get('empresa')}, raw_val={raw_val}, valor={valor}, tipo_agente={tipo_agente}")

        if valor_abs > 1e-6:
            if tipo_agente == "Gerador-EER":
                logging.info("Selecionado: COM_VALOR_GERADOR")
                return variants.get("COM_VALOR_GERADOR", {}), "COM_VALOR_GERADOR"
            logging.info("Selecionado: COM_VALOR_OUTROS")
            return variants.get("COM_VALOR_OUTROS", {}), "COM_VALOR_OUTROS"

        if tipo_agente == "Gerador-EER":
            logging.info("Selecionado: SKIP (Gerador-EER com valor 0)")
            return {}, "SKIP"

        logging.info("Selecionado: ZERO_VALOR")
        return variants.get("ZERO_VALOR", {}), "ZERO_VALOR"

    first_key = next(iter(variants), "default")
    return variants.get(first_key, report_config), first_key


def render_email_from_template(report_type: str, row: Dict[str, Any], common: Dict[str, Any], cfg: Dict[str, Any], auto_send: bool = False) -> Optional[Dict[str, Any]]:
    templates = load_email_templates()
    
    template_key = "LFRES" if report_type.startswith("LFRES") else report_type
    if template_key not in templates:
        raise ReportProcessingError(f"Template para '{template_key}' não encontrado.")
    report_cfg = templates[template_key]

    context = {**row, **common, **cfg}
    
    context["empresa"] = row.get("Empresa")
    context["mesext"] = common.get("month_long")
    context["mes"] = common.get("month_num")
    context["ano"] = common.get("year")
    context["data"] = row.get("Data")
    
    raw_valor = row.get("Valor", 0)
    parsed_valor = _parse_brazilian_number(raw_valor)
    context["valor"] = parsed_valor
    
    logging.info(f"Processando {context.get('empresa')} - Tipo: {report_type} - Valor original: '{raw_valor}' -> Parseado: {parsed_valor}")

    tipo_agente = str(row.get("TipoAgente", "")).strip()
    
    if template_key.startswith("LFRES"):
        situacao = str(row.get("Situacao", "")).strip()

        data_linha = row.get("Data")
        if data_linha is not None and not pd.isna(data_linha) and str(data_linha).strip() != "":
            context["data"] = data_linha
        else:
            try:
                df_raw_data_lfres = _load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], -1)
                data_debito = df_raw_data_lfres.iloc[26, 0]
                data_credito = df_raw_data_lfres.iloc[26, 1]
                context["data"] = data_credito if situacao == "Crédito" else data_debito
            except Exception as e:
                logging.warning(f"LFRES: Não foi possível extrair a data do Excel: {e}")
                context["data"] = None

        context["TipoAgente"] = tipo_agente
        
        logging.info(f"LFRES: TipoAgente='{tipo_agente}', Valor={parsed_valor}, Situacao='{situacao}'")

    if report_type == "LFN001":
        context["ValorLiquidacao"] = _parse_brazilian_number(row.get("ValorLiquidacao", 0))
        context["ValorLiquidado"] = _parse_brazilian_number(row.get("ValorLiquidado", 0))
        context["ValorInadimplencia"] = _parse_brazilian_number(row.get("ValorInadimplencia", 0))

    if report_type == "SUM001":
        try:
            df_raw_sum = _load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], -1)
            data_debito, data_credito = df_raw_sum.iloc[23, 0], df_raw_sum.iloc[23, 1]
        except Exception:
            data_debito, data_credito = None, None
        
        situacao = str(row.get("Situacao", "")).strip()
        if situacao == "Crédito":
            context["texto1"] = "crédito"
            context["texto2"] = "ressaltamos que esse crédito está sujeito ao rateio de inadimplência dos agentes devedores da Câmara, conforme Resolução ANEEL nº 552, de 14/10/2002."
            context["data"] = data_credito
        elif situacao == "Débito":
            context["texto1"] = "débito"
            context["texto2"] = "teoricamente a conta possui o saldo necessário, mas recomendamos verificar e disponibilizar o valor com 1 (um) dia útil de antecedência."
            context["data"] = data_debito
        else:
            context["texto1"], context["texto2"] = "transação", "verifique os dados na planilha."
        context["valor"] = abs(parsed_valor)

    selected_template, variant_name = resolve_variant(template_key, report_cfg, context)
    
    logging.info(f"Variante selecionada: {variant_name}")
    
    if variant_name == "SKIP":
        logging.info(f"Pulando {context.get('empresa')} (Gerador-EER com valor zero)")
        return None

    for key in ["valor", "ValorLiquidacao", "ValorLiquidado", "ValorInadimplencia"]:
        if key in context and context[key] is not None:
            context[key] = _format_currency(context[key])
    
    if "data" in context and context["data"] is not None:
        context["data"] = _format_date(context["data"])

    attachments = []
    filename = _build_filename(str(row.get("Empresa","")), report_type, common["month_long"].upper(), str(common.get("year","")))
    if cfg.get("pdfs_dir"):
        path = _find_attachment(cfg["pdfs_dir"], filename)
        if path: 
            attachments.append(path)
            logging.info(f"Anexo encontrado: {filename}")
        else:
            logging.warning(f"Anexo não encontrado: {filename}")

    if report_type == "GFN001":
        filename_sum = _build_filename(str(row.get("Empresa","")), "SUM001", common["month_long"].upper(), str(common.get("year","")))
        base_dir = Path(cfg.get("pdfs_dir", ""))
        memoria_calc_dir = base_dir.parent.parent / "Sumário" / "SUM001 - Memória_de_Cálculo"
        sum_path = _find_attachment(str(memoria_calc_dir), filename_sum)
        if sum_path: 
            attachments.append(sum_path)
            logging.info(f"Anexo SUM001 encontrado: {filename_sum}")

    subject_tpl = selected_template.get("subject_template", "")
    body_tpl = selected_template.get("body_html", "")

    if report_type == "LFN001":
        situacao_lfn = str(row.get("Situacao","")).strip()
        if situacao_lfn == "Crédito":
            body_tpl = selected_template.get("body_html_credit", body_tpl)
        else:
            body_tpl = selected_template.get("body_html_debit", body_tpl)

    env = Environment(loader=BaseLoader())
    def normalize(s: str): 
        """Normaliza placeholders {var} para {{ var }}"""
        return re.sub(r"\{(\w+)\}", r"{{ \1 }}", s) if isinstance(s, str) else s
    
    parsed_body = env.parse(normalize(body_tpl))
    undeclared_vars = meta.find_undeclared_variables(parsed_body)
    for k in undeclared_vars: 
        if k not in context:
            context[k] = f"[{k} N/D]"
            logging.warning(f"   ⚠️  Placeholder não encontrado: {k}")

    subject = env.from_string(normalize(subject_tpl)).render(context)
    body = env.from_string(normalize(body_tpl)).render(context)
    
    result = {
        "subject": sanitize_subject(subject), 
        "body": sanitize_html(body), 
        "attachments": attachments,
        "missing_placeholders": list(undeclared_vars),
        "attachment_warnings": []
    }

    if auto_send:
        result["body"] += f"<p>Atenciosamente,</p><p><strong>{common['analyst']}</strong></p>"
        _create_outlook_draft(row.get("Email", ""), result["subject"], result["body"], result["attachments"])
    
    return result

def _load_and_process_data(cfg: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Carrega e processa os dados do Excel (dados e contatos).
    
    Args:
        cfg: Configurações do relatório
        
    Returns:
        Tupla com (df_dados, df_contatos)
    """
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
    """
    Processa relatórios e cria rascunhos de e-mail no Outlook.
    
    Args:
        report_type: Tipo do relatório
        analyst: Nome do analista
        month: Mês (nome por extenso)
        year: Ano
        
    Returns:
        Lista de dicionários com resultados do processamento
    """
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    if not cfg:
        raise ReportProcessingError(f"Configuração para '{report_type}' não encontrada.")
    
    logging.info(f"Iniciando processamento para Relatório: {report_type}, Analista: {analyst}, Mês/Ano: {month}/{year}")

    cfg.update(build_report_paths(report_type, year, month))
    
    df_dados, df_contatos = _load_and_process_data(cfg)
    
    df_merged = pd.merge(df_dados, df_contatos, on="Empresa", how="left")
    df_filtered = df_merged[df_merged["Analista"] == analyst].copy()
    
    if df_filtered.empty: 
        logging.warning(f"Nenhum dado encontrado para analista '{analyst}'")
        return []

    df_filtered["Email"] = df_filtered["Email"].fillna("EMAIL_NAO_ENCONTRADO")

    common_data = {
        "analyst": analyst,
        "month_long": month.title(),
        "month_num": {m: f"{i+1:02d}" for i, m in enumerate(MESES)}.get(month.upper()),
        "year": year
    }

    results, created_count = [], 0
    
    for idx, row in df_filtered.iterrows():
        try:
            logging.info(f"\n{'='*60}")
            logging.info(f"Processando linha {idx+1}/{len(df_filtered)}: {row.get('Empresa', 'N/A')}")
            
            email_data = render_email_from_template(
                report_type, 
                row.to_dict(), 
                common_data, 
                cfg, 
                auto_send=True
            )
            
            if email_data:
                created_count += 1
                results.append({
                    "empresa": row["Empresa"],
                    "data": row.get("Data", "N/A"),
                    "valor": _format_currency(row.get("Valor", 0)),
                    "email": row["Email"], 
                    "anexos_count": len(email_data.get("attachments", [])), 
                    "created_count": created_count
                })
                logging.info(f"E-mail criado com sucesso para {row['Empresa']}")
            else:
                logging.info(f"E-mail pulado para {row['Empresa']}")
                
        except Exception as e:
            logging.error(f"❌ Erro ao processar linha para {row.get('Empresa', 'Empresa desconhecida')}: {e}")
            continue
    
    logging.info(f"\n{'='*60}")
    logging.info(f"Processamento concluído: {created_count} e-mails criados de {len(df_filtered)} empresas")
    logging.info(f"{'='*60}\n")
            
    return results


@st.cache_data(show_spinner=False)
def preview_dados(report_type: str, analyst: str, month: str, year: str) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Carrega dados para pré-visualização sem enviar e-mails.
    
    Args:
        report_type: Tipo do relatório
        analyst: Nome do analista
        month: Mês (nome por extenso)
        year: Ano
        
    Returns:
        Tupla com (DataFrame filtrado, configurações)
    """
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    if not cfg:
        raise ReportProcessingError(f"'{report_type}' não encontrado nas configurações.")

    report_paths = build_report_paths(report_type, year, month)
    cfg.update(report_paths)
    
    df_dados, df_contatos = _load_and_process_data(cfg)
    
    df_merged = pd.merge(df_dados, df_contatos, on="Empresa", how="left")
    df_filtered = df_merged[df_merged["Analista"] == analyst].copy()
    
    if df_filtered.empty:
        raise ReportProcessingError(f"Nenhum registro encontrado para o analista '{analyst}'")

    df_filtered["Email"] = df_filtered["Email"].fillna("EMAIL_NAO_ENCONTRADO")
    
    return df_filtered, cfg