import pandas as pd
import sys
import pythoncom
import re
import base64
import streamlit as st
import mimetypes
import requests
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from jinja2 import Environment, BaseLoader, meta
import logging
from src.config.config import MESES
from src.config.config_manager import load_configs, build_report_paths
from src.utils.security_utils import sanitize_html, sanitize_subject
from src.utils.data_utils import parse_brazilian_number, format_currency, format_date
from src.utils.file_utils import load_excel_data, find_attachment, load_email_templates
from src.handlers.report_handlers import REPORT_HANDLERS
from src.utils.file_utils import ReportProcessingError

try:
    import win32com.client as win32
    WIN32_AVAILABLE = sys.platform == "win32"
except ImportError:
    WIN32_AVAILABLE = False
    logging.warning("win32com não disponível. Modo de simulação ativado.")

# Nova função para criar rascunho via Graph API
def create_graph_draft(access_token: str, recipient: str, subject: str, body: str, attachments: List[Path]) -> bool:
    """Cria um rascunho de e-mail na caixa do usuário logado via MS Graph API."""
    if not access_token:
        st.error("Token de acesso inválido ou ausente.")
        logging.error("Tentativa de criar rascunho sem token de acesso.")
        return False

    graph_url = "https://graph.microsoft.com/v1.0/me/messages"
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }

    # Prepara a lista de destinatários (lidando com múltiplos e-mails separados por ';')
    to_recipients_list = []
    if recipient:
        addresses = [addr.strip() for addr in recipient.split(';') if addr.strip()]
        if addresses:
             to_recipients_list = [{"emailAddress": {"address": addr}} for addr in addresses]
        else:
             logging.warning(f"Nenhum destinatário válido encontrado em: {recipient}")
             # Decide se quer falhar ou criar rascunho sem destinatário
             # return False # Descomente para falhar se não houver destinatário


    email_payload = {
        "subject": subject,
        "importance": "Normal",
        "body": {
            "contentType": "HTML",
            "content": body
        },
        # Só adiciona toRecipients se a lista não estiver vazia
        **({"toRecipients": to_recipients_list} if to_recipients_list else {}),
        "attachments": []
    }

    # Adicionar anexos
    total_attachment_size = 0
    MAX_ATTACHMENT_SIZE_MB = 25 # Limite comum, ajuste se necessário
    for attachment_path in attachments:
        if attachment_path and attachment_path.exists():
            try:
                file_size = attachment_path.stat().st_size
                if total_attachment_size + file_size > MAX_ATTACHMENT_SIZE_MB * 1024 * 1024:
                    logging.warning(f"Anexo {attachment_path.name} excede o limite de tamanho total e não será adicionado.")
                    st.warning(f"Anexo {attachment_path.name} excede o limite de tamanho e não foi adicionado.")
                    continue # Pula para o próximo anexo

                content_bytes = attachment_path.read_bytes()
                content_b64 = base64.b64encode(content_bytes).decode('utf-8')
                mime_type, _ = mimetypes.guess_type(attachment_path.name)
                email_payload["attachments"].append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment_path.name,
                    "contentType": mime_type or "application/octet-stream",
                    "contentBytes": content_b64
                })
                total_attachment_size += file_size
            except Exception as e:
                logging.error(f"Erro ao processar anexo {attachment_path.name}: {e}")
                st.warning(f"Não foi possível anexar {attachment_path.name}: {e}")
        else:
            logging.warning(f"Anexo não encontrado ou inválido ao preparar API: {attachment_path}")

    try:
        # Enviar request para criar o rascunho (salva na pasta Rascunhos)
        response = requests.post(graph_url, headers=headers, json=email_payload)

        if response.status_code == 201: # 201 Created significa sucesso
            logging.info(f"Rascunho criado com sucesso para {recipient or 'sem destinatário'}")
            return True
        else:
            # Log detalhado do erro da API
            error_details = response.json().get('error', {})
            error_message = error_details.get('message', 'Erro desconhecido da API Graph.')
            logging.error(f"Erro ao criar rascunho via Graph API ({response.status_code}) para {recipient}: {response.text}")
            st.error(f"Erro da API ao criar rascunho ({response.status_code}): {error_message}")
            return False
    except requests.exceptions.RequestException as e:
        logging.error(f"Erro de conexão com a API Graph ao criar rascunho: {e}")
        st.error(f"Erro de conexão ao tentar criar rascunho: {e}")
        return False

def create_outlook_draft(recipient: str, subject: str, body: str, attachments: List[Path]) -> None:
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

def build_filename(company: str, report_type: str, month: str, year: str) -> str:
    company_clean = str(company).strip()
    company_part = re.sub(r"[\s_-]+", "_", company_clean).upper()
    report_part = str(report_type).upper()
    month_part = str(month).lower()[:3]
    year_part = str(year)[-2:]
    return f"{company_part}_{report_part}_{month_part}_{year_part}.pdf"

def resolve_variant(report_type: str, report_config: Dict[str, Any], context: Dict[str, Any]) -> Tuple[Dict[str, Any], str]:
    if "variants" not in report_config:
        return report_config, "default"

    variants = report_config["variants"]

    if report_type.startswith("LFRES"):
        raw_val = context.get("valor", 0.0)
        try:
            valor = float(raw_val)
        except (ValueError, TypeError):
            try:
                valor = parse_brazilian_number(raw_val)
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

    first_key = next(iter(variants), "Padrao")
    return variants.get(first_key, report_config), first_key

def render_email_from_template(report_type: str, row: Dict[str, Any], common: Dict[str, Any], cfg: Dict[str, Any], auto_send: bool = False) -> Optional[Dict[str, Any]]:
    templates = load_email_templates()
    template_key = "LFRES" if report_type.startswith("LFRES") else report_type
    report_cfg = templates.get(template_key)
    if not report_cfg:
        raise ReportProcessingError(f"Template para '{template_key}' não encontrado.")

    context = {**row, **common, **cfg}
    context.update({
        "empresa": row.get("Empresa"),
        "mesext": common.get("month_long"),
        "mes": common.get("month_num"),
        "ano": common.get("year"),
        "data": row.get("Data"),
        "assinatura": common.get("analyst"),
        "valor": parse_brazilian_number(row.get("Valor", 0))
    })
    logging.info(f"Processando {context['empresa']} - Tipo: {report_type} - Valor original: '{row.get('Valor', 0)}' -> Parseado: {context['valor']}")

    if report_type in REPORT_HANDLERS:
        handler_func = REPORT_HANDLERS[report_type]
    context = handler_func(context, row, cfg, report_type=report_type, parsed_valor=context['valor'])

    selected_template, variant_name = resolve_variant(template_key, report_cfg, context)
    
    logging.info(f"Variante selecionada: {variant_name}")
    
    if variant_name == "SKIP":
        logging.info(f"Pulando {context.get('empresa')} (Gerador-EER com valor zero)")
        return None

    for key in ["valor", "ValorLiquidacao", "ValorLiquidado", "ValorInadimplencia"]:
        if key in context and context[key] is not None:
            context[key] = format_currency(context[key])
    
    if "data" in context and context["data"] is not None:
        context["data"] = format_date(context["data"])
    
    for key in ["data", "dataaporte", "data_liquidacao"]:
        if key in context and context.get(key) is not None:
            context[key] = format_date(context[key])

    attachments = []
    filename = build_filename(str(row.get("Empresa","")), report_type, common["month_long"].upper(), str(common.get("year","")))
    if cfg.get("pdfs_dir"):
        path = find_attachment(cfg["pdfs_dir"], filename)
        if path: 
            attachments.append(path)
            logging.info(f"Anexo encontrado: {filename}")
        else:
            logging.warning(f"Anexo não encontrado: {filename}")

    if report_type == "GFN001":
        filename_sum = build_filename(str(row.get("Empresa","")), "SUM001", common["month_long"].upper(), str(common.get("year","")))
        base_dir = Path(cfg.get("pdfs_dir", ""))
        memoria_calc_dir = base_dir.parent.parent / "Sumário" / "SUM001 - Memória_de_Cálculo"
        sum_path = find_attachment(str(memoria_calc_dir), filename_sum)
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

    logging.info(f"Contexto final para renderização: {context}")

    env = Environment(loader=BaseLoader())
    def normalize(s: str): 
        """Normaliza placeholders {var} para {{ var }}"""
        return re.sub(r"\{(\w+)\}", r"{{ \1 }}", s) if isinstance(s, str) else s
    
    parsed_body = env.parse(normalize(body_tpl))
    undeclared_vars = meta.find_undeclared_variables(parsed_body)
    for k in undeclared_vars: 
        if k not in context:
            context[k] = f"[{k} N/D]"
            logging.warning(f"Placeholder não encontrado: {k}")

    subject = env.from_string(normalize(subject_tpl)).render(context)
    body = env.from_string(normalize(body_tpl)).render(context)
    
    result = {
        "subject": sanitize_subject(subject), 
        "body": sanitize_html(body), 
        "attachments": attachments,
        "missing_placeholders": list(undeclared_vars),
        "attachment_warnings": []
    }

    attachment_names = [p.name for p in attachments]
    logging.info(f"Anexos a serem incluídos no e-mail: {attachment_names}")

    if auto_send:
        access_token = st.session_state.get("ms_token", {}).get("access_token")
        if access_token:
            if "<p>Atenciosamente," not in result["body"]:
                 result["body"] += f"<br><p>Atenciosamente,</p><p><strong>{common['analyst']}</strong></p>"

            recipient_email = row.get("Email", "")

            success = create_graph_draft(
                access_token,
                recipient_email,
                result["subject"],
                result["body"],
                result["attachments"]
            )
            if not success:
                 logging.error(f"Falha ao criar rascunho via Graph API para {row.get('Empresa')}")

        else:
            st.warning("Login necessário para enviar/criar rascunhos.")
            logging.warning(f"Tentativa de envio sem token para {row.get('Empresa')}")


    return result

def load_and_process_data(cfg: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    header = int(cfg.get("header_row", 0))

    logging.info(f"Carregando dados de: {cfg['excel_dados']}")
    df_dados = load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], header)

    logging.info(f"Carregando contatos de: {cfg['excel_contatos']}")
    df_contatos = load_excel_data(cfg["excel_contatos"], cfg["sheet_contatos"], 0)
    
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
    if not cfg:
        raise ReportProcessingError(f"Configuração para '{report_type}' não encontrada.")
    
    logging.info(f"Iniciando processamento para Relatório: {report_type}, Analista: {analyst}, Mês/Ano: {month}/{year}")

    cfg.update(build_report_paths(report_type, year, month))
    
    df_dados, df_contatos = load_and_process_data(cfg)
    
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

    results, created_count, error_count = [], 0, 0
    
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
                    "valor": format_currency(row.get("Valor", 0)),
                    "email": row["Email"], 
                    "anexos_count": len(email_data.get("attachments", [])), 
                    "created_count": created_count
                })
                logging.info(f"E-mail criado com sucesso para {row['Empresa']}")
            else:
                logging.info(f"E-mail pulado para {row['Empresa']}")
                
        except Exception as e:
            error_count += 1 
            logging.error(f"Erro ao processar linha para {row.get('Empresa', 'Empresa desconhecida')}: {e}")
            continue
    
    final_message = f"Processamento concluído: {created_count} e-mails criados de {len(df_filtered)} empresas."
    if error_count > 0:
        final_message += f" {error_count} empresas falharam."
    
    logging.info(final_message)
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
    
    df_dados, df_contatos = load_and_process_data(cfg)
    
    df_merged = pd.merge(df_dados, df_contatos, on="Empresa", how="left")
    df_filtered = df_merged[df_merged["Analista"] == analyst].copy()
    
    if df_filtered.empty:
        raise ReportProcessingError(f"Nenhum registro encontrado para o analista '{analyst}'")

    df_filtered["Email"] = df_filtered["Email"].fillna("EMAIL_NAO_ENCONTRADO")
    
    return df_filtered, cfg