import pandas as pd
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

    to_recipients_list = []
    if recipient:
        addresses = [addr.strip() for addr in recipient.split(';') if addr.strip() and '@' in addr] # Basic check
        if addresses:
             to_recipients_list = [{"emailAddress": {"address": addr}} for addr in addresses]
        else:
             logging.warning(f"Nenhum destinatário válido encontrado em: {recipient}")
             st.warning(f"Tentando criar rascunho sem destinatário válido para linha com '{recipient}'.")


    email_payload = {
        "subject": subject,
        "importance": "Normal",
        "body": {
            "contentType": "HTML",
            "content": body
        },
        **({"toRecipients": to_recipients_list} if to_recipients_list else {}),
        "attachments": []
    }

    total_attachment_size = 0
    MAX_ATTACHMENT_SIZE_MB = 25
    for attachment_path in attachments:
        if attachment_path and attachment_path.exists():
            try:
                file_size = attachment_path.stat().st_size
                if total_attachment_size + file_size > MAX_ATTACHMENT_SIZE_MB * 1024 * 1024:
                    logging.warning(f"Anexo {attachment_path.name} excede o limite de tamanho total e não será adicionado.")
                    st.warning(f"Anexo {attachment_path.name} excede o limite de tamanho e não foi adicionado.")
                    continue

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
                logging.error(f"Erro ao processar anexo {attachment_path.name}: {e}", exc_info=True)
                st.warning(f"Não foi possível processar o anexo {attachment_path.name}: {e}")
        else:
            logging.warning(f"Anexo não encontrado ou inválido ao preparar API: {attachment_path}")

    try:
        response = requests.post(graph_url, headers=headers, json=email_payload)

        if response.status_code == 201: # 201 Created significa sucesso
            logging.info(f"Rascunho criado com sucesso para {recipient or 'sem destinatário'}")
            return True
        else:
            error_details = response.json().get('error', {})
            error_message = error_details.get('message', 'Erro desconhecido da API Graph.')
            logging.error(f"Erro ao criar rascunho via Graph API ({response.status_code}) para {recipient}: {response.text}")
            st.error(f"Erro da API ao criar rascunho ({response.status_code}): {error_message}")
            return False
    except requests.exceptions.RequestException as e:
        logging.error(f"Erro de conexão com a API Graph ao criar rascunho: {e}")
        st.error(f"Erro de conexão ao tentar criar rascunho: {e}")
        return False
    except Exception as e:
        logging.error(f"Erro inesperado em create_graph_draft: {e}", exc_info=True)
        st.error(f"Erro inesperado ao criar rascunho: {e}")
        return False

def render_email_from_template(report_type: str, row: Dict[str, Any], common: Dict[str, Any], cfg: Dict[str, Any]) -> Optional[Dict[str, Any]]:
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
    logging.info(f"Processando {context.get('empresa', 'N/A')} - Tipo: {report_type} - Valor original: '{row.get('Valor', 0)}' -> Parseado: {context.get('valor', 'N/A')}")

    if report_type in REPORT_HANDLERS:
        handler_func = REPORT_HANDLERS[report_type]
        try:
            context = handler_func(context, row, cfg, report_type=report_type, parsed_valor=context.get('valor'))
        except Exception as e:
             logging.error(f"Erro no handler {report_type} para {context.get('empresa')}: {e}", exc_info=True)
             st.warning(f"Erro ao preparar dados específicos ({report_type}) para {context.get('empresa')}: {e}")


    selected_template, variant_name = resolve_variant(template_key, report_cfg, context)
    logging.info(f"Variante selecionada para {context.get('empresa')}: {variant_name}")

    if variant_name == "SKIP":
        logging.info(f"Pulando {context.get('empresa')} (lógica da variante SKIP)")
        return None

    for key in ["valor", "ValorLiquidacao", "ValorLiquidado", "ValorInadimplencia"]:
        if key in context and context[key] is not None:
            try:
                 if not isinstance(context[key], str) or "R$" not in context[key]:
                      context[key] = format_currency(context[key])
            except Exception:
                 logging.warning(f"Não foi possível formatar '{key}' como moeda para {context.get('empresa')}.")
                 context[key] = str(context[key])

    date_keys = ["data", "dataaporte", "data_liquidacao"]
    for key in date_keys:
        if key in context and context.get(key) is not None:
             context[key] = format_date(context[key])

    attachments = []
    filename = build_filename(str(row.get("Empresa","Desconhecida")), report_type, common.get("month_long", "").upper(), str(common.get("year","")))
    if cfg.get("pdfs_dir"):
        try:
            path = find_attachment(cfg["pdfs_dir"], filename)
            if path:
                attachments.append(path)
                logging.info(f"Anexo principal encontrado para {context.get('empresa')}: {filename}")
        except Exception as e:
            logging.error(f"Erro ao procurar anexo principal para {context.get('empresa')}: {e}")
    if report_type == "GFN001":
        try:
            filename_sum = build_filename(str(row.get("Empresa","Desconhecida")), "SUM001", common.get("month_long", "").upper(), str(common.get("year","")))
            base_dir = Path(cfg.get("pdfs_dir", ""))
            if base_dir.parent and base_dir.parent.parent:
                memoria_calc_dir = base_dir.parent.parent / "Sumário" / "SUM001 - Memória_de_Cálculo"
                sum_path = find_attachment(str(memoria_calc_dir), filename_sum)
                if sum_path:
                    attachments.append(sum_path)
                    logging.info(f"Anexo SUM001 encontrado para {context.get('empresa')}: {filename_sum}")
            else:
                 logging.warning(f"Não foi possível determinar o diretório pai para buscar o anexo SUM001 (base_dir: {cfg.get('pdfs_dir')}).")
        except Exception as e:
            logging.error(f"Erro ao procurar anexo SUM001 para {context.get('empresa')}: {e}")


    subject_tpl = selected_template.get("subject_template", f"{report_type} - {context.get('empresa')}") # Default mais seguro
    body_tpl = selected_template.get("body_html", "")

    if report_type == "LFN001":
        situacao_lfn = str(row.get("Situacao","")).strip()
        if situacao_lfn == "Crédito":
            body_tpl = selected_template.get("body_html_credit", body_tpl)
        elif situacao_lfn == "Débito":
            body_tpl = selected_template.get("body_html_debit", body_tpl)

    logging.debug(f"Contexto final para renderização ({context.get('empresa')}): {context}")

    env = Environment(loader=BaseLoader())
    def normalize(s: str):
        return re.sub(r"\{(\w+)\}", r"{{ \1 }}", s) if isinstance(s, str) else s

    missing_placeholders = []
    try:
        parsed_body = env.parse(normalize(body_tpl))
        undeclared_vars = meta.find_undeclared_variables(parsed_body)
        missing_placeholders = list(undeclared_vars)
        for k in undeclared_vars:
            if k not in context:
                context[k] = f"[{k} N/D]"
                logging.warning(f"Placeholder '{k}' não encontrado no contexto para {context.get('empresa')}.")
    except Exception as e:
        logging.error(f"Erro ao analisar template Jinja2 para {context.get('empresa')}: {e}")

        body_tpl = f"<p>Erro ao processar o template do e-mail: {e}</p>"

    try:
        subject = env.from_string(normalize(subject_tpl)).render(context)
        body = env.from_string(normalize(body_tpl)).render(context)
    except Exception as e:
        logging.error(f"Erro ao renderizar template Jinja2 para {context.get('empresa')}: {e}", exc_info=True)
        subject = f"ERRO NO TEMPLATE - {report_type} - {context.get('empresa')}"
        body = f"<p>Ocorreu um erro ao gerar o corpo deste e-mail a partir do template.</p><p>Erro: {e}</p>"

    assinatura_html = f"<br><p>Atenciosamente,</p><p><strong>{common.get('analyst', 'Equipe DGCA')}</strong></p>"
    if "<p>Atenciosamente," not in body:
         body += assinatura_html


    result = {
        "subject": sanitize_subject(subject),
        "body": sanitize_html(body),
        "attachments": attachments,
        "missing_placeholders": missing_placeholders,
        "attachment_warnings": []
    }

    return result

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

    # Lógica específica para SUM001
    if report_type == "SUM001":
        logic = report_config.get("logic", {})
        if logic and "variant_selector" in logic and "conditions" in logic:
            selector_value = str(context.get(logic["variant_selector"], "")).strip()
            logging.info(f"SUM001 Seletor valor={selector_value}, contexto={context}")
            variant_name = logic["conditions"].get(selector_value, logic["conditions"].get("default", "padrao"))
            variant = variants.get(variant_name, {})
            if not variant:
                logging.warning(f"Variante {variant_name} não encontrada para {selector_value}, usando padrão")
                variant = variants.get("padrao", {})
            logging.info(f"SUM001 Variante selecionada para {context.get('empresa')}: {variant_name} (selector={selector_value})")
            return variant, variant_name

    # Lógica específica para LFRES
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
    """
    Processa relatórios, renderiza e-mails e tenta criar rascunhos via API Graph.

    Returns:
        Lista de dicionários, um para cada e-mail *efetivamente* criado com sucesso pela API.
    """
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    if not cfg:
        logging.error(f"Configuração para '{report_type}' não encontrada ao iniciar process_reports.")
        raise ReportProcessingError(f"Configuração para '{report_type}' não encontrada.")

    logging.info(f"Iniciando processamento para Relatório: {report_type}, Analista: {analyst}, Mês/Ano: {month}/{year}")

    try:
        cfg.update(build_report_paths(report_type, year, month))
        df_dados, df_contatos = load_and_process_data(cfg)
    except FileNotFoundError as e:
         logging.error(f"Erro ao carregar arquivos base (Excel): {e}", exc_info=True)
         raise ReportProcessingError(f"Erro ao carregar arquivos de dados/contato: {e}")
    except Exception as e:
         logging.error(f"Erro inesperado ao carregar/processar dados iniciais: {e}", exc_info=True)
         raise ReportProcessingError(f"Erro inesperado ao carregar dados: {e}")


    df_merged = pd.merge(df_dados, df_contatos, on="Empresa", how="left")
    if "Analista" not in df_merged.columns:
        logging.error("Coluna 'Analista' não encontrada no DataFrame mesclado. Verifique mapeamentos e planilha de contatos.")
        raise ReportProcessingError("Coluna 'Analista' ausente nos dados. Verifique a configuração e a planilha de contatos.")

    df_filtered = df_merged[df_merged["Analista"] == analyst].copy()

    if df_filtered.empty:
        logging.warning(f"Nenhum dado encontrado para o analista '{analyst}' após filtro.")
        return []

    missing_email_mask = df_filtered["Email"].isna() | (df_filtered["Email"].str.strip() == "")
    if missing_email_mask.any():
        missing_companies = df_filtered.loc[missing_email_mask, "Empresa"].tolist()
        logging.warning(f"E-mail não encontrado para as seguintes empresas: {', '.join(missing_companies)}")
        df_filtered.loc[missing_email_mask, "Email"] = "EMAIL_NAO_ENCONTRADO"

    common_data = {
        "analyst": analyst,
        "month_long": month.title(),
        "month_num": {m.upper(): f"{i+1:02d}" for i, m in enumerate(MESES)}.get(month.upper(), "??"),
        "year": year
    }

    results_success = []
    created_count = 0
    render_errors = 0
    api_errors = 0
    skipped_count = 0

    access_token = st.session_state.get("ms_token", {}).get("access_token")
    if not access_token:
        st.error("Erro: Usuário não autenticado ou sessão expirada. Não é possível criar rascunhos.")
        logging.error("Tentativa de processar relatórios sem token de acesso válido na sessão.")
        return []

    for idx, row in df_filtered.iterrows():
        try:
            logging.info(f"--- Processando Linha {idx+1}/{len(df_filtered)}: {row.get('Empresa', 'N/A')} ---")

            email_data = render_email_from_template(
                report_type,
                row.to_dict(),
                common_data,
                cfg
            )

            if email_data is None:
                skipped_count += 1
                logging.info(f"E-mail pulado (lógica interna, ex: variante SKIP) para {row.get('Empresa')}")
                continue

            recipient_email = row.get("Email", "")
            valid_recipients = [addr.strip() for addr in recipient_email.split(';') if addr.strip() and '@' in addr]
            if not valid_recipients:
                 logging.warning(f"Nenhum destinatário válido para {row.get('Empresa')} ('{recipient_email}'). Pulando chamada da API.")
                 st.warning(f"E-mail para {row.get('Empresa')} não pôde ser criado (sem destinatário válido).")
                 api_errors += 1
                 continue


            success = create_graph_draft(
                access_token,
                recipient_email,
                email_data["subject"],
                email_data["body"],
                email_data["attachments"]
            )

            if success:
                created_count += 1
                results_success.append({
                    "empresa": row.get("Empresa", "N/A"),
                    "data": format_date(row.get("Data")),
                    "valor": format_currency(row.get("Valor", 0)),
                    "email": recipient_email,
                    "anexos_count": len(email_data.get("attachments", [])),
                    "created_count": created_count
                })
            else:
                api_errors += 1
                logging.error(f"Falha na chamada da API Graph para criar rascunho para {row.get('Empresa')}")

        except ReportProcessingError as rpe:
             render_errors += 1
             logging.error(f"Erro de processamento (ReportProcessingError) para {row.get('Empresa', 'Empresa desconhecida')}: {rpe}", exc_info=True)
             st.warning(f"Erro ao processar {row.get('Empresa', 'Empresa desconhecida')}: {rpe}")
             continue
        except Exception as e:
            render_errors += 1
            logging.error(f"Erro GERAL inesperado ao processar linha para {row.get('Empresa', 'Empresa desconhecida')}: {e}", exc_info=True)
            st.error(f"Erro inesperado ao processar {row.get('Empresa', 'Empresa desconhecida')}: {e}")
            continue

    total_processed = len(df_filtered)
    final_message = f"Processamento concluído: {created_count} de {total_processed} rascunhos criados com sucesso."
    if skipped_count > 0:
        final_message += f" {skipped_count} e-mails foram pulados intencionalmente."
    if render_errors > 0:
        final_message += f" {render_errors} falharam durante a preparação/renderização."
    if api_errors > 0:
        final_message += f" {api_errors} falharam durante a criação via API (verifique logs/mensagens de erro acima)."

    logging.info(final_message)
    logging.info(f"{'='*60}\n")
    return results_success

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