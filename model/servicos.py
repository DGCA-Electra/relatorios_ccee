import pandas as pd
import re
import os
import base64
import mimetypes
import requests
import shutil
import tempfile
import subprocess
import tempfile
import logging
from pathlib import Path as caminho
from typing import Dict, List, Any, Optional, Tuple
from jinja2 import Environment, BaseLoader, meta
from apps.relatorios_ccee.configuracoes.constantes import MESES
from apps.relatorios_ccee.configuracoes.gerenciador import carregar_configuracoes, construir_caminhos_relatorio
from apps.relatorios_ccee.model.seguranca import sanitizar_html, sanitizar_assunto
from apps.relatorios_ccee.model.utils_dados import converter_numero_br, formatar_moeda, formatar_data
from apps.relatorios_ccee.model.arquivos import ler_dados_excel, encontrar_anexo, carregar_templates_email, ErroProcessamento
from .relatorios import PROCESSADORES_RELATORIO, processador_generico_relatorio

def criar_rascunho_graph(token_acesso: str, destinatario: str, assunto: str, corpo: str, anexos: List[caminho]) -> bool:
    """Cria um rascunho de e-mail na caixa do usuário logado via MS Graph API.

    Raises:
        ErroProcessamento: Em caso de falha na criação do rascunho ou ausência de token.
    """
    if not token_acesso:
        logging.error("Tentativa de criar rascunho sem token de acesso.")
        raise ErroProcessamento("Token de acesso inválido ou ausente.")
    graph_url = "https://graph.microsoft.com/v1.0/me/messages"
    headers = {
        'Authorization': 'Bearer ' + token_acesso,
        'Content-Type': 'application/json'
    }
    lista_destinatarios = []
    if destinatario:
        enderecos = [addr.strip() for addr in destinatario.split(';') if addr.strip() and '@' in addr] # Checagem básica
        if enderecos:
             lista_destinatarios = [{"emailAddress": {"address": addr}} for addr in enderecos]
        else:
             logging.warning(f"Nenhum destinatário válido encontrado em: {destinatario}")
             st.warning(f"Tentando criar rascunho sem destinatário válido para linha com '{destinatario}'.")
    
    payload_email = {
        "subject": assunto,
        "importance": "Normal",
        "body": {
            "contentType": "HTML",
            "content": corpo
        },
        **({"toRecipients": lista_destinatarios} if lista_destinatarios else {}),
        "attachments": []
    }
    tamanho_total_anexos = 0
    LIMITE_TAMANHO_ANEXO_MB = 25
    for caminho_anexo in anexos:
        if caminho_anexo and caminho_anexo.exists():
            caminho_temporario_str = None
            try:
                origem_str = str(caminho_anexo.resolve())
                
                fd, caminho_temporario_str = tempfile.mkstemp(suffix=".pdf")
                os.close(fd)

                comando = f'copy /Y /B "{origem_str}" "{caminho_temporario_str}"'
                
                processo = subprocess.run(
                    comando, 
                    shell=True, 
                    stdout=subprocess.PIPE, 
                    stderr=subprocess.PIPE
                )
                
                if processo.returncode != 0 or not os.path.exists(caminho_temporario_str) or os.path.getsize(caminho_temporario_str) == 0:
                    erro_msg = processo.stderr.decode("cp850", errors="ignore") or "Erro desconhecido no copy"
                    logging.error(f"CMD Copy falhou para {caminho_anexo.name}: {erro_msg}")
                    st.warning(f"Windows não conseguiu copiar o anexo {caminho_anexo.name}. Erro: {erro_msg}")
                    continue

                tamanho_arquivo = os.path.getsize(caminho_temporario_str)
                if tamanho_total_anexos + tamanho_arquivo > LIMITE_TAMANHO_ANEXO_MB * 1024 * 1024:
                    logging.warning(f"Anexo {caminho_anexo.name} excede o limite.")
                    st.warning(f"Anexo {caminho_anexo.name} muito grande, ignorado.")
                    try: os.remove(caminho_temporario_str)
                    except: pass
                    continue

                with open(caminho_temporario_str, "rb") as f:
                    conteudo_bytes = f.read()

                try: os.remove(caminho_temporario_str)
                except: pass 

                conteudo_b64 = base64.b64encode(conteudo_bytes).decode('utf-8')
                tipo_mime, _ = mimetypes.guess_type(caminho_anexo.name)

                payload_email["attachments"].append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": caminho_anexo.name,
                    "contentType": tipo_mime or "application/octet-stream",
                    "contentBytes": conteudo_b64
                })
                
                tamanho_total_anexos += tamanho_arquivo

            except Exception as e:
                logging.error(f"Erro CRÍTICO ao processar anexo {caminho_anexo.name}: {e}", exc_info=True)
                st.warning(f"Erro ao anexar {caminho_anexo.name}: {e}")
                if caminho_temporario_str and os.path.exists(caminho_temporario_str):
                    try: os.remove(caminho_temporario_str)
                    except: pass
        else:
             logging.warning(f"Anexo não encontrado ou caminho inválido: {caminho_anexo}")
    try:
        response = requests.post(graph_url, headers=headers, json=payload_email)
        if response.status_code == 201:
            logging.info(f"Rascunho criado com sucesso para {destinatario or 'sem destinatário'}")
            return True
        else:
            detalhes_erro = response.json().get('error', {})
            mensagem_erro = detalhes_erro.get('message', 'Erro desconhecido da API Graph.')
            logging.error(f"Erro ao criar rascunho via Graph API ({response.status_code}) para {destinatario}: {response.text}")
            raise ErroProcessamento(f"Erro da API ao criar rascunho ({response.status_code}): {mensagem_erro}")
    except requests.exceptions.RequestException as e:
        logging.error(f"Erro de conexão com a API Graph ao criar rascunho: {e}")
        raise ErroProcessamento(f"Erro de conexão ao tentar criar rascunho: {e}")
    except Exception as e:
        logging.error(f"Erro inesperado em create_graph_draft: {e}", exc_info=True)
        raise ErroProcessamento(f"Erro inesperado ao criar rascunho: {e}")
def renderizar_email_modelo(tipo_relatorio: str, row: Dict[str, Any], dados_comuns: Dict[str, Any], config: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    templates = carregar_templates_email()
    template_key = "LFRES" if tipo_relatorio.startswith("LFRES") else tipo_relatorio
    report_config = templates.get(template_key)
    if not report_config:
        raise ErroProcessamento(f"Template para '{template_key}' não encontrado.")
    context = {**row, **dados_comuns, **config}
    context.update({
        "empresa": row.get("Empresa"),
        "mesext": dados_comuns.get("mes_long"),
        "mes": dados_comuns.get("mes_num"),
        "ano": dados_comuns.get("ano"),
        "data": row.get("Data"),
        "assinatura": dados_comuns.get("analista"),
        "valor": converter_numero_br(row.get("Valor", 0))
    })
    logging.info(f"Processando {context.get('empresa', 'N/A')} - Tipo: {tipo_relatorio} - Valor original: '{row.get('Valor', 0)}' -> Parseado: {context.get('valor', 'N/A')}")
    if tipo_relatorio in PROCESSADORES_RELATORIO:
        handler_func = PROCESSADORES_RELATORIO[tipo_relatorio]
        try:
            context = handler_func(context, row, config, tipo_relatorio=tipo_relatorio, parsed_valor=context.get('valor'))
        except Exception as e:
             logging.error(f"Erro no handler específico {tipo_relatorio}: {e}")
    else:
        try:
            context = processador_generico_relatorio(context, row, config, tipo_relatorio=tipo_relatorio)
        except Exception as e:
             logging.error(f"Erro no handler genérico {tipo_relatorio}: {e}")
    selected_template, variant_name = definir_variante_template(template_key, report_config, context)
    logging.info(f"Variante selecionada para {context.get('empresa')}: {variant_name}")
    if variant_name == "SKIP":
        logging.info(f"Pulando {context.get('empresa')} (lógica da variante SKIP)")
        return None
    for key in ["valor", "ValorLiquidacao", "ValorLiquidado", "ValorInadimplencia"]:
        if key in context and context[key] is not None:
            try:
                 if not isinstance(context[key], str) or "R$" not in context[key]:
                      context[key] = formatar_moeda(context[key])
            except Exception:
                 logging.warning(f"Não foi possível formatar '{key}' como moeda para {context.get('empresa')}.")
                 context[key] = str(context[key])
    date_keys = ["data", "dataaporte", "data_liquidacao"]
    for key in date_keys:
        if key in context and context.get(key) is not None:
             context[key] = formatar_data(context[key])
    anexos = []
    nome_arquivo = gerar_nome_arquivo(str(row.get("Empresa","Desconhecida")), tipo_relatorio, dados_comuns.get("mes_long", "").upper(), str(dados_comuns.get("ano","")))
    main_cache = config.get("_pdf_cache_main", {})
    caminho = main_cache.get(nome_arquivo.upper())
    if caminho:
        anexos.append(caminho)
        logging.info(f"Anexo principal encontrado (Cache) para {context.get('empresa')}: {nome_arquivo}")
    elif config.get("diretorio_pdfs"):
        try:
            caminho = encontrar_anexo(config["diretorio_pdfs"], nome_arquivo)
            if caminho:
                anexos.append(caminho)
                logging.info(f"Anexo principal encontrado (Disco) para {context.get('empresa')}: {nome_arquivo}")
            else:
                 logging.debug(f"Anexo não encontrado no diretório: {nome_arquivo}")
        except Exception as e:
            logging.error(f"Erro ao procurar anexo principal para {context.get('empresa')}: {e}")
    else:
        logging.debug(f"Anexo não encontrado no cache e sem diretório configurado: {nome_arquivo}")
    if tipo_relatorio == "GFN001":
        nome_arquivo_sum = gerar_nome_arquivo(str(row.get("Empresa","Desconhecida")), "SUM001", dados_comuns.get("mes_long", "").upper(), str(dados_comuns.get("ano","")))
        sum_cache = config.get("_pdf_cache_sumario", {})
        sum_caminho = sum_cache.get(nome_arquivo_sum.upper())
        if sum_caminho:
            anexos.append(sum_caminho)
            logging.info(f"Anexo SUM001 encontrado (Cache): {nome_arquivo_sum}")
        else:
            logging.debug(f"Anexo SUM001 não encontrado no cache: {nome_arquivo_sum}")
    assunto_tpl = selected_template.get("assunto_template", f"{tipo_relatorio} - {context.get('empresa')}") # Default mais seguro
    corpo_tpl = selected_template.get("corpo_html", "")
    if tipo_relatorio == "LFN001":
        situacao_lfn = str(row.get("Situacao","")).strip().lower()
        logging.info(f"LFN001 Debug - Empresa: {context.get('empresa')}, Situacao (raw): '{row.get('Situacao')}', Situacao (norm): '{situacao_lfn}'")
        if "crédito" in situacao_lfn or "credito" in situacao_lfn:
            corpo_tpl = selected_template.get("corpo_html_credit", corpo_tpl)
        elif "débito" in situacao_lfn or "debito" in situacao_lfn:
            corpo_tpl = selected_template.get("corpo_html_debit", corpo_tpl)
    logging.debug(f"Contexto final para renderização ({context.get('empresa')}): {context}")
    env = Environment(loader=BaseLoader())
    def normalize(s: str):
        return re.sub(r"\{(\w+)\}", r"{{ \1 }}", s) if isinstance(s, str) else s
    variaveis_ausentes = []
    try:
        parsed_corpo = env.parse(normalize(corpo_tpl))
        vars_nao_declaradas = meta.find_undeclared_variables(parsed_corpo)
        variaveis_ausentes = list(vars_nao_declaradas)
        for k in vars_nao_declaradas:
            if k not in context:
                context[k] = f"[{k} N/D]"
                logging.warning(f"Placeholder '{k}' não encontrado no contexto para {context.get('empresa')}.")
    except Exception as e:
        logging.error(f"Erro ao analisar template Jinja2 para {context.get('empresa')}: {e}")
        corpo_tpl = f"<p>Erro ao processar o template do e-mail: {e}</p>"
    try:
        assunto = env.from_string(normalize(assunto_tpl)).render(context)
        corpo = env.from_string(normalize(corpo_tpl)).render(context)
    except Exception as e:
        logging.error(f"Erro ao renderizar template Jinja2 para {context.get('empresa')}: {e}", exc_info=True)
        assunto = f"ERRO NO TEMPLATE - {tipo_relatorio} - {context.get('empresa')}"
        corpo = f"<p>Ocorreu um erro ao gerar o corpo deste e-mail a partir do template.</p><p>Erro: {e}</p>"
    assinatura_html = f"<br><p>Atenciosamente,</p><p><strong>{dados_comuns.get('analista', 'Equipe DGCA')}</strong></p>"
    if "<p>Atenciosamente," not in corpo:
         corpo += assinatura_html
    result = {
        "assunto": sanitizar_assunto(assunto),
        "corpo": sanitizar_html(corpo),
        "anexos": anexos,
        "variaveis_ausentes": variaveis_ausentes,
        "attachment_warnings": [],
        "final_data": context
    }
    return result
def gerar_nome_arquivo(company: str, tipo_relatorio: str, mes: str, ano: str) -> str:
    company_clean = str(company).strip()
    company_part = re.sub(r"[\s_-]+", "_", company_clean).upper()
    report_part = str(tipo_relatorio).upper()
    mes_part = str(mes).lower()[:3]
    ano_part = str(ano)[-2:]
    return f"{company_part}_{report_part}_{mes_part}_{ano_part}.pdf"
def definir_variante_template(tipo_relatorio: str, report_config: Dict[str, Any], context: Dict[str, Any]) -> Tuple[Dict[str, Any], str]:
    if "variantes" not in report_config:
        return report_config, "default"
    variantes = report_config["variantes"]
    # Lógica específica para SUM001
    if tipo_relatorio == "SUM001":
        logica = report_config.get("logica", {})
        if logica and "seletor_variante" in logica and "condicoes" in logica:
            selector_value = str(context.get(logica["seletor_variante"], "")).strip()
            logging.info(f"SUM001 Seletor valor={selector_value}, contexto={context}")
            variant_name = logica["condicoes"].get(selector_value, logica["condicoes"].get("default", "padrao"))
            variant = variantes.get(variant_name, {})
            if not variant:
                logging.warning(f"Variante {variant_name} não encontrada para {selector_value}, usando padrão")
                variant = variantes.get("padrao", {})

            merged = {**report_config, **variant}

            if 'variantes' in merged and isinstance(merged['variantes'], dict):
                merged.pop('variantes', None)
            logging.info(f"SUM001 Variante selecionada para {context.get('empresa')}: {variant_name} (selector={selector_value})")
            return merged, variant_name
    # Lógica específica para LFRES
    if tipo_relatorio.startswith("LFRES"):
        raw_val = context.get("valor", 0.0)
        try:
            valor = float(raw_val)
        except (ValueError, TypeError):
            try:
                valor = converter_numero_br(raw_val)
            except Exception:
                valor = 0.0
        valor_abs = abs(valor)
        tipo_agente = str(context.get("TipoAgente", "")).strip()
        logging.info(f"Resolve_variant LFRES - empresa={context.get('empresa')}, raw_val={raw_val}, valor={valor}, tipo_agente={tipo_agente}")
        if valor_abs > 1e-6:
            if tipo_agente == "Gerador-EER":
                logging.info("Selecionado: COM_VALOR_GERADOR")
                return variantes.get("COM_VALOR_GERADOR", {}), "COM_VALOR_GERADOR"
            logging.info("Selecionado: COM_VALOR_OUTROS")
            return variantes.get("COM_VALOR_OUTROS", {}), "COM_VALOR_OUTROS"
        if tipo_agente == "Gerador-EER":
            logging.info("Selecionado: SKIP (Gerador-EER com valor 0)")
            return {}, "SKIP"
        logging.info("Selecionado: ZERO_VALOR")
        return variantes.get("ZERO_VALOR", {}), "ZERO_VALOR"
    first_key = next(iter(variantes), "Padrao")
    return variantes.get(first_key, report_config), first_key
def carregar_e_processar_dados(config: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    cabecalho = int(config.get("linha_cabecalho", 0))
    logging.info(f"Carregando dados de: {config['excel_dados']}")
    df_dados = ler_dados_excel(config["excel_dados"], config["planilha_dados"], cabecalho)
    logging.info(f"Carregando contatos de: {config['excel_contatos']}")
    df_contatos = ler_dados_excel(config["excel_contatos"], config["planilha_contatos"], 0)
    column_mapping = dict(item.split(":") for item in config["colunas_dados"].split(","))
    df_dados.rename(columns=column_mapping, inplace=True)
    df_contatos.rename(columns={
        "AGENTE": "Empresa", 
        "ANALISTA": "Analista", 
        "E-MAILS RELATÓRIOS CCEE": "Email"
    }, inplace=True)
    return df_dados, df_contatos
def _indexar_diretorio(directory: str) -> Dict[str, caminho]:
    """
    Lista todos os arquivos PDF de um diretório e retorna um dicionário
    { "NOME_DO_ARQUIVO.PDF": caminho_Completo } para busca rápida (O(1)).
    A chave é armazenada em MAIÚSCULO para garantir busca case-insensitive.
    """
    if not directory:
        return {}
    caminho_obj = caminho(directory)
    if not caminho_obj.exists():
        logging.warning(f"Tentativa de indexar diretório inexistente: {directory}")
        return {}
    cache = {f.name.upper(): f for f in caminho_obj.glob("*.pdf")}
    logging.info(f"Diretório indexado: {directory} ({len(cache)} arquivos encontrados)")
    return cache
def _preparar_dados_relatorio(tipo_relatorio: str, analista: str, mes: str, ano: str, user_info: Optional[Dict[str, Any]] = None) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Função interna para carregar configs e dados.
    A decisão de usar caminho de REDE ou LOCAL agora é feita automaticamente pelo config_manager.
    """
    all_configs = carregar_configuracoes()
    config = all_configs.get(tipo_relatorio)
    if not config:
        raise ErroProcessamento(f"Configuração para '{tipo_relatorio}' não encontrada.")
    email_usuario = ""
    if user_info:
        email_usuario = user_info.get("userPrincipalName", "")
    username_rede = email_usuario.split("@")[0] if (email_usuario and "@" in email_usuario) else None
    if username_rede:
        logging.info(f"Usuário identificado para preferência de caminhos: {username_rede}")
    try:
        caminhos = construir_caminhos_relatorio(tipo_relatorio, ano, mes, username=username_rede)
        config.update(caminhos)
        df_dados, df_contatos = carregar_e_processar_dados(config)
    except FileNotFoundError as e:
        logging.error(f"Arquivos não encontrados (nem rede nem local): {e}")
        raise ErroProcessamento(f"Arquivos base não encontrados. Verifique a existência das pastas ou arquivos Excel.")
    except Exception as e:
         logging.error(f"Erro inesperado ao carregar dados iniciais: {e}", exc_info=True)
         raise ErroProcessamento(f"Erro inesperado ao carregar dados: {e}")
    if "diretorio_pdfs" in config:
        config["_pdf_cache_main"] = _indexar_diretorio(config["diretorio_pdfs"])
        if tipo_relatorio == "GFN001":
            base_dir = caminho(config["diretorio_pdfs"])
            try:
                if base_dir.parent and base_dir.parent.parent:
                    sum_dir = base_dir.parent.parent / "Sumário" / "SUM001 - Memória_de_Cálculo"
                    if sum_dir.exists():
                        config["_pdf_cache_sumario"] = _indexar_diretorio(str(sum_dir))
                    else:
                        logging.warning(f"Diretório de sumário não encontrado: {sum_dir}")
            except Exception as e:
                logging.error(f"Erro ao tentar indexar diretório de sumários: {e}")
    df_merged = pd.merge(df_dados, df_contatos, on="Empresa", how="left")
    if "Analista" not in df_merged.columns:
        raise ErroProcessamento("Coluna 'Analista' ausente nos dados. Verifique a configuração e a planilha de contatos.")
    df_filtrado = df_merged[df_merged["Analista"] == analista].copy()
    if df_filtrado.empty:
        logging.warning(f"Nenhum dado encontrado para o analista '{analista}' após filtro.")
        return df_filtrado, config
    missing_email_mask = df_filtrado["Email"].isna() | (df_filtrado["Email"].str.strip() == "")
    if missing_email_mask.any():
        df_filtrado.loc[missing_email_mask, "Email"] = "EMAIL_NAO_ENCONTRADO"
    return df_filtrado, config
def informa_processos(tipo_relatorio: str, analista: str, mes: str, ano: str, token_acesso: str, user_info: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """
    Processa relatórios, renderiza e-mails e tenta criar rascunhos via API Graph.
    """
    logging.info(f"Iniciando processamento: {tipo_relatorio}, Analista: {analista}, {mes}/{ano}")
    df_filtrado, config = _preparar_dados_relatorio(tipo_relatorio, analista, mes, ano, user_info=user_info)
    if df_filtrado.empty:
        return []
    dados_comuns = {
        "analista": analista,
        "mes_long": mes.title(),
        "mes_num": {m.upper(): f"{i+1:02d}" for i, m in enumerate(MESES)}.get(mes.upper(), "??"),
        "ano": ano
    }
    results_success = []
    contagem_criados = 0
    render_errors = 0
    api_errors = 0
    skipped_count = 0
    if not token_acesso:
        logging.error("Erro: Token de acesso ausente ao tentar enviar rascunhos.")
        raise ErroProcessamento("Usuário não autenticado. Não é possível criar rascunhos.")
    for idx, row in df_filtrado.iterrows():
        try:
            logging.info(f"--- Processando Linha {idx+1}/{len(df_filtrado)}: {row.get('Empresa', 'N/A')} ---")
            dados_email = renderizar_email_modelo(tipo_relatorio, row.to_dict(), dados_comuns, config)
            if dados_email is None:
                skipped_count += 1
                continue
            destinatario_email = row.get("Email", "")
            # Validação simples antes de chamar a API
            if not destinatario_email or "EMAIL_NAO_ENCONTRADO" in destinatario_email:
                 logging.warning(f"E-mail inválido para {row.get('Empresa')}. Pulando.")
                 api_errors += 1
                 continue
            try:
                criar_rascunho_graph(
                    token_acesso,
                    destinatario_email,
                    dados_email["assunto"],
                    dados_email["corpo"],
                    dados_email["anexos"]
                )
                contagem_criados += 1
                data_final = dados_email.get("final_data", {}).get("data") or row.get("Data")
                results_success.append({
                    "empresa": row.get("Empresa", "N/A"),
                    "data": formatar_data(data_final),
                    "valor": formatar_moeda(row.get("Valor", 0)),
                    "email": destinatario_email,
                    "contagem_anexos": len(dados_email.get("anexos", [])),
                    "contagem_criados": contagem_criados
                })
            except ErroProcessamento as e:
                api_errors += 1
                logging.error(f"Falha ao criar rascunho para {row.get('Empresa')}: {e}")
        except ErroProcessamento as rpe:
             render_errors += 1
             logging.error(f"Erro processamento: {rpe}")
             continue
        except Exception as e:
            render_errors += 1
            logging.error(f"Erro inesperado: {e}")
            continue
    logging.info(f"Fim do processamento. Criados: {contagem_criados}. Erros Render: {render_errors}. Erros API: {api_errors}")
    return results_success
def visualizar_previa_dados(tipo_relatorio: str, analista: str, mes: str, ano: str, user_info: Optional[Dict[str, Any]] = None) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Carrega dados para pré-visualização.
    """
    df_filtrado, config = _preparar_dados_relatorio(tipo_relatorio, analista, mes, ano, user_info=user_info)
    if df_filtrado.empty:
        raise ErroProcessamento(f"Nenhum registro encontrado para o analista '{analista}'")
    return df_filtrado, config