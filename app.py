import streamlit as st
import pandas as pd
import config.config as config
import services
from mail.email_utils import montar_corpo_email, enviar_email
import logging
import openpyxl
from typing import Dict, Any, Optional
import os
from utils.dataframe_utils import tratar_valores_df
import streamlit.components.v1 as components
from jinja2 import Environment, FileSystemLoader
import json

LOG_DIR = 'logs'
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, 'app.log')
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s', encoding='utf-8')

def registrar_log(mensagem: str) -> None:
    """Registra uma mensagem no log."""
    logging.info(mensagem)

st.set_page_config(
    page_title="Envio de Relatórios CCEE",
    page_icon="static/icon.png",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={
        'About': "Aplicação para automação de envio de e-mails DGCA."
    }
)

def init_state():
    defaults = {
        "report_type": "GFN001",
        "analyst": "Artur Bello Rodrigues",
        "month": "JANEIRO",
        "year": 2025
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def safe_join_emails(email_field):
    if not email_field:
        return ""
    if isinstance(email_field, list):
        return "; ".join(e.strip() for e in email_field if e)
    return "; ".join([e.strip() for e in str(email_field).split(';') if e.strip()])

def show_main_page() -> None:
    """Renderiza a página principal de envio de relatórios."""
    all_configs = config.load_configs()
    report_types = list(all_configs.keys())

    init_state()

    tipo = st.session_state.report_type
    analista_final = st.session_state.analyst
    mes = st.session_state.month
    ano = st.session_state.year

    st.title("⚡Envio de Relatórios CCEE - DGCA")
    st.info("💡 **Dica:** Você pode enviar relatórios para qualquer analista. Isso é útil durante férias ou ausências.")

    st.header("Parâmetros de Envio")
    c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
    with c1:
        st.session_state.report_type = st.selectbox("Tipo de Relatório", options=report_types, index=report_types.index(tipo) if tipo in report_types else 0)
    with c2:
        st.session_state.analyst = st.selectbox("Analista", options=config.ANALISTAS, index=config.ANALISTAS.index(analista_final) if analista_final in config.ANALISTAS else 0)
    with c3:
        st.session_state.month = st.selectbox("Mês", options=config.MESES, index=config.MESES.index(mes) if mes in config.MESES else 0)
    with c4:
        st.session_state.year = st.selectbox("Ano", options=config.ANOS, index=config.ANOS.index(str(ano)) if str(ano) in config.ANOS else 0)
    
    col1, col2 = st.columns(2)
    if col1.button("📊 Visualizar Dados", use_container_width=True):
        st.session_state.preview_trigger = True
    if col2.button("📧 Enviar E-mails", use_container_width=True, type="primary"):
        st.session_state.send_trigger = True

    if st.session_state.get("send_trigger"):
        with st.spinner("Enviando e-mails... Aguarde. Janelas do Outlook podem abrir para revisão."):
            try:
                results = services.process_reports(report_type=tipo, analyst=analista_final, month=mes, year=str(ano))
                st.session_state.results = results
                st.success(f"✅ E-mails processados e rascunhos criados no Outlook para {len(results)} empresas.")
            except services.ReportProcessingError as e:
                st.error(f"❌ Erro no envio: {e}")
        st.session_state.send_trigger = False

    if st.session_state.get("preview_trigger"):
        with st.spinner("Carregando dados para visualização... Por favor, aguarde."):
            try:
                df_filtered, cfg = services.preview_dados(report_type=tipo, analyst=analista_final, month=mes, year=str(ano))
                st.session_state.raw_preview_data = df_filtered
                st.session_state.preview_cfg = cfg
                st.session_state.form_data = {'tipo': tipo, 'analista': analista_final, 'mes': mes, 'ano': ano}
                st.success(f'✅ Dados carregados com sucesso! {len(df_filtered)} empresas encontradas para {analista_final}.')
            except services.ReportProcessingError as e:
                st.error(f"❌ Erro de processamento: {e}")
        st.session_state.preview_trigger = False

    def render_email_preview(subject: str, body_html: str):
        html = f"""
        <html><head><style>
          body {{ color: black; font-family: Inter, Arial, sans-serif; }}
          @media (prefers-color-scheme: dark) {{ body {{ color: white; }} }}
        </style></head><body><h4>{subject}</h4><hr/>{body_html}</body></html>
        """
        components.html(html, height=400, scrolling=True)

    if 'raw_preview_data' in st.session_state:
        df_raw = st.session_state.raw_preview_data
        cfg = st.session_state.get('preview_cfg', {})
        
        if not df_raw.empty:
            st.subheader(f"Dados para {tipo} - {mes}/{ano} - {analista_final}")
            df_display = tratar_valores_df(df_raw.copy())
            st.dataframe(df_display.reset_index(drop=True), use_container_width=True)
            
            st.subheader("Pré-visualização do E-mail")
            preview_limit = min(5, len(df_raw))
            for idx in range(preview_limit):
                dados_empresa = df_raw.iloc[idx].to_dict()

                if 'Email' in dados_empresa:
                    dados_empresa['Email'] = safe_join_emails(dados_empresa['Email'])
                
                common = {
                    'month_long': mes.title(),
                    'month_num': {m.upper(): f"{i+1:02d}" for i, m in enumerate(config.MESES)}.get(mes.upper(), '00'),
                    'year': str(ano),
                    'analyst': analista_final,
                }
                try:
                    rendered = services.render_email_from_template(tipo, dados_empresa, common, cfg, auto_send=False)
                    with st.expander(f"Prévia #{idx+1} - {dados_empresa.get('Empresa','')}", expanded=False):
                        render_email_preview(rendered['subject'], rendered['body'])
                except Exception as e:
                    st.warning(f"Falha ao renderizar template para {dados_empresa.get('Empresa','')}: {e}")

        if st.button("🗑️ Limpar Visualização", key="limpar_preview"):
            del st.session_state.raw_preview_data
            if 'preview_cfg' in st.session_state: del st.session_state.preview_cfg
            st.rerun()

    if 'results' in st.session_state and st.session_state.results:
        results = st.session_state.results
        form = st.session_state.get('form_data', {})
        st.header(f"📤 Resultado do Envio - {form.get('tipo', 'N/A')} - {form.get('mes', 'N/A')}/{form.get('ano', 'N/A')}")
        
        total_created = results[-1]['created_count'] if results else 0
        col1, col2 = st.columns(2)
        col1.metric("Empresas Processadas", len(results))
        col2.metric("E-mails Criados", total_created)

        df_results = pd.DataFrame(results)
        df_to_show = df_results[['empresa', 'data', 'valor', 'email', 'anexos_count']].rename(columns={
            'empresa': 'Empresa', 'data': 'Data', 'valor': 'Valor', 'email': 'E-mail', 'anexos_count': 'Anexos'
        })
        st.dataframe(df_to_show, use_container_width=True, hide_index=True)
        
        if st.button("🗑️ Limpar Resultados", key="limpar_results"):
            del st.session_state.results
            st.rerun()

def show_config_page() -> None:
    """Renderiza a página de configurações."""
    st.title("⚙️ Configurações do Sistema")
    
    st.info("Aqui você pode ajustar a estrutura das planilhas e o mapeamento de colunas para cada tipo de relatório. Os caminhos dos arquivos são montados automaticamente.")
    
    current_configs = config.load_configs()

    with st.form("config_form"):
        st.subheader("📋 Configurações dos Relatórios")
        
        tab_names = list(current_configs.keys())
        tabs = st.tabs(tab_names)
        
        for i, (report_type, config_data) in enumerate(current_configs.items()):
            with tabs[i]:
                st.subheader(f"Configurações para {report_type}")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    sheet_dados = st.text_input(
                        "Aba dos Dados",
                        value=config_data.get('sheet_dados', ''),
                        key=f"sheet_dados_{report_type}",
                        help="Nome da aba que contém os dados do relatório"
                    )
                    
                    sheet_contatos = st.text_input(
                        "Aba dos Contatos",
                        value=config_data.get('sheet_contatos', ''),
                        key=f"sheet_contatos_{report_type}",
                        help="Nome da aba que contém os contatos"
                    )
                
                with col2:
                    header_row = st.number_input(
                        "Linha do Cabeçalho",
                        value=int(config_data.get('header_row', 0)),
                        min_value=0,
                        key=f"header_row_{report_type}",
                        help="Número da linha que contém os cabeçalhos das colunas"
                    )
                
                data_columns = st.text_area(
                    "Mapeamento de Colunas",
                    value=config_data.get('data_columns', ''),
                    height=100,
                    key=f"data_columns_{report_type}",
                    help="Formato: NomeNoExcel:NomePadrão,OutraColuna:OutroNome"
                )
                
                current_configs[report_type].update({
                    'sheet_dados': sheet_dados,
                    'sheet_contatos': sheet_contatos,
                    'header_row': header_row,
                    'data_columns': data_columns
                })
        
        if st.form_submit_button("💾 Salvar Configurações"):
            try:
                config.save_configs(current_configs)
                st.success("✅ Configurações salvas com sucesso!")
            except Exception as e:
                st.error(f"❌ Erro ao salvar configurações: {e}")
                registrar_log(f"Erro ao salvar configurações: {e}")

    st.divider()
    st.subheader("🧩 Templates de E-mail")
    st.caption("Edite os templates usados para assunto, corpo e anexos. As alterações são salvas em email_templates.json.")
    try:
        templates_json = services.load_email_templates()
    except Exception as e:
        st.error(f"Erro ao carregar templates: {e}")
        templates_json = {}

    def json_dumps_pretty(obj: Any) -> str:
        try:
            return json.dumps(obj, ensure_ascii=False, indent=2)
        except Exception:
            return "{}"

    tab_names = list(templates_json.keys()) if templates_json else []
    if tab_names:
        tabs = st.tabs(tab_names)
        for i, key in enumerate(tab_names):
            with tabs[i]:
                st.markdown(f"**Report:** `{key}`")
                cfg = templates_json.get(key, {})
                has_variants = isinstance(cfg.get('variants'), dict)
                variant_keys = list(cfg.get('variants', {}).keys()) if has_variants else []
                target_label = 'Variante' if has_variants else 'Template'
                selected_variant = None
                if has_variants:
                    selected_variant = st.selectbox(target_label, variant_keys, key=f"var_{key}")
                    edit_block = cfg['variants'][selected_variant]
                else:
                    selected_variant = "default"
                    edit_block = cfg

                st.caption("Modo simples (HTML como no Outlook). Use Modo avançado para editar o JSON cru.")
                with st.expander("Modo avançado (JSON)"):
                    editable = st.text_area("JSON do Template", value=json_dumps_pretty(cfg), height=200, key=f"tpl_json_{key}_{selected_variant}")
                    if st.button("Salvar JSON", key=f"save_json_{key}_{selected_variant}"):
                        try:
                            parsed = json.loads(editable)
                            templates_json[key] = parsed
                            services.save_email_templates(templates_json)
                            st.success("JSON salvo.")
                        except Exception as e:
                            st.error(f"JSON inválido: {e}")

                col1, col2 = st.columns([2,1])
                with col1:
                    subj = st.text_input("Assunto (subject_template)", value=edit_block.get('subject_template', ''), key=f"subj_{key}_{selected_variant}")
                with col2:
                    send_mode = st.selectbox("Modo de envio", options=["display","send"], index=0 if (edit_block.get('send_mode','display').startswith('display')) else 1, key=f"send_{key}_{selected_variant}")

                body = st.text_area("Corpo do e-mail (HTML)", value=edit_block.get('body_html') or edit_block.get('body_html_credit') or edit_block.get('body_html_debit') or '', height=200, key=f"body_{key}_{selected_variant}")
                attachments_str = "\n".join(edit_block.get('attachments', []))
                attachments_edit = st.text_area("Anexos (um por linha)", value=attachments_str, height=100, key=f"att_{key}_{selected_variant}")

                if st.button("Salvar Template", key=f"save_simple_{key}_{selected_variant}"):
                    new_block = dict(edit_block)
                    new_block['subject_template'] = subj
                    if 'body_html_credit' in new_block or 'body_html_debit' in new_block:
                        new_block['body_html'] = body
                    else:
                        new_block['body_html'] = body
                    new_block['attachments'] = [ln.strip() for ln in attachments_edit.splitlines() if ln.strip()]
                    new_block['send_mode'] = send_mode
                    if has_variants:
                        cfg['variants'][selected_variant] = new_block
                        templates_json[key] = cfg
                    else:
                        templates_json[key] = new_block
                    try:
                        services.save_email_templates(templates_json)
                        st.success("Template salvo.")
                    except Exception as e:
                        st.error(f"Falha ao salvar: {e}")

                live_preview = st.checkbox("Prévia instantânea", value=False, key=f"liveprev_{key}_{selected_variant}")
                if live_preview:
                    sample = st.session_state.get('preview_data')
                    if sample is not None and not sample.empty:
                        row = sample.iloc[0].to_dict()
                        meses_map = {m.upper(): f"{i+1:02d}" for i, m in enumerate(config.MESES)}
                        common = {
                            'month_long': st.session_state.month.title() if 'month' in st.session_state else '',
                            'month_num': meses_map.get(st.session_state.month.upper(), '00') if 'month' in st.session_state else '',
                            'year': st.session_state.year if 'year' in st.session_state else '',
                        }
                        try:
                            rendered = services.render_email_from_template(key, row, common, auto_send=False)
                            components.html(f"<h4>{rendered['subject']}</h4>" + rendered['body'], height=350, scrolling=True)
                        except Exception as e:
                            st.error(f"Falha na renderização: {e}")
                if st.button("Pré-visualizar", key=f"prev_simple_{key}_{selected_variant}"):
                    sample = st.session_state.get('preview_data')
                    if sample is not None and not sample.empty:
                        row = sample.iloc[0].to_dict()
                        meses_map = {m.upper(): f"{i+1:02d}" for i, m in enumerate(config.MESES)}
                        common = {
                            'month_long': st.session_state.month.title() if 'month' in st.session_state else '',
                            'month_num': meses_map.get(st.session_state.month.upper(), '00') if 'month' in st.session_state else '',
                            'year': st.session_state.year if 'year' in st.session_state else '',
                        }
                        try:
                            rendered = services.render_email_from_template(key, row, common, auto_send=False)
                            components.html(f"<h4>{rendered['subject']}</h4>" + rendered['body'], height=350, scrolling=True)
                            if rendered['missing_placeholders']:
                                st.warning("Placeholders ausentes: " + ", ".join(rendered['missing_placeholders']))
                            if rendered['attachment_warnings']:
                                st.warning("\n".join(rendered['attachment_warnings']))
                        except Exception as e:
                            st.error(f"Falha na renderização: {e}")
                    else:
                        st.warning("Carregue dados na página 'Envio de Relatórios' para usar a prévia.")

def main() -> None:
    """Função principal da aplicação."""
    st.image("static/logo.png", width=250)

    st.sidebar.title("🧭 Navegação")
    page_options = ["Envio de Relatórios", "Configurações"]
    page = st.sidebar.radio("Escolha a página:", page_options, label_visibility="hidden", key="sidebar_radio")
    
    if page == "Envio de Relatórios":
        show_main_page()
    else:
        show_config_page()

if __name__ == "__main__":
    main()
    st.sidebar.info("Aplicação desenvolvida para automação de envio de e-mails DGCA.")
    st.sidebar.warning("Nota: Ao processar, janelas do Outlook podem abrir para sua revisão. Isso é esperado.")