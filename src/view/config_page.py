import streamlit as st
import json
import logging
from typing import Any
import streamlit.components.v1 as components
import src.config.config as config
from src.config.config_manager import load_configs, save_configs
from src.utils.file_utils import load_email_templates, save_email_templates
import src.services as services

def show_config_page() -> None:
    """Renderiza a página de configurações."""
    st.title("⚙️ Configurações do Sistema")
    
    st.info("Aqui você pode ajustar a estrutura das planilhas e o mapeamento de colunas para cada tipo de relatório.")
    
    current_configs = load_configs()

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
                save_configs(current_configs)
                st.success("✅ Configurações salvas com sucesso!")
            except Exception as e:
                st.error(f"❌ Erro ao salvar configurações: {e}")
                logging.error(f"Erro ao salvar configurações: {e}")

    st.divider()
    st.subheader("🧩 Templates de E-mail")
    st.caption("Edite os templates usados para assunto, corpo e anexos.")
    try:
        templates_json = load_email_templates()
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
                            save_email_templates(templates_json)
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
                        save_email_templates(templates_json)
                        st.success("Template salvo.")
                    except Exception as e:
                        st.error(f"Falha ao salvar: {e}")

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