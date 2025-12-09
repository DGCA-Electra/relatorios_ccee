import streamlit as st
import pandas as pd
import json
import logging
import string
from typing import Any
import src.config.config as config
from src.config.config_manager import load_configs, save_configs
from src.utils.file_utils import load_email_templates, save_email_templates
import src.services as services

def col_letter_to_index(letter: str) -> int:
    """Converte 'A' para 0, 'B' para 1, 'AA' para 26, etc."""
    letter = letter.upper().strip()
    num = 0
    for c in letter:
        if c in string.ascii_uppercase:
            num = num * 26 + (ord(c) - ord('A') + 1)
    return num - 1 if num > 0 else 0

def show_config_page() -> None:
    """Renderiza a página de configurações."""
    st.title("⚙️ Configurações do Sistema")
    
    st.info("Gerencie os relatórios existentes ou cadastre novos tipos de relatórios para o robô processar.")
    
    current_configs = load_configs()

    tab_edit, tab_new, tab_templates = st.tabs(["✏️ Editar Configurações", "➕ Criar Novo Relatório", "📧 Templates de E-mail"])

    with tab_edit:
        st.caption("Edite os caminhos e mapeamentos de relatórios que já estão cadastrados.")
        with st.form("config_form"):
            tab_names = list(current_configs.keys())
            if not tab_names:
                st.warning("Nenhum relatório configurado.")
            else:
                tabs = st.tabs(tab_names)
                
                for i, (report_type, config_data) in enumerate(current_configs.items()):
                    with tabs[i]:
                        st.subheader(f"Configurações para {report_type}")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            sheet_dados = st.text_input("Nome da Aba de Dados", value=config_data.get('sheet_dados', ''), key=f"sd_{report_type}")
                            sheet_contatos = st.text_input("Nome da Aba de Contatos", value=config_data.get('sheet_contatos', ''), key=f"sc_{report_type}")
                        with col2:
                            header_row_val = int(config_data.get('header_row', 0))
                            header_row = st.number_input("Linha do Cabeçalho (0 = Linha 1)", value=header_row_val, min_value=0, key=f"hr_{report_type}")
                        
                        data_columns = st.text_area("Mapeamento de Colunas (Texto)", value=config_data.get('data_columns', ''), height=70, key=f"dc_{report_type}", help="Legado: ColunaExcel:NomePadrao")
                        
                        current_configs[report_type].update({
                            'sheet_dados': sheet_dados,
                            'sheet_contatos': sheet_contatos,
                            'header_row': header_row,
                            'data_columns': data_columns
                        })
            
            if st.form_submit_button("💾 Salvar Alterações"):
                try:
                    save_configs(current_configs)
                    st.success("✅ Configurações atualizadas com sucesso!")
                except Exception as e:
                    st.error(f"❌ Erro ao salvar: {e}")

    with tab_new:
        st.header("Cadastrar Novo Relatório")
        st.markdown("Use este assistente para ensinar o robô a ler um novo arquivo Excel.")

        with st.form("new_report_form"):
            c_code, c_dummy = st.columns([1, 2])
            with c_code:
                new_code = st.text_input("Código do Relatório (Sigla)", placeholder="Ex: VENDA01", help="Use uma sigla única, sem espaços.").upper().strip()

            st.divider()
            st.subheader("1. Onde estão os dados?")
            c1, c2, c3 = st.columns(3)
            with c1:
                new_sheet_dados = st.text_input("Nome da Aba (Dados)", placeholder="Ex: Planilha1", help="Copie exatamente o nome da aba no Excel.")
            with c2:
                new_sheet_contatos = st.text_input("Nome da Aba (Contatos)", value="Planilha1", help="Onde estão os e-mails dos clientes?")
            with c3:
                excel_header_line = st.number_input("Em qual linha começa o cabeçalho?", min_value=1, value=1, help="Olhe no Excel o número da linha onde estão os títulos (Nome, Valor, etc).")

            st.divider()
            st.subheader("2. Relacione as Colunas (Mapeamento)")
            st.info("Diga qual coluna do seu Excel corresponde aos campos que o sistema precisa.")
            
            df_map_template = pd.DataFrame([
                {"Coluna no Excel": "Agente", "Campo no Sistema": "Empresa"},
                {"Coluna no Excel": "Valor Total", "Campo no Sistema": "Valor"},
                {"Coluna no Excel": "E-mail Contato", "Campo no Sistema": "Email"},
                {"Coluna no Excel": "Data Vcto", "Campo no Sistema": "Data"},
            ])
            
            column_config = {
                "Campo no Sistema": st.column_config.SelectboxColumn(
                    "Campo no Sistema",
                    help="Como o sistema deve entender essa coluna?",
                    options=["Empresa", "Valor", "Email", "Data", "ValorLiquidacao", "ValorLiquidado", "ValorInadimplencia", "Outro"],
                    required=True
                )
            }
            
            edited_map = st.data_editor(
                df_map_template, 
                column_config=column_config, 
                num_rows="dynamic", 
                use_container_width=True,
                key="editor_mapping"
            )

            st.divider()
            st.subheader("3. Onde salvar os arquivos?")
            st.markdown("Use as variáveis: `{ano}`, `{ano_mes}` (ex: 202501), `{mes_abrev}` (ex: jan).")
            path_excel = st.text_input("Caminho do Excel", value="{sharepoint_root}/{ano}/{ano_mes}/Relatorio.xlsx")
            path_pdf = st.text_input("Pasta dos PDFs (Opcional)", value="{sharepoint_root}/{ano}/{ano_mes}/PDFs")

            st.divider()
            st.subheader("4. Dados Soltos (Opcional)")
            st.markdown("Precisa pegar uma data ou valor que está fora da tabela principal? (Ex: Uma data na célula B3).")
            
            df_extra_template = pd.DataFrame([
                {"Nome da Variável": "", "Linha Excel": 1, "Coluna Excel (A, B...)": "A"},
            ])
            
            edited_extra = st.data_editor(
                df_extra_template, 
                num_rows="dynamic", 
                use_container_width=True,
                key="editor_extra"
            )

            submitted_new = st.form_submit_button("✨ Criar Relatório", type="primary")

            if submitted_new:
                if not new_code:
                    st.error("O Código do Relatório é obrigatório.")
                elif new_code in current_configs:
                    st.error(f"O relatório '{new_code}' já existe.")
                else:
                    map_list = []
                    for _, row in edited_map.iterrows():
                        col_excel = str(row["Coluna no Excel"]).strip()
                        col_sys = str(row["Campo no Sistema"]).strip()
                        if col_excel and col_sys:
                            map_list.append(f"{col_excel}:{col_sys}")
                    final_data_columns = ",".join(map_list)

                    extra_fields_list = []
                    for _, row in edited_extra.iterrows():
                        var_name = str(row["Nome da Variável"]).strip()
                        row_excel = int(row["Linha Excel"])
                        col_letter = str(row["Coluna Excel (A, B...)"]).strip()
                        
                        if var_name and col_letter:
                            extra_fields_list.append({
                                "name": var_name,
                                "row": row_excel - 1,
                                "col": col_letter_to_index(col_letter)
                            })

                    new_config = {
                        "sheet_dados": new_sheet_dados,
                        "sheet_contatos": new_sheet_contatos,
                        "header_row": excel_header_line - 1,
                        "data_columns": final_data_columns,
                        "path_template": {
                            "excel_dados": path_excel,
                            "pdfs_dir": path_pdf
                        },
                        "extra_fields": extra_fields_list
                    }
                    
                    try:
                        current_configs[new_code] = new_config
                        save_configs(current_configs)
                        templates = load_email_templates()
                        placeholders = ["empresa", "mes", "ano"] + [x['name'] for x in extra_fields_list]
                        templates[new_code] = {
                            "subject_template": f"{new_code} - Relatório - {{empresa}}",
                            "body_html": f"<p>Segue relatório referente a {new_code}.</p>",
                            "attachments": [],
                            "placeholders": placeholders,
                            "send_mode": "display"
                        }
                        save_email_templates(templates)

                        st.balloons()
                        st.success(f"Relatório '{new_code}' criado com sucesso! Recarregue a página para vê-lo na lista.")
                        
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")

    with tab_templates:
        st.caption("Edite os templates de e-mail (Assunto e Corpo HTML).")
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
                    st.markdown(f"**Editando:** `{key}`")
                    cfg = templates_json.get(key, {})
                    has_variants = isinstance(cfg.get('variants'), dict)
                    variant_keys = list(cfg.get('variants', {}).keys()) if has_variants else []
                    
                    col_sel, col_act = st.columns([3, 1])
                    selected_variant = "default"
                    edit_block = cfg
                    
                    with col_sel:
                        if has_variants:
                            selected_variant = st.selectbox(f"Variante ({key})", variant_keys, key=f"var_{key}")
                            edit_block = cfg['variants'][selected_variant]
                        else:
                            st.info("Este relatório usa um template único.")

                    with st.expander("🛠️ Modo Avançado (JSON)"):
                        editable = st.text_area("JSON Completo", value=json_dumps_pretty(cfg), height=200, key=f"json_{key}")
                        if st.button("Salvar JSON", key=f"btn_json_{key}"):
                            try:
                                parsed = json.loads(editable)
                                templates_json[key] = parsed
                                save_email_templates(templates_json)
                                st.success("JSON salvo.")
                            except Exception as e:
                                st.error(f"JSON inválido: {e}")

                    c_subj, c_mode = st.columns([3, 1])
                    with c_subj:
                        subj = st.text_input("Assunto do E-mail", value=edit_block.get('subject_template', ''), key=f"subj_{key}_{selected_variant}")
                    with c_mode:
                        mode_opts = ["display", "send"]
                        curr_mode = edit_block.get('send_mode', 'display')
                        curr_idx = 0 if 'display' in curr_mode else 1
                        send_mode = st.selectbox("Modo de Envio", options=mode_opts, index=curr_idx, key=f"sm_{key}_{selected_variant}", help="'Display' cria rascunho. 'Send' envia direto (cuidado).")

                    body = st.text_area("Corpo do E-mail (HTML)", value=edit_block.get('body_html') or edit_block.get('body_html_credit') or '', height=250, key=f"body_{key}_{selected_variant}", help="Você pode usar HTML básico aqui.")
                    
                    atts = edit_block.get('attachments', [])
                    attachments_str = "\n".join(atts) if isinstance(atts, list) else ""
                    attachments_edit = st.text_area("Anexos Adicionais (Links ou Caminhos - 1 por linha)", value=attachments_str, height=100, key=f"att_{key}_{selected_variant}")

                    if st.button("💾 Salvar Template", key=f"save_{key}_{selected_variant}"):
                        new_block = dict(edit_block)
                        new_block['subject_template'] = subj
                        new_block['body_html'] = body
                        new_block['attachments'] = [ln.strip() for ln in attachments_edit.splitlines() if ln.strip()]
                        new_block['send_mode'] = send_mode
                        
                        if has_variants:
                            cfg['variants'][selected_variant] = new_block
                        else:
                            cfg.update(new_block)
                            
                        templates_json[key] = cfg
                        try:
                            save_email_templates(templates_json)
                            st.success("Template salvo com sucesso!")
                        except Exception as e:
                            st.error(f"Erro ao salvar: {e}")