import streamlit as st
import pandas as pd
import json
import logging
import string
import apps.relatorios_ccee.configuracoes.constantes as config
import apps.relatorios_ccee.model.servicos as services

from typing import Any
from apps.relatorios_ccee.configuracoes.gerenciador import carregar_configuracoes, salvar_configuracoes
from apps.relatorios_ccee.model.arquivos import carregar_templates_email, salvar_templates_email

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
    current_configs = carregar_configuracoes()
    tab_edit, tab_new, tab_templates = st.tabs(["✏️ Editar Configurações", "➕ Criar Novo Relatório", "📧 Templates de E-mail"])
    with tab_edit:
        st.caption("Edite os caminhos e mapeamentos de relatórios que já estão cadastrados.")
        with st.form("config_form"):
            tab_names = list(current_configs.keys())
            if not tab_names:
                st.warning("Nenhum relatório configurado.")
            else:
                tabs = st.tabs(tab_names)
                for i, (tipo_relatorio, config_data) in enumerate(current_configs.items()):
                    with tabs[i]:
                        st.subheader(f"Configurações para {tipo_relatorio}")
                        col1, col2 = st.columns(2)
                        with col1:
                            planilha_dados = st.text_input("Nome da Aba de Dados", value=config_data.get('planilha_dados', ''), key=f"sd_{tipo_relatorio}")
                            planilha_contatos = st.text_input("Nome da Aba de Contatos", value=config_data.get('planilha_contatos', ''), key=f"sc_{tipo_relatorio}")
                        with col2:
                            linha_cabecalho_val = int(config_data.get('linha_cabecalho', 0))
                            linha_cabecalho = st.number_input("Linha do Cabeçalho (0 = Linha 1)", value=linha_cabecalho_val, min_value=0, key=f"hr_{tipo_relatorio}")
                        colunas_dados = st.text_area("Mapeamento de Colunas (Texto)", value=config_data.get('colunas_dados', ''), height=70, key=f"dc_{tipo_relatorio}", help="Legado: ColunaExcel:NomePadrao")
                        current_configs[tipo_relatorio].update({
                            'planilha_dados': planilha_dados,
                            'planilha_contatos': planilha_contatos,
                            'linha_cabecalho': linha_cabecalho,
                            'colunas_dados': colunas_dados
                        })
            if st.form_submit_button("💾 Salvar Alterações"):
                try:
                    salvar_configuracoes(current_configs)
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
            config_colunas = {
                "Campo no Sistema": st.column_config.SelectboxColumn(
                    "Campo no Sistema",
                    help="Como o sistema deve entender essa coluna?",
                    options=["Empresa", "Valor", "Email", "Data", "ValorLiquidacao", "ValorLiquidado", "ValorInadimplencia", "Outro"],
                    required=True
                )
            }
            mapa_editado = st.data_editor(
                df_map_template, 
                column_config=config_colunas, 
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
                    for _, row in mapa_editado.iterrows():
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
                        "planilha_dados": new_sheet_dados,
                        "planilha_contatos": new_sheet_contatos,
                        "linha_cabecalho": excel_header_line - 1,
                        "colunas_dados": final_data_columns,
                        "modelo_caminho": {
                            "excel_dados": path_excel,
                            "diretorio_pdfs": path_pdf
                        },
                        "extra_fields": extra_fields_list
                    }
                    try:
                        current_configs[new_code] = new_config
                        salvar_configuracoes(current_configs)
                        templates = carregar_templates_email()
                        variaveis = ["empresa", "mes", "ano"] + [x['name'] for x in extra_fields_list]
                        templates[new_code] = {
                            "assunto_template": f"{new_code} - Relatório - {{empresa}}",
                            "corpo_html": f"<p>Segue relatório referente a {new_code}.</p>",
                            "anexos": [],
                            "variaveis": variaveis,
                            "modo_envio": "display"
                        }
                        salvar_templates_email(templates)
                        st.balloons()
                        st.success(f"Relatório '{new_code}' criado com sucesso! Recarregue a página para vê-lo na lista.")
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")
    with tab_templates:
        st.caption("Edite os templates de e-mail (Assunto e Corpo HTML).")
        try:
            templates_json = carregar_templates_email()
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
                    tem_variantes = isinstance(cfg.get('variantes'), dict)
                    chaves_variantes = list(cfg.get('variantes', {}).keys()) if tem_variantes else []
                    col_sel, col_act = st.columns([3, 1])
                    variante_selecionada = "default"
                    bloco_edicao = cfg
                    with col_sel:
                        if tem_variantes:
                            variante_selecionada = st.selectbox(f"Variante ({key})", chaves_variantes, key=f"var_{key}")
                            bloco_edicao = cfg['variantes'][variante_selecionada]
                        else:
                            st.info("Este relatório usa um template único.")
                    with st.expander("🛠️ Modo Avançado (JSON)"):
                        editable = st.text_area("JSON Completo", value=json_dumps_pretty(cfg), height=200, key=f"json_{key}")
                        if st.button("Salvar JSON", key=f"btn_json_{key}"):
                            try:
                                parsed = json.loads(editable)
                                templates_json[key] = parsed
                                salvar_templates_email(templates_json)
                                st.success("JSON salvo.")
                            except Exception as e:
                                st.error(f"JSON inválido: {e}")
                    c_subj, c_mode = st.columns([3, 1])
                    with c_subj:
                        subj = st.text_input("Assunto do E-mail", value=bloco_edicao.get('assunto_template', ''), key=f"subj_{key}_{variante_selecionada}")
                    with c_mode:
                        opcoes_modo = ["display", "send"]
                        modo_atual = bloco_edicao.get('modo_envio', 'display')
                        idx_atual = 0 if 'display' in modo_atual else 1
                        modo_envio = st.selectbox("Modo de Envio", options=opcoes_modo, index=idx_atual, key=f"sm_{key}_{variante_selecionada}", help="'Display' cria rascunho. 'Send' envia direto (cuidado).")
                    corpo = st.text_area("Corpo do E-mail (HTML)", value=bloco_edicao.get('corpo_html') or bloco_edicao.get('corpo_html_credit') or '', height=250, key=f"body_{key}_{variante_selecionada}", help="Você pode usar HTML básico aqui.")
                    anexos = bloco_edicao.get('anexos', [])
                    anexos_str = "\n".join(anexos) if isinstance(anexos, list) else ""
                    anexos_edit = st.text_area("Anexos Adicionais (Links ou Caminhos - 1 por linha)", value=anexos_str, height=100, key=f"att_{key}_{variante_selecionada}")
                    if st.button("💾 Salvar Template", key=f"save_{key}_{variante_selecionada}"):
                        novo_bloco = dict(bloco_edicao)
                        novo_bloco['assunto_template'] = subj
                        novo_bloco['corpo_html'] = corpo
                        novo_bloco['anexos'] = [ln.strip() for ln in anexos_edit.splitlines() if ln.strip()]
                        novo_bloco['modo_envio'] = modo_envio
                        if tem_variantes:
                            cfg['variantes'][variante_selecionada] = novo_bloco
                        else:
                            cfg.update(novo_bloco)
                        templates_json[key] = cfg
                        try:
                            salvar_templates_email(templates_json)
                            st.success("Template salvo com sucesso!")
                        except Exception as e:
                            st.error(f"Erro ao salvar: {e}")