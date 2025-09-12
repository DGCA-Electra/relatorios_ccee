import streamlit as st
import pandas as pd
import config
import services
from mail.email_utils import montar_corpo_email, enviar_email
import logging
import openpyxl
from typing import Dict, Any, Optional
import os
from utils.dataframe_utils import tratar_valores_df
from config import REPORT_DISPLAY_COLUMNS

# Configuração básica de logging
logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

# Função para registrar logs
def registrar_log(mensagem: str) -> None:
    """Registra uma mensagem no log."""
    logging.info(mensagem)

st.set_page_config(
    page_title="Envio de Relatórios CCEE",
    page_icon="static/icon.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

def show_main_page() -> None:
    # Carregar template HTML para e-mail
    with open('templates/email_template.html', 'r', encoding='utf-8') as f:
        email_template = f.read()
    """Renderiza a página principal de envio de relatórios."""
    st.title("📊 Envio de Relatórios CCEE - DGCA")
    
    st.info("💡 **Dica:** Você pode enviar relatórios para qualquer analista. Isso é útil durante férias ou ausências, quando um analista precisa enviar relatórios para outro.")
    
    all_configs = config.load_configs()
    report_types = list(all_configs.keys())
        import streamlit.components.v1 as components
        from jinja2 import Environment, FileSystemLoader

        # Inicializa st.session_state com valores padrão
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

        init_state()

    with st.form("report_form"):
        st.subheader("Parâmetros de Envio")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            tipo_form = st.selectbox("Tipo de Relatório", options=report_types, key="form_tipo")
        with col2:
            analista_form = st.selectbox("Analista", options=config.ANALISTAS, key="form_analista")
        with col3:
            mes_form = st.selectbox("Mês", options=config.MESES, key="form_mes")
        with col4:
            ano_form = st.selectbox("Ano", options=config.ANOS, key="form_ano")
        # Dois botões separados
        col1, col2 = st.columns(2)
        with col1:
            preview_submitted = st.form_submit_button("👁️ Visualizar Dados", use_container_width=True)
        with col2:
            send_submitted = st.form_submit_button("📧 Enviar E-mails", use_container_width=True)

        # Inputs organizados no sidebar

        # Botões de ação

        # Pré-visualização de e-mail
        if 'preview_data' in st.session_state and st.session_state.preview_data is not None:
            df_preview = st.session_state.preview_data
            if not df_preview.empty:
                st.subheader("Pré-visualização do E-mail")
                dados_empresa = df_preview.iloc[0].to_dict()
                with open('templates/email_template.html', 'r', encoding='utf-8') as f:
                    email_template = f.read()
                from mail.email_utils import montar_corpo_email
                corpo_preview = montar_corpo_email(dados_empresa, email_template)
                st.markdown(corpo_preview, unsafe_allow_html=True)

        def render_main_parameters():
            st.header("Parâmetros de Envio")
            c1, c2, c3, c4 = st.columns([2,2,2,1])
            with c1:
                st.session_state.report_type = st.selectbox("Tipo de Relatório", options=report_types, index=report_types.index(st.session_state.report_type) if st.session_state.report_type in report_types else 0)
            with c2:
                st.session_state.analyst = st.selectbox("Analista", options=config.ANALISTAS, index=config.ANALISTAS.index(st.session_state.analyst) if st.session_state.analyst in config.ANALISTAS else 0)
            with c3:
                st.session_state.month = st.selectbox("Mês", options=config.MESES, index=config.MESES.index(st.session_state.month) if st.session_state.month in config.MESES else 0)
            with c4:
                st.session_state.year = st.selectbox("Ano", options=config.ANOS, index=config.ANOS.index(st.session_state.year) if st.session_state.year in config.ANOS else 0)
            col1, col2 = st.columns(2)
            with col1:
                if st.button("👁️ Visualizar Dados", use_container_width=True):
                    st.session_state.preview_trigger = True
            with col2:
                if st.button("📧 Enviar E-mails", use_container_width=True):
                    st.session_state.send_trigger = True

        render_main_parameters()

        # Sidebar apenas navegação
        with st.sidebar:
            st.title("🧭 Navegação")
            page_options = ["Envio de Relatórios", "Configurações"]
            page = st.radio("Menu", page_options)
            st.markdown("---")
            st.caption("Aplicativo para automação de envio de relatórios")

        # Usar o analista selecionado
        analista_final = st.session_state.analyst
        tipo = st.session_state.report_type
        mes = st.session_state.month
        ano = st.session_state.year

        # Verificar se o analista é válido
        if not analista_final or analista_final not in config.ANALISTAS:
            st.error("❌ Analista inválido. Selecione um analista válido.")
            return

        # Verificar se o tipo de relatório é válido
        if not tipo or tipo not in report_types:
            st.error("❌ Tipo de relatório inválido. Selecione um tipo válido.")
            return

        # Verificar se o mês e ano são válidos
        if not mes or mes not in config.MESES:
            st.error("❌ Mês inválido. Selecione um mês válido.")
            return

        if not ano or ano not in config.ANOS:
            st.error("❌ Ano inválido. Selecione um ano válido.")
            return

        # Função para sanitizar e-mails
        def safe_join_emails(email_field):
            if not email_field:
                return ""
            if isinstance(email_field, list):
                return "; ".join(e.strip() for e in email_field if e)
            return "; ".join([e.strip() for e in str(email_field).split(';') if e.strip()])

        # Função para renderizar pré-visualização do e-mail
        def render_email_preview(context):
            env = Environment(loader=FileSystemLoader("templates"))
            template = env.get_template("email_template.html")
            html = template.render(**context)
            components.html(html, height=400, scrolling=True)

        # Processar visualização de dados
        if st.session_state.get("preview_trigger"):
            with st.spinner("Carregando dados para visualização... Por favor, aguarde."):
                try:
                    df_filtered, df_preview = services.preview_dados(
                        report_type=tipo, 
                        analyst=analista_final, 
                        month=mes, 
                        year=ano
                    )
                    df_filtered = tratar_valores_df(df_filtered)
                    st.session_state.preview_data = df_filtered
                    st.session_state.form_data = {'tipo': tipo, 'analista': analista_final, 'mes': mes, 'ano': ano}
                    st.success(f'✅ Dados carregados com sucesso! {len(df_filtered)} empresas encontradas para {analista_final}.')
                except services.ReportProcessingError as e:
                    st.error(f"❌ Erro de processamento: {e}")
            st.session_state.preview_trigger = False

        # Visualização dos dados
        if 'preview_data' in st.session_state and st.session_state.preview_data is not None:
            df_preview = st.session_state.preview_data
            if not df_preview.empty:
                title = f"Dados para {tipo} - {mes}/{ano} - {analista_final}"
                st.subheader(title)
                st.dataframe(df_preview.reset_index(drop=True), use_container_width=True)
                # Pré-visualização do e-mail
                st.subheader("Pré-visualização do E-mail")
                dados_empresa = df_preview.iloc[0].to_dict()
                # Sanitizar campos
                for k in dados_empresa:
                    if dados_empresa[k] is None:
                        dados_empresa[k] = "N/A"
                if 'Email' in dados_empresa:
                    dados_empresa['Email'] = safe_join_emails(dados_empresa['Email'])
                render_email_preview(dados_empresa)
    # Usar o analista selecionado
    analista_final = analista

    # Verificar se o analista é válido
    if not analista_final or analista_final not in config.ANALISTAS:
        st.error("❌ Analista inválido. Selecione um analista válido.")
        return

    # Verificar se o tipo de relatório é válido
    if not tipo or tipo not in report_types:
        st.error("❌ Tipo de relatório inválido. Selecione um tipo válido.")
        return

    # Verificar se o mês e ano são válidos
    if not mes or mes not in config.MESES:
        st.error("❌ Mês inválido. Selecione um mês válido.")
        return

    if not ano or ano not in config.ANOS:
        st.error("❌ Ano inválido. Selecione um ano válido.")
        return



    # Processar visualização de dados
    if preview_submitted:
        with st.spinner("Carregando dados para visualização... Por favor, aguarde."):
            try:
                df_filtered, df_preview = services.preview_dados(
                    report_type=tipo, 
                    analyst=analista_final, 
                    month=mes, 
                    year=ano
                )
                # DRY: tratamento centralizado
                df_filtered = tratar_valores_df(df_filtered)
                st.session_state.preview_data = df_filtered
                st.session_state.form_data = {'tipo': tipo, 'analista': analista_final, 'mes': mes, 'ano': ano}
                st.success(f'✅ Dados carregados com sucesso! {len(df_filtered)} empresas encontradas para {analista_final}.')
            except services.ReportProcessingError as e:
                st.error(f"❌ Erro de processamento: {e}")
            except FileNotFoundError as e:
                st.error(f"❌ Arquivo não encontrado: {e}")
                st.info("💡 Verifique se os caminhos dos arquivos estão corretos e se os arquivos existem.")
            except PermissionError as e:
                st.error(f"❌ {str(e)}")
                st.info("💡 Tente fechar outros programas que possam estar usando o arquivo e aguarde a sincronização completa do OneDrive/SharePoint.")
            except ValueError as e:
                st.error(f"❌ Erro de configuração: {e}")
                st.info("💡 Verifique as configurações do relatório na aba 'Configurações'.")
            except Exception as e:
                st.error(f"❌ Erro inesperado: {e}")
                registrar_log(f"Erro inesperado em preview: {e}")


    # Processar envio de e-mails
    if send_submitted:
        if 'preview_data' not in st.session_state or st.session_state.preview_data is None:
            st.error("❌ Primeiro visualize os dados antes de enviar os e-mails.")
            return
        # Verificar se os dados de preview correspondem aos parâmetros atuais
        form_data = st.session_state.get('form_data', {})
        if (form_data.get('tipo') != tipo or 
            form_data.get('analista') != analista_final or 
            form_data.get('mes') != mes or 
            form_data.get('ano') != ano):
            st.error("❌ Os dados de visualização não correspondem aos parâmetros atuais. Visualize os dados novamente.")
            return
        with st.spinner("Processando relatórios e gerando e-mails... Por favor, aguarde."):
            try:
                # DRY: tratamento centralizado
                df_email = st.session_state.preview_data.copy()
                df_email = tratar_valores_df(df_email)
                st.session_state.preview_data = df_email
                # Montar e enviar e-mails para cada empresa, com tratamento de exceções/logs
                results = []
                for idx, row in df_email.iterrows():
                    dados_empresa = row.to_dict()
                    corpo = montar_corpo_email(dados_empresa, email_template)
                    try:
                        enviar_email(
                            destinatario=dados_empresa.get('Email', ''),
                            assunto=f"Relatório {tipo} - {dados_empresa.get('Empresa', '')}",
                            corpo=corpo,
                            anexos=[] # Adicionar lógica de anexos se necessário
                        )
                        status = 'sucesso'
                        erro = ''
                    except Exception as e:
                        status = 'falha'
                        erro = str(e)
                        registrar_log(f"Erro ao enviar e-mail para {dados_empresa.get('Empresa', '')}: {e}")
                        st.error(f"❌ Falha ao enviar e-mail para {dados_empresa.get('Empresa', '')}: {e}")
                    results.append({
                        'empresa': dados_empresa.get('Empresa', ''),
                        'data': dados_empresa.get('Data', ''),
                        'valor': dados_empresa.get('Valor', ''),
                        'email': dados_empresa.get('Email', ''),
                        'anexos_count': 0,
                        'created_count': idx+1,
                        'status': status,
                        'erro': erro
                    })
                st.session_state.results = results
                created_count = sum(1 for r in results if r['status'] == 'sucesso')
                st.success(f'✅ {created_count} de {len(results)} e-mails foram gerados com sucesso! Verifique seu Outlook.')
            except services.ReportProcessingError as e:
                st.error(f"❌ Erro de processamento: {e}")
            except FileNotFoundError as e:
                st.error(f"❌ Arquivo não encontrado: {e}")
                st.info("💡 Verifique se os caminhos dos arquivos estão corretos e se os arquivos existem.")
            except ValueError as e:
                st.error(f"❌ Erro de configuração: {e}")
                st.info("💡 Verifique as configurações do relatório na aba 'Configurações'.")
            except Exception as e:
                st.error(f"❌ Erro inesperado: {e}")
                registrar_log(f"Erro inesperado em envio: {e}")


    # Mostrar dados de visualização
    if 'preview_data' in st.session_state and st.session_state.preview_data is not None:
        df_filtered = st.session_state.preview_data
        form = st.session_state.get('form_data', {})
        if not form:
            st.error("❌ Dados do formulário não encontrados.")
            return
        st.header(f"📈 Dados para {form.get('tipo', 'N/A')} - {form.get('mes', 'N/A')}/{form.get('ano', 'N/A')} - {form.get('analista', 'N/A')}")
        col1, col2 = st.columns(2)
        col2.metric("Analista", form.get('analista', 'N/A'))

        # DRY: lógica consolidada para LFRES001/LFRES002
        if form.get('tipo') in ['LFRES001', 'LFRES002']:
            # Adiciona colunas se não existirem
            if 'Situacao' not in df_filtered.columns:
                df_filtered['Situacao'] = ''
            if 'Valor' not in df_filtered.columns:
                df_filtered['Valor'] = 0
            if 'TipoAgente' not in df_filtered.columns:
                df_filtered['TipoAgente'] = df_filtered.apply(
                    lambda row: 'Gerador-EER' if row.get('Empresa','') == 'PCH PONTE BRANCA' else 'Consumidor',
                    axis=1
                )
            # Data para LFRES001
            if form.get('tipo') == 'LFRES001':
                data_debito = ''
                try:
                    paths = config.build_report_paths(form.get('tipo'), form.get('ano'), form.get('mes'))
                    excel_path = paths.get('excel_dados')
                    if excel_path and os.path.exists(excel_path):
                        wb = openpyxl.load_workbook(excel_path, data_only=True)
                        ws = wb.active
                        cell_value = ws.cell(row=27, column=1).value
                        if cell_value:
                            data_debito = str(cell_value)
                except Exception as e:
                    st.warning(f"Não foi possível extrair a data do aporte: {e}")
                df_filtered['Data'] = data_debito if data_debito else ''
                df_filtered['Data_Formatada'] = df_filtered['Data'].apply(lambda x: services._format_date(x) if x else '')
            else:
                if 'Data' in df_filtered.columns:
                    df_filtered['Data_Formatada'] = df_filtered['Data'].apply(lambda x: services._format_date(x) if pd.notna(x) and str(x).strip() != '' else '')
                else:
                    df_filtered['Data_Formatada'] = ''
            # Selecionar colunas relevantes
            display_cols = ["Empresa", "Email", "TipoAgente", "Valor", "Data_Formatada", "Situacao"]
            columns_to_show = [col for col in display_cols if col in df_filtered.columns]
            df_display = df_filtered[columns_to_show].copy() if columns_to_show else pd.DataFrame()
            rename_map = {
                "Empresa": "Empresa",
                "Email": "E-mail",
                "Valor": "Valor (R$)",
                "Data_Formatada": "Data",
                "TipoAgente": "Tipo do Agente",
                "Situacao": "Situação"
            }
            df_display.columns = [rename_map.get(col, col) for col in df_display.columns]
            # Métricas
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total de Empresas", len(df_filtered))
            with col2:
                total_valor = df_filtered['Valor'].sum() if 'Valor' in df_filtered.columns else 0
                st.metric("Valor Total", services._format_currency(total_valor))
            with col3:
                situacao = df_filtered['Situacao'].iloc[0] if not df_filtered.empty else "Não informado"
                st.metric("Situação", situacao)
            # Detalhes por tipo de agente
            if 'TipoAgente' in df_filtered.columns:
                st.subheader("📊 Detalhes por Tipo de Agente")
                for tipo in df_filtered['TipoAgente'].unique():
                    df_tipo = df_filtered[df_filtered['TipoAgente'] == tipo]
                    with st.expander(f"{tipo} ({len(df_tipo)} empresas)"):
                        valor_total = df_tipo['Valor'].sum()
                        st.markdown(f"**Valor Total**: {services._format_currency(valor_total)}")
                        if not df_tipo.empty:
                            st.markdown(f"**Situação**: {df_tipo['Situacao'].iloc[0]}")
                            if 'Data' in df_tipo.columns:
                                st.markdown("**Data**: " + services._format_date(df_tipo['Data'].iloc[0]))
            # Função para destacar a situação
            def highlight_situation(val):
                if val == "Débito":
                    return 'color: #ff4b4b'
                elif val == "Crédito":
                    return 'color: #28a745'
                return ''
            styles = [
                dict(selector="th", props=[("font-size", "12px"), ("text-align", "left"), ("padding", "5px"), ("background-color", "#f0f2f6"), ("font-weight", "bold")]),
                dict(selector="td", props=[("font-size", "11px"), ("text-align", "left"), ("padding", "5px"), ("white-space", "nowrap")]),
                dict(selector="tr:hover", props=[("background-color", "#f5f5f5")]),
            ]
            styled_df = df_display.style.set_table_styles(styles)
            if "Situação" in df_display.columns:
                styled_df = styled_df.applymap(highlight_situation, subset=['Situação'])
            st.dataframe(styled_df, use_container_width=True, hide_index=True)
        # Botão para limpar dados de visualização
        if st.button("🗑️ Limpar Visualização", key="limpar_preview"):
            del st.session_state.preview_data
            st.rerun()

    # Mostrar resultados de envio
    if 'results' in st.session_state and st.session_state.results:
        results = st.session_state.results
        form = st.session_state.get('form_data', {})
        
        if not form:
            st.error("❌ Dados do formulário não encontrados.")
            return
        
        st.header(f"📤 Resultado do Envio - {form.get('tipo', 'N/A')} - {form.get('mes', 'N/A')}/{form.get('ano', 'N/A')} - {form.get('analista', 'N/A')}")
        
        total_processed = len(results)
        total_created = results[-1]['created_count'] if results else 0
        
        col1, col2 = st.columns(2)
        col1.metric("Empresas Processadas", total_processed)
        col2.metric("E-mails Criados", total_created)

        if results:
            df_results = pd.DataFrame(results)
            df_to_show = df_results[['empresa', 'data', 'valor', 'email', 'anexos_count']].rename(columns={
                'empresa': 'Empresa', 'data': 'Data', 'valor': 'Valor',
                'email': 'E-mail', 'anexos_count': 'Anexos'
            })
            st.dataframe(df_to_show, use_container_width=True, hide_index=True)
        
        # Botão para limpar resultados
        if st.button("🗑️ Limpar Resultados", key="limpar_results"):
            del st.session_state.results
            st.rerun()

def show_config_page() -> None:
    """Renderiza a página de configurações."""
    st.title("⚙️ Configurações do Sistema")
    
    # Informações principais em uma linha
    st.info("Aqui você pode ajustar a estrutura das planilhas e o mapeamento de colunas para cada tipo de relatório. Os caminhos dos arquivos são montados automaticamente.")
    
    current_configs = config.load_configs()

    with st.form("config_form"):
        # Agrupar configurações por categoria
        st.subheader("📋 Configurações dos Relatórios")
        
        # Usar tabs para organizar melhor
        tab_names = list(current_configs.keys())
        tabs = st.tabs(tab_names)
        
        for i, (report_type, config_data) in enumerate(current_configs.items()):
            with tabs[i]:
                st.subheader(f"Configurações para {report_type}")
                
                # Campos de configuração
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
                
                # Mapeamento de colunas
                data_columns = st.text_area(
                    "Mapeamento de Colunas",
                    value=config_data.get('data_columns', ''),
                    height=100,
                    key=f"data_columns_{report_type}",
                    help="Formato: NomeNoExcel:NomePadrão,OutraColuna:OutroNome"
                )
                
                # Atualizar configuração
                current_configs[report_type].update({
                    'sheet_dados': sheet_dados,
                    'sheet_contatos': sheet_contatos,
                    'header_row': header_row,
                    'data_columns': data_columns
                })
        
        # Botão de salvar
        if st.form_submit_button("💾 Salvar Configurações"):
            try:
                config.save_configs(current_configs)
                st.success("✅ Configurações salvas com sucesso!")
            except Exception as e:
                st.error(f"❌ Erro ao salvar configurações: {e}")
                registrar_log(f"Erro ao salvar configurações: {e}")

def main() -> None:
    """Função principal da aplicação."""
    st.image("static/logo.png", width=250)

    # Navegação principal
    st.sidebar.title("🧭 Navegação")
    page_options = ["Envio de Relatórios", "Configurações"]
    page = st.sidebar.radio("Escolha a página:", page_options, label_visibility="hidden", key="sidebar_radio")
    
    if page == "Envio de Relatórios":
        show_main_page()
    else:  # Configurações
        show_config_page()

if __name__ == "__main__":
    main()
    st.sidebar.info("Aplicação desenvolvida para automação de envio de e-mails DGCA.")
    st.sidebar.warning("Nota: Ao processar, janelas do Outlook podem abrir para sua revisão. Isso é esperado.")