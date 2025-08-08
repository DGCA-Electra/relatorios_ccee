import streamlit as st
import pandas as pd
import config
import services
import logging
from typing import Dict, Any, Optional
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
    """Renderiza a página principal de envio de relatórios."""
    st.title("📊 Envio de Relatórios CCEE - DGCA")
    
    st.info("💡 **Dica:** Você pode enviar relatórios para qualquer analista. Isso é útil durante férias ou ausências, quando um analista precisa enviar relatórios para outro.")
    
    all_configs = config.load_configs()
    report_types = list(all_configs.keys())

    with st.form("report_form"):
        st.subheader("Parâmetros de Envio")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            tipo = st.selectbox("Tipo de Relatório", options=report_types, key="sb_tipo")
        with col2:
            analista = st.selectbox("Analista", options=config.ANALISTAS, key="sb_analista")
        with col3:
            mes = st.selectbox("Mês", options=config.MESES, key="sb_mes")
        with col4:
            ano = st.selectbox("Ano", options=config.ANOS, key="sb_ano")
        
        # Dois botões separados
        col1, col2 = st.columns(2)
        with col1:
            preview_submitted = st.form_submit_button("👁️ Visualizar Dados", use_container_width=True)
        with col2:
            send_submitted = st.form_submit_button("📧 Enviar E-mails", use_container_width=True)

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
                st.session_state.preview_data = df_filtered
                st.session_state.form_data = {'tipo': tipo, 'analista': analista_final, 'mes': mes, 'ano': ano}
                
                st.success(f'✅ Dados carregados com sucesso! {len(df_filtered)} empresas encontradas para {analista_final}.')

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
                # Tratamento global para valores nulos, NaN e zero antes do envio de e-mail
                df_email = st.session_state.preview_data.copy()
                for col in df_email.columns:
                    if ("valor" in col.lower()) or (col in ["ValorLiquidacao", "ValorLiquidado", "ValorInadimplencia", "Valor"]):
                        df_email[col] = df_email[col].apply(lambda x: "0" if pd.isna(x) or x in [None, 0, 0.0, "0", "0.0", "nan", "None"] else services._format_currency(x))
                st.session_state.preview_data = df_email

                results = services.process_reports(
                    report_type=tipo, 
                    analyst=analista_final, 
                    month=mes, 
                    year=ano
                )
                st.session_state.results = results
                created_count = results[-1]['created_count'] if results else 0
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

        total_empresas = len(df_filtered)

        col1, col2 = st.columns(2)
        col1.metric("Empresas Encontradas", total_empresas)
        col2.metric("Analista", form.get('analista', 'N/A'))


        # Preparar dados para exibição
        df_to_show = df_filtered.copy()

        # Selecionar colunas relevantes para o tipo de relatório
        display_cols = REPORT_DISPLAY_COLUMNS.get(form.get('tipo', ''), ["Empresa", "Email", "Valor"])
        columns_to_show = [col for col in display_cols if col in df_to_show.columns]
        df_display = df_to_show[columns_to_show].copy()

        # Renomear colunas para exibição (opcional, se quiser nomes amigáveis)
        rename_map = {
            "Empresa": "Empresa",
            "Email": "E-mail",
            "Valor": "Valor",
            "ValorLiquidacao": "Valor a Liquidar",
            "ValorLiquidado": "Valor Liquidado",
            "ValorInadimplencia": "Inadimplência",
            "Data": "Data",
            "Data_Debito_Credito": "Data Débito/Crédito"
        }
        df_display.columns = [rename_map.get(col, col) for col in df_display.columns]

        # Tratamento global para valores nulos, NaN e zero em colunas de valor
        for col in df_display.columns:
            if ("valor" in col.lower()) or (col in ["Valor a Liquidar", "Valor Liquidado", "Inadimplência"]):
                df_display[col] = df_display[col].apply(lambda x: "0" if pd.isna(x) or x in [None, 0, 0.0, "0", "0.0", "nan", "None"] else services._format_currency(x))

        st.dataframe(df_display, use_container_width=True, hide_index=True)

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
                'empresa': 'Empresa', 'data': 'Data Aporte', 'valor': 'Valor',
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