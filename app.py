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
    
    # Mostrar indicador de modo de teste se estiver ativo
    if st.session_state.get('analista_teste'):
        st.warning(f"🧪 Modo de teste ativo: Testando como {st.session_state['analista_teste']}")
        if st.button("Desativar modo de teste", key="desativar_teste_main"):
            del st.session_state['analista_teste']
            st.rerun()
    
    all_configs = config.load_configs()
    report_types = list(all_configs.keys())

    with st.form("report_form"):
        st.subheader("Parâmetros de Envio")
        col1, col2, col3 = st.columns(3)
        with col1:
            tipo = st.selectbox("Tipo de Relatório", options=report_types, key="sb_tipo")
        with col2:
            mes = st.selectbox("Mês", options=config.MESES, key="sb_mes")
        with col3:
            ano = st.selectbox("Ano", options=config.ANOS, key="sb_ano")
        
        # Dois botões separados
        col1, col2 = st.columns(2)
        with col1:
            preview_submitted = st.form_submit_button("👁️ Visualizar Dados", use_container_width=True)
        with col2:
            send_submitted = st.form_submit_button("📧 Enviar E-mails", use_container_width=True)

    # Usar analista de teste se estiver ativo, senão usar o login do usuário
    analista = st.session_state.get('analista_teste', st.session_state.get('login_usuario'))
    login_usuario = st.session_state.get('login_usuario')

    if not login_usuario:
        st.error("❌ Login do usuário não encontrado. Faça login novamente.")
        return

    # Processar visualização de dados
    if preview_submitted:
        with st.spinner("Carregando dados para visualização... Por favor, aguarde."):
            try:
                df_filtered, df_preview = services.preview_dados(
                    report_type=tipo, 
                    analyst=analista, 
                    month=mes, 
                    year=ano,
                    login_usuario=login_usuario
                )
                st.session_state.preview_data = df_filtered
                st.session_state.form_data = {'tipo': tipo, 'analista': analista, 'mes': mes, 'ano': ano}
                
                st.success(f'✅ Dados carregados com sucesso! {len(df_filtered)} empresas encontradas para {analista}.')

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
        if 'preview_data' not in st.session_state:
            st.error("❌ Primeiro visualize os dados antes de enviar os e-mails.")
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
                    analyst=analista, 
                    month=mes, 
                    year=ano,
                    login_usuario=login_usuario
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
        form = st.session_state.form_data

        st.header(f"📈 Dados para {form['tipo']} - {form['mes']}/{form['ano']} - {form['analista']}")

        total_empresas = len(df_filtered)

        col1, col2 = st.columns(2)
        col1.metric("Empresas Encontradas", total_empresas)
        col2.metric("Analista", form['analista'])


        # Preparar dados para exibição
        df_to_show = df_filtered.copy()

        # Selecionar colunas relevantes para o tipo de relatório
        display_cols = REPORT_DISPLAY_COLUMNS.get(form['tipo'], ["Empresa", "Email", "Valor"])
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
        form = st.session_state.form_data
        
        st.header(f"📤 Resultado do Envio - {form['tipo']} - {form['mes']}/{form['ano']} - {form['analista']}")
        
        total_processed = len(results)
        total_created = results[-1]['created_count'] if results else 0
        
        col1, col2 = st.columns(2)
        col1.metric("Empresas Processadas", total_processed)
        col2.metric("E-mails Criados", total_created)

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

def show_test_page() -> None:
    """Renderiza a página de teste de analistas."""
    st.title("🧪 Teste de Analistas")
    
    st.info("Use esta página para testar o sistema como se fosse outro analista. Isso é útil para verificar se os dados estão sendo carregados corretamente para diferentes analistas.")
    
    # Mostrar status atual
    if st.session_state.get('analista_teste'):
        st.success(f"✅ **Modo de teste ativo:** {st.session_state['analista_teste']}")
    
    # Interface de teste
    st.markdown("---")
    st.subheader("🎯 Selecionar Analista para Teste")
    
    # Layout melhorado para seleção
    col1, col2 = st.columns([3, 1])
    with col1:
        analista_teste = st.selectbox(
            "Selecione o analista:",
            options=config.ANALISTAS,
            key="analista_teste_select",
            help="Escolha o analista cujo contexto você quer simular"
        )
    with col2:
        st.markdown("")  # Espaçamento para alinhar com o selectbox
        st.markdown("")  # Espaçamento adicional
        if st.button("🔬 Ativar Teste", key="testar_analista_btn", use_container_width=True):
            st.session_state['analista_teste'] = analista_teste
            st.success(f"✅ Modo de teste ativado para: {analista_teste}")
            st.info("Agora você pode usar a aba 'Envio de Relatórios' para testar como se fosse este analista.")
            st.rerun()
    
    # Desativar teste
    if st.session_state.get('analista_teste'):
        st.markdown("---")
        st.subheader("🛑 Gerenciar Modo de Teste")
        
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.warning(f"**Modo de teste ativo:** {st.session_state['analista_teste']}")
        with col2:
            st.markdown("")  # Espaçamento para alinhar
            st.markdown("")  # Espaçamento adicional
            if st.button("❌ Desativar Teste", key="desativar_teste_btn", use_container_width=True):
                del st.session_state['analista_teste']
                st.success("✅ Modo de teste desativado.")
                st.rerun()
        with col3:
            st.markdown("")  # Espaçamento para alinhar
            st.markdown("")  # Espaçamento adicional
            if st.button("🔄 Limpar Cache", key="limpar_cache_btn", use_container_width=True):
                st.session_state.clear()
                st.success("✅ Cache limpo. Faça login novamente.")
                st.rerun()

def validate_login(login_usuario: str) -> bool:
    """
    Valida se o login do usuário é válido.
    
    Args:
        login_usuario: Login do usuário
        
    Returns:
        True se o login é válido, False caso contrário
    """
    if not login_usuario or not login_usuario.strip():
        return False
    
    # Verificar se o login tem formato válido (nome.sobrenome)
    if '.' not in login_usuario:
        return False
    
    return True

def get_user_paths(login_usuario: str) -> Dict[str, str]:
    """
    Obtém os caminhos do usuário usando a nova função do config.
    
    Args:
        login_usuario: Login do usuário
        
    Returns:
        Dicionário com os caminhos do usuário
    """
    try:
        return config.get_user_paths(login_usuario)
    except Exception as e:
        st.error(f"Erro ao obter caminhos do usuário: {e}")
        return {}

def main() -> None:
    """Função principal da aplicação."""
    st.image("static/logo.png", width=250)

    # Tela de login
    if 'login_usuario' not in st.session_state or not st.session_state['login_usuario']:
        with st.form("login_form"):
            st.subheader("🔐 Login")
            login_usuario = st.text_input(
                "Informe seu login de rede (ex: nome.sobrenome)",
                value='',
                key="login_usuario_input_form",
                help="Digite apenas o seu login de rede, sem @dominio."
            )
            submitted = st.form_submit_button("Entrar")
            
        if submitted and login_usuario:
            login_clean = login_usuario.strip().lower()
            
            if not validate_login(login_clean):
                st.error("❌ Formato de login inválido. Use o formato: nome.sobrenome")
            else:
                st.session_state['login_usuario'] = login_clean
                
                # Obter e validar caminhos do usuário
                user_paths = get_user_paths(login_clean)
                if user_paths:
                    st.session_state['raiz_sharepoint'] = user_paths.get('raiz_sharepoint', '')
                    st.session_state['contratos_email_path'] = user_paths.get('contratos_email_path', '')
                    st.rerun()
                else:
                    st.error("❌ Não foi possível configurar os caminhos do usuário.")
        
        st.stop()

    # Botão de logout com key único
    if st.sidebar.button("🚪 Logout", key="logout_btn"):
        st.session_state['login_usuario'] = ''
        st.rerun()

    # Verificar se os caminhos estão configurados
    login_usuario = st.session_state.get('login_usuario')
    if not st.session_state.get('raiz_sharepoint') or not st.session_state.get('contratos_email_path'):
        user_paths = get_user_paths(login_usuario)
        if user_paths:
            st.session_state['raiz_sharepoint'] = user_paths.get('raiz_sharepoint', '')
            st.session_state['contratos_email_path'] = user_paths.get('contratos_email_path', '')

    # Se for admin, mostra navegação
    if st.session_state['login_usuario'] == 'malik.mourad':
        st.sidebar.title("🧭 Navegação")
        page_options = ["Envio de Relatórios", "Configurações", "Teste de Analistas"]
        page = st.sidebar.radio("Escolha a página:", page_options, label_visibility="hidden", key="sidebar_radio")
        
        if page == "Envio de Relatórios":
            show_main_page()
        elif page == "Configurações":
            show_config_page()
        else:  # Teste de Analistas
            show_test_page()
    else:
        # Usuário comum só vê a tela principal, sem navegação lateral
        show_main_page()

if __name__ == "__main__":
    main()

st.sidebar.info("Aplicação desenvolvida para automação de envio de e-mails DGCA.")
st.sidebar.warning("Nota: Ao processar, janelas do Outlook podem abrir para sua revisão. Isso é esperado.")