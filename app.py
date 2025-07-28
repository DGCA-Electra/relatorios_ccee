import streamlit as st
import pandas as pd
import config
import services
import logging

# Configuração básica de logging
logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

# Função para registrar logs
def registrar_log(mensagem):
    logging.info(mensagem)

st.set_page_config(
    page_title="Envio de Relatórios CCEE",
    page_icon="static/icon.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

def show_main_page():
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
        submitted = st.form_submit_button("Pré-visualizar Dados")

    # Usar analista de teste se estiver ativo, senão usar o login do usuário
    analista = st.session_state.get('analista_teste', st.session_state.get('login_usuario'))

    if submitted:
        with st.spinner("Processando relatórios e gerando e-mails... Por favor, aguarde."):
            try:
                results = services.process_reports(
                    report_type=tipo, analyst=analista, month=mes, year=ano
                )
                st.session_state.results = results
                st.session_state.form_data = {'tipo': tipo, 'analista': analista, 'mes': mes, 'ano': ano}
                
                created_count = results[-1]['created_count'] if results else 0
                st.success(f'{created_count} de {len(results)} e-mails foram gerados com sucesso! Verifique seu Outlook.')

            except (FileNotFoundError, services.ReportProcessingError, Exception) as e:
                st.error(f"Ocorreu um erro: {e}")
                if 'results' in st.session_state:
                    del st.session_state.results

    if 'results' in st.session_state and st.session_state.results:
        results = st.session_state.results
        form = st.session_state.form_data
        
        st.header(f"📈 Resultado para {form['tipo']} - {form['mes']}/{form['ano']} - {form['analista']}")
        
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

def show_config_page():
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
        
        for i, (tipo, cfg) in enumerate(current_configs.items()):
            with tabs[i]:
                st.markdown(f"**Configurações para: {tipo}**")
                
                # Estrutura das planilhas
                st.markdown("##### 📊 Estrutura das Planilhas")
                col1, col2, col3 = st.columns(3)
                with col1:
                    cfg['sheet_dados'] = st.text_input(
                        "Aba Dados", 
                        value=cfg.get('sheet_dados', ''), 
                        key=f"sheetd_{tipo}",
                        help="Nome da aba que contém os dados do relatório"
                    )
                with col2:
                    cfg['sheet_contatos'] = st.text_input(
                        "Aba Contatos", 
                        value=cfg.get('sheet_contatos', ''), 
                        key=f"sheetc_{tipo}",
                        help="Nome da aba que contém os contatos de email"
                    )
                with col3:
                    cfg['header_row'] = st.number_input(
                        "Linha Cabeçalho", 
                        min_value=0, 
                        value=cfg.get('header_row', 0), 
                        key=f"header_{tipo}",
                        help="Número da linha onde está o cabeçalho (inicia em 0)"
                    )

                # Mapeamento de colunas
                st.markdown("##### 🗺️ Mapeamento de Colunas")
                exemplos_mapeamento = {
                    "GFN001": "Agente:Empresa,Garantia Avulsa (R$):Valor",
                    "SUM001": "Agente:Empresa,Garantia Avulsa (R$):Valor",
                    "LFN001": "Agente:Empresa,Débito/Crédito:Situacao,Valor a Liquidar (R$):ValorLiquidacao,Valor Liquidado (R$):ValorLiquidado,Inadimplência (R$):ValorInadimplencia",
                    "LFRES": "Agente:Empresa,Data do Débito:Data,Valor do Débito (R$):Valor,Tipo Agente:TipoAgente",
                    "LEMBRETE": "Agente:Empresa,Garantia Avulsa (R$):Valor",
                    "LFRCAP": "Agente:Empresa,Data do Débito:Data,Valor do Débito (R$):Valor",
                    "RCAP": "Agente:Empresa,Data:Data,Valor do Débito (R$):Valor"
                }
                exemplo = exemplos_mapeamento.get(tipo, "NomeNoExcel:NomePadrao,...")
                st.caption(f"💡 Exemplo: {exemplo}")
                cfg['data_columns'] = st.text_area(
                    "Mapeamento de Colunas", 
                    value=cfg.get('data_columns', ''), 
                    key=f"map_{tipo}", 
                    height=80,
                    help="Formato: NomeNoExcel:NomePadrao,NomeNoExcel2:NomePadrao2"
                )

        # Botão de salvar centralizado
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            submitted = st.form_submit_button("💾 Salvar Todas as Configurações", use_container_width=True)

    if submitted:
        config.save_configs(current_configs)
        st.success("✅ Configurações salvas com sucesso!")
        st.balloons()

def show_test_page():
    """Renderiza a página de teste para administrador."""
    st.title("🧪 Teste de Envio como Outros Analistas")
    
    # Informações principais
    st.info("Como administrador, você pode testar o envio de relatórios como se fosse outro analista, sem precisar fazer logout.")
    
    # Status do teste atual
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

def main():
    st.image("static/logo.png", width=250)

    # Tela de login
    if 'login_usuario' not in st.session_state or not st.session_state['login_usuario']:
        with st.form("login_form"):
            st.subheader("Identificação do Usuário")
            login_usuario = st.text_input(
                "Informe seu login de rede (ex: malik.sobrenome)",
                value='',
                key="login_usuario_input_form",
                help="Digite apenas o seu login de rede, sem @dominio."
            )
            submitted = st.form_submit_button("Entrar")
        if submitted and login_usuario:
            st.session_state['login_usuario'] = login_usuario.strip().lower()
            st.rerun()
        st.stop()

    # Botão de logout com key único
    if st.sidebar.button("Logout", key="logout_btn"):
        st.session_state['login_usuario'] = ''
        st.rerun()

    # Montar diretório raiz do SharePoint automaticamente
    raiz_sharepoint = f"C:/Users/{st.session_state['login_usuario']}/ELECTRA COMERCIALIZADORA DE ENERGIA S.A/GE - ECE/DGCA/DGA/CCEE/Relatórios CCEE"
    st.session_state['raiz_sharepoint'] = raiz_sharepoint

    # Se for admin, mostra navegação
    if st.session_state['login_usuario'] == 'malik.mourad':
        st.sidebar.title("Navegação")
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