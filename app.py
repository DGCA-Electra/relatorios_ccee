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

    # ... resto do fluxo de pré-visualização e envio ...

def show_config_page():
    """Renderiza a página de configurações."""
    st.title("⚙️ Configurações do Sistema")
    st.info("Aqui você pode ajustar a estrutura das planilhas e o mapeamento de colunas para cada tipo de relatório. Os caminhos dos arquivos são montados automaticamente.")
    
    current_configs = config.load_configs()

    with st.form("config_form"):
        for tipo, cfg in current_configs.items():
            with st.expander(f"📋 Configurações para: {tipo}"):
                st.subheader(f"Estrutura das Planilhas - {tipo}")
                col1, col2, col3 = st.columns([2, 2, 1])
                with col1:
                    cfg['sheet_dados'] = st.text_input(f"Aba Dados", value=cfg.get('sheet_dados', ''), key=f"sheetd_{tipo}")
                with col2:
                    cfg['sheet_contatos'] = st.text_input(f"Aba Contatos", value=cfg.get('sheet_contatos', ''), key=f"sheetc_{tipo}")
                with col3:
                    cfg['header_row'] = st.number_input(f"Linha Cabeçalho (inicia em 0)", min_value=0, value=cfg.get('header_row', 0), key=f"header_{tipo}")

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
                label_mapeamento = f"Mapeamento de Colunas (Exemplo: {exemplo})"
                cfg['data_columns'] = st.text_area(label_mapeamento, value=cfg.get('data_columns', ''), key=f"map_{tipo}", height=100)

        submitted = st.form_submit_button("💾 Salvar Todas as Configurações", use_container_width=True)

    if submitted:
        config.save_configs(current_configs)
        st.success("Configurações salvas com sucesso!")
        st.balloons()

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
        page_options = ["Envio de Relatórios", "Configurações"]
        page = st.sidebar.radio("Escolha a página:", page_options, label_visibility="hidden", key="sidebar_radio")
        if page == "Envio de Relatórios":
            show_main_page()
        else:
            show_config_page()
    else:
        # Usuário comum só vê a tela principal, sem navegação lateral
        show_main_page()

if __name__ == "__main__":
    main()

st.sidebar.info("Aplicação desenvolvida para automação de envio de e-mails DGCA.")
st.sidebar.warning("Nota: Ao processar, janelas do Outlook podem abrir para sua revisão. Isso é esperado.")