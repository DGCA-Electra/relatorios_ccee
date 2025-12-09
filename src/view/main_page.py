import streamlit as st
import pandas as pd
import logging
import src.config.config as config
from src.config.config_manager import load_configs
import src.services as services
from src.utils.dataframe_utils import tratar_valores_df
import streamlit.components.v1 as components
from datetime import datetime

def init_state():
    mes_atual_idx = datetime.now().month - 1
    mes_nome = config.MESES[mes_atual_idx] if 0 <= mes_atual_idx < len(config.MESES) else config.MESES[0]

    defaults = {
        "report_type": "GFN001",
        "analyst": "Artur Bello Rodrigues",
        "month": mes_nome,
        "year": str(datetime.now().year)
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
    all_configs = load_configs() 
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
        if "ms_token" not in st.session_state or not st.session_state["ms_token"].get("access_token"):
            st.warning("Por favor, faça o login com sua conta Microsoft para enviar e-mails.")
        else:
            st.session_state.send_trigger = True
            st.rerun()

    if st.session_state.get("send_trigger"):
        access_token = st.session_state.get("ms_token", {}).get("access_token")
        if access_token:
            with st.spinner("Criando rascunhos na sua caixa de e-mail... Aguarde."):
                try:
                    results = services.process_reports(
                        report_type=tipo,
                        analyst=analista_final,
                        month=mes,
                        year=str(ano)
                    )
                    st.session_state.results = results
                    st.success(f"✅ Rascunhos criados com sucesso na sua caixa de e-mail para {len(results)} empresas.")
                except services.ReportProcessingError as e:
                    st.error(f"❌ Erro no processamento: {e}")
                except Exception as e:
                     st.error(f"❌ Ocorreu um erro inesperado durante o envio: {e}")
                     logging.exception("Erro inesperado durante process_reports:")
        else:
            st.error("Erro: Sua sessão expirou ou o login falhou. Por favor, faça login novamente.")

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
                    rendered = services.render_email_from_template(tipo, dados_empresa, common, cfg)
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
        
        base_columns = ['empresa', 'email', 'anexos_count']
        display_names = {
            'empresa': 'Empresa',
            'email': 'E-mail',
            'anexos_count': 'Anexos',
            'data': 'Data',
            'valor': 'Valor',
            'data_liquidacao': 'Data Liquidação',
            'dataaporte': 'Data Aporte',
            'ValorLiquidacao': 'Valor Liquidação',
            'ValorLiquidado': 'Valor Liquidado',
            'ValorInadimplencia': 'Valor Inadimplência',
            'situacao': 'Situação'
        }
        
        report_specific_columns = {
            'SUM001': ['data_liquidacao', 'valor', 'situacao'],
            'LFN001': ['data', 'ValorLiquidacao', 'ValorLiquidado', 'ValorInadimplencia'],
            'GFN001': ['dataaporte', 'valor'],
            'LFRES001': ['data', 'valor'],
            'LFRCAP001': ['dataaporte', 'valor'],
            'RCAP002': ['dataaporte', 'valor']
        }
        
        report_type = st.session_state.report_type
        specific_columns = report_specific_columns.get(report_type, ['data', 'valor'])
        
        columns_to_show = base_columns + specific_columns
        
        existing_columns = [col for col in columns_to_show if col in df_results.columns]
        df_to_show = df_results[existing_columns].rename(columns={
            col: display_names.get(col, col) for col in existing_columns
        })
        
        st.dataframe(df_to_show, use_container_width=True, hide_index=True)
        
        if st.button("🗑️ Limpar Resultados", key="limpar_results"):
            del st.session_state.results
            st.rerun()