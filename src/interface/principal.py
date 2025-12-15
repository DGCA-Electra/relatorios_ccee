import streamlit as st
import pandas as pd
import logging
import streamlit.components.v1 as components
from datetime import datetime

import src.configuracoes.constantes as config
from src.configuracoes.gerenciador import carregar_configuracoes
import src.servicos as services
from src.utilitarios.tabelas import tratar_valores_df

def iniciar_estado_sessao():
    mes_atual_idx = datetime.now().month - 1
    mes_nome = config.MESES[mes_atual_idx] if 0 <= mes_atual_idx < len(config.MESES) else config.MESES[0]
    padroes = {
        "tipo_relatorio": "GFN001",
        "analista": "Artur Bello Rodrigues",
        "mes": mes_nome,
        "ano": str(datetime.now().year)
    }
    for k, v in padroes.items():
        if k not in st.session_state:
            st.session_state[k] = v

def unir_emails_seguro(campo_email):
    if not campo_email:
        return ""
    if isinstance(campo_email, list):
        return "; ".join(e.strip() for e in campo_email if e)
    return "; ".join([e.strip() for e in str(campo_email).split(';') if e.strip()])

def exibir_pagina_principal() -> None:
    """Renderiza a página principal de envio de relatórios."""
    todas_configuracoes = carregar_configuracoes() 
    tipos_relatorio = list(todas_configuracoes.keys())
    iniciar_estado_sessao()
    
    tipo = st.session_state.tipo_relatorio
    analista_final = st.session_state.analista
    mes = st.session_state.mes
    ano = st.session_state.ano
    
    st.title("⚡Envio de Relatórios CCEE - DGCA")
    st.info("💡 **Dica:** Você pode enviar relatórios para qualquer analista. Isso é útil durante férias ou ausências.")
    st.header("Parâmetros de Envio")
    c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
    with c1:
        st.session_state.tipo_relatorio = st.selectbox("Tipo de Relatório", options=tipos_relatorio, index=tipos_relatorio.index(tipo) if tipo in tipos_relatorio else 0)
    with c2:
        st.session_state.analista = st.selectbox("Analista", options=config.ANALISTAS, index=config.ANALISTAS.index(analista_final) if analista_final in config.ANALISTAS else 0)
    with c3:
        st.session_state.mes = st.selectbox("Mês", options=config.MESES, index=config.MESES.index(mes) if mes in config.MESES else 0)
    with c4:
        st.session_state.ano = st.selectbox("Ano", options=config.ANOS, index=config.ANOS.index(str(ano)) if str(ano) in config.ANOS else 0)
    
    col1, col2 = st.columns(2)
    
    if col1.button("📊 Visualizar Dados", use_container_width=True):
        st.session_state.gatilho_previa = True
    
    if col2.button("📧 Enviar E-mails", use_container_width=True, type="primary"):
        if "ms_token" not in st.session_state or not st.session_state["ms_token"].get("access_token"):
            st.warning("Por favor, faça o login com sua conta Microsoft para enviar e-mails.")
        else:
            st.session_state.gatilho_envio = True
            st.rerun()

    if st.session_state.get("gatilho_envio"):
        token_acesso = st.session_state.get("ms_token", {}).get("access_token")
        if token_acesso:
            with st.spinner("Criando rascunhos na sua caixa de e-mail... Aguarde."):
                try:
                    resultados = services.informa_processos(
                        tipo_relatorio=tipo,
                        analista=analista_final,
                        mes=mes,
                        ano=str(ano)
                    )
                    st.session_state.resultados = resultados
                    st.success(f"✅ Rascunhos criados com sucesso na sua caixa de e-mail para {len(resultados)} empresas.")
                except services.ErroProcessamento as e:
                    st.error(f"❌ Erro no processamento: {e}")
                except Exception as e:
                     st.error(f"❌ Ocorreu um erro inesperado durante o envio: {e}")
                     logging.exception("Erro inesperado durante informa_processos:")
        else:
            st.error("Erro: Sua sessão expirou ou o login falhou. Por favor, faça login novamente.")
        st.session_state.gatilho_envio = False

    if st.session_state.get("gatilho_previa"):
        with st.spinner("Carregando dados para visualização... Por favor, aguarde."):
            try:
                df_filtrado, config_previa_dados = services.visualizar_previa_dados(
                    tipo_relatorio=tipo, 
                    analista=analista_final, 
                    mes=mes, 
                    ano=str(ano)
                )
                st.session_state.dados_previa_brutos = df_filtrado
                st.session_state.config_previa = config_previa_dados
                st.session_state.dados_formulario = {'tipo': tipo, 'analista': analista_final, 'mes': mes, 'ano': ano}
                st.success(f'✅ Dados carregados com sucesso! {len(df_filtrado)} empresas encontradas para {analista_final}.')
            except services.ErroProcessamento as e:
                st.error(f"❌ Erro de processamento: {e}")
        st.session_state.gatilho_previa = False

    def exibir_previa_email(assunto: str, corpo_html: str):
        html = f"""
        <html><head><style>
          body {{ color: black; font-family: Inter, Arial, sans-serif; }}
          @media (prefers-color-scheme: dark) {{ body {{ color: white; }} }}
        </style></head><body><h4>{assunto}</h4><hr/>{corpo_html}</body></html>
        """
        components.html(html, height=400, scrolling=True)

    if 'dados_previa_brutos' in st.session_state:
        df_bruto = st.session_state.dados_previa_brutos
        cfg = st.session_state.get('config_previa', {})
        if not df_bruto.empty:
            st.subheader(f"Dados para {tipo} - {mes}/{ano} - {analista_final}")
            df_exibicao = tratar_valores_df(df_bruto.copy())
            st.dataframe(df_exibicao.reset_index(drop=True), use_container_width=True)
            
            st.subheader("Pré-visualização do E-mail")
            limite_visualizacao = min(5, len(df_bruto))
            for idx in range(limite_visualizacao):
                dados_empresa = df_bruto.iloc[idx].to_dict()
                if 'Email' in dados_empresa:
                    dados_empresa['Email'] = unir_emails_seguro(dados_empresa['Email'])
                    
                dados_comuns = {
                    'mes_long': mes.title(),
                    'mes_num': {m.upper(): f"{i+1:02d}" for i, m in enumerate(config.MESES)}.get(mes.upper(), '00'),
                    'ano': str(ano),
                    'analista': analista_final,
                }
                
                try:
                    renderizado = services.renderizar_email_modelo(tipo, dados_empresa, dados_comuns, cfg)
                    with st.expander(f"Prévia #{idx+1} - {dados_empresa.get('Empresa','')}", expanded=False):
                        exibir_previa_email(renderizado['assunto'], renderizado['corpo'])
                except Exception as e:
                    st.warning(f"Falha ao renderizar template para {dados_empresa.get('Empresa','')}: {e}")

        if st.button("🗑️ Limpar Visualização", key="limpar_preview"):
            del st.session_state.dados_previa_brutos
            if 'config_previa' in st.session_state: del st.session_state.config_previa
            st.rerun()

    if 'resultados' in st.session_state and st.session_state.resultados:
        resultados = st.session_state.resultados
        formulario = st.session_state.get('dados_formulario', {})
        st.header(f"📤 Resultado do Envio - {formulario.get('tipo', 'N/A')} - {formulario.get('mes', 'N/A')}/{formulario.get('ano', 'N/A')}")
        
        total_criados = resultados[-1]['contagem_criados'] if resultados else 0
        
        col1, col2 = st.columns(2)
        col1.metric("Empresas Processadas", len(resultados))
        col2.metric("E-mails Criados", total_criados)
        
        df_resultados = pd.DataFrame(resultados)
        colunas_base = ['empresa', 'email', 'anexos_count']
        
        nomes_exibicao = {
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
        
        colunas_especificas_relatorio = {
            'SUM001': ['data_liquidacao', 'valor', 'situacao'],
            'LFN001': ['data', 'ValorLiquidacao', 'ValorLiquidado', 'ValorInadimplencia'],
            'GFN001': ['dataaporte', 'valor'],
            'LFRES001': ['data', 'valor'],
            'LFRCAP001': ['dataaporte', 'valor'],
            'RCAP002': ['dataaporte', 'valor']
        }
        
        tipo_relatorio = st.session_state.tipo_relatorio
        colunas_especificas = colunas_especificas_relatorio.get(tipo_relatorio, ['data', 'valor'])
        colunas_para_mostrar = colunas_base + colunas_especificas
        
        colunas_existentes = [col for col in colunas_para_mostrar if col in df_resultados.columns]
        df_exibicao = df_resultados[colunas_existentes].rename(columns={
            col: nomes_exibicao.get(col, col) for col in colunas_existentes
        })
        
        st.dataframe(df_exibicao, use_container_width=True, hide_index=True)

        if st.button("🗑️ Limpar Resultados", key="limpar_resultados"):
            del st.session_state.resultados
            st.rerun()