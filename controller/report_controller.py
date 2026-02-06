import logging
import streamlit as st
from apps.relatorios_ccee.model import servicos
from typing import Any
from typing import List, Dict, Any, Tuple
from apps.relatorios_ccee.model.arquivos import ErroProcessamento
from apps.relatorios_ccee.configuracoes.constantes import MESES


def criar_rascunhos(tipo_relatorio: str, analista: str, mes: str, ano: str) -> List[Dict[str, Any]]:
    """Orquestra o processamento de relatórios e criação de rascunhos via Graph.

    Args:
        tipo_relatorio: Código do relatório (ex: 'GFN001').
        analista: Nome do analista.
        mes: Nome do mês.
        ano: Ano.

    Returns:
        Lista de dicionários com resultados por empresa.

    Raises:
        ErroProcessamento: Se ocorrer erro no processamento.
    """
    token_acesso = st.session_state.get("ms_token", {}).get("access_token")
    user_info = st.session_state.get("user_info")
    if not token_acesso:
        logging.error("Tentativa de envio sem token de acesso presente na sessão.")
        raise ErroProcessamento("Usuário não autenticado. Faça login para enviar e-mails.")
    try:
        resultados = servicos.informa_processos(tipo_relatorio, analista, mes, ano, token_acesso, user_info=user_info)
        return resultados
    except ErroProcessamento:
        raise
    except Exception as e:
        logging.exception("Erro inesperado em criar_rascunhos:")
        raise ErroProcessamento(f"Erro inesperado ao criar rascunhos: {e}")


def visualizar_previa(tipo_relatorio: str, analista: str, mes: str, ano: str) -> Tuple[Any, Dict[str, Any]]:
    """Carrega dados para pré-visualização sem realizar efeitos colaterais.

    Retorna o DataFrame filtrado e a configuração utilizada.
    """
    user_info = st.session_state.get("user_info")
    try:
        df, cfg = servicos.visualizar_previa_dados(tipo_relatorio, analista, mes, ano, user_info=user_info)
        return df, cfg
    except ErroProcessamento:
        raise
    except Exception as e:
        logging.exception("Erro inesperado em visualizar_previa:")
        raise ErroProcessamento(f"Erro inesperado ao carregar pré-visualização: {e}")


def _build_dados_comuns(analista: str, mes: str, ano: str) -> Dict[str, Any]:
    """Cria o dicionário `dados_comuns` usado para renderização no Model.

    Mantém a lógica de padronização de mês/ano centralizada no controller.
    """
    from apps.relatorios_ccee.configuracoes.constantes import MESES
    return {
        "analista": analista,
        "mes_long": mes.title(),
        "mes_num": {m.upper(): f"{i+1:02d}" for i, m in enumerate(MESES)}.get(mes.upper(), "??"),
        "ano": ano
    }


def renderizar_email_preview(tipo_relatorio: str, row: Dict[str, Any], analista: str, mes: str, ano: str, config: Dict[str, Any]) -> Dict[str, Any]:
    """Renderiza o template de e-mail para exibição no front-end (apenas formatação, sem efeitos colaterais).

    Args:
        tipo_relatorio: Código do relatório.
        row: Dicionário com dados da linha (empresa).
        dados_comuns: Dados comuns como mes/ano/analista.
        config: Configuração do relatório.

    Returns:
        Dicionário com chaves: 'assunto', 'corpo', 'anexos', 'variaveis_ausentes', 'final_data'.

    Raises:
        ErroProcessamento: Se ocorrer falha ao renderizar.
    """
    dados_comuns = _build_dados_comuns(analista, mes, ano)
    try:
        resultado = servicos.renderizar_email_modelo(tipo_relatorio, row, dados_comuns, config)
        return resultado
    except Exception as e:
        logging.exception("Erro ao renderizar preview de e-mail:")
        raise ErroProcessamento(f"Falha ao renderizar preview do e-mail: {e}")
