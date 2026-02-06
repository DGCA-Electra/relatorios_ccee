import pandas as pd
from pathlib import Path
from typing import Dict, Any, Optional
import json
import logging

TEMPLATES_JSON_PATH = "src/configuracoes/email_templates.json"
class ErroProcessamento(Exception):
    pass

def ler_dados_excel(caminho_excel: str, nome_planilha: str, linha_cabecalho: int) -> pd.DataFrame:

    """Carrega dados de uma planilha Excel."""

    if not Path(caminho_excel).exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_excel}")
    if linha_cabecalho == -1:
        return pd.read_excel(Path(caminho_excel), sheet_name=nome_planilha, header=None)
    return pd.read_excel(Path(caminho_excel), sheet_name=nome_planilha, header=linha_cabecalho)

def encontrar_anexo(diretorio_pdf: str, nome_arquivo: str) -> Optional[Path]:

    """Procura por um arquivo PDF no diretório especificado."""

    caminho_anexo = Path(diretorio_pdf) / nome_arquivo
    if caminho_anexo.exists():
        return caminho_anexo
    logging.warning(f"Anexo não encontrado no caminho principal: {caminho_anexo}")
    return None

def carregar_templates_email() -> Dict[str, Any]:

    try:
        with open(TEMPLATES_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise ErroProcessamento(f"Falha ao carregar {TEMPLATES_JSON_PATH}: {e}")

def salvar_templates_email(dados: Dict[str, Any]) -> None:

    """Salva os templates de e-mail no arquivo JSON."""
    try:
        with open(TEMPLATES_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise ErroProcessamento(f"Falha ao salvar {TEMPLATES_JSON_PATH}: {e}")