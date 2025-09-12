"""
data/excel_utils.py
Funções utilitárias para leitura e manipulação de arquivos Excel.
"""
import pandas as pd
import openpyxl
from pathlib import Path

def ler_dados_excel(excel_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    """
    Carrega dados de uma planilha Excel.
    """
    import logging
    if not Path(excel_path).exists():
        logging.error(f"Arquivo não encontrado: {excel_path}")
        raise FileNotFoundError(f"Arquivo não encontrado: {excel_path}")
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row)
        return df
    except Exception as e:
        logging.error(f"Erro ao ler Excel {excel_path}: {e}")
        raise

def ler_celula_excel(excel_path: str, row: int, col: int) -> str:
    """
    Lê o valor de uma célula específica de um arquivo Excel.
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    return str(ws.cell(row=row, column=col).value)
