# seu_projeto/config.py

import json
from pathlib import Path

CONFIG_FILE = Path('config_relatorios.json')
ANALISTAS = [
    'Artur Bello Rodrigues', 'Camila Padovan Baptista', 'Cassiana Unruh',
    'Isabela Loredo', 'Jorge Ferreira', 'Tiago Padilha Foletto'
]
MESES = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
         'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
ANOS = [str(y) for y in range(2025, 2030)]

DEFAULT_CONFIGS = {
    "GFN001": {
        "sheet_dados": "GFN003 - Garantia Financeira po",
        "sheet_contatos": "Planilha1",
        "header_row": 30,
        "data_columns": "Agente:Empresa,Garantia Avulsa (R$):Valor"
    },
    "SUM001": {
        "sheet_dados": "GFN003 - Garantia Financeira po",
        "sheet_contatos": "Planilha1",
        "header_row": 30,
        "data_columns": "Agente:Empresa,Garantia Avulsa (R$):Valor"
    },
    "LFN001": {
        "sheet_dados": "LFN004 - Liquidação Financeira ",
        "sheet_contatos": "Planilha1",
        "header_row": 31,
        "data_columns": "Agente:Empresa,Débito/Crédito:Situacao,Valor a Liquidar (R$):ValorLiquidacao,Valor Liquidado (R$):ValorLiquidado,Inadimplência (R$):ValorInadimplencia"
    },
    "LFRES": {
        "sheet_dados": "LFRES002 - Liquidação de Energi",
        "sheet_contatos": "Planilha1",
        "header_row": 42,
        "data_columns": "Agente:Empresa,Data do Débito:Data,Valor do Débito (R$):Valor,Tipo Agente:TipoAgente"
    },
    "LEMBRETE": {
        "sheet_dados": "GFN003 - Garantia Financeira po",
        "sheet_contatos": "Planilha1",
        "header_row": 30,
        "data_columns": "Agente:Empresa,Garantia Avulsa (R$):Valor"
    },
    "LFRCAP": {
        "sheet_dados": "LFRCAP002 - Liquidação de Reser",
        "sheet_contatos": "Planilha1",
        "header_row": 30,
        "data_columns": "Agente:Empresa,Data do Débito:Data,Valor do Débito (R$):Valor"
    },
    "RCAP": {
        "sheet_dados": "Sheet1",
        "sheet_contatos": "Planilha1",
        "header_row": 4,
        "data_columns": "Agente:Empresa,Data:Data,Valor do Débito (R$):Valor"
    }
}

def load_configs() -> dict:
    if not CONFIG_FILE.exists():
        save_configs(DEFAULT_CONFIGS)
        return DEFAULT_CONFIGS.copy()
    
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            loaded_configs = json.load(f)
            for key, value in DEFAULT_CONFIGS.items():
                if key not in loaded_configs:
                    loaded_configs[key] = value
            return loaded_configs
    except (json.JSONDecodeError, IOError):
        return DEFAULT_CONFIGS.copy()

def save_configs(configs: dict):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(configs, f, indent=4, ensure_ascii=False)