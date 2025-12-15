from pathlib import Path
from datetime import datetime

CONFIG_FILE = Path('src/configuracoes/config_relatorios.json')

ANALISTAS = [
    'Artur Bello Rodrigues', 'Camila Padovan Baptista', 'Cassiana Unruh',
    'Isabela Loredo', 'Tiago Padilha Foletto'
]
MESES = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
         'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
ANOS = [str(y) for y in range(datetime.now().year, datetime.now().year + 5)]

PATH_CONFIGS = {
    "sharepoint_root": "ELECTRA COMERCIALIZADORA DE ENERGIA S.A/GE - ECE/DGCA/DGA/CCEE/Relatórios CCEE",
    "contatos_email": "ELECTRA COMERCIALIZADORA DE ENERGIA S.A/GE - ECE/DGCA/DGC/Macro/Contatos de E-mail para Macros.xlsx",
    "user_base": "C:/Users"
}

DEFAULT_CONFIGS = {
    "GFN001": {
        "planilha_dados": "GFN003 - Garantia Financeira po",
        "planilha_contatos": "Planilha1",
        "linha_cabecalho": 30,
        "colunas_dados": "Agente:Empresa,Garantia Avulsa (R$):Valor",
        "modelo_caminho": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Garantia Financeira/GFN003 - Excel/ELECTRA_ENERGY_GFN003_{mes_abrev}_{ano_2dig}.xlsx",
            "diretorio_pdfs": "{sharepoint_root}/{ano}/{ano_mes}/Garantia Financeira/GFN001"
        }
    },
    "SUM001": {
        "planilha_dados": "LFN004 - Liquidação Financeira ",
        "planilha_contatos": "Planilha1",
        "linha_cabecalho": 31,
        "colunas_dados": "Agente:Empresa,Débito/Crédito:Situacao,Valor a Liquidar (R$):Valor",
        "modelo_caminho": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação Financeira/LFN004/ELECTRA ENERGY LFN004 {mes_abrev}.{ano_2dig}.xlsx",
            "diretorio_pdfs": "{sharepoint_root}/{ano}/{ano_mes}/Sumário/SUM001"
        }
    },
    "LFN001": {
        "planilha_dados": "LFN004 - Liquidação Financeira ",
        "planilha_contatos": "Planilha1",
        "linha_cabecalho": 31,
        "colunas_dados": "Agente:Empresa,Débito/Crédito:Situacao,Valor a Liquidar (R$):ValorLiquidacao,Valor Liquidado (R$):ValorLiquidado,Inadimplência (R$):ValorInadimplencia",
        "modelo_caminho": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação Financeira/LFN004/ELECTRA ENERGY LFN004 {mes_abrev}.{ano_2dig} (pós).xlsx",
            "diretorio_pdfs": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação Financeira/LFN001"
        }
    },
    "LFRES001": {
        "planilha_dados": "LFRES002 - Liquidação de Energi",
        "planilha_contatos": "Planilha1",
        "linha_cabecalho": 42,
        "colunas_dados": "Agente:Empresa,Data do Débito:Data,Valor a Liquidar (R$):Valor,Tipo do Agente:TipoAgente",
        "modelo_caminho": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação da Energia de Reserva/LFRES002/ELECTRA_ENERGY_LFRES002_{mes_abrev}_{ano_2dig}.xlsx",
            "diretorio_pdfs": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação da Energia de Reserva/LFRES001"
        }
    },
    "GFN - LEMBRETE": {
        "planilha_dados": "GFN003 - Garantia Financeira po",
        "planilha_contatos": "Planilha1",
        "linha_cabecalho": 30,
        "colunas_dados": "Agente:Empresa,Garantia Avulsa (R$):Valor",
        "modelo_caminho": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Garantia Financeira/GFN003 - Excel/ELECTRA_ENERGY_GFN003_{mes_abrev}_{ano_2dig}.xlsx",
            "diretorio_pdfs": "{sharepoint_root}/{ano}/{ano_mes}/Garantia Financeira/GFN001"
        }
    },
    "LFRCAP001": {
        "planilha_dados": "LFRCAP002 - Liquidação de Reser",
        "planilha_contatos": "Planilha1",
        "linha_cabecalho": 30,
        "colunas_dados": "Agente:Empresa,Data do Débito:Data,Valor do Débito (R$):Valor",
        "modelo_caminho": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação de Reserva de Capacidade/LFRCAP002/ELECTRA_ENERGY_LFRCAP002_{mes_abrev}_{ano_2dig}.xlsx",
            "diretorio_pdfs": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação de Reserva de Capacidade/LFRCAP001"
        }
    },
    "RCAP002": {
        "planilha_dados": "Sheet1",
        "planilha_contatos": "Planilha1",
        "linha_cabecalho": 4,
        "colunas_dados": "Sigla do Agente:Empresa,\"(S) ERCAP_C am\":Valor,Data:Data",
        "modelo_caminho": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Reserva de Capacidade/RCAP002 - Consulta Dinamica/RCAP002 {mes_abrev}.{ano_2dig}.xlsx",
            "diretorio_pdfs": "{sharepoint_root}/{ano}/{ano_mes}/Reserva de Capacidade/RCAP002"
        }
    }
}


