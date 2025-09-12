import json
import os
from pathlib import Path
from typing import Dict, Any, Optional

CONFIG_FILE = Path('config_relatorios.json')
ANALISTAS = [
    'Artur Bello Rodrigues', 'Camila Padovan Baptista', 'Cassiana Unruh',
    'Isabela Loredo', 'Jorge Ferreira', 'Tiago Padilha Foletto'
]
MESES = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
         'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
ANOS = [str(y) for y in range(2025, 2030)]

# Configurações de caminho centralizadas
PATH_CONFIGS = {
    "sharepoint_root": "ELECTRA COMERCIALIZADORA DE ENERGIA S.A/GE - ECE/DGCA/DGA/CCEE/Relatórios CCEE",
    "contatos_email": "ELECTRA COMERCIALIZADORA DE ENERGIA S.A/GE - ECE/DGCA/DGC/Macro/Contatos de E-mail para Macros.xlsx",
    "user_base": "C:/Users"
}

# Mapeamento das colunas relevantes por tipo de relatório
REPORT_DISPLAY_COLUMNS = {
    "SUM001": ["Empresa", "Email", "Valor", "Data_Debito_Credito"],
    "LFN001": ["Empresa", "Email", "ValorLiquidacao", "ValorLiquidado", "ValorInadimplencia"],
    "GFN001": ["Empresa", "Email", "Valor"],
    "LFRES001": ["Empresa", "Email", "TipoAgente", "Valor", "Data"],
    "LFRCAP001": ["Empresa", "Email", "Valor", "Data"],
    "RCAP002": ["Empresa", "Email", "Valor", "Data"],
    "GFN - LEMBRETE": ["Empresa", "Email", "Valor"]
}

# Configurações padrão dos relatórios
DEFAULT_CONFIGS = {
    "GFN001": {
        "sheet_dados": "GFN003 - Garantia Financeira po",
        "sheet_contatos": "Planilha1",
        "header_row": 30,
        "data_columns": "Agente:Empresa,Garantia Avulsa (R$):Valor",
        "path_template": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Garantia Financeira/GFN003 - Excel/ELECTRA_ENERGY_GFN003_{mes_abrev}_{ano_2dig}.xlsx",
            "pdfs_dir": "{sharepoint_root}/{ano}/{ano_mes}/Garantia Financeira/GFN001"
        }
    },
    "SUM001": {
        "sheet_dados": "LFN004 - Liquidação Financeira ",
        "sheet_contatos": "Planilha1",
        "header_row": 31,
        "data_columns": "Agente:Empresa,Débito/Crédito:Situacao,Valor a Liquidar (R$):Valor",
        "path_template": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação Financeira/LFN004/ELECTRA ENERGY LFN004 {mes_abrev}.{ano_2dig}.xlsx",
            "pdfs_dir": "{sharepoint_root}/{ano}/{ano_mes}/Sumário/SUM001"
        }
    },
    "LFN001": {
        "sheet_dados": "LFN004 - Liquidação Financeira ",
        "sheet_contatos": "Planilha1",
        "header_row": 31,
        "data_columns": "Agente:Empresa,Débito/Crédito:Situacao,Valor a Liquidar (R$):ValorLiquidacao,Valor Liquidado (R$):ValorLiquidado,Inadimplência (R$):ValorInadimplencia",
        "path_template": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação Financeira/LFN004/ELECTRA ENERGY LFN004 {mes_abrev}.{ano_2dig} (pós).xlsx",
            "pdfs_dir": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação Financeira/LFN001"
        }
    },
    "LFRES001": {
        "sheet_dados": "LFRES002 - Liquidação de Energi",
        "sheet_contatos": "Planilha1",
        "header_row": 42,
        "data_columns": "Agente:Empresa,Data do Débito:Data,Valor do Débito (R$):Valor,Tipo do Agente:TipoAgente",
        "path_template": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação da Energia de Reserva/LFRES002/ELECTRA_ENERGY_LFRES002_{mes_abrev}_{ano_2dig}.xlsx",
            "pdfs_dir": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação da Energia de Reserva/LFRES001"
        }
    },
    "GFN - LEMBRETE": {
        "sheet_dados": "GFN003 - Garantia Financeira po",
        "sheet_contatos": "Planilha1",
        "header_row": 30,
        "data_columns": "Agente:Empresa,Garantia Avulsa (R$):Valor",
        "path_template": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Garantia Financeira/GFN003 - Excel/ELECTRA_ENERGY_GFN003_{mes_abrev}_{ano_2dig}.xlsx",
            "pdfs_dir": "{sharepoint_root}/{ano}/{ano_mes}/Garantia Financeira/GFN001"
        }
    },
    "LFRCAP001": {
        "sheet_dados": "LFRCAP002 - Liquidação de Reser",
        "sheet_contatos": "Planilha1",
        "header_row": 30,
        "data_columns": "Agente:Empresa,Data do Débito:Data,Valor do Débito (R$):Valor",
        "path_template": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação de Reserva de Capacidade/LFRCAP002/ELECTRA_ENERGY_LFRCAP002_{mes_abrev}_{ano_2dig}.xlsx",
            "pdfs_dir": "{sharepoint_root}/{ano}/{ano_mes}/Liquidação de Reserva de Capacidade/LFRCAP001"
        }
    },
    "RCAP002": {
        "sheet_dados": "Sheet1",
        "sheet_contatos": "Planilha1",
        "header_row": 4,
        "data_columns": "Sigla do Agente:Empresa,\"(S) ERCAP_C am\":Valor,Data:Data",
        "path_template": {
            "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Reserva de Capacidade/RCAP002 - Consulta Dinamica/RCAP002 {mes_abrev}.{ano_2dig}.xlsx",
            "pdfs_dir": "{sharepoint_root}/{ano}/{ano_mes}/Reserva de Capacidade/RCAP002"
        }
    }
}

def get_user_paths() -> Dict[str, str]:
    """
    Gera os caminhos padrão para o sistema.
    
    Returns:
        Dicionário com os caminhos configurados
    """
    # Usar caminho padrão fixo baseado no usuário atual do sistema
    import os
    current_user = os.getenv('USERNAME', 'malik.mourad')  # Fallback para malik.mourad
    user_base = f"{PATH_CONFIGS['user_base']}/{current_user}"
    
    return {
        "raiz_sharepoint": f"{user_base}/{PATH_CONFIGS['sharepoint_root']}",
        "contratos_email_path": f"{user_base}/{PATH_CONFIGS['contatos_email']}"
    }

def build_report_paths(report_type: str, ano: str, mes: str) -> Dict[str, str]:
    """
    Constrói os caminhos específicos para um relatório.
    
    Args:
        report_type: Tipo do relatório (ex: GFN001)
        ano: Ano do relatório (ex: 2025)
        mes: Mês do relatório (ex: JUNHO)
        
    Returns:
        Dicionário com os caminhos do relatório
    """
    if report_type not in DEFAULT_CONFIGS:
        raise ValueError(f"Tipo de relatório '{report_type}' não reconhecido")
    
    # Mapeamento de meses para formato numérico
    meses_map = {
        'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'ABRIL': '04',
        'MAIO': '05', 'JUNHO': '06', 'JULHO': '07', 'AGOSTO': '08',
        'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11', 'DEZEMBRO': '12'
    }
    
    mes_num = meses_map.get(mes.upper(), '01')
    ano_mes = f"{ano}{mes_num}"
    mes_abrev = mes.lower()[:3]  # jun, mai, abr, etc.
    ano_2dig = ano[-2:]  # 25, 26, etc.
    
    # Obter caminhos do usuário
    user_paths = get_user_paths()
    
    # Obter template de caminhos do relatório
    path_template = DEFAULT_CONFIGS[report_type].get("path_template", {})
    
    # Construir caminhos substituindo as variáveis
    paths = {}
    for key, template in path_template.items():
        paths[key] = template.format(
            sharepoint_root=user_paths["raiz_sharepoint"],
            ano=ano,
            ano_mes=ano_mes,
            mes_abrev=mes_abrev,
            ano_2dig=ano_2dig
        )
    
    # Adicionar caminho dos contatos
    paths["excel_contatos"] = user_paths["contratos_email_path"]
    
    return paths

def load_configs() -> Dict[str, Any]:
    """
    Carrega as configurações do arquivo JSON ou cria com valores padrão.
    
    Returns:
        Dicionário com todas as configurações
    """
    if not CONFIG_FILE.exists():
        save_configs(DEFAULT_CONFIGS)
        return DEFAULT_CONFIGS.copy()
    
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            loaded_configs = json.load(f)
            
            # Garantir que todas as configurações padrão estejam presentes
            for key, value in DEFAULT_CONFIGS.items():
                if key not in loaded_configs:
                    loaded_configs[key] = value
                else:
                    # Mesclar configurações existentes com padrões
                    for default_key, default_value in value.items():
                        if default_key not in loaded_configs[key]:
                            loaded_configs[key][default_key] = default_value
            
            return loaded_configs
    except (json.JSONDecodeError, IOError) as e:
        print(f"Erro ao carregar configurações: {e}. Usando configurações padrão.")
        return DEFAULT_CONFIGS.copy()

def save_configs(configs: Dict[str, Any]) -> None:
    """
    Salva as configurações no arquivo JSON.
    
    Args:
        configs: Dicionário com as configurações a serem salvas
    """
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(configs, f, indent=4, ensure_ascii=False)
    except IOError as e:
        print(f"Erro ao salvar configurações: {e}")

def validate_config(config: Dict[str, Any], report_type: str) -> bool:
    """
    Valida se uma configuração está completa e correta.
    
    Args:
        config: Configuração a ser validada
        report_type: Tipo do relatório
        
    Returns:
        True se a configuração é válida, False caso contrário
    """
    required_fields = ["sheet_dados", "sheet_contatos", "header_row", "data_columns"]
    
    for field in required_fields:
        if field not in config:
            print(f"Campo obrigatório '{field}' não encontrado em {report_type}")
            return False
    
    # Validar se header_row é um número
    try:
        int(config["header_row"])
    except (ValueError, TypeError):
        print(f"header_row deve ser um número em {report_type}")
        return False
    
    return True