import json
from pathlib import Path
from typing import Dict, Any
from src.config.config import DEFAULT_CONFIGS, PATH_CONFIGS, CONFIG_FILE

def get_user_paths() -> Dict[str, str]:
    """
    Gera os caminhos padrão para o sistema.
    
    Returns:
        Dicionário com os caminhos configurados
    """
    import os
    current_user = os.getenv('USERNAME', 'malik.mourad')
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
    
    meses_map = {
        'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'ABRIL': '04',
        'MAIO': '05', 'JUNHO': '06', 'JULHO': '07', 'AGOSTO': '08',
        'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11', 'DEZEMBRO': '12'
    }
    
    mes_num = meses_map.get(mes.upper(), '01')
    ano_mes = f"{ano}{mes_num}"
    mes_abrev = mes.lower()[:3]
    ano_2dig = ano[-2:]
    
    user_paths = get_user_paths()
    
    path_template = DEFAULT_CONFIGS[report_type].get("path_template", {})
    
    paths = {}
    for key, template in path_template.items():
        paths[key] = template.format(
            sharepoint_root=user_paths["raiz_sharepoint"],
            ano=ano,
            ano_mes=ano_mes,
            mes_abrev=mes_abrev,
            ano_2dig=ano_2dig
        )
    
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
            
            for key, value in DEFAULT_CONFIGS.items():
                if key not in loaded_configs:
                    loaded_configs[key] = value
                else:
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
    
    try:
        int(config["header_row"])
    except (ValueError, TypeError):
        print(f"header_row deve ser um número em {report_type}")
        return False
    
    return True