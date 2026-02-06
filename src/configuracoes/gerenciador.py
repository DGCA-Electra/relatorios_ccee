import json
import os
from pathlib import Path
from typing import Dict, Any
from src.configuracoes.constantes import DEFAULT_CONFIGS, PATH_CONFIGS, CONFIG_FILE
def obter_caminhos_brutos_usuario(username: str) -> Dict[str, str]:
    """Gera os caminhos base para um usuário específico (sem validar existência)."""
    user_base = f"{PATH_CONFIGS['user_base']}/{username}"
    return {
        "raiz_sharepoint": f"{user_base}/{PATH_CONFIGS['sharepoint_root']}",
        "contratos_email_path": f"{user_base}/{PATH_CONFIGS['contatos_email']}"
    }
def resolver_melhores_caminhos(preferred_username: str = None) -> Dict[str, str]:
    """
    Tenta encontrar os caminhos válidos.
    1. Tenta o usuário preferencial (da rede/login).
    2. Se o caminho não existir, faz fallback para o usuário local (os.environ).
    """
    if preferred_username:
        paths = obter_caminhos_brutos_usuario(preferred_username)
        if os.path.exists(paths["raiz_sharepoint"]):
            return paths
        print(f"Aviso: Caminho de rede para '{preferred_username}' não encontrado. Tentando local...")
    local_user = os.getenv('USERNAME', 'malik.mourad')
    paths = obter_caminhos_brutos_usuario(local_user)
    return paths
def construir_caminhos_relatorio(report_type: str, ano: str, mes: str, username: str = None) -> Dict[str, str]:
    """
    Constrói os caminhos específicos para um relatório, resolvendo automaticamente
    se deve usar caminho de rede ou local.
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
    user_paths = resolver_melhores_caminhos(username)
    modelo_caminho = DEFAULT_CONFIGS[report_type].get("modelo_caminho", {})
    caminhos = {}
    for chave, template in modelo_caminho.items():
        caminhos[chave] = template.format(
            sharepoint_root=user_paths["raiz_sharepoint"],
            ano=ano,
            ano_mes=ano_mes,
            mes_abrev=mes_abrev,
            ano_2dig=ano_2dig
        )
    caminhos["excel_contatos"] = user_paths["contratos_email_path"]
    return caminhos
def carregar_configuracoes() -> Dict[str, Any]:
    """
    Carrega as configurações do arquivo JSON ou cria com valores padrão.
    Returns:
        Dicionário com todas as configurações
    """
    if not CONFIG_FILE.exists():
        salvar_configuracoes(DEFAULT_CONFIGS)
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
def salvar_configuracoes(configs: Dict[str, Any]) -> None:
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
def validar_configuracao(config: Dict[str, Any], report_type: str) -> bool:
    """
    Valida se uma configuração está completa e correta.
    Args:
        config: Configuração a ser validada
        report_type: Tipo do relatório
    Returns:
        True se a configuração é válida, False caso contrário
    """
    campos_obrigatorios = ["planilha_dados", "planilha_contatos", "linha_cabecalho", "colunas_dados"]
    for campo in campos_obrigatorios:
        if campo not in config:
            print(f"Campo obrigatório '{campo}' não encontrado em {report_type}")
            return False
    try:
        int(config["linha_cabecalho"])
    except (ValueError, TypeError):
        print(f"linha_cabecalho deve ser um número em {report_type}")
        return False
    return True