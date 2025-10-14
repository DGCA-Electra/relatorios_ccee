import pandas as pd
from pathlib import Path
from typing import Dict, Any, Optional
import json
import logging

TEMPLATES_JSON_PATH = "src/config/email_templates.json"

class ReportProcessingError(Exception):
    pass

def load_excel_data(excel_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    """Carrega dados de uma planilha Excel."""
    if not Path(excel_path).exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {excel_path}")
    if header_row == -1:
        return pd.read_excel(Path(excel_path), sheet_name=sheet_name, header=None)
    return pd.read_excel(Path(excel_path), sheet_name=sheet_name, header=header_row)

def find_attachment(pdf_dir: str, filename: str) -> Optional[Path]:
    """Procura por um arquivo PDF no diretório especificado."""
    attachment_path = Path(pdf_dir) / filename
    if attachment_path.exists():
        return attachment_path
    logging.warning(f"Anexo não encontrado no caminho principal: {attachment_path}")
    return None

def load_email_templates() -> Dict[str, Any]:
    try:
        with open(TEMPLATES_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao carregar {TEMPLATES_JSON_PATH}: {e}")

def save_email_templates(data: Dict[str, Any]) -> None:
    """Salva os templates de e-mail no arquivo JSON."""
    try:
        with open(TEMPLATES_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise ReportProcessingError(f"Falha ao salvar {TEMPLATES_JSON_PATH}: {e}")
