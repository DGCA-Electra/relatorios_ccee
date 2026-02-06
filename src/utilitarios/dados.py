import pandas as pd
from typing import Dict, Any, Optional
import re
def converter_numero_br(val: Any) -> float:
    """Converte 'R$ 1.234,56' ou '(1.234,56)' ou 1234.56 para float. Retorna 0.0 em erro."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s == "":
        return 0.0
    is_neg = False
    if s.startswith("(") and s.endswith(")"):
        is_neg = True
        s = s[1:-1]
    s = s.replace("R$", "").replace("r$", "").replace("\xa0", "").replace(" ", "")
    s = s.replace(".", "")
    s = s.replace(",", ".")
    s = re.sub(r"[^0-9\.-]", "", s)
    try:
        n = float(s) if s not in ("", "-", ".") else 0.0
        return -n if is_neg else n
    except Exception:
        return 0.0
def formatar_moeda(value: Any) -> str:
    try:
        val = float(value)
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "R$ 0,00"
def formatar_data(date_value: Any) -> str:
    """Formata uma data para o formato brasileiro (dd/mm/aaaa)."""
    try:
        if date_value is None or pd.isna(date_value):
            return "Data não informada"
        return pd.to_datetime(date_value).strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return "Data Inválida"
