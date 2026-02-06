from __future__ import annotations
import re
from pathlib import Path
from typing import Iterable, Tuple
import bleach

TAGS_PERMITIDAS = list(bleach.sanitizer.ALLOWED_TAGS) + [
    "p",
    "br",
    "hr",
    "strong",
    "em",
    "ul",
    "ol",
    "li",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
]

ATRIBUTOS_PERMITIDOS = {
    "a": ["href", "title", "target", "rel"],
    #"*": ["style"],
}

def sanitizar_html(html: str) -> str:
    if not isinstance(html, str):
        return ""
    return bleach.clean(html, tags=TAGS_PERMITIDAS, attributes=ATRIBUTOS_PERMITIDOS, strip=True)

def sanitizar_assunto(assunto: str) -> str:
    if not isinstance(assunto, str):
        return ""
    return re.sub(r"[\r\n]", " ", assunto).strip()

PADRAO_EMAIL = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")

def validar_email(endereco: str) -> bool:
    if not isinstance(endereco, str):
        return False
    if "\n" in endereco or "\r" in endereco:
        return False
    return PADRAO_EMAIL.match(endereco.strip()) is not None

def validar_lista_emails(valor: str) -> Tuple[str, list[str]]:
    avisos: list[str] = []
    if not valor:
        return "", avisos
    partes = [p.strip() for p in valor.split(";") if p.strip()]
    validos = []
    for p in partes:
        if validar_email(p):
            validos.append(p)
        else:
            avisos.append(f"E-mail inválido/descartado: {p}")
    return "; ".join(validos), avisos

def caminho_eh_seguro(diretorios_base: Iterable[str], caminho_alvo: Path) -> bool:
    try:
        alvo_resolvido = caminho_alvo.resolve(strict=False)
        for base in diretorios_base:
            if alvo_resolvido.is_relative_to(Path(base).resolve(strict=False)):
                return True
    except Exception:
        return False
    return False

def dentro_limite_tamanho(caminho: Path, max_mb: int = 20) -> bool:
    try:
        return caminho.stat().st_size <= max_mb * 1024 * 1024
    except Exception:
        return False