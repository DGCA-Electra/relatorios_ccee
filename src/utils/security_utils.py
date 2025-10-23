from __future__ import annotations
import re
from pathlib import Path
from typing import Iterable, Tuple
import bleach


ALLOWED_TAGS = list(bleach.sanitizer.ALLOWED_TAGS) + [
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
ALLOWED_ATTRS = {
    "a": ["href", "title", "target", "rel"],
    #"*": ["style"],
}


def sanitize_html(html: str) -> str:
    if not isinstance(html, str):
        return ""
    return bleach.clean(html, tags=ALLOWED_TAGS, attributes=ALLOWED_ATTRS, strip=True)


def sanitize_subject(subject: str) -> str:
    if not isinstance(subject, str):
        return ""
    return re.sub(r"[\r\n]", " ", subject).strip()


EMAIL_PATTERN = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


def validate_email(addr: str) -> bool:
    if not isinstance(addr, str):
        return False
    if "\n" in addr or "\r" in addr:
        return False
    return EMAIL_PATTERN.match(addr.strip()) is not None


def validate_email_list(value: str) -> Tuple[str, list[str]]:
    warnings: list[str] = []
    if not value:
        return "", warnings
    parts = [p.strip() for p in value.split(";") if p.strip()]
    valid = []
    for p in parts:
        if validate_email(p):
            valid.append(p)
        else:
            warnings.append(f"E-mail inválido/descartado: {p}")
    return "; ".join(valid), warnings


def is_safe_path(base_dirs: Iterable[str], target_path: Path) -> bool:
    try:
        resolved_target = target_path.resolve(strict=False)
        for base in base_dirs:
            if resolved_target.is_relative_to(Path(base).resolve(strict=False)):
                return True
    except Exception:
        return False
    return False


def within_size_limit(path: Path, max_mb: int = 20) -> bool:
    try:
        return path.stat().st_size <= max_mb * 1024 * 1024
    except Exception:
        return False