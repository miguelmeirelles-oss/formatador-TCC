"""Utilitários de normalização de texto reutilizados pelos outros módulos."""
from __future__ import annotations

import re
import unicodedata


def sem_acento(s: str) -> str:
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def normalizar(s: str) -> str:
    """Maiúsculas, sem acento, espaços colapsados -- para comparação de chaves."""
    s = sem_acento(s).upper()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def contar_palavras(texto: str) -> int:
    palavras = re.findall(r"[A-Za-zÀ-ÿ0-9]+(?:[-'][A-Za-zÀ-ÿ0-9]+)*", texto)
    return len(palavras)


# Ano plausível de publicação (1500-2099), com sufixo de letra opcional
# (2020a, 2020b, para desambiguar obras do mesmo autor no mesmo ano),
# isolado por limites de palavra -- evita capturar 4 dígitos no meio de
# números maiores (ex.: "NBR 14720/2002" não deve casar "1472"), códigos de
# norma, ISBN etc.
RE_ANO = re.compile(r"\b(?:1[5-9]\d{2}|20\d{2})[a-z]?\b")
