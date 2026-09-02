from typing import Optional
import re
from models import ValidationResult


def _format_author_apa7(name: str) -> str:
    """
    Converte um nome para o formato APA7: Sobrenome, I.I.

    Formatos aceitos:
      - 'Britnell, S.J.'  -> 'Britnell, S.J.'  (Crossref - ja correto)
      - 'Britnell, SJ'    -> 'Britnell, S.J.'  (PubMed legado)
      - 'Joao Silva'      -> 'Silva, J.'
      - 'Britnell SJ'     -> 'Britnell, S.J.'  (PubMed sem virgula)
    """
    name = name.strip()
    if not name:
        return ""

    if "," in name:
        parts = name.split(",", 1)
        family = parts[0].strip()
        given_raw = parts[1].strip() if len(parts) > 1 else ""
    else:
        tokens = name.split()
        if len(tokens) == 1:
            return tokens[0]
        last = tokens[-1]
        # Iniciais coladas sem virgula: "Britnell SJ"
        if last.isupper() and len(last) <= 4 and last.isalpha():
            family = tokens[0]
            given_raw = last
        else:
            # "Nome Sobrenome" -> inverte
            family = tokens[-1]
            given_raw = " ".join(tokens[:-1])

    # Se o given_raw ja esta no formato de iniciais pontuadas (S.J., M.M., L.)
    # normaliza para APA7 com espaco apos cada ponto: S.J. -> S. J.
    compact = given_raw.replace(" ", "")
    if re.match(r"^([A-Z]\.)+$", compact):
        spaced = " ".join(ch + "." for ch in compact.replace(".", "") if ch)
        return f"{family}, {spaced}" if spaced else family

    # Caso contrario, abrevia cada parte
    initials = ""
    for part in given_raw.split():
        part = part.rstrip(".")
        if not part:
            continue
        # Iniciais coladas sem ponto: "SJ" -> "S. J."
        if part.isupper() and len(part) > 1 and part.isalpha():
            for ch in part:
                initials += ch.upper() + ". "
        else:
            initials += part[0].upper() + ". "

    if initials:
        return f"{family}, {initials.strip()}"
    return family


def _format_authors_apa7(authors: list) -> str:
    """
    Formata uma lista de autores no estilo APA7.
    - 1 autor:  Sobrenome, I.
    - 2-20:     Sobrenome, I., ..., & SobrenomeN, IN.
    - 21+:      primeiros 19, ..., & ultimo
    """
    if not authors:
        return ""

    formatted = [_format_author_apa7(a) for a in authors if a.strip()]

    if len(formatted) == 1:
        return formatted[0]

    if len(formatted) <= 20:
        return ", ".join(formatted[:-1]) + ", & " + formatted[-1]

    # 21+ autores: 19 primeiros + ... + ultimo
    return ", ".join(formatted[:19]) + ", ... & " + formatted[-1]


def format_apa7(result: ValidationResult) -> Optional[str]:
    """
    Monta uma referencia no formato APA 7a edicao a partir de um ValidationResult.
    Retorna None se nao houver dados suficientes (sem titulo encontrado).
    """
    if not result.title_found:
        return None

    parts = []

    # Autores
    authors_str = _format_authors_apa7(result.authors_found)
    if authors_str:
        parts.append(authors_str)

    # Ano
    year = result.year_found
    if year:
        parts.append(f"({year})")
    else:
        parts.append("(s.d.)")

    # Titulo
    title = result.title_found.strip().rstrip(".")
    parts.append(f"{title}.")

    # Periodico, Volume(Numero), Paginas
    journal_part = ""
    if result.journal_found:
        journal_part = result.journal_found.strip()

    vol_issue = ""
    if result.volume_found:
        vol_issue = result.volume_found
        if result.issue_found:
            vol_issue += f"({result.issue_found})"

    if journal_part and vol_issue:
        journal_part = f"{journal_part}, {vol_issue}"
    elif vol_issue:
        journal_part = vol_issue

    if result.pages_found:
        if journal_part:
            journal_part += f", {result.pages_found}"
        else:
            journal_part = result.pages_found

    if journal_part:
        parts.append(f"{journal_part}.")

    # DOI / URL
    if result.doi_found:
        doi = result.doi_found.strip()
        if not doi.startswith("http"):
            doi = f"https://doi.org/{doi}"
        parts.append(doi)
    elif result.url:
        parts.append(result.url)

    return " ".join(parts)
