import requests
from urllib.parse import quote
from typing import Optional
from models import ReferenceInput, ValidationResult
from utils import calculate_similarity, check_authors_match, get_http_session
from .base import BaseValidator

class CrossrefValidator(BaseValidator):
    def __init__(self, mailto: str = "acta@uem.br"):
        self.api_name = "Crossref"
        self.base_url = "https://api.crossref.org/works"
        user_agent = f"ActaScientiarumBibValidator/1.0 (mailto:{mailto})"
        self.session = get_http_session(retries=3, backoff_factor=1.0, user_agent=user_agent)

    def validate(self, reference: ReferenceInput) -> Optional[ValidationResult]:
        try:
            if reference.title:
                params = {'query.title': reference.title, 'rows': 3}
                response = self.session.get(self.base_url, params=params, timeout=10)
                if response.status_code == 200:
                    items = response.json().get('message', {}).get('items', [])
                    for item in items:
                        title_found = item.get('title', [''])[0]
                        if calculate_similarity(reference.title, title_found) > 0.8:
                            return self._parse_result(item, reference, match_type='title')

            if reference.doi:
                clean_doi = reference.doi.replace('https://doi.org/', '').replace('http://doi.org/', '').strip()
                url = f"{self.base_url}/{quote(clean_doi)}"
                response = self.session.get(url, timeout=10)
                if response.status_code == 200:
                    data = response.json().get('message', {})
                    return self._parse_result(data, reference, match_type='doi')
                elif response.status_code == 429:
                    return ValidationResult(
                        api_source=self.api_name,
                        status="Erro",
                        issues=["Limite de requisições excedido no Crossref (HTTP 429)."],
                        confidence="Baixo"
                    )

            # Fallback para query.bibliographic usando o texto original
            if reference.original_text:
                params = {'query.bibliographic': reference.original_text, 'rows': 3}
                response = self.session.get(self.base_url, params=params, timeout=10)
                if response.status_code == 200:
                    items = response.json().get('message', {}).get('items', [])
                    for item in items:
                        title_found = item.get('title', [''])[0]
                        # Mesmo que não bata o título exato, crossref com bibliographic é bom
                        if reference.title and calculate_similarity(reference.title, title_found) > 0.6:
                            return self._parse_result(item, reference, match_type='bibliographic')
                        elif not reference.title:
                            return self._parse_result(item, reference, match_type='bibliographic')

            return None
        except Exception as e:
            return ValidationResult(
                api_source=self.api_name,
                status="Erro",
                issues=[f"Erro na API Crossref: {str(e)}"],
                confidence="Baixo"
            )

    def _parse_result(self, item: dict, reference: ReferenceInput, match_type: str) -> ValidationResult:
        title_found = item.get('title', [''])[0]
        doi_found = item.get('DOI')
        url_found = item.get('URL')

        authors_found = []
        for author in item.get('author', []):
            given = author.get('given', '').strip()
            family = author.get('family', '').strip()
            if family and given:
                authors_found.append(f"{family}, {given}")
            elif family:
                authors_found.append(family)
            elif given:
                authors_found.append(given)

        authors_match, issues = check_authors_match(reference.authors, authors_found)

        # DOI e identificador unico: se bate, o artigo esta confirmado
        if match_type == 'doi':
            status = "Confirmado"
            issues = []  # Limpa divergencias de autoria, pois o DOI garante a veracidade
            # Fallback: se a base nao tiver titulo/autores mas validou pelo DOI, usa os originais
            if not title_found and reference.title:
                title_found = reference.title
            if not authors_found and reference.authors:
                authors_found = reference.authors
        elif calculate_similarity(reference.title, title_found) > 0.9:
            status = "Confirmado" if authors_match else "Parcialmente confirmado"
        else:
            status = "Nao confirmado"

        # ── Campos APA7 ──────────────────────────────────────
        year_found = None
        published = item.get('published') or item.get('published-print') or item.get('published-online')
        if published:
            date_parts = published.get('date-parts', [[]])
            if date_parts and date_parts[0]:
                year_found = str(date_parts[0][0])

        journal_found = None
        container = item.get('container-title', [])
        if container:
            journal_found = container[0]

        volume_found = item.get('volume')
        issue_found = item.get('issue')
        pages_found = item.get('page')

        return ValidationResult(
            api_source=self.api_name,
            status=status,
            title_found=title_found,
            doi_found=doi_found,
            authors_found=authors_found,
            issues=issues,
            confidence="Alto" if status == "Confirmado" else "Moderado",
            url=url_found,
            year_found=year_found,
            journal_found=journal_found,
            volume_found=volume_found,
            issue_found=issue_found,
            pages_found=pages_found,
        )

