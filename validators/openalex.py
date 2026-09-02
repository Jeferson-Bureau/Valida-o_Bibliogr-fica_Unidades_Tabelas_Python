import requests
from urllib.parse import quote
from typing import Optional
from models import ReferenceInput, ValidationResult
from utils import calculate_similarity, check_authors_match, get_http_session
from .base import BaseValidator

class OpenAlexValidator(BaseValidator):
    def __init__(self, mailto: str = "acta@uem.br"):
        self.api_name = "OpenAlex"
        self.base_url = "https://api.openalex.org/works"
        user_agent = f"ActaScientiarumBibValidator/1.0 (mailto:{mailto})"
        self.session = get_http_session(retries=3, backoff_factor=1.0, user_agent=user_agent)

    def validate(self, reference: ReferenceInput) -> Optional[ValidationResult]:
        try:
            if reference.title:
                params = {'filter': f'title.search:{reference.title}', 'per_page': 3}
                response = self.session.get(self.base_url, params=params, timeout=10)
                if response.status_code == 200:
                    results = response.json().get('results', [])
                    if results:
                        for item in results:
                            if calculate_similarity(reference.title, item.get('title', '')) > 0.8:
                                return self._parse_result(item, reference, match_type='title')

                # Fallback: busca geral no OpenAlex por título
                params_search = {'search': reference.title, 'per_page': 3}
                response_search = self.session.get(self.base_url, params=params_search, timeout=10)
                if response_search.status_code == 200:
                    results = response_search.json().get('results', [])
                    for item in results:
                        if calculate_similarity(reference.title, item.get('title', '')) > 0.75:
                            return self._parse_result(item, reference, match_type='title')

            if reference.doi:
                doi_clean = reference.doi.replace('https://doi.org/', '').replace('http://doi.org/', '').strip()
                url = f"{self.base_url}/https://doi.org/{quote(doi_clean, safe='')}"
                response = self.session.get(url, timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    return self._parse_result(data, reference, match_type='doi')
                elif response.status_code == 429:
                    return ValidationResult(
                        api_source=self.api_name,
                        status="Erro",
                        issues=["Limite de requisições excedido no OpenAlex (HTTP 429)."],
                        confidence="Baixo"
                    )

            return None
        except Exception as e:
            return ValidationResult(
                api_source=self.api_name,
                status="Erro",
                issues=[f"Erro na API OpenAlex: {str(e)}"],
                confidence="Baixo"
            )

    def _parse_result(self, item: dict, reference: ReferenceInput, match_type: str = 'title') -> ValidationResult:
        title_found = item.get('title')
        doi_found = item.get('doi')
        if doi_found:
            doi_found = doi_found.replace('https://doi.org/', '')

        authors_found = []
        for authorship in item.get('authorships', []):
            author = authorship.get('author', {})
            display_name = author.get('display_name', '').strip()
            if not display_name:
                continue
            # OpenAlex retorna "Nome Sobrenome" -> converte para "Sobrenome, Nome"
            parts = display_name.split()
            if len(parts) >= 2:
                family = parts[-1]
                given = " ".join(parts[:-1])
                authors_found.append(f"{family}, {given}")
            else:
                authors_found.append(display_name)

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
        elif reference.title and title_found and calculate_similarity(reference.title, title_found) > 0.9:
            status = "Confirmado" if authors_match else "Parcialmente confirmado"
        else:
            status = "Parcialmente confirmado"

        # ── Campos APA7 ──────────────────────────────────────
        year_found = str(item.get('publication_year')) if item.get('publication_year') else None

        journal_found = None
        primary_location = item.get('primary_location') or {}
        source = primary_location.get('source') or {}
        journal_found = source.get('display_name')

        biblio = item.get('biblio') or {}
        volume_found = biblio.get('volume')
        issue_found = biblio.get('issue')
        first_page = biblio.get('first_page')
        last_page = biblio.get('last_page')
        pages_found = None
        if first_page and last_page:
            pages_found = f"{first_page}-{last_page}"
        elif first_page:
            pages_found = first_page

        return ValidationResult(
            api_source=self.api_name,
            status=status,
            title_found=title_found,
            doi_found=doi_found,
            authors_found=authors_found,
            issues=issues,
            confidence="Alto" if status == "Confirmado" else "Moderado",
            url=item.get('id'),
            year_found=year_found,
            journal_found=journal_found,
            volume_found=volume_found,
            issue_found=issue_found,
            pages_found=pages_found,
        )

