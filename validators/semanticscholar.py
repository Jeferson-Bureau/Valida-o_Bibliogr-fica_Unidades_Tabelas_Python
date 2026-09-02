import requests
from typing import Optional
from models import ReferenceInput, ValidationResult
from urllib.parse import quote
from utils import calculate_similarity, check_authors_match, get_http_session
from .base import BaseValidator

class SemanticScholarValidator(BaseValidator):
    def __init__(self):
        self.api_name = "Semantic Scholar"
        self.paper_url = "https://api.semanticscholar.org/graph/v1/paper"
        self.search_url = "https://api.semanticscholar.org/graph/v1/paper/search"
        self.session = get_http_session(retries=3, backoff_factor=1.5)

    def validate(self, reference: ReferenceInput) -> Optional[ValidationResult]:
        try:
            fields = "title,authors,url,externalIds,year,venue,journal"
            if reference.title:
                params = {
                    'query': reference.title,
                    'limit': 3,
                    'fields': fields
                }
                response = self.session.get(self.search_url, params=params, timeout=10)
                if response.status_code == 200:
                    data = response.json().get('data', [])
                    if data:
                        for item in data:
                            if calculate_similarity(reference.title, item.get('title', '')) > 0.8:
                                return self._parse_result(item, reference, match_type='title')
                elif response.status_code == 429:
                    return ValidationResult(
                        api_source=self.api_name,
                        status="Erro",
                        issues=["Limite de requisições excedido no Semantic Scholar (HTTP 429)."],
                        confidence="Baixo"
                    )

            if reference.doi:
                clean_doi = reference.doi.replace('https://doi.org/', '').replace('http://doi.org/', '').strip()
                url = f"{self.paper_url}/DOI:{quote(clean_doi)}?fields={fields}"
                response = self.session.get(url, timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    return self._parse_result(data, reference, match_type='doi')
                elif response.status_code == 429:
                    return ValidationResult(
                        api_source=self.api_name,
                        status="Erro",
                        issues=["Limite de requisições excedido no Semantic Scholar (HTTP 429)."],
                        confidence="Baixo"
                    )

            return None
        except Exception as e:
            return ValidationResult(
                api_source=self.api_name,
                status="Erro",
                issues=[f"Erro na API Semantic Scholar: {str(e)}"],
                confidence="Baixo"
            )

    def _parse_result(self, item: dict, reference: ReferenceInput, match_type: str = 'title') -> ValidationResult:
        title_found = item.get('title')
        external_ids = item.get('externalIds', {}) or {}
        doi_found = external_ids.get('DOI')

        authors_found = []
        for author in item.get('authors', []):
            name = author.get('name', '').strip()
            if not name:
                continue
            # Semantic Scholar retorna "Nome Sobrenome" -> converte para "Sobrenome, Nome"
            parts = name.split()
            if len(parts) >= 2:
                family = parts[-1]
                given = " ".join(parts[:-1])
                authors_found.append(f"{family}, {given}")
            else:
                authors_found.append(name)

        authors_match, issues = check_authors_match(reference.authors, authors_found)

        if match_type == 'doi':
            status = "Confirmado"
            issues = []
            if not title_found and reference.title:
                title_found = reference.title
            if not authors_found and reference.authors:
                authors_found = reference.authors
        else:
            status = "Confirmado" if authors_match else "Parcialmente confirmado"

        # ── Campos APA7 ──────────────────────────────────────
        year_found = str(item['year']) if item.get('year') else None

        # 'venue' é string simples; 'journal' é dict com name, volume, pages
        journal_info = item.get('journal') or {}
        journal_found = journal_info.get('name') or item.get('venue') or None
        volume_found = str(journal_info['volume']) if journal_info.get('volume') else None
        pages_found = journal_info.get('pages') or None

        return ValidationResult(
            api_source=self.api_name,
            status=status,
            title_found=title_found,
            doi_found=doi_found,
            authors_found=authors_found,
            issues=issues,
            confidence="Alto" if status == "Confirmado" else "Moderado",
            url=item.get('url'),
            year_found=year_found,
            journal_found=journal_found,
            volume_found=volume_found,
            issue_found=None,   # Semantic Scholar não retorna issue
            pages_found=pages_found,
        )
