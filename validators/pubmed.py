import requests
from typing import Optional
from models import ReferenceInput, ValidationResult
import time
from utils import calculate_similarity, check_authors_match, get_http_session
from .base import BaseValidator

class PubMedValidator(BaseValidator):
    def __init__(self):
        self.api_name = "PubMed"
        self.search_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
        self.summary_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
        self.session = get_http_session(retries=3, backoff_factor=1.0)

    def validate(self, reference: ReferenceInput) -> Optional[ValidationResult]:
        try:
            if reference.title:
                query = f"{reference.title}[Title]"
                match_type = 'title'
            elif reference.doi:
                clean_doi = reference.doi.replace('https://doi.org/', '').replace('http://doi.org/', '').strip()
                query = f"{clean_doi}[Location ID]"
                match_type = 'doi'
            else:
                return None

            search_params = {
                'db': 'pubmed',
                'term': query,
                'retmode': 'json'
            }
            search_response = self.session.get(self.search_url, params=search_params, timeout=10)
            if search_response.status_code == 200:
                id_list = search_response.json().get('esearchresult', {}).get('idlist', [])
                if id_list:
                    pubmed_id = id_list[0]
                    # Cortesia para API da NCBI sem chave (máx 3 req/s)
                    time.sleep(0.35)
                    summary_params = {
                        'db': 'pubmed',
                        'id': pubmed_id,
                        'retmode': 'json'
                    }
                    summary_response = self.session.get(self.summary_url, params=summary_params, timeout=10)
                    if summary_response.status_code == 200:
                        result = summary_response.json().get('result', {})
                        docsum = result.get(pubmed_id, {})
                        return self._parse_result(docsum, reference, pubmed_id, match_type)
                elif match_type == 'doi' and reference.title:
                    # Se busca por DOI falhar no PubMed, tenta fallback por Título
                    time.sleep(0.35)
                    fallback_params = {'db': 'pubmed', 'term': f"{reference.title}[Title]", 'retmode': 'json'}
                    fallback_resp = self.session.get(self.search_url, params=fallback_params, timeout=10)
                    if fallback_resp.status_code == 200:
                        fb_ids = fallback_resp.json().get('esearchresult', {}).get('idlist', [])
                        if fb_ids:
                            pubmed_id = fb_ids[0]
                            time.sleep(0.35)
                            summary_response = self.session.get(self.summary_url, params={'db': 'pubmed', 'id': pubmed_id, 'retmode': 'json'}, timeout=10)
                            if summary_response.status_code == 200:
                                result = summary_response.json().get('result', {})
                                docsum = result.get(pubmed_id, {})
                                return self._parse_result(docsum, reference, pubmed_id, match_type='title')

            return None
        except Exception as e:
            return ValidationResult(
                api_source=self.api_name,
                status="Erro",
                issues=[f"Erro na API PubMed: {str(e)}"],
                confidence="Baixo"
            )

    def _parse_result(self, docsum: dict, reference: ReferenceInput, pubmed_id: str, match_type: str = 'title') -> ValidationResult:
        title_found = docsum.get('title', '')
        doi_found = None
        for art_id in docsum.get('articleids', []):
            if art_id.get('idtype') == 'doi':
                doi_found = art_id.get('value')
                break

        authors_found = []
        for author in docsum.get('authors', []):
            # Filtra apenas autores reais (ignora autores coletivos/institucionais)
            if author.get('authtype', '').lower() not in ('author', ''):
                continue
            raw = author.get('name', '').strip()
            if not raw:
                continue
            # PubMed retorna "Sobrenome SJ" (sem virgula, iniciais coladas)
            # Converte para "Sobrenome, S.J." (sem espaco entre iniciais)
            tokens = raw.split()
            if len(tokens) >= 2:
                surname = tokens[0]
                initials_raw = tokens[-1]  # ex: "SJ", "MM", "JV"
                # Expande iniciais coladas: "SJ" -> "S. J."
                initials_expanded = ". ".join(initials_raw.upper()) + "."
                authors_found.append(f"{surname}, {initials_expanded}")
            else:
                authors_found.append(raw)

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
        else:
            status = "Confirmado" if authors_match else "Parcialmente confirmado"

        # ── Campos APA7 ──────────────────────────────────────
        # pubdate pode ser '2017 Dec' ou '2017'
        pubdate = docsum.get('pubdate', '')
        year_found = pubdate.split()[0] if pubdate else None

        journal_found = docsum.get('source') or None
        volume_found = docsum.get('volume') or None
        issue_found = docsum.get('issue') or None
        pages_found = docsum.get('pages') or None

        return ValidationResult(
            api_source=self.api_name,
            status=status,
            title_found=title_found,
            doi_found=doi_found,
            authors_found=authors_found,
            issues=issues,
            confidence="Alto" if status == "Confirmado" else "Moderado",
            url=f"https://pubmed.ncbi.nlm.nih.gov/{pubmed_id}/",
            year_found=year_found,
            journal_found=journal_found,
            volume_found=volume_found,
            issue_found=issue_found,
            pages_found=pages_found,
        )
