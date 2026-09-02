from pydantic import BaseModel
from typing import List, Optional

class ReferenceInput(BaseModel):
    id: str
    original_text: str
    title: Optional[str] = None
    doi: Optional[str] = None
    authors: List[str] = []
    year: Optional[str] = None

class ValidationResult(BaseModel):
    api_source: str
    status: str # "Confirmado", "Parcialmente confirmado", "Não confirmado", "Possível Referência Falsa"
    title_found: Optional[str] = None
    doi_found: Optional[str] = None
    authors_found: List[str] = []
    issues: List[str] = []
    confidence: str # "Baixo", "Moderado", "Alto"
    url: Optional[str] = None
    # Campos adicionais para formatação APA7
    year_found: Optional[str] = None
    journal_found: Optional[str] = None
    volume_found: Optional[str] = None
    issue_found: Optional[str] = None
    pages_found: Optional[str] = None
