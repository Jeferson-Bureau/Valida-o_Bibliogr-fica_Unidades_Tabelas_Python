import unidecode
import re
import difflib

def normalize_text(text: str) -> str:
    """Remove acentos, caracteres especiais e transforma em minúsculas."""
    if not text:
        return ""
    text = unidecode.unidecode(text).lower()
    text = re.sub(r'[^a-z0-9]', '', text)
    return text

def get_http_session(retries: int = 3, backoff_factor: float = 1.0, user_agent: str = None) -> requests.Session:
    """Retorna uma sessão do requests configurada com retentativas e User-Agent adequado."""
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util import Retry

    session = requests.Session()
    retry_strategy = Retry(
        total=retries,
        backoff_factor=backoff_factor,
        status_forcelist=[429, 500, 502, 503, 504],
        raise_on_status=False
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    if not user_agent:
        user_agent = "ActaScientiarumValidator/1.0 (mailto:acta@uem.br)"
    session.headers.update({"User-Agent": user_agent})

    return session


def normalize_author_name(name: str) -> str:
    """Simplifica nome de autor para suportar diferentes formatos de abreviação.
    Ex: 'Silva, J.' -> 'silva j', 'João Silva' -> 'joao silva'
    """
    if not name:
        return ""
    name = unidecode.unidecode(name).lower()
    # Remove pontos e vírgulas
    name = name.replace('.', '').replace(',', '')
    return " ".join(name.split())

def calculate_similarity(str1: str, str2: str) -> float:
    """Calcula a similaridade entre duas strings usando SequenceMatcher."""
    s1 = normalize_text(str1)
    s2 = normalize_text(str2)
    if not s1 or not s2: 
        return 0.0
    if s1 == s2: 
        return 1.0
    if s1 in s2 or s2 in s1: 
        return 0.9 # heurística para contimento
    
    return difflib.SequenceMatcher(None, s1, s2).ratio()

def check_authors_match(input_authors: list[str], found_authors: list[str]) -> tuple[bool, list[str]]:
    """
    Verifica se os autores de entrada estão presentes nos autores encontrados na base.
    Retorna um booleano (se todos/maioria batem) e uma lista de issues (se faltarem autores).
    """
    if not input_authors and not found_authors:
        return True, []
        
    issues = []
    matches = 0
    
    found_tokens_list = []
    for f_author in found_authors:
        f_norm = unidecode.unidecode(f_author).lower().replace('.', '').replace(',', '')
        found_tokens_list.append(set(f_norm.split()))
        
    for in_author in input_authors:
        in_norm = unidecode.unidecode(in_author).lower().replace('.', '').replace(',', '')
        in_tokens = set(in_norm.split())
        
        found_match = False
        for f_tokens in found_tokens_list:
            # Simplificação: se a interseção dos nomes/sobrenomes for significativa
            # Ex: "Vaswani, Ashish" -> {"vaswani", "ashish"}
            # "Ashish Vaswani" -> {"ashish", "vaswani"}
            # A intersecção será 2.
            common = in_tokens.intersection(f_tokens)
            if len(common) >= 1 and len(in_tokens) > 0:
                # Tem pelo menos uma palavra em comum (ex: sobrenome)
                found_match = True
                break
        
        if found_match:
            matches += 1
        else:
            issues.append(f"Autor não encontrado: '{in_author}'")
            
    # Heurística: se acertou pelo menos todos os autores solicitados
    is_match = matches == len(input_authors) and len(input_authors) > 0
    return is_match, issues
