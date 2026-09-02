import re

def parse_abnt_reference(text: str) -> dict:
    """
    Tenta extrair os autores, o título e o DOI de uma referência bibliográfica (ABNT ou APA).
    Retorna {'authors': [...], 'title': '...', 'doi': '...'}
    """
    result = {'authors': [], 'title': None, 'doi': None}
    if not text:
        return result
        
    # Extração de DOI com regex
    doi_match = re.search(r'(10\.\d{4,9}/[-._;()/:a-zA-Z0-9]+)', text)
    if doi_match:
        result['doi'] = doi_match.group(1).rstrip('.')
        
    # Tentar heurística para APA (Ano entre parênteses)
    # Ex: Coledam, D. H. C., Arruda, G. A. (2023). Título.
    year_match = re.search(r'\(+\s*\d{4}[a-z]?\s*\)+', text)
    if year_match:
        authors_str = text[:year_match.start()].strip(' .')
        # Separa por ponto e vírgula seguido de espaço (., ) ou &
        authors_raw = [a.replace('&', '').strip() for a in re.split(r'\.,\s+|\s+&\s+', authors_str)]
        result['authors'] = [a for a in authors_raw if len(a) > 2]
        
        after_year = text[year_match.end():].strip(' .')
        title_end = after_year.find('. ')
        if title_end != -1:
            result['title'] = after_year[:title_end].strip()
        else:
            result['title'] = after_year.strip()
            
        # Se encontrou autores e título, podemos retornar
        if result['authors'] and result['title']:
            return result
            
    # Fallback ABNT tradicional (Autores em maiúsculas separados por ponto)
    parts = text.split('. ')
    if len(parts) >= 2:
        authors_part = parts[0]
        # verifica se tem aparência de autor (mais letras maiúsculas que minúsculas)
        if sum(1 for c in authors_part if c.isupper()) >= sum(1 for c in authors_part if c.islower()):
            authors_raw = authors_part.split(';')
            for a in authors_raw:
                if a.strip():
                    result['authors'].append(a.strip())
            
            result['title'] = parts[1].strip()
            
    return result
