import docx
import re
from typing import List, Tuple
from docx.document import Document
from docx.text.paragraph import Paragraph
import unidecode

def extract_references_from_docx(filepath: str) -> Tuple[Document, List]:
    """
    Lê um arquivo .docx, localiza a seção de referências e retorna:
    - O objeto Document (para salvarmos com modificações depois)
    - Uma lista de tuplas (parágrafo, texto_da_referência)
      onde `parágrafo` é o objeto Paragraph do python-docx correspondente,
      garantindo que anotações nunca atinjam o corpo do texto.
    """
    doc = docx.Document(filepath)

    references = []
    in_references_section = False

    ref_patterns = [
        r'^referências',
        r'^referencias',
        r'^references'
    ]

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        if not in_references_section:
            text_lower = text.lower()
            for pattern in ref_patterns:
                if re.match(pattern, text_lower):
                    in_references_section = True
                    break
            continue

        if in_references_section:
            if len(text) < 15:
                continue
            # Ignora linhas que parecem novo cabeçalho de seção (todo maiúsculo e curto)
            if text.isupper() and len(text.split()) < 4:
                continue
            references.append((para, text))

    # Ordena a lista em memória alfabeticamente sem tocar na estrutura XML do documento
    if references:
        references.sort(key=lambda item: unidecode.unidecode(item[1]).strip().lower())

    return doc, references
