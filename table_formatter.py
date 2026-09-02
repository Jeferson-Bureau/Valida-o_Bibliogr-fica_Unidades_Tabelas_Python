import docx
from table_style import TableStyleSpec, ACTA_SCIENTIARUM_STYLE, apply_table_style, apply_to_all_tables

NOME_FONTE = "PT Serif"
TAM_FONTE = 8  # pt
MARGEM_LATERAL_INCHES = 0.05

def format_table_acta(table, font_name=NOME_FONTE, font_size_pt=TAM_FONTE, margin_inches=MARGEM_LATERAL_INCHES):
    """
    Aplica a formatação completa da Acta Scientiarum a uma tabela python-docx.
    Utiliza o módulo refinado table_style.py.
    """
    estilo = TableStyleSpec()
    estilo.font.name = font_name
    estilo.font.size_pt = font_size_pt
    # 0.05 in -> 3.6 pt (~3.5 pt)
    estilo.cell_margin_left_pt = margin_inches * 72.0
    estilo.cell_margin_right_pt = margin_inches * 72.0
    
    apply_table_style(table, estilo)


def format_all_tables_in_docx(doc_or_path, font_name=NOME_FONTE, font_size_pt=TAM_FONTE) -> int:
    """
    Percorre e formata todas as tabelas em um documento Document ou caminho de arquivo.
    Retorna o número de tabelas formatadas.
    """
    if isinstance(doc_or_path, str):
        doc = docx.Document(doc_or_path)
    else:
        doc = doc_or_path

    estilo = TableStyleSpec()
    estilo.font.name = font_name
    estilo.font.size_pt = font_size_pt

    return apply_to_all_tables(doc, estilo)

