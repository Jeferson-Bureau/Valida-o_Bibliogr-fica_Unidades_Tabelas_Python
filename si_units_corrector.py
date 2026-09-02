"""
si_units_corrector.py — Corretor de Unidades Científicas para Python/python-docx
Porta da Macro VBA "Sistema Internacional de Unidades" v2.1
Normas: APA 7ª Edição + Sistema Internacional de Unidades (SI / BIPM)

Estratégia:
  - Processa runs individuais para substituições simples.
  - Para padrões que cruzam runs (ex: número em um run, unidade em outro),
    reconstrói o texto do parágrafo, aplica e redistribui no primeiro run.
  - Protege parágrafos após o título "Referências" (qualquer grafia).
"""

import re
from typing import List
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run

# ──────────────────────────────────────────────────────────────────────────────
# Caracteres especiais (mesmos do VBA)
# ──────────────────────────────────────────────────────────────────────────────
DEG   = "\u00B0"       # °  símbolo de grau correto
ORD   = "\u00BA"       # º  ordinal (errado para grau)
MU    = "\u03BC"       # μ  micro
OHM   = "\u03A9"       # Ω  Ohm
SM1   = "\u207B\u00B9" # ⁻¹
SM2   = "\u207B\u00B2" # ⁻²
SM3   = "\u207B\u00B3" # ⁻³
SUP2  = "\u00B2"       # ²
SUP3  = "\u00B3"       # ³
SUPn  = "\u207F"       # ⁿ
NBSP  = "\u00A0"       # espaço não quebrável
PH_DEG = "\uE000"      # placeholder PUA para °C durante bloco F

# ──────────────────────────────────────────────────────────────────────────────
# Lista de unidades SI simples (para bloco H: espaçamento número-unidade)
# ──────────────────────────────────────────────────────────────────────────────
SI_UNITS = {
    "kg","g","mg","ng","pg","t",
    "km","m","cm","mm","nm","pm",
    "kL","L","mL","dL",
    "GHz","MHz","kHz","Hz",
    "GPa","MPa","kPa","Pa",
    "MN","kN","N",
    "MJ","kJ","J","cal","kcal",
    "MW","kW","W",
    "MV","kV","V","mV",
    f"G{OHM}",f"M{OHM}",f"k{OHM}",OHM,"mA","A",
    "T","mT",
    "mmol","mol",
    "Bq","Gy","Sv",
    "bar","atm","mmHg","Torr",
    "h","min","s","ms","ns","ps",
    "K",
    # micro (Unicode)
    f"{MU}m",f"{MU}g",f"{MU}L",f"{MU}s",f"{MU}mol",
}

# Constrói regex para bloco H (número imediatamente seguido de unidade sem espaço)
_unit_re = re.compile(
    r'(\d)(' + '|'.join(re.escape(u) for u in sorted(SI_UNITS, key=len, reverse=True)) + r')(?=\s|$|[,;.)\]])',
)


def _apply_all(text: str) -> str:
    """Aplica todas as substituições de texto a uma string."""

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO A — Pré-limpeza e normalização de símbolos
    # ──────────────────────────────────────────────────────────────────────────
    # A1. Espaço não quebrável → espaço comum
    text = text.replace(NBSP, " ")

    # A2. Ordinal "º" → símbolo de grau "°" apenas antes de C
    text = text.replace(f"{ORD}C", f"{DEG}C")

    # A3. Operadores full-width → ASCII
    text = text.replace("\uFF1D", "=").replace("\uFF1C", "<").replace("\uFF1E", ">")

    # A4. Ohm por extenso
    text = re.sub(r'(?i) ohm\b', f' {OHM}', text)
    text = text.replace(f"kO", f"k{OHM}").replace(f"MO", f"M{OHM}").replace(f"GO", f"G{OHM}")

    # A5. "u" ASCII → μ (micro SI) quando precedido de número
    pairs_micro = [
        (r'(\d)\s?um\b',   r'\1 ' + MU + 'm'),
        (r'(\d)\s?ug\b',   r'\1 ' + MU + 'g'),
        (r'(\d)\s?uL\b',   r'\1 ' + MU + 'L'),
        (r'(\d)\s?us\b',   r'\1 ' + MU + 's'),
        (r'(\d)\s?umol\b', r'\1 ' + MU + 'mol'),
    ]
    for pat, rep in pairs_micro:
        text = re.sub(pat, rep, text)

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO B — Plurais indevidos em símbolos de unidades
    # ──────────────────────────────────────────────────────────────────────────
    plural_map = {
        "kgs":"kg","mgs":"mg","ngs":"ng","kms":"km","cms":"cm","mms":"mm","nms":"nm",
        "mLs":"mL","dLs":"dL","mins":"min",
        "kHzs":"kHz","MHzs":"MHz","GHzs":"GHz",
        "kPas":"kPa","MPas":"MPa","GPas":"GPa",
        "kNs":"kN","MNs":"MN","kJs":"kJ","MJs":"MJ","kWs":"kW","MWs":"MW",
        "mVs":"mV","kVs":"kV","mAs":"mA",
        "mols":"mol","mmols":"mmol","atms":"atm","Torrs":"Torr","mTs":"mT",
    }
    for wrong, correct in plural_map.items():
        text = re.sub(rf'\b{re.escape(wrong)}\b', correct, text)

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO C — Prefixos soltos (ex: "k m" → "km")
    # ──────────────────────────────────────────────────────────────────────────
    prefix_pairs = [
        ("k m","km"),("k g","kg"),("k Hz","kHz"),("k Pa","kPa"),("k N","kN"),
        ("k J","kJ"),("k W","kW"),("k V","kV"),
        ("M Hz","MHz"),("M Pa","MPa"),("M N","MN"),("M J","MJ"),("M W","MW"),("M V","MV"),
        ("G Hz","GHz"),("G Pa","GPa"),
        ("m m","mm"),("m L","mL"),("m V","mV"),("m A","mA"),("m T","mT"),
        ("m mol","mmol"),("m g","mg"),
        ("n m","nm"),("n g","ng"),("n s","ns"),
        (f"{MU} m",f"{MU}m"),(f"{MU} g",f"{MU}g"),(f"{MU} L",f"{MU}L"),
        (f"{MU} s",f"{MU}s"),(f"{MU} mol",f"{MU}mol"),
    ]
    for wrong, correct in prefix_pairs:
        text = re.sub(rf'(\d)\s?{re.escape(wrong)}', r'\1 ' + correct, text)

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO D — Espaços múltiplos
    # ──────────────────────────────────────────────────────────────────────────
    text = re.sub(r'  +', ' ', text)

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO E — Porcentagem (sem espaço antes de %)
    #   Correto: 45%    Incorreto: 45 %
    # ──────────────────────────────────────────────────────────────────────────
    text = re.sub(r'(\d)\s+%', r'\1%', text)

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO F — Temperatura e ângulos
    # ──────────────────────────────────────────────────────────────────────────
    # F1. Proteger °C com placeholder
    text = text.replace(f"{DEG}C", f"{PH_DEG}C")

    # F2. Ângulos: remover espaço antes de °, ′, ″
    PRIME  = "\u2032"  # ′ minuto
    DPRIME = "\u2033"  # ″ segundo
    text = re.sub(r'(\d)\s' + re.escape(DEG),    lambda m: m.group(1) + DEG,    text)
    text = re.sub(r'(\d)\s' + re.escape(PRIME),  lambda m: m.group(1) + PRIME,  text)
    text = re.sub(r'(\d)\s' + re.escape(DPRIME), lambda m: m.group(1) + DPRIME, text)

    # F3. Restaurar °C
    text = text.replace(f"{PH_DEG}C", f"{DEG}C")

    # F4. Celsius: remover espaço entre número e °C ou ° C
    #   Correto: 37°C    Incorreto: 37 °C, 37 ° C, 37 ºC
    text = re.sub(r'(\d)\s*' + re.escape(DEG) + r'\s*C\b', lambda m: m.group(1) + DEG + 'C', text)

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO G — Unidades compostas (barra → expoente negativo)
    # ──────────────────────────────────────────────────────────────────────────
    compound_map = [
        # Massa/volume
        (f"{MU}g/mL",  f"{MU}g mL{SM1}"), (f"{MU}g/dL",  f"{MU}g dL{SM1}"),
        (f"{MU}g/L",   f"{MU}g L{SM1}"),   (f"{MU}g/g",   f"{MU}g g{SM1}"),
        (f"{MU}g/100 g", f"{MU}g 100g{SM1}"), (f"{MU}g/100g", f"{MU}g 100g{SM1}"),
        ("ng/mL",  f"ng mL{SM1}"),  ("ng/L",  f"ng L{SM1}"),
        ("mg/mL",  f"mg mL{SM1}"),  ("mg/dL", f"mg dL{SM1}"),
        ("mg/ml",  f"mg mL{SM1}"),  ("mg/dl", f"mg dL{SM1}"),
        ("mg/L",   f"mg L{SM1}"),   ("mg/kg", f"mg kg{SM1}"),
        ("U/mL",   f"U mL{SM1}"),   ("U/ml",  f"U mL{SM1}"),
        ("U/mg",   f"U mg{SM1}"),
        ("10n/q",  f"10{SUPn} q{SM1}"),
        ("mg/100 g",f"mg 100g{SM1}"), ("mg/100g",f"mg 100g{SM1}"),
        ("g/mL",   f"g mL{SM1}"),   ("g/dL",  f"g dL{SM1}"),
        ("g/L",    f"g L{SM1}"),    ("g/kg",  f"g kg{SM1}"),
        # Substância/volume
        (f"{MU}mol/L",  f"{MU}mol L{SM1}"),
        ("nmol/L", f"nmol L{SM1}"),
        ("mmol/mL",f"mmol mL{SM1}"),("mmol/L", f"mmol L{SM1}"),
        ("mol/mL", f"mol mL{SM1}"), ("mol/L",  f"mol L{SM1}"),
        (f"mol/m{SUP3}",f"mol m{SM3}"), ("mol/m3", f"mol m{SM3}"),
        # Velocidade / aceleração
        (f"m/s{SUP2}", f"m s{SM2}"), ("m/s2", f"m s{SM2}"),
        ("km/h",  f"km h{SM1}"),  ("km/s", f"km s{SM1}"),
        ("m/s",   f"m s{SM1}"),
        # Força / pressão / energia / potência
        (f"N/m{SUP2}", f"N m{SM2}"), ("N/m2", f"N m{SM2}"),
        (f"W/m{SUP2}", f"W m{SM2}"), ("W/m2", f"W m{SM2}"),
        ("N/m",   f"N m{SM1}"),   ("Pa/m",  f"Pa m{SM1}"),
        ("kJ/mol/nm", f"kJ mol{SM1} nm{SM1}"),
        ("J/mol", f"J mol{SM1}"), ("J/kg",  f"J kg{SM1}"),
        ("J/g",   f"J g{SM1}"),   ("W/kg",  f"W kg{SM1}"),
        ("Kcal/g", f"Kcal g{SM1}"),
        # Densidade / Massa por área
        (f"kg/m{SUP3}", f"kg m{SM3}"), ("kg/m3", f"kg m{SM3}"),
        (f"kg/m{SUP2}", f"kg m{SM2}"), ("kg/m2", f"kg m{SM2}"),
        (f"g/cm{SUP3}", f"g cm{SM3}"), ("g/cm3", f"g cm{SM3}"),
        # Fluxo / débito
        ("mL/min",f"mL min{SM1}"),("L/min", f"L min{SM1}"),
        ("mL/h",  f"mL h{SM1}"),  ("L/h",   f"L h{SM1}"),
        ("mL/kg", f"mL kg{SM1}"), ("L/kg",  f"L kg{SM1}"),
        ("min/day",f"min day{SM1}"),("min/dia",f"min dia{SM1}"),
        ("min/d", f"min d{SM1}"),
        # Compostos / antioxidantes (case-insensitive tratado abaixo)
        ("GAE/kg",  f"GAE kg{SM1}"),  ("GAE/g",    f"GAE g{SM1}"),
        ("GAE/100 g",f"GAE 100g{SM1}"),("GAE/100g",f"GAE 100g{SM1}"),
        ("QE/100 g",f"QE 100g{SM1}"), ("QE/100g",  f"QE 100g{SM1}"),
        ("TEAC/100 g",f"TEAC 100g{SM1}"),("TEAC/100g",f"TEAC 100g{SM1}"),
        ("catechins/100 g",f"catechins 100g{SM1}"),("catechins/100g",f"catechins 100g{SM1}"),
        ("Trolox/kg",f"Trolox kg{SM1}"),
        ("ug/g",  f"{MU}g g{SM1}"),
        ("ug/L",  f"{MU}g L{SM1}"),  ("ug/mL", f"{MU}g mL{SM1}"),
        ("ug/100 g",f"{MU}g 100g{SM1}"),("ug/100g",f"{MU}g 100g{SM1}"),
        # Microbiologia: CFU (Colony Forming Units) → UFC mL⁻¹ etc.
        # Variações com CFU (inglês) → UFC (português) + expoente negativo
        ("CFU/mL",  f"UFC mL{SM1}"),  ("CFU/ml",  f"UFC mL{SM1}"),
        ("CFU/L",   f"UFC L{SM1}"),
        ("CFU/g",   f"UFC g{SM1}"),
        (f"CFU/cm{SUP2}", f"UFC cm{SM2}"), ("CFU/cm2", f"UFC cm{SM2}"),
        (f"CFU/m{SUP2}",  f"UFC m{SM2}"),  ("CFU/m2",  f"UFC m{SM2}"),
        ("CFU/mL/h",f"UFC mL{SM1} h{SM1}"),
        # Variações com UFC (já em português, mas com barra)
        ("UFC/mL",  f"UFC mL{SM1}"),  ("UFC/ml",  f"UFC mL{SM1}"),
        ("UFC/L",   f"UFC L{SM1}"),
        ("UFC/g",   f"UFC g{SM1}"),
        (f"UFC/cm{SUP2}", f"UFC cm{SM2}"), ("UFC/cm2", f"UFC cm{SM2}"),
        (f"UFC/m{SUP2}",  f"UFC m{SM2}"),  ("UFC/m2",  f"UFC m{SM2}"),
        # Temperatura / taxa
        (f"{DEG}C/min", f"{DEG}C min{SM1}"),
    ]
    for wrong, correct in compound_map:
        text = text.replace(wrong, correct)

    # G1b. m/m com regex (word boundary evita falsos positivos em cm/min etc.)
    text = re.sub(r'\bm/m\b', f'm m{SM1}', text)

    # G2. Pontos de produto → espaço
    for dot in ["\u00B7", "\u22C5", "\u2219", "\u2022"]:
        text = text.replace(dot, " ")

    text = re.sub(r'  +', ' ', text)

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO H — Espaço entre número e unidade SI
    # ──────────────────────────────────────────────────────────────────────────
    text = _unit_re.sub(r'\1 \2', text)

    # H2. Unidades com μ (Unicode)
    for u_suffix in ["m","g","L","s","mol"]:
        text = re.sub(rf'(\d){re.escape(MU)}{re.escape(u_suffix)}', r'\1 ' + MU + u_suffix, text)

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO I — Operadores matemáticos (espaço antes e depois)
    # Aplica apenas quando o operador está entre caracteres no mesmo run.
    # O caso cross-run (± em run separado) é tratado por _fix_cross_run_operators.
    # ──────────────────────────────────────────────────────────────────────────
    OPERATORS = ["=", "<", ">", "±", "~", "≤", "≥", "≠", "≈"]
    for op in OPERATORS:
        e = re.escape(op)
        # Só espaceja se houver pelo menos um caractere não-espaço em cada lado
        text = re.sub(rf'(\S)\s*{e}\s*(\S)', lambda m: f'{m.group(1)} {op} {m.group(2)}', text)
        
        # Remove espaços ao redor do operador se estiver após parêntese/colchete aberto (ex: "( < " -> "(<")
        text = text.replace(f"( {op} ", f"({op}")
        text = text.replace(f"[ {op} ", f"[{op}")

    text = re.sub(r'  +', ' ', text)

    # ──────────────────────────────────────────────────────────────────────────
    # BLOCO J — Limpeza final de pontuação
    # ──────────────────────────────────────────────────────────────────────────
    text = re.sub(r' ([.,;:])', r'\1', text)
    # Nota: NÃO aplicar strip() aqui — runs isolados com operadores (ex: '±')
    # precisam dos espaços nas bordas para separar do run vizinho.

    return text


def _is_references_heading(para: Paragraph) -> bool:
    """Retorna True se o parágrafo for o título da seção Referências."""
    text = para.text.strip()
    return bool(re.fullmatch(r'refer[eê]ncias?', text, re.IGNORECASE))


def _process_run(run: Run) -> None:
    """Aplica correções ao texto de um único run."""
    original = run.text
    if not original:
        return
    corrected = _apply_all(original)
    if corrected != original:
        run.text = corrected


OPERATOR_CHARS = {"±", "=", "<", ">", "~", "≤", "≥", "≠", "≈"}


def _fix_cross_run_operators(para: Paragraph) -> None:
    """
    Garante espaço antes/depois de operadores matemáticos que estão
    em runs separados dos números ao redor.

    Abordagem: insere exatamente UM espaço no final do run anterior
    e UM espaço no início do run seguinte, sem duplicar.
    """
    runs = para.runs
    for i, run in enumerate(runs):
        txt = run.text
        if not txt:
            continue

        # Verifica se este run contém algum operador
        contains_op = any(op in txt for op in OPERATOR_CHARS)
        if not contains_op:
            continue

        # Garante exatamente 1 espaço no final do run anterior
        if i > 0:
            prev = runs[i - 1]
            if prev.text and not prev.text.endswith(" "):
                prev.text = prev.text + " "

        # Garante exatamente 1 espaço no início do run seguinte
        if i < len(runs) - 1:
            nxt = runs[i + 1]
            if nxt.text and not nxt.text.startswith(" "):
                nxt.text = " " + nxt.text

        # Remove espaços nas bordas do run do operador (evita duplo espaço
        # quando o Bloco I já adicionou espaços internos como " ± "):
        # o espaçamento fica exclusivamente nos runs vizinhos.
        run.text = txt.strip()


def _merge_split_compound_units(para: Paragraph) -> None:
    """
    No Word, expoentes formatados (ex: 'kg/m' normal e '2' sobrescrito) ficam em runs separados.
    Isso impede que o _apply_all identifique a unidade inteira.
    Esta função funde o expoente no run da base antes do processamento.
    """
    for i in range(len(para.runs) - 1):
        r1 = para.runs[i]
        r2 = para.runs[i+1]
        if not r1.text or not r2.text:
            continue
            
        if r1.text.endswith(("/m", "/cm", "/s", "/m ", "/cm ", "/s ")):
            match = re.match(r'^(\s*)([23²³])(?:\s|[,.;)]|$)', r2.text)
            if match:
                spaces = match.group(1)
                exp = match.group(2)
                r1.text = r1.text + spaces + exp
                r2.text = r2.text[len(spaces) + 1:]


def _fix_cross_run_spacing(para: Paragraph) -> None:
    """
    Remove espaços indevidos entre números e os símbolos % ou °C que acabaram
    separados em runs diferentes.
    """
    runs = para.runs
    for i in range(len(runs) - 1):
        r1 = runs[i]
        r2 = runs[i+1]
        if not r1.text or not r2.text:
            continue
            
        t2 = r2.text.lstrip()
        if t2.startswith(("%", "°C", "\u00b0C")):
            t1 = r1.text.rstrip()
            if t1 and t1[-1].isdigit():
                r1.text = t1
                r2.text = t2


def _fix_cross_run_compounds(para: Paragraph) -> None:
    """
    Aplica substituições de unidades compostas (compound_map) ao nível do
    parágrafo inteiro, resolvendo padrões que ficaram divididos entre runs.

    Estratégia:
      1. Concatena o texto de todos os runs.
      2. Aplica as mesmas substituições do compound_map.
      3. Se houve alteração, redistribui o texto corrigido nos runs,
         preservando os limites originais de caracteres (e portanto a
         formatação de cada run).
    """
    runs = para.runs
    if not runs:
        return

    full_text = "".join(r.text or "" for r in runs)
    corrected = full_text

    # Aplica compound_map (mesma lista do _apply_all Bloco G)
    compound_map = [
        # Padrões com múltiplas barras (devem vir primeiro)
        ("kJ/mol/nm", f"kJ mol{SM1} nm{SM1}"),
        ("CFU/mL/h", f"UFC mL{SM1} h{SM1}"),
        # Temperatura / taxa
        (f"{DEG}C/min", f"{DEG}C min{SM1}"),
    ]
    for wrong, correct in compound_map:
        corrected = corrected.replace(wrong, correct)
    # m/m com regex (word boundary evita falsos positivos)
    corrected = re.sub(r'\bm/m\b', f'm m{SM1}', corrected)

    if corrected == full_text:
        return

    # Redistribui o texto corrigido respeitando os limites dos runs
    # Quando o texto encolhe/cresce, o excedente é absorvido pelo último run
    pos = 0
    for i, run in enumerate(runs):
        orig_len = len(run.text or "")
        if i < len(runs) - 1:
            run.text = corrected[pos:pos + orig_len]
            pos += orig_len
        else:
            # Último run recebe o restante
            run.text = corrected[pos:]


def _process_paragraph(para: Paragraph) -> None:
    """Processa todos os runs de um parágrafo."""
    # 0ª passagem: fundir unidades compostas que foram divididas por formatação
    _merge_split_compound_units(para)
    
    # 1ª passagem: substituições run a run
    for run in para.runs:
        _process_run(run)
    # 1.5ª passagem: unidades compostas que cruzam runs
    _fix_cross_run_compounds(para)
    # 2ª passagem: espaçamento cross-run para operadores em runs separados
    _fix_cross_run_operators(para)
    # 3ª passagem: espaçamento cross-run para % e °C colados no número
    _fix_cross_run_spacing(para)
    # 4ª passagem: colapsar espaços duplos que possam surgir nas bordas dos runs
    for run in para.runs:
        if run.text and "  " in run.text:
            run.text = re.sub(r"  +", " ", run.text)


def correct_si_units(doc: Document) -> int:
    """
    Aplica todas as correções de unidades SI ao documento.
    Para quando encontra a seção "Referências" (protegida).
    Retorna o número de parágrafos processados.
    """
    processed = 0
    in_references = False

    # Processar corpo principal
    for para in doc.paragraphs:
        if _is_references_heading(para):
            in_references = True
        if in_references:
            continue
        _process_paragraph(para)
        processed += 1

    # Processar cabeçalhos e rodapés de todas as seções
    for section in doc.sections:
        for header_footer in [
            section.header, section.footer,
            section.even_page_header, section.even_page_footer,
            section.first_page_header, section.first_page_footer,
        ]:
            try:
                for para in header_footer.paragraphs:
                    _process_paragraph(para)
                    processed += 1
            except Exception:
                pass

    # Processar tabelas (corpo principal, antes de Referências)
    in_references = False
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _process_paragraph(para)
                    processed += 1

    return processed
