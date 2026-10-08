"""
TPEF — Extracción de denominaciones equivalentes
=================================================
Autor: C. O. Rocha Ochoa
Descripción:
    Extrae las denominaciones equivalentes codificadas dentro de cada entrada del corpus mediante seis patrones de reconocimiento (P1-P6) aplicados sobre los campos de lema y definición.
    Patrones:
      P1 — Acrónimo + guion + nombre completo en el lema
           Ej.: «BADLAR - Buenos Aires Deposits of Large Amount Rate»
      P2 — Nombre completo + acrónimo entre paréntesis en el lema
           Ej.: «Producto Bruto Interno (PBI)»
      P3 — Expansión parentética de sigla en la definición
           Ej.: «La CBU (Cuenta Bancaria Uniforme) es...»
      P4 — Cadena de sinónimos introducida por «se llama X, Y o Z»
           en la definición
      P5 — Sinónimo introducido por construcciones de
           Identificación en la definición: «también
           conocido/llamado/denominado como X»
      P6 — Sinónimo declarado léxicamente mediante el término
           «sinónimo» en lema o definición: «es sinónimo de X»
Entrada:  Glosarios_2026.xlsx
Salida:   tpef_denominaciones_equivalentes.xlsx
          tpef_denominaciones_resumen.xlsx
Requisitos:
    pip install openpyxl pandas
"""
import re
import unicodedata
from collections import defaultdict
from openpyxl import load_workbook
import pandas as pd
# ── Configuración ─────────────────────────────────────────────────────────────
INPUT_FILE   = 'Glosarios_2026.xlsx'
OUTPUT_FILE  = 'tpef_denominaciones_equivalentes.xlsx'
SUMMARY_FILE = 'tpef_denominaciones_resumen.xlsx'
COUNTRY_LABELS = {
    'AR':'Argentina','BO':'Bolivia','CL':'Chile','CO':'Colombia',
    'CR':'Costa Rica','CU':'Cuba','DO':'Rep. Dominicana','EC':'Ecuador',
    'ES':'España','GT':'Guatemala','HN':'Honduras','MX':'México',
    'NI':'Nicaragua','PA':'Panamá','PE':'Perú','PY':'Paraguay',
    'SV':'El Salvador','VE':'Venezuela'
}
# ── Función de normalización ──────────────────────────────────────────────────
def normalize(s):
    """Convierte a minúsculas y elimina diacríticos."""
    if not s:
        return ''
    s = str(s).strip().lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s
# ── Patrones de extracción ────────────────────────────────────────────────────
def extract_denominations(lema, defn):
    """
    Aplica los seis patrones de extracción sobre el lema y la
    definición de una entrada. Devuelve una lista de denominaciones
    con el lema principal en primer lugar y las variantes extraídas
    a continuación, junto con una lista de los patrones activados.
    Una entrada activa un patrón como máximo una vez: la primera
    denominación adicional extraída por un patrón impide que el
    mismo patrón añada más variantes de menor calidad.
    """
    denominations = [lema]
    patterns_activated = []
    # ── P1: ACRONYM - Full Name en el lema ───────────────────────────────────
    # Formato: «SIGLA - Nombre completo» o «SIGLA – Nombre completo»
    m1 = re.match(r'^([A-ZÁÉÍÓÚÑ]{2,12})\s*[-–]\s*(.{5,80})$', lema)
    if m1:
        acr  = m1.group(1).strip()
        full = m1.group(2).strip()
        if acr not in denominations:
            denominations.append(acr)
        if full not in denominations:
            denominations.append(full)
        patterns_activated.append('P1')
    # ── P2: Full Name (ACRONYM) en el lema ───────────────────────────────────
    # Formato: «Nombre completo (SIGLA)»
    m2 = re.match(
        r'^(.{4,70}?)\s*\(([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ0-9\-]{1,10})\)\s*$', lema
    )
    if m2:
        full = m2.group(1).strip()
        acr  = m2.group(2).strip()
        if full not in denominations:
            denominations.append(full)
        if acr not in denominations:
            denominations.append(acr)
        patterns_activated.append('P2')
    # ── P3: Expansión parentética de sigla en la definición ──────────────────
    # Formato: «SIGLA (Nombre completo)» dentro de la definición
    for m in re.finditer(
        r'\b([A-ZÁÉÍÓÚÑ]{2,10})\s*\(([A-ZÁÉÍÓÚÑA-Za-záéíóúñ][^)\]]{4,60})\)',
        defn
    ):
        acr  = m.group(1).strip()
        full = m.group(2).strip()
        if acr not in denominations and acr != lema[:len(acr)]:
            denominations.append(acr)
            if 'P3' not in patterns_activated:
                patterns_activated.append('P3')
        if full not in denominations and len(full) > 4:
            denominations.append(full)
    # ── P4: Cadena de sinónimos «se llama X, Y o Z» en la definición ─────────
    m4 = re.match(r'^[Ss]e\s+llama[n]?\s+(.+?)[,\.]\s', defn)
    if m4:
        chain = m4.group(1)
        parts = re.split(r'\s+o\s+|,\s*', chain)
        for p in parts:
            p = p.strip().strip('"\'«»')
            if p and 2 < len(p) <= 60 and len(p.split()) <= 6 \
                    and p not in denominations:
                denominations.append(p)
                if 'P4' not in patterns_activated:
                    patterns_activated.append('P4')
    # ── P5: «también conocido/llamado/denominado como X» en la definición ────
    for m in re.finditer(
        r'[Tt]ambién\s+(?:se\s+)?(?:conocid[ao]s?|llamad[ao]s?|'
        r'denominad[ao]s?)\s+(?:como\s+)?[«"\']?'
        r'([A-ZÁÉÍÓÚÑa-záéíóúñ][^,\.;\(\)«»"\']{2,50})[»"\']?',
        defn
    ):
        candidate = m.group(1).strip().rstrip('.,;)')
        if candidate and len(candidate.split()) <= 7 \
                and candidate not in denominations:
            denominations.append(candidate)
            if 'P5' not in patterns_activated:
                patterns_activated.append('P5')
    # ── P6: «sinónimo de X» o «es sinónimo de X» en lema o definición ────────
    p6_patterns = [
        r'[Ss]inónimo\s+de\s+([A-ZÁÉÍÓÚÑa-záéíóúñ][^,\.;\(\)]{2,50})',
        r'[Ee]s\s+sinónimo\s+de\s+([A-ZÁÉÍÓÚÑa-záéíóúñ][^,\.;\(\)]{2,50})',
        r'[Tt]ambién\s+denominado\s+([A-ZÁÉÍÓÚÑa-záéíóúñ][^,\.;\(\)]{2,50})',
        r'[Ll]lamado\s+también\s+([A-ZÁÉÍÓÚÑa-záéíóúñ][^,\.;\(\)]{2,50})',
        r'[Dd]enominado\s+también\s+([A-ZÁÉÍÓÚÑa-záéíóúñ][^,\.;\(\)]{2,50})',
    ]
    search_text = lema + ' ' + defn
    for pat in p6_patterns:
        for m in re.finditer(pat, search_text):
            candidate = m.group(1).strip().rstrip('.,;)')
            if candidate and len(candidate.split()) <= 7 \
                    and candidate not in denominations:
                denominations.append(candidate)
                if 'P6' not in patterns_activated:
                    patterns_activated.append('P6')
    return denominations, patterns_activated
# ── Script principal ──────────────────────────────────────────────────────────
def main():
    print("TPEF — Extracción de denominaciones equivalentes")
    print("="*55)
    wb = load_workbook(INPUT_FILE, read_only=True)
    print(f"Corpus: {len(wb.sheetnames)} glosarios")
    rows_out   = []
    sheet_stats = defaultdict(lambda: {
        'total': 0, 'multi': 0, 'extra_dens': 0,
        'P1':0,'P2':0,'P3':0,'P4':0,'P5':0,'P6':0
    })
    corpus_total = 0
    corpus_multi  = 0
    corpus_extra  = 0
    pattern_totals = defaultdict(int)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        country = sheet_name[:2]
        for row in ws.iter_rows(values_only=True):
            if not (row[0] and str(row[0]).strip()):
                continue
            lema = str(row[0]).strip()
            defn = str(row[1]).strip() if row[1] else ''
            dens, pats = extract_denominations(lema, defn)
            n_extra = len(dens) - 1
            corpus_total  += 1
            sheet_stats[sheet_name]['total'] += 1
            if n_extra > 0:
                corpus_multi  += 1
                corpus_extra  += n_extra
                sheet_stats[sheet_name]['multi'] += 1
                sheet_stats[sheet_name]['extra_dens'] += n_extra
                for p in pats:
                    sheet_stats[sheet_name][p] += 1
                    pattern_totals[p] += 1
            # Output row: one row per denomination
            for i, den in enumerate(dens):
                rows_out.append({
                    'Glosario':      sheet_name,
                    'País':          COUNTRY_LABELS.get(country, country),
                    'Lema principal': lema,
                    'Denominación':  den,
                    'Posición':      i,          # 0 = lema principal
                    'Es variante':   'Sí' if i > 0 else 'No',
                    'Patrones':      '+'.join(pats) if pats else '',
                    'N variantes':   n_extra,
                })
    # ── Exportar entradas con denominaciones ──────────────────────────────────
    df = pd.DataFrame(rows_out)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nDenominaciones extraídas → {OUTPUT_FILE}")
    # ── Resumen global ────────────────────────────────────────────────────────
    corpus_dens = corpus_total + corpus_extra
    print(f"\nTotal entradas del corpus:         {corpus_total:>7,}")
    print(f"Entradas con >1 denominación:      {corpus_multi:>7,} "
          f"({corpus_multi/corpus_total*100:.1f}%)")
    print(f"Denominaciones adicionales (extra): {corpus_extra:>7,}")
    print(f"Total denominaciones:              {corpus_dens:>7,} "
          f"(+{corpus_extra/corpus_total*100:.1f}% respecto a entradas)")
    print(f"\nPatrones activados:")
    for p in ['P1','P2','P3','P4','P5','P6']:
        n = pattern_totals[p]
        print(f"  {p}: {n:,} ({n/corpus_total*100:.1f}%)")
    # ── Resumen por glosario ──────────────────────────────────────────────────
    summary_rows = []
    for sheet_name in wb.sheetnames:
        s = sheet_stats[sheet_name]
        tot = s['total']
        if tot == 0:
            continue
        pct_increase = s['extra_dens'] / tot * 100 if tot > 0 else 0
        summary_rows.append({
            'Glosario':          sheet_name,
            'País':              COUNTRY_LABELS.get(sheet_name[:2], ''),
            'Entradas':          tot,
            'Entradas c/>1 den.': s['multi'],
            '% entradas c/>1':   round(s['multi']/tot*100, 1),
            'Den. adicionales':  s['extra_dens'],
            'Total den.':        tot + s['extra_dens'],
            '% aumento':         round(pct_increase, 1),
            'P1':                s['P1'],
            'P2':                s['P2'],
            'P3':                s['P3'],
            'P4':                s['P4'],
            'P5':                s['P5'],
            'P6':                s['P6'],
        })
    df_summary = pd.DataFrame(summary_rows)
    df_summary.to_excel(SUMMARY_FILE, index=False)
    print(f"\nResumen por glosario → {SUMMARY_FILE}")
    print("\nFinalizado.")
if __name__ == '__main__':
    main()
