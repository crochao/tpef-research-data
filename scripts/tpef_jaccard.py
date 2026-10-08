"""
TPEF — Normalización y cálculo del índice de similitud de Jaccard
=================================================================
Autor: C. O. Rocha Ochoa
Descripción:
    Calcula el índice de similitud de Jaccard entre todos los pares
    de subcorpus nacionales del corpus TPEF, a partir de los conjuntos
    de lemas únicos normalizados de cada país. Produce:
    (1) Una matriz de similitud de 18 × 18 exportada como Excel.
    (2) Un mapa de calor en formato PNG.
    (3) Una tabla de estadísticas globales (media, máximo, mínimo,
        pares más y menos convergentes).
    (4) Una tabla del núcleo terminológico: lemas ordenados por
        número de glosarios en los que aparecen.
    El índice de Jaccard entre dos conjuntos A y B se define como:
    donde A y B son los conjuntos de lemas únicos normalizados de
    dos subcorpus nacionales. Los valores oscilan entre 0 (ningún
    lema compartido) y 1 (conjuntos idénticos).
Entrada:  Glosarios_2026.xlsx
Salida:   tpef_jaccard_matriz.xlsx
          tpef_jaccard_heatmap.png
          tpef_nucleo_terminologico.xlsx
Requisitos:
    pip install openpyxl pandas numpy matplotlib seaborn
"""
import unicodedata
import re
from collections import defaultdict, Counter
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import load_workbook
# ── Configuración ─────────────────────────────────────────────────────────────
INPUT_FILE    = 'Glosarios_2026.xlsx'
MATRIX_FILE   = 'tpef_jaccard_matriz.xlsx'
HEATMAP_FILE  = 'tpef_jaccard_heatmap.png'
NUCLEUS_FILE  = 'tpef_nucleo_terminologico.xlsx'
ISO3 = {
    'AR':'ARG','BO':'BOL','CL':'CHL','CO':'COL','CR':'CRI','CU':'CUB',
    'DO':'DOM','EC':'ECU','ES':'ESP','GT':'GTM','HN':'HND','MX':'MEX',
    'NI':'NIC','PA':'PAN','PE':'PER','PY':'PRY','SV':'SLV','VE':'VEN'
}
COUNTRY_LABELS = {
    'AR':'Argentina','BO':'Bolivia','CL':'Chile','CO':'Colombia',
    'CR':'Costa Rica','CU':'Cuba','DO':'Rep. Dominicana','EC':'Ecuador',
    'ES':'España','GT':'Guatemala','HN':'Honduras','MX':'México',
    'NI':'Nicaragua','PA':'Panamá','PE':'Perú','PY':'Paraguay',
    'SV':'El Salvador','VE':'Venezuela'
}
# ── Paso 1: Normalización ─────────────────────────────────────────────────────
def normalize(s):
    """
    Normaliza un lema para comparación:
      - Elimina espacios iniciales y finales
      - Convierte a minúsculas
      - Elimina diacríticos (NFD + filtro de categoría Mn)
      - Unifica espacios internos y comillas tipográficas
    """
    if not s:
        return ''
    s = str(s).strip().lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    s = re.sub(r'\s+', ' ', s)
    s = s.translate(str.maketrans({
        '‘': "'", '’': "'", '‚': "'", '‛': "'",
        '“': '"', '”': '"', '„': '"', '‟': '"', '«': '"', '»': '"',
    }))
    return s
# ── Paso 2: Carga y construcción de conjuntos de lemas ───────────────────────
def load_corpus(filepath):
    """
    Carga el corpus y construye:
      - sheet_unique:   {nombre_glosario: set(lemas normalizados)}
      - country_unique: {codigo_pais:     set(lemas normalizados)}
      - country_order:  lista de países en el orden del archivo
      - lema_glossary_count: Counter con el número de glosarios
                             en que aparece cada lema
    """
    wb = load_workbook(filepath, read_only=True)
    sheet_unique   = {}
    country_unique = defaultdict(set)
    country_order  = []
    seen_countries = set()
    lema_glossary_count = Counter()
    for sheet_name in wb.sheetnames:
        country = sheet_name[:2]
        if country not in seen_countries:
            country_order.append(country)
            seen_countries.add(country)
        ws = wb[sheet_name]
        lemas = set()
        for row in ws.iter_rows(values_only=True):
            if row[0] and str(row[0]).strip():
                lema_n = normalize(str(row[0]))
                lemas.add(lema_n)
        sheet_unique[sheet_name] = lemas
        country_unique[country] |= lemas
        for lema in lemas:
            lema_glossary_count[lema] += 1
    return sheet_unique, country_unique, country_order, lema_glossary_count, wb.sheetnames
# ── Paso 3: Cálculo de la matriz de Jaccard ───────────────────────────────────
def jaccard(set_a, set_b):
    """Calcula el índice de Jaccard entre dos conjuntos."""
    union = set_a | set_b
    if not union:
        return 0.0
    return len(set_a & set_b) / len(union)
def build_jaccard_matrix(country_unique, country_order):
    """
    Construye la matriz de similitud de Jaccard de N × N
    para todos los pares de subcorpus nacionales.
    """
    n = len(country_order)
    matrix = np.zeros((n, n))
    for i, c1 in enumerate(country_order):
        for j, c2 in enumerate(country_order):
            matrix[i, j] = jaccard(country_unique[c1], country_unique[c2])
    return matrix
# ── Paso 4: Mapa de calor ─────────────────────────────────────────────────────
def plot_heatmap(matrix, labels, output_file):
    """
    Genera el mapa de calor de la matriz de Jaccard.
    La diagonal (autocomparación, J = 1) se colorea en rojo oscuro
    y se enmascara para el cálculo de la escala de colores.
    """
    n = len(labels)
    mask = np.eye(n, dtype=bool)
    fig, ax = plt.subplots(figsize=(13, 10))
    sns.heatmap(
        matrix, ax=ax, mask=mask,
        cmap='Reds', vmin=0, vmax=0.25,
        xticklabels=labels, yticklabels=labels,
        linewidths=0.4, linecolor='#f5e8e8',
        annot=True, fmt='.2f', annot_kws={'size': 7.5},
        cbar_kws={'label': 'Similitud de Jaccard', 'shrink': 0.7}
    )
    # Colorear diagonal en rojo oscuro
    for i in range(n):
        ax.add_patch(plt.Rectangle(
            (i, i), 1, 1,
            fill=True, color='#8b0000', alpha=0.85, zorder=3
        ))
    ax.set_title(
        'Convergencia léxica entre subcorpus nacionales\n'
        '(índice de similitud de Jaccard, 0–1)',
        fontsize=13, fontweight='bold', pad=16
    )
    ax.tick_params(axis='x', rotation=45, labelsize=9)
    ax.tick_params(axis='y', rotation=0,  labelsize=9)
    plt.tight_layout()
    plt.savefig(output_file, dpi=180, bbox_inches='tight')
    plt.close()
    print(f"  Mapa de calor guardado en: {output_file}")
# ── Script principal ──────────────────────────────────────────────────────────
def main():
    print("TPEF — Normalización y cálculo del índice de Jaccard")
    print("="*55)
    # ── Carga ─────────────────────────────────────────────────────────────────
    print("\nCargando corpus...")
    sheet_unique, country_unique, country_order, \
        lema_count, all_sheets = load_corpus(INPUT_FILE)
    n_glossaries = len(all_sheets)
    n_countries  = len(country_order)
    global_union = set()
    for lemas in sheet_unique.values():
        global_union |= lemas
    total_entries = sum(len(s) for s in sheet_unique.values())
    print(f"  Glosarios:      {n_glossaries}")
    print(f"  Países:         {n_countries}")
    print(f"  Entradas brutas (suma únicos por glosario): {total_entries:,}")
    print(f"  Lemas únicos globales (unión): {len(global_union):,}")
    # ── Matriz de Jaccard ─────────────────────────────────────────────────────
    print("\nCalculando matriz de Jaccard...")
    countries_sorted = sorted(country_order, key=lambda c: ISO3.get(c, c))
    labels = [ISO3.get(c, c) for c in countries_sorted]
    country_unique_sorted = {c: country_unique[c] for c in countries_sorted}
    matrix = build_jaccard_matrix(country_unique_sorted, countries_sorted)
    # Exportar matriz como Excel
    df_matrix = pd.DataFrame(matrix, index=labels, columns=labels).round(3)
    df_matrix.to_excel(MATRIX_FILE)
    print(f"  Matriz exportada: {MATRIX_FILE}")
    # ── Estadísticas globales ─────────────────────────────────────────────────
    print("\nEstadísticas globales:")
    off_diag = matrix[~np.eye(len(countries_sorted), dtype=bool)]
    print(f"  Media (fuera de la diagonal):  {off_diag.mean():.3f}")
    print(f"  Máximo:                        {off_diag.max():.3f}")
    print(f"  Mínimo:                        {off_diag.min():.3f}")
    # Pares extremos
    pairs = []
    n = len(countries_sorted)
    for i in range(n):
        for j in range(i+1, n):
            pairs.append((matrix[i,j], labels[i], labels[j]))
    pairs.sort(reverse=True)
    print("\n  Top 5 pares con mayor convergencia:")
    for v, a, b in pairs[:5]:
        print(f"    {a}-{b}: {v:.3f}")
    print("\n  Top 5 pares con menor convergencia (excluido 0):")
    non_zero = [(v,a,b) for v,a,b in pairs if v > 0]
    for v,a,b in non_zero[-5:]:
        print(f"    {a}-{b}: {v:.3f}")
    # Hapax legomena
    hapax = sum(1 for v in lema_count.values() if v == 1)
    print(f"\n  Hapax legomena (1 solo glosario): {hapax:,} "
          f"({hapax/len(global_union)*100:.1f}%)")
    # ── Mapa de calor ─────────────────────────────────────────────────────────
    print("\nGenerando mapa de calor...")
    plot_heatmap(matrix, labels, HEATMAP_FILE)
    # ── Núcleo terminológico ──────────────────────────────────────────────────
    print("\nGenerando tabla del núcleo terminológico...")
    nucleus_rows = []
    for lema, count in lema_count.most_common():
        nucleus_rows.append({
            'Lema':          lema,
            'N glosarios':   count,
            '% glosarios':   round(count / n_glossaries * 100, 1),
            'Cobertura':     (
                'Núcleo amplio (≥50%)'  if count >= n_glossaries * 0.50 else
                'Núcleo medio (25–49%)' if count >= n_glossaries * 0.25 else
                'Periferia (<25%)'
            ),
        })
    df_nucleus = pd.DataFrame(nucleus_rows)
    df_nucleus.to_excel(NUCLEUS_FILE, index=False)
    print(f"  Núcleo terminológico exportado: {NUCLEUS_FILE}")
    print("\n  Top 10 lemas más compartidos:")
    for _, row in df_nucleus.head(10).iterrows():
        print(f"    {row['Lema']:<35} {row['N glosarios']:>3} glosarios "
              f"({row['% glosarios']}%)")
    print("\nFinalizado.")
if __name__ == '__main__':
    main()
