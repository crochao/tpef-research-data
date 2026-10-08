"""
TPEF — Estimación del inventario conceptual por clustering TF-IDF
=================================================================
Autor: C. O. Rocha Ochoa
Descripción:
    Estima el número de conceptos distintos del corpus mediante
    agrupamiento por similitud definicional. El procedimiento combina
    vectorización TF-IDF de las definiciones con cálculo de similitud
    del coseno y agrupamiento Union-Find a dos umbrales (0,75 y 0,60),
    produciendo un rango de estimación del inventario conceptual.
    El algoritmo opera en cuatro pasos:
      1. Selección de definiciones utilizables (≥ 20 caracteres)
      2. Vectorización TF-IDF con unigramas y bigramas
      3. Cálculo de similitud del coseno por lotes
      4. Agrupamiento Union-Find a dos umbrales
    Los grupos resultantes son candidatos a equivalencia conceptual
    y requieren validación terminológica manual para confirmar qué
    entradas corresponden al mismo concepto.
Entrada:  Glosarios_2026.xlsx
Salida:   tpef_conceptos_grupos_075.xlsx   (umbral conservador)
          tpef_conceptos_grupos_060.xlsx   (umbral amplio)
          tpef_conceptos_resumen.xlsx      (estadísticas globales)
Requisitos:
    pip install openpyxl pandas numpy scikit-learn
"""
import unicodedata
from collections import defaultdict
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
# ── Configuración ─────────────────────────────────────────────────────────────
INPUT_FILE      = 'Glosarios_2026.xlsx'
OUTPUT_075      = 'tpef_conceptos_grupos_075.xlsx'
OUTPUT_060      = 'tpef_conceptos_grupos_060.xlsx'
OUTPUT_SUMMARY  = 'tpef_conceptos_resumen.xlsx'
# Umbrales de similitud del coseno
THRESHOLD_CONSERVATIVE = 0.75   # umbral conservador
THRESHOLD_BROAD        = 0.60   # umbral amplio
# Longitud mínima de definición utilizable (en caracteres)
MIN_DEFN_LENGTH = 20
# Tamaño del lote para el cálculo de similitud (gestión de memoria)
BATCH_SIZE = 500
# ── Paso 1: Normalización ─────────────────────────────────────────────────────
def normalize(s):
    """Convierte a minúsculas y elimina diacríticos."""
    if not s:
        return ''
    s = str(s).strip().lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s
# ── Paso 2: Carga y selección de definiciones ─────────────────────────────────
def load_definitions(filepath):
    """
    Carga el corpus y construye un diccionario {lema_normalizado: definicion}
    conservando, para cada lema único, la definición más extensa disponible
    entre todos los glosarios que lo incluyen.
    """
    wb = load_workbook(filepath, read_only=True)
    lemma_defn = {}       # {lema_norm: definicion_original}
    lemma_sources = defaultdict(list)  # {lema_norm: [glosarios]}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            if not (row[0] and str(row[0]).strip()):
                continue
            lema      = str(row[0]).strip()
            lema_norm = normalize(lema)
            defn      = str(row[1]).strip() if row[1] else ''
            # Conservar la definición más extensa
            if len(defn) > len(lemma_defn.get(lema_norm, '')):
                lemma_defn[lema_norm] = defn
            lemma_sources[lema_norm].append(sheet_name)
    # Filtrar: solo definiciones utilizables
    lemma_defn_filtered = {
        k: v for k, v in lemma_defn.items()
        if len(v) >= MIN_DEFN_LENGTH
    }
    print(f"  Lemas únicos en el corpus:          {len(lemma_defn):,}")
    print(f"  Lemas con definición utilizable:    {len(lemma_defn_filtered):,}")
    print(f"  Lemas sin definición o muy corta:   "
          f"{len(lemma_defn) - len(lemma_defn_filtered):,}")
    return lemma_defn_filtered, lemma_sources
# ── Paso 3: Vectorización TF-IDF ─────────────────────────────────────────────
def vectorize(definitions):
    """
    Transforma las definiciones en vectores TF-IDF.
    Parámetros del vectorizador:
      - ngram_range=(1,2): unigramas y bigramas
      - sublinear_tf=True: escala TF logarítmica (1 + log(tf))
      - min_df=1:          incluir todos los términos
      - token_pattern:     secuencias alfabéticas ≥ 3 caracteres con diacríticos
    """
    vectorizer = TfidfVectorizer(
        ngram_range   = (1, 2),
        sublinear_tf  = True,
        min_df        = 1,
        strip_accents = None,
        analyzer      = 'word',
        token_pattern = r'(?u)\b[a-záéíóúüñ]{3,}\b',
    )
    matrix = vectorizer.fit_transform(definitions)
    print(f"  Matriz TF-IDF: {matrix.shape[0]:,} definiciones × "
          f"{matrix.shape[1]:,} características")
    return matrix
# ── Paso 4: Union-Find ────────────────────────────────────────────────────────
class UnionFind:
    """
    Estructura Union-Find con compresión de caminos.
    Agrupa los índices de lemas cuya similitud supera el umbral.
    """
    def __init__(self, n):
        self.parent = list(range(n))
    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x
    def union(self, x, y):
        px, py = self.find(x), self.find(y)
        if px != py:
            self.parent[px] = py
    def clusters(self, n):
        groups = defaultdict(list)
        for i in range(n):
            groups[self.find(i)].append(i)
        return list(groups.values())
# ── Paso 5: Clustering por similitud del coseno ───────────────────────────────
def cluster_by_similarity(matrix, lemmas, threshold, batch_size=BATCH_SIZE):
    """
    Agrupa los lemas cuya similitud del coseno supera el umbral
    mediante Union-Find. El cálculo se realiza por lotes para
    gestionar la memoria sobre corpus grandes.
    Devuelve la lista de grupos (cada grupo es una lista de lemas).
    """
    n  = matrix.shape[0]
    uf = UnionFind(n)
    n_pairs_merged = 0
    for start in range(0, n, batch_size):
        end        = min(start + batch_size, n)
        batch      = matrix[start:end]
        sim_matrix = cosine_similarity(batch, matrix)
        for i, row in enumerate(sim_matrix):
            global_i = start + i
            similar  = np.where(row >= threshold)[0]
            for j in similar:
                if j != global_i:
                    uf.union(global_i, j)
                    n_pairs_merged += 1
        if (start // batch_size) % 5 == 0:
            print(f"    Procesado lote {start//batch_size + 1} / "
                  f"{(n + batch_size - 1) // batch_size}")
    raw_clusters = uf.clusters(n)
    groups = [[lemmas[i] for i in cluster] for cluster in raw_clusters]
    multi  = [g for g in groups if len(g) > 1]
    print(f"  Pares fusionados:           {n_pairs_merged:,}")
    print(f"  Grupos totales:             {len(groups):,}")
    print(f"  Grupos con >1 lema:         {len(multi):,}")
    print(f"  Estimación de conceptos:    {len(groups):,}")
    return groups, multi
# ── Exportar grupos ───────────────────────────────────────────────────────────
def export_groups(groups, multi, lemma_defn, lemma_sources, output_file,
                  threshold):
    """
    Exporta los grupos de candidatos a equivalencia conceptual.
    Cada fila representa un lema con su grupo asignado.
    """
    rows = []
    group_id = 1
    # Primero los grupos con más de un miembro (candidatos a equivalencia)
    for group in sorted(multi, key=len, reverse=True):
        for lema in sorted(group):
            rows.append({
                'Grupo ID':       group_id,
                'Tamaño grupo':   len(group),
                'Tipo':           'Candidato a equivalencia',
                'Lema':           lema,
                'Definición':     lemma_defn.get(lema, '')[:300],
                'Glosarios':      ', '.join(lemma_sources.get(lema, [])),
                'N glosarios':    len(lemma_sources.get(lema, [])),
            })
        group_id += 1
    # Luego los singletons (lemas sin equivalente)
    singletons = [g for g in groups if len(g) == 1]
    for group in singletons:
        lema = group[0]
        rows.append({
            'Grupo ID':       group_id,
            'Tamaño grupo':   1,
            'Tipo':           'Singleton',
            'Lema':           lema,
            'Definición':     lemma_defn.get(lema, '')[:300],
            'Glosarios':      ', '.join(lemma_sources.get(lema, [])),
            'N glosarios':    len(lemma_sources.get(lema, [])),
        })
        group_id += 1
    df = pd.DataFrame(rows)
    df.to_excel(output_file, index=False)
    print(f"  Exportado: {output_file} ({len(rows):,} filas)")
    return df
# ── Script principal ──────────────────────────────────────────────────────────
def main():
    print("TPEF — Estimación del inventario conceptual por clustering TF-IDF")
    print("="*65)
    # Paso 1: Carga
    print("\n[1] Cargando corpus y seleccionando definiciones...")
    lemma_defn, lemma_sources = load_definitions(INPUT_FILE)
    lemmas      = list(lemma_defn.keys())
    definitions = [lemma_defn[l] for l in lemmas]
    n_lemmas    = len(lemmas)
    # Paso 2: Vectorización
    print("\n[2] Vectorizando definiciones con TF-IDF...")
    matrix = vectorize(definitions)
    # Paso 3: Clustering conservador (umbral 0,75)
    print(f"\n[3] Clustering con umbral conservador ({THRESHOLD_CONSERVATIVE})...")
    groups_075, multi_075 = cluster_by_similarity(
        matrix, lemmas, THRESHOLD_CONSERVATIVE
    )
    # Paso 4: Clustering amplio (umbral 0,60)
    print(f"\n[4] Clustering con umbral amplio ({THRESHOLD_BROAD})...")
    groups_060, multi_060 = cluster_by_similarity(
        matrix, lemmas, THRESHOLD_BROAD
    )
    # Exportar grupos
    print("\n[5] Exportando resultados...")
    df_075 = export_groups(
        groups_075, multi_075, lemma_defn, lemma_sources,
        OUTPUT_075, THRESHOLD_CONSERVATIVE
    )
    df_060 = export_groups(
        groups_060, multi_060, lemma_defn, lemma_sources,
        OUTPUT_060, THRESHOLD_BROAD
    )
    # Resumen comparativo
    summary_rows = [
        {
            'Umbral':                    THRESHOLD_CONSERVATIVE,
            'Descripción':               'Conservador',
            'Lemas con definición':      n_lemmas,
            'Grupos (conceptos est.)':   len(groups_075),
            'Grupos multi-lema':         len(multi_075),
            'Reducción vs. lemas (%)':   round(
                (1 - len(groups_075) / n_lemmas) * 100, 1),
        },
        {
            'Umbral':                    THRESHOLD_BROAD,
            'Descripción':               'Amplio',
            'Lemas con definición':      n_lemmas,
            'Grupos (conceptos est.)':   len(groups_060),
            'Grupos multi-lema':         len(multi_060),
            'Reducción vs. lemas (%)':   round(
                (1 - len(groups_060) / n_lemmas) * 100, 1),
        },
    ]
    df_summary = pd.DataFrame(summary_rows)
    df_summary.to_excel(OUTPUT_SUMMARY, index=False)
    print(f"  Resumen: {OUTPUT_SUMMARY}")
    # Estadísticas finales
    print(f"\n{'='*65}")
    print(f"RESUMEN DE ESTIMACIÓN")
    print(f"{'='*65}")
    print(f"  Lemas únicos en el corpus:             {n_lemmas:,}")
    print(f"  Umbral 0,75 → {len(groups_075):,} conceptos estimados "
          f"(reducción: {(1-len(groups_075)/n_lemmas)*100:.1f}%)")
    print(f"  Umbral 0,60 → {len(groups_060):,} conceptos estimados "
          f"(reducción: {(1-len(groups_060)/n_lemmas)*100:.1f}%)")
    print(f"  Rango estimado: {len(groups_060):,} – {len(groups_075):,} conceptos")
    print(f"\nNota: Los grupos exportados son candidatos a equivalencia")
    print(f"conceptual y requieren validación terminológica manual.")
    print(f"\nFinalizado.")
if __name__ == '__main__':
    main()
