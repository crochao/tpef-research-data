"""
TPEF — B.3 Verificación de geosinónimos en esTenTen23 vía API de Sketch Engine
Autor: C. O. Rocha Ochoa
Método: concordancias con muestreo (evita timeout en freqs)
Atributo de variedad: doc.tld
"""
import requests
import pandas as pd
import time
import os
from collections import defaultdict, Counter
USERNAME = os.getenv('SKETCHENGINE_USERNAME', '')
API_KEY = os.getenv('SKETCHENGINE_API_KEY', '')
BASE_URL = 'https://app.sketchengine.eu/bonito/run.cgi/'
CORPNAME = 'preloaded/estenten23_fl6'
AUTH = (USERNAME, API_KEY)
# Número de concordancias a muestrear por lema (aumentar = más preciso pero más lento)
SAMPLE_SIZE = 1000
# ============================================================
# GEOSYNONYM_PAIRS — Lista consolidada final (58 pares)
# Fuente: revisión de lemas exclusivos por país (VE-CVV, SV-BVES,
# AR-BCRA y los 18 países del corpus TPEF) + hallazgos aportados
# directamente sobre el corpus.
# ============================================================
GEOSYNONYM_PAIRS = [
    # ============================================================
    # SECCIÓN A — CLÚSTERES MULTIRREGIONALES (3 o más variantes)
    # ============================================================
    # [CLÚSTER 1] Extracto/estado de cuenta bancaria
    # ESP: extracto de cuenta | ARG: resumen de cuenta | LAT (resto): estado de cuenta
    ('extracto de cuenta', 'estado de cuenta', 'ESP', 'LAT'),
    ('resumen de cuenta', 'extracto de cuenta', 'ARG', 'ESP'),
    ('resumen de cuenta', 'estado de cuenta', 'ARG', 'LAT'),
    # [CLÚSTER 2] Cheque certificado/bancario (el más numeroso: 6 variantes)
    # ESP: cheque conformado, cheque bancario | ARG: cheque cancelatorio
    # MEX+HND+GTM: cheque de caja | COL+ECU+GTM+SLV: cheque de gerencia | CHL: vale vista
    ('cheque conformado', 'cheque de caja', 'ESP', 'MEX+HND'),
    ('cheque de gerencia', 'cheque de caja', 'COL', 'MEX+HND'),
    ('cheque cancelatorio', 'cheque de caja', 'ARG', 'MEX+HND+GTM'),
    ('cheque cancelatorio', 'cheque de gerencia', 'ARG', 'COL+ECU+GTM+SLV'),
    ('cheque cancelatorio', 'cheque bancario', 'ARG', 'ESP'),
    ('cheque cancelatorio', 'vale vista', 'ARG', 'CHL'),
    # [CLÚSTER 3] Domiciliación de pagos / débito directo
    # ARG: débito directo | ESP: domiciliación bancaria | ECU: débito automático
    ('débito directo', 'domiciliación bancaria', 'ARG', 'ESP'),
    ('débito directo', 'débito automático', 'ARG', 'ECU'),
    # [CLÚSTER 4] Agente bancario tercerizado (corresponsalía)
    # BOL: corresponsalía | COL+MEX: corresponsal bancario | COL+DOM: banco corresponsal
    ('corresponsalía', 'corresponsal bancario', 'BOL', 'COL+MEX'),
    ('corresponsalía', 'banco corresponsal', 'BOL', 'COL+DOM'),
    # [CLÚSTER 5] Registro/central de riesgo crediticio
    # CHL: boletín comercial | DOM+MEX: buró de crédito | COL: centrales de riesgo, datacrédito
    ('boletín comercial', 'buró de crédito', 'CHL', 'DOM+MEX'),
    ('boletín comercial', 'central de riesgo', 'CHL', 'COL'),
    ('boletín comercial', 'datacrédito', 'CHL', 'COL'),
    # [CLÚSTER 6] Indicador regulatorio de coste total del crédito
    # ESP: TAE | CHL: carga anual equivalente (CAE) | MEX: CAT
    ('carga anual equivalente', 'TAE', 'CHL', 'ESP'),
    ('carga anual equivalente', 'CAT', 'CHL', 'MEX'),
    ('CAT', 'TAE', 'MEX', 'ESP'),
    # [CLÚSTER 7] Ajustador/perito de seguros
    # BOL: liquidador de reclamos | ESP: perito, tasador de daños
    # MEX: ajustador de seguros | CHL: liquidador de siniestros
    ('liquidador de reclamos', 'perito', 'BOL', 'ESP'),
    ('liquidador de reclamos', 'tasador de daños', 'BOL', 'ESP'),
    ('liquidador de reclamos', 'ajustador de seguros', 'BOL', 'MEX'),
    ('liquidador de reclamos', 'liquidador de siniestros', 'BOL', 'CHL'),
    # [CLÚSTER 8] Tipo de cambio de referencia oficial
    # BOL: dólar oficial | CHL: dólar observado | COL+SLV: TRM
    ('dólar oficial', 'dólar observado', 'BOL', 'CHL'),
    ('dólar oficial', 'TRM', 'BOL', 'COL+SLV'),
    # [CLÚSTER 9] Cuenta de depósito de nómina/sueldo
    # ARG: cuenta sueldo | CHL: cuenta vista | ESP: cuenta nómina | MEX: cuenta de nómina, cuenta eje
    ('cuenta sueldo', 'cuenta vista', 'ARG', 'CHL'),
    ('cuenta sueldo', 'cuenta nómina', 'ARG', 'ESP'),
    ('cuenta sueldo', 'cuenta de nómina', 'ARG', 'MEX'),
    ('cuenta sueldo', 'cuenta eje', 'ARG', 'MEX'),
    # ============================================================
    # SECCIÓN B — PARES BINARIOS SIMPLES (por dominio semántico)
    # ============================================================
    # --- Macroeconomía y estadística ---
    ('producto interior bruto', 'producto interno bruto', 'ESP', 'LAT'),
    ('PIB', 'PBI', 'ESP+LAT', 'ARG+PER'),
    ('valor agregado', 'valor añadido', 'COL+LAT', 'ESP+DOM'),
    # --- Banca minorista y cuentas ---
    ('cuenta de cheques', 'cuenta corriente', 'MEX', 'RESTO'),
    ('cuenta a la vista', 'cuenta vista', 'ESP+ARG+BOL', 'CHL'),
    ('cuentacorrentista', 'cuentahabiente', 'ARG+CHL+ECU', 'HND+MEX'),
    ('sobregiro', 'descubierto', 'LAT', 'ESP'),
    ('dólar paralelo', 'dólar informal', 'BOL', 'CHL'),
    ('pasanaku', 'caja de ahorro informal', 'BOL', 'MEX'),
    # --- Crédito y financiación ---
    ('préstamo hipotecario', 'crédito hipotecario', 'ESP+ARG', 'LAT'),
    ('préstamo personal', 'crédito personal', 'ESP+ARG', 'LAT'),
    ('préstamo de consumo', 'crédito de consumo', 'ESP', 'LAT'),
    ('tipo de interés', 'tasa de interés', 'ESP', 'LAT'),
    ('coste', 'costo', 'ESP', 'LAT'),
    # --- Mercado de valores e inversión ---
    ('titulización', 'titularización', 'ESP', 'LAT'),
    ('fondo común de inversión', 'fondo de inversión', 'ARG', 'RESTO'),
    ('carteras colectivas', 'fondo de inversión', 'COL', 'BOL+CRI+DOM+ECU+ESP+MEX+PAN+PER+PRY+SLV+VEN'),
    ('comisionista de bolsa', 'corredor de bolsa', 'COL', 'SLV'),
    ('central de valores', 'caja de valores', 'NIC', 'ARG+PRY+VEN'),
    ('opción sobre futuros', 'opción de futuros', 'PRY', 'ESP+NIC'),
    ('operaciones de pase o reporte', 'operación de reporto', 'PRY', 'BOL+NIC'),
    ('bono a descuento', 'bono cupón cero', 'VEN', 'ARG+ESP+MEX+PRY'),
    ('palanqueo financiero', 'apalancamiento financiero', 'COL', 'ARG+BOL+CHL+ECU+ESP+MEX+PAN'),
    ('inversor', 'inversionista', 'ESP', 'LAT'),
    ('accionista', 'socio accionista', 'ESP', 'LAT'),
    # --- Seguros y previsión social ---
    ('seguro de cesantía', 'seguro de desempleo', 'CHL', 'ESP+MEX'),
    ('utilidad', 'ganancia', 'LAT', 'ESP'),
    # --- Estructura corporativa y jurídica ---
    ('filial', 'subsidiaria', 'ESP', 'LAT'),
    ('persona moral', 'persona jurídica', 'MEX', 'BOL+CHL+ESP'),
]
# ============================================================
# TLD_GROUPS — mapeo de cada etiqueta de variedad usada en
# GEOSYNONYM_PAIRS a los dominios de nivel superior (doc.tld)
# correspondientes en esTenTen23. Generado automáticamente a
# partir de las 36 etiquetas distintas que aparecen en la lista.
# ============================================================
TLD_GROUPS = {
    'ARG': ['ar'],
    'ARG+BOL+CHL+ECU+ESP+MEX+PAN': ['ar', 'bo', 'cl', 'ec', 'es', 'mx', 'pa'],
    'ARG+CHL+ECU': ['ar', 'cl', 'ec'],
    'ARG+ESP+MEX+PRY': ['ar', 'es', 'mx', 'py'],
    'ARG+PER': ['ar', 'pe'],
    'ARG+PRY+VEN': ['ar', 'py', 've'],
    'BOL': ['bo'],
    'BOL+CHL+ESP': ['bo', 'cl', 'es'],
    'BOL+CRI+DOM+ECU+ESP+MEX+PAN+PER+PRY+SLV+VEN': ['bo', 'cr', 'do', 'ec', 'es', 'mx', 'pa', 'pe', 'py', 'sv', 've'],
    'BOL+NIC': ['bo', 'ni'],
    'CHL': ['cl'],
    'COL': ['co'],
    'COL+DOM': ['co', 'do'],
    'COL+ECU+GTM+SLV': ['co', 'ec', 'gt', 'sv'],
    'COL+LAT': ['co', 'ar', 'bo', 'cl', 'cr', 'cu', 'do', 'ec', 'gt', 'hn', 'mx', 'ni', 'pa', 'pe', 'py', 'sv', 've', 'uy'],
    'COL+MEX': ['co', 'mx'],
    'COL+SLV': ['co', 'sv'],
    'DOM+MEX': ['do', 'mx'],
    'ECU': ['ec'],
    'ESP': ['es'],
    'ESP+ARG': ['es', 'ar'],
    'ESP+ARG+BOL': ['es', 'ar', 'bo'],
    'ESP+DOM': ['es', 'do'],
    'ESP+LAT': ['es', 'ar', 'bo', 'cl', 'co', 'cr', 'cu', 'do', 'ec', 'gt', 'hn', 'mx', 'ni', 'pa', 'pe', 'py', 'sv', 've', 'uy'],
    'ESP+MEX': ['es', 'mx'],
    'ESP+NIC': ['es', 'ni'],
    'HND+MEX': ['hn', 'mx'],
    'LAT': ['ar', 'bo', 'cl', 'co', 'cr', 'cu', 'do', 'ec', 'gt', 'hn', 'mx', 'ni', 'pa', 'pe', 'py', 'sv', 've', 'uy'],
    'MEX': ['mx'],
    'MEX+HND': ['mx', 'hn'],
    'MEX+HND+GTM': ['mx', 'hn', 'gt'],
    'NIC': ['ni'],
    'PRY': ['py'],
    'RESTO': ['es', 'ar', 'bo', 'cl', 'co', 'cr', 'cu', 'do', 'ec', 'gt', 'hn', 'mx', 'ni', 'pa', 'pe', 'py', 'sv', 've', 'uy'],
    'SLV': ['sv'],
    'VEN': ['ve'],
}
def api_get(endpoint, params, timeout=90):
    for attempt in range(3):
        try:
            r = requests.get(
                BASE_URL + endpoint,
                params={**params, 'format': 'json', 'corpname': CORPNAME},
                auth=AUTH, timeout=timeout
            )
            if r.status_code == 200:
                data = r.json()
                # Sketch Engine puede responder HTTP 200 con un cuerpo JSON
                # que contiene un campo 'error' cuando la consulta CQL es
                # inválida o hace referencia a un atributo inexistente. Sin
                # esta comprobación, ese error se confunde silenciosamente
                # con una búsqueda legítima de cero resultados.
                if isinstance(data, dict) and data.get('error'):
                    print(
                        f"  ⚠ Error de la API para q={params.get('q')!r}: "
                        f"{data['error']}"
                    )
                return data
            print(f"  Error {r.status_code}: {r.text[:100]}")
            time.sleep(3 * (attempt + 1))
        except requests.exceptions.Timeout:
            print(f"  Timeout (intento {attempt+1}/3)")
            time.sleep(3 * (attempt + 1))
        except Exception as e:
            print(f"  Excepción: {e}")
            time.sleep(3)
    return None
def build_cql(lemma, by_word=False):
    """
    Construye la consulta CQL para un lema, ya sea de una sola
    palabra o multipalabra. [lemma="..."] solo casa con un token,
    por lo que los términos multipalabra ('tipo de interés') deben
    expresarse como una secuencia de tokens concatenados:
    [lemma="tipo"][lemma="de"][lemma="interés"].
    by_word=True construye la consulta sobre el atributo posicional
    'lc' (forma de palabra en minúsculas) en vez de sobre el lema.
    Es necesaria como respaldo porque el lema de un sustantivo plural
    o un adjetivo en la entrada de la lista no coincide con el lema
    canónico que usa el corpus (p. ej. 'carteras' vs. lema real
    'cartera'; 'colectivas' vs. lema real 'colectivo'), lo que
    produce cero concordancias aunque la colocación exista.
    Nota: se probó primero con el modificador %c ([word="..." %c] e
    [word="..."%c]) para una búsqueda insensible a mayúsculas sobre
    'word', pero la API de esTenTen23 rechaza ese modificador en
    cualquier posición ("unexpected character %"). El atributo 'lc'
    evita el problema por completo, ya que compara directamente
    contra la forma ya normalizada a minúsculas sin necesitar ningún
    modificador adicional.
    """
    palabras = lemma.strip().split()
    if by_word:
        attr = 'lc="%s"'
        palabras = [p.lower() for p in palabras]
    else:
        attr = 'lemma="%s"'
    tokens = ''.join('[' + (attr % p) + ']' for p in palabras)
    return f'q{tokens}'
def get_tld_distribution(lemma, sample=SAMPLE_SIZE):
    """
    Obtiene la distribución de TLDs para un lema (uni o multipalabra)
    usando concordancias con muestreo aleatorio. Primero intenta la
    búsqueda por lema; si no devuelve resultados, reintenta por forma
    de palabra literal (ver build_cql).
    Devuelve (tld_counts, metodo), donde metodo es 'lema' o
    'forma_palabra' según cuál haya producido resultados.
    """
    conc = api_get('view', {
        'q': build_cql(lemma, by_word=False),
        'pagesize': sample,
        'refs': '=doc.tld',
        'attrs': 'word',
        'fromp': 1,
        'random': 1,  # muestreo aleatorio
    }, timeout=90)
    metodo = 'lema'
    if not conc or not conc.get('Lines'):
        conc = api_get('view', {
            'q': build_cql(lemma, by_word=True),
            'pagesize': sample,
            'refs': '=doc.tld',
            'attrs': 'word',
            'fromp': 1,
            'random': 1,
        }, timeout=90)
        metodo = 'forma_palabra'
    tld_counts = Counter()
    if not conc:
        return tld_counts, metodo
    lines = conc.get('Lines', [])
    for line in lines:
        refs = line.get('Refs', [])
        for ref in refs:
            val = str(ref).strip().lower().strip('.')
            if val:
                tld_counts[val] += 1
    return tld_counts, metodo
def group_freq(dist, tlds):
    return sum(dist.get(t, 0) for t in tlds)
def asymmetry_index(fa_A, fa_B, fb_A, fb_B):
    tot_A = fa_A + fb_A
    tot_B = fa_B + fb_B
    if tot_A == 0 or tot_B == 0:
        return None
    return round(abs(fa_A / tot_A - fa_B / tot_B), 3)
def classify(idx):
    if idx is None:
        return 'Indeterminado'
    if idx >= 0.40:
        return 'GC ✓'
    if idx >= 0.20:
        return 'GP ~'
    return 'VnD ✗'
def main():
    if not USERNAME or not API_KEY:
        raise RuntimeError(
            'Defina SKETCHENGINE_USERNAME y SKETCHENGINE_API_KEY antes de ejecutar.'
        )
    print("TPEF — Verificación de geosinonimia en esTenTen23")
    print(f"Método: concordancias con muestreo (n={SAMPLE_SIZE})")
    print(f"Total de pares a verificar: {len(GEOSYNONYM_PAIRS)}")
    print("=" * 65)
    # Verificar conexión
    print("\nVerificando conexión...")
    info = api_get('corp_info', {'corpname': CORPNAME})
    if info:
        print(f" ✓ Corpus: {info.get('name', '?')}")
    else:
        print(" ✗ No se pudo conectar.")
        return
    # Verificar que refs=doc.tld funciona
    print("\nVerificando doc.tld con 'banco' (palabra frecuente)...")
    test, _ = get_tld_distribution('banco', sample=100)
    if test:
        print(f" ✓ TLDs encontrados: {dict(test.most_common(5))}")
    else:
        print(" ✗ Sin resultados. Verificar atributo doc.tld.")
        return
    results = []
    for den_a, den_b, var_a_lbl, var_b_lbl in GEOSYNONYM_PAIRS:
        print(f"\n── '{den_a}' / '{den_b}' ──")
        tlds_a = TLD_GROUPS.get(var_a_lbl, [var_a_lbl.lower()])
        tlds_b = TLD_GROUPS.get(var_b_lbl, [var_b_lbl.lower()])
        dist_a, metodo_a = get_tld_distribution(den_a)
        time.sleep(1)
        dist_b, metodo_b = get_tld_distribution(den_b)
        time.sleep(1)
        total_a = sum(dist_a.values())
        total_b = sum(dist_b.values())
        print(f"  '{den_a}' ({metodo_a}): {total_a} concordancias | TLDs: {dict(dist_a.most_common(5))}")
        print(f"  '{den_b}' ({metodo_b}): {total_b} concordancias | TLDs: {dict(dist_b.most_common(5))}")
        fa_A = group_freq(dist_a, tlds_a)
        fa_B = group_freq(dist_a, tlds_b)
        fb_A = group_freq(dist_b, tlds_a)
        fb_B = group_freq(dist_b, tlds_b)
        idx = asymmetry_index(fa_A, fa_B, fb_A, fb_B)
        cls = classify(idx)
        print(f"  '{den_a}' → {var_a_lbl}: {fa_A} | {var_b_lbl}: {fa_B}")
        print(f"  '{den_b}' → {var_a_lbl}: {fb_A} | {var_b_lbl}: {fb_B}")
        print(f"  Índice asimetría: {idx} → {cls}")
        results.append({
            'Denominación A': den_a,
            'Denominación B': den_b,
            'Variedad A': var_a_lbl,
            'Variedad B': var_b_lbl,
            'Método A': metodo_a,
            'Método B': metodo_b,
            'Total conc. A': total_a,
            'Total conc. B': total_b,
            'Frec. A en Var.A': fa_A,
            'Frec. A en Var.B': fa_B,
            'Frec. B en Var.A': fb_A,
            'Frec. B en Var.B': fb_B,
            'Índice asimetría': idx,
            'Clasificación': cls,
        })
    df = pd.DataFrame(results)
    df.to_excel('tpef_geosinonimia_resultados.xlsx', index=False)
    print(f"\n{'='*65}")
    print("Resultados en: tpef_geosinonimia_resultados.xlsx")
    print("\nResumen:")
    print(df.groupby('Clasificación').size().to_string())
    print("\nTabla completa:")
    cols = ['Denominación A', 'Denominación B', 'Índice asimetría', 'Clasificación']
    print(df[cols].to_string(index=False))
if __name__ == '__main__':
    main()
