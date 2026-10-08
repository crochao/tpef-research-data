"""
TPEF — Clasificador refinado de subdominios temáticos

Mejoras frente a la versión inicial
1. Coincidencia por palabras o expresiones completas, no por subcadenas.
2. Puntuación acumulativa, sin depender del orden de los subdominios.
3. Mayor peso del lema, de las expresiones poliléxicas y de los términos fuertes.
4. Reglas de desambiguación para voces especialmente polisémicas.
5. Salida auditable con puntuación, margen y palabras clave activadas.
"""

import re
import unicodedata
import argparse
from collections import defaultdict
from functools import lru_cache
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


DEFAULT_INPUT_FILE = Path("Glosarios_2026.xlsx")
DEFAULT_OUTPUT_FILE = Path("tpef_subdominios_clasificados_refinado.xlsx")
DEFAULT_SUMMARY_FILE = Path("tpef_subdominios_resumen_refinado.xlsx")

COUNTRY_LABELS = {
    "AR": "Argentina", "BO": "Bolivia", "CL": "Chile",
    "CO": "Colombia", "CR": "Costa Rica", "CU": "Cuba",
    "DO": "Rep. Dominicana", "EC": "Ecuador", "ES": "España",
    "GT": "Guatemala", "HN": "Honduras", "MX": "México",
    "NI": "Nicaragua", "PA": "Panamá", "PE": "Perú",
    "PY": "Paraguay", "SV": "El Salvador", "VE": "Venezuela",
}


def normalize(value):
    """Normaliza mayúsculas, diacríticos, guiones y espacios."""
    if value is None:
        return ""
    text = unicodedata.normalize("NFD", str(value).strip().lower())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[-‐‑‒–—_/]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


@lru_cache(maxsize=None)
def normalize_keyword(keyword):
    return normalize(keyword)


def pluralizable_token_pattern(token):
    """Admite singular y plural regular sin volver a subcadenas libres."""
    escaped = re.escape(token)
    if token.endswith("ion") and len(token) > 3:
        return re.escape(token[:-2]) + r"(?:on|ones)"
    if token.endswith("z") and len(token) > 1:
        return re.escape(token[:-1]) + r"(?:z|ces)"
    if token.endswith(("a", "e", "i", "o", "u")):
        return escaped + r"s?"
    return escaped + r"(?:es)?"


@lru_cache(maxsize=None)
def keyword_pattern(keyword):
    tokens = normalize_keyword(keyword).split()
    body = r"\s+".join(pluralizable_token_pattern(token) for token in tokens)
    return re.compile(rf"(?<![a-z0-9]){body}(?![a-z0-9])")


def contains(text, keyword):
    """Comprueba unidades completas y plurales regulares."""
    return bool(keyword_pattern(keyword).search(text))


# strong contiene términos de alta capacidad discriminante.
# support contiene términos que requieren acumulación o apoyo contextual.
LEXICON = {
    "SisFin": {
        "strong": [
            "sistema financiero", "sistema bancario", "banco central",
            "entidad financiera", "institucion financiera",
            "intermediario financiero", "superintendencia financiera",
            "superintendencia de bancos", "autoridad supervisora",
            "organismo regulador", "regulador financiero", "supervision bancaria",
            "regulacion prudencial", "encaje bancario", "encaje legal",
            "reservas internacionales", "coeficiente de solvencia",
            "adecuacion de capital", "basilea", "bancarizacion",
            "inclusion financiera", "microfinanzas", "sucursal bancaria",
            "agencia calificadora", "central de riesgos",
        ],
        "support": [
            "banco", "banca", "sucursal", "cooperativa", "regulacion",
            "supervision", "solvencia", "liquidez bancaria", "intermediacion",
            "institucion", "entidad", "organismo", "autoridad", "fintech",
        ],
    },
    "CrePre": {
        "strong": [
            "credito", "prestamo", "credito hipotecario", "prestamo hipotecario",
            "microcredito", "linea de credito", "tarjeta de credito",
            "tasa de interes", "tipo de interes", "tasa efectiva anual",
            "tasa nominal anual", "costo financiero total", "coste del credito",
            "costo del credito", "capacidad de pago", "historial crediticio",
            "buro de credito", "riesgo de credito", "refinanciacion",
            "refinanciamiento", "reestructuracion de deuda", "sobreendeudamiento",
            "amortizacion", "cuota de prestamo", "capital prestado",
            "leasing", "arrendamiento financiero", "factoring",
            "credito prendario", "credito de consumo", "credito comercial",
        ],
        "support": [
            "deuda", "acreedor", "deudor", "garantia", "aval", "fianza",
            "prenda", "hipoteca", "interes", "cuota", "mora", "morosidad",
            "insolvencia", "quita", "incobrable", "pagare",
            "arrendador", "arrendatario", "endosar", "vencimiento",
            "financiamiento", "endeudamiento",
        ],
    },
    "AhoDep": {
        "strong": [
            "ahorro", "deposito bancario", "deposito a plazo",
            "deposito a termino", "deposito a la vista", "plazo fijo",
            "cuenta de ahorro", "caja de ahorro", "cuenta corriente",
            "cuenta bancaria", "cuenta vista", "cuenta remunerada",
            "cuenta de cheques", "cuenta sueldo", "cuenta nomina",
            "certificado de deposito", "fondo de garantia de depositos",
            "seguro de depositos", "libreta de ahorro", "estado de cuenta",
            "extracto bancario", "saldo disponible",
        ],
        "support": [
            "deposito", "ahorrador", "depositante", "libreta", "cdt",
            "cuenta", "saldo de la cuenta", "entidad depositaria",
        ],
    },
    "InvMer": {
        "strong": [
            "mercado de valores", "mercado de capitales", "mercado bursatil",
            "bolsa de valores", "bolsa de comercio", "mercado primario",
            "mercado secundario", "fondo de inversion", "fondo mutuo",
            "fondo cotizado", "cartera de inversion", "portafolio de inversion",
            "renta fija", "renta variable", "instrumento financiero",
            "activo financiero", "titulo valor", "titulo de deuda",
            "valor mobiliario", "valores negociables", "oferta publica",
            "oferta publica de adquisicion",
            "oferta publica de venta", "indice bursatil", "dow jones",
            "capitalizacion bursatil", "ganancia de capital", "plusvalia",
            "corredor de bolsa", "comisionista de bolsa", "agente de bolsa",
            "casa de bolsa", "corretaje", "negociacion bursatil",
            "cotizacion", "volatilidad", "trading", "precio spot",
            "mercado de futuros", "venta en corto", "papel comercial",
            "private equity", "venture capital", "capital de riesgo",
            "inversion", "inversor", "inversionista", "accion", "bono",
            "prima de opcion", "prima de emision",
        ],
        "support": [
            "dividendo", "rentabilidad", "rendimiento", "cartera", "portafolio",
            "etf", "derivado", "opcion", "futuro", "swap", "warrant",
            "emisor", "emision", "subasta", "negociacion", "broker",
            "underwriter", "commodity", "especulacion", "duracion",
            "vencimiento", "letra", "titulo", "valor",
        ],
    },
    "SegPre": {
        "strong": [
            "seguro", "contrato de seguro", "seguro de vida", "seguro de salud",
            "seguro de desempleo", "seguro de desgravamen", "poliza de seguro",
            "prima de seguro", "compania aseguradora", "fondo de pensiones",
            "plan de pensiones", "sistema previsional", "prevision social",
            "seguridad social", "ahorro previsional", "jubilacion",
            "pension de vejez", "pension de invalidez", "renta vitalicia",
            "administradora de fondos de pensiones", "reaseguro",
        ],
        "support": [
            "aseguradora", "asegurador", "asegurado", "tomador", "poliza",
            "siniestro", "cobertura", "beneficiario", "pension", "afp",
            "afore", "prevision", "actuarial", "jubilado", "retiro",
            "cesantia", "indemnizacion", "desgravamen",
        ],
    },
    "ConEco": {
        "strong": [
            "producto interno bruto", "producto interior bruto",
            "producto bruto interno", "balanza de pagos", "balanza comercial",
            "politica monetaria", "politica fiscal", "tipo de cambio",
            "crecimiento economico", "ciclo economico", "actividad economica",
            "indicador macroeconomico", "costo de oportunidad", "coste de oportunidad",
            "costo marginal", "coste marginal", "costos fijos", "costos variables",
            "economia de mercado", "competencia perfecta", "competencia imperfecta",
            "poder adquisitivo", "oferta monetaria", "masa monetaria",
            "reservas internacionales", "deuda publica", "deficit fiscal",
            "superavit fiscal", "planificacion financiera", "finanzas personales",
            "presupuesto personal", "ingreso personal", "educacion financiera",
            "estado financiero", "estados financieros", "estado contable",
            "estados contables", "balance general", "patrimonio neto",
            "capital contable", "ejercicio contable", "resultado contable",
            "analisis financiero", "indicador financiero",
        ],
        "support": [
            "pib", "pbi", "pnb", "inflacion", "deflacion", "devaluacion",
            "revaluacion", "fluctuacion", "paridad", "divisa", "moneda", "euro",
            "exportacion", "importacion", "deficit", "superavit", "recesion",
            "productividad", "macroeconomia", "microeconomia", "economia",
            "mercado", "oferta", "demanda", "precio", "costo", "coste",
            "salario", "nomina", "sueldo", "remuneracion", "desempleo", "empleo",
            "consumo", "produccion", "empresa", "capital", "trabajo", "tierra",
            "impuesto", "tributo", "iva", "renta", "ingreso", "gasto",
            "presupuesto", "depreciacion", "pyme", "mipyme", "monopolio",
            "oligopolio", "dumping", "lucro", "elasticidad", "escasez",
            "eficiencia", "competencia", "utilidad", "ganancia", "perdida",
        ],
    },
    "MedDig": {
        "strong": [
            "medio de pago", "sistema de pagos", "instrumento de pago",
            "tarjeta de credito", "tarjeta de debito", "tarjeta prepago",
            "transferencia bancaria", "transferencia electronica",
            "pago electronico", "pago digital", "pago en linea",
            "dinero electronico", "dinero digital", "moneda digital",
            "billetera digital", "billetera electronica", "cajero automatico",
            "terminal punto de venta", "debito automatico", "orden de pago",
            "domiciliacion bancaria", "banca movil", "banca electronica",
            "home banking", "e banking", "firma electronica", "firma digital",
            "autenticacion multifactor", "ciberseguridad", "seguridad informatica",
            "fraude cibernetico", "fraude electronico", "phishing", "smishing",
            "ransomware", "criptomoneda", "criptoactivo", "blockchain",
        ],
        "support": [
            "tarjeta", "cheque", "transferencia", "pago", "debito", "cajero",
            "efectivo", "billetera", "wallet", "bitcoin", "token", "nfc", "pos",
            "giro", "remesa", "compensacion", "cbu", "cvv", "contracargo",
            "chargeback", "debin", "autenticacion", "ciberataque", "malware",
            "encriptacion", "firewall", "firma", "transaccion electronica",
        ],
    },
    "ConJur": {
        "strong": [
            "contrato", "clausula contractual", "persona juridica",
            "persona fisica", "persona natural", "marco juridico",
            "acto juridico", "instrumento legal", "escritura publica",
            "poder notarial", "sociedad anonima", "sociedad limitada",
            "sociedad comercial", "proteccion al consumidor",
            "derechos del consumidor", "consumidor financiero",
            "usuario financiero", "defensor del cliente", "lavado de dinero",
            "lavado de activos", "blanqueo de capitales", "fideicomiso",
            "administrador concursal", "proceso concursal", "debida diligencia",
            "beneficiario final", "secreto bancario", "cumplimiento normativo",
        ],
        "support": [
            "obligacion", "rescision", "incumplimiento", "notaria", "patrimonio",
            "herencia", "mandato", "domicilio", "clausula", "sancion", "multa",
            "normativa", "reglamento", "ley", "decreto", "resolucion", "arbitraje",
            "quiebra", "disolucion", "sociedad", "acta", "escritura", "trust",
            "quorum", "asamblea", "directorio", "accionista", "fusion",
            "fiduciario", "sindico", "consumidor", "contratante",
        ],
    },
}


# Reglas para términos cuyo dominio depende del contexto.
CONTEXT_RULES = [
    # término, dominio, contexto que debe acompañarlo
    ("pagare", "InvMer", ["mercado", "valor", "titulo", "emision", "negociacion"]),
    ("pagare", "CrePre", ["credito", "prestamo", "deuda", "pago", "acreedor"]),
    ("emision", "InvMer", ["accion", "bono", "titulo", "valor", "oferta publica"]),
    ("emision", "ConEco", ["moneda", "dinero", "banco central"]),
    ("subasta", "InvMer", ["valor", "bono", "accion", "titulo", "mercado"]),
    ("subasta", "ConJur", ["judicial", "remate", "embargo"]),
    ("vencimiento", "CrePre", ["credito", "prestamo", "cuota", "deuda"]),
    ("vencimiento", "InvMer", ["bono", "titulo", "valor", "opcion", "futuro"]),
    ("duracion", "InvMer", ["bono", "cartera", "renta fija", "sensibilidad"]),
    ("firma", "MedDig", ["electronica", "digital", "autenticacion"]),
    ("firma", "ConJur", ["contrato", "documento", "escritura"]),
    ("liquidacion", "MedDig", ["pago", "compensacion", "transferencia", "operacion"]),
    ("liquidacion", "ConJur", ["sociedad", "quiebra", "concurso", "disolucion"]),
    ("fiduciario", "ConJur", ["fideicomiso", "contrato", "patrimonio"]),
    ("fideicomiso", "InvMer", ["inversion", "valor", "titularizacion"]),
    ("nomina", "AhoDep", ["cuenta", "deposito", "banco"]),
    ("nomina", "ConEco", ["salario", "sueldo", "remuneracion", "empleo"]),
    ("prima", "InvMer", ["opcion", "accion", "bono", "futuro"]),
    ("prima", "SegPre", ["seguro", "poliza", "siniestro", "asegurador", "asegurado"]),
    ("tarifa", "SegPre", ["seguro", "prima", "riesgo", "asegurador"]),
    ("tarifa", "ConEco", ["precio", "servicio", "impuesto", "transporte"]),
]


# Lemas breves y generales que pueden clasificarse cuando constituyen la
# entrada completa. Funcionan como respaldo y nunca vencen a una evidencia
# temática fuerte encontrada en el lema o en la definición.
EXACT_LEMMA_FALLBACK = {
    "ConEco": [
        "bien", "liquidez", "activo", "pasivo", "gravamen",
        "activo fijo", "balance", "base imponible", "hiperinflacion",
        "ano fiscal", "sector publico", "activo circulante",
        "activo corriente", "activo intangible", "apalancamiento financiero",
        "beneficio", "bien intermedio", "capitalismo", "fondo de maniobra",
        "industria", "necesidad", "sector privado", "analisis economico",
        "comercio", "comercio exterior", "deflactor", "equilibrio",
        "estado de resultados", "flujo de caja", "terminos de intercambio",
        "arancel", "dinero", "indice", "apalancamiento", "auditoria",
        "apreciacion", "cash flow", "cliente", "franquicia", "ratio",
        "subsidio", "ano base", "ano comercial", "balance consolidado",
        "base liquidable", "bien inferior", "bien libre", "bien no transable",
        "bien normal", "bien publico", "bien transable", "cambio climatico",
        "comercializacion", "competitividad", "contribucion", "contribuyente",
        "convertibilidad", "debe", "duopolio", "eficacia", "emprendimiento",
        "estado", "estanflacion", "evasion fiscal", "exencion fiscal",
        "indicadores de coyuntura", "infraestructura", "innovacion", "inventario",
        "margen neto", "mercosur", "monopsonio", "nacionalizacion", "nafta",
        "nivel de vida", "privatizacion", "provision", "recursos naturales",
        "riqueza", "sector informal", "sector primario", "sector secundario",
        "sector terciario", "tipo impositivo", "zona de libre comercio",
        "zona franca", "empresario", "recursos propios", "auditoria interna",
        "hecho imponible", "iliquidez", "intercambio", "memoria",
        "recursos ajenos", "recursos humanos", "retencion",
        "roa return on assets", "roe", "roe return on equity", "utilities",
        "activo neto", "activo eventual", "activo liquido", "aduana",
        "agente economico", "agio", "aguinaldo", "alicuota",
        "analisis cuantitativo", "ano tributario", "auditor", "auditoria externa",
        "autofinanciacion", "barreras a la entrada", "beneficios no distribuidos",
        "bien complementario", "bien sustituto", "break even",
        "cadena de suministro", "cartel", "competidor", "compra",
        "conglomerado", "consumismo", "contabilidad", "criterio de caja",
        "coeficiente de gini", "curva de lorenz", "declaracion tributaria",
        "desarrollo sostenible", "desempleado", "deseo", "desgravacion",
        "desregulacion", "devengado", "dinero en circulacion",
        "dinero mercancia", "doble imposicion", "dolar", "dolarizacion",
        "econometria", "empleador", "emprendedor", "equidad",
        "instituto nacional de estadistica y geografia inegi",
    ],
    "ConJur": [
        "embargo", "avaluo", "usufructo", "bien inmueble", "bien mueble",
        "condonacion", "consejo de administracion", "escision",
        "gobierno corporativo", "holding", "patente", "prorroga", "tasacion",
        "fraude", "exencion", "bancarrota", "catastro", "circular", "cohecho",
        "comunidad de bienes", "consorcio", "donacion", "licencia", "licitacion",
        "novacion", "nuda propiedad", "paraiso fiscal", "razon social",
        "riesgo legal", "subsidiaria", "estafa", "informacion reservada",
        "operacion acordeon", "queja", "absorcion", "adjudicacion",
        "boletin oficial del estado boe", "buena fe", "caso fortuito",
        "codigo de etica", "confidencialidad", "contrabando", "dano emergente",
        "declaracion complementaria", "declaracion jurada",
        "defensoria legal gratuita", "desfalco",
    ],
    "InvMer": [
        "spread", "capitalizacion", "al portador", "nominativo",
        "bajo la par", "a la par", "benchmark", "bull market", "bear market",
        "call", "dealer", "custodia", "gestion pasiva", "nasdaq",
        "negociable", "obligacionista", "orden limitada", "orden por lo mejor",
        "repo", "accionariado", "ask", "bid", "calificacion de riesgo",
        "comision de suscripcion", "ibex 35", "informacion privilegiada",
        "market makers", "minusvalia", "operacion cruzada", "operaciones a plazo",
        "overnight", "pay out", "permuta financiera", "reporto", "spot",
        "desmaterializacion", "a la orden", "adr", "analisis fundamental",
        "bursatil", "caballero blanco", "caballero negro",
        "cesion temporal de activos", "chartismo", "comision sobre resultados",
        "commodities", "corro", "criterios asg", "custodio", "deceval",
        "desmaterializado", "diversificacion", "endosatario", "endoso",
    ],
    "CrePre": [
        "avalado", "colateral", "afianzar", "carga financiera", "fiador",
        "financiacion", "pignoracion",
    ],
    "SisFin": [
        "riesgo de liquidez", "riesgo operativo", "riesgo operacional",
        "riesgo sistemico", "encaje", "fmi fondo monetario internacional",
        "fondo monetario internacional fmi", "g20",
        "organizacion para la cooperacion y el desarrollo economicos ocde",
        "comite de auditoria", "intermediario", "oficina de representacion",
        "autorregulacion", "bancarizar", "casa de cambio", "control interno",
    ],
    "SegPre": [
        "deducible", "accidente", "mutualidad", "coaseguro",
    ],
    "MedDig": [
        "billete alterado", "hardware", "big data", "certificado digital",
        "ciberamenaza", "ciberdelincuente", "ciberespacio", "criptografia",
        "comprobante", "envios de dinero",
    ],
}


INSTITUTIONAL_WORDS = [
    "asociacion", "instituto", "organismo", "agencia", "comite", "comision",
    "consejo", "federacion", "confederacion", "camara", "banco", "corporacion",
    "entidad", "autoridad", "ministerio", "secretaria", "superintendencia",
]


def keyword_weight(keyword, field, strength, exact=False):
    tokens = len(normalize_keyword(keyword).split())
    if field == "lema" and strength == "strong":
        base = 12
    elif field == "lema":
        base = 6
    elif strength == "strong":
        base = 5
    else:
        base = 1
    if tokens > 1:
        base += min(tokens - 1, 2)
    if exact and field == "lema":
        base += 4
    return base


def is_institutional_acronym(lema, definition):
    raw = re.sub(r"[\s\-.,:;()0-9]", "", str(lema))
    if len(raw) < 2:
        return False
    upper_ratio = sum(ch.isupper() for ch in raw) / len(raw)
    return upper_ratio >= 0.70 and any(contains(definition, w) for w in INSTITUTIONAL_WORDS)


def classify(lema, definition):
    lemma_n = normalize(lema)
    defn_n = normalize(definition)
    combined = f"{lemma_n} {defn_n}".strip()
    scores = defaultdict(int)
    evidence = defaultdict(list)
    evidence_level = defaultdict(int)
    defn_support = defaultdict(list)
    context_activation = defaultdict(set)

    for domain, levels in LEXICON.items():
        for strength, keywords in levels.items():
            for keyword in keywords:
                if contains(lemma_n, keyword):
                    exact = bool(keyword_pattern(keyword).fullmatch(lemma_n))
                    points = keyword_weight(keyword, "lema", strength, exact)
                    scores[domain] += points
                    evidence[domain].append(f"L:{keyword}(+{points})")
                    evidence_level[domain] = max(
                        evidence_level[domain],
                        3 if strength == "strong" else (2 if exact else 1),
                    )
                if contains(defn_n, keyword):
                    points = keyword_weight(keyword, "defn", strength)
                    if strength == "strong":
                        scores[domain] += points
                        evidence[domain].append(f"D:{keyword}(+{points})")
                        evidence_level[domain] = max(evidence_level[domain], 2)
                    else:
                        defn_support[domain].append((keyword, points))

    # Las coincidencias generales de una definición aportan un máximo de
    # cuatro puntos por dominio. De este modo, una enumeración incidental no
    # puede imponerse a una denominación específica en el lema.
    for domain, matches in defn_support.items():
        remaining = 4
        for keyword, points in matches:
            if remaining <= 0:
                break
            awarded = min(points, remaining)
            scores[domain] += awarded
            evidence[domain].append(f"D:{keyword}(+{awarded})")
            remaining -= awarded

    for domain, keywords in EXACT_LEMMA_FALLBACK.items():
        for keyword in keywords:
            if keyword_pattern(keyword).fullmatch(lemma_n):
                scores[domain] += 6
                evidence[domain].append(f"E:{keyword}(+6)")
                evidence_level[domain] = max(evidence_level[domain], 1)

    for trigger, domain, contexts in CONTEXT_RULES:
        if contains(combined, trigger) and any(contains(combined, c) for c in contexts):
            scores[domain] += 7
            evidence[domain].append(f"C:{trigger}(+7)")
            evidence_level[domain] = max(evidence_level[domain], 2)
            context_activation[trigger].add(domain)

    if is_institutional_acronym(lema, defn_n):
        scores["SisFin"] += 7
        evidence["SisFin"].append("R:acronimo_institucional(+7)")
        evidence_level["SisFin"] = max(evidence_level["SisFin"], 2)

    # La jerarquía de la evidencia precede a la suma de puntos.
    # 3 = término fuerte en el lema; 2 = término fuerte en la definición o
    # regla contextual; 1 = término de apoyo en el lema; 0 = apoyo definicional.
    ranked = sorted(
        scores.items(),
        key=lambda item: (-evidence_level[item[0]], -item[1], item[0]),
    )
    if not ranked or ranked[0][1] < 5:
        return "NC", "ninguno", 0, 0, ""

    winner, top_score = ranked[0]
    second_score = ranked[1][1] if len(ranked) > 1 else 0
    top_level = evidence_level[winner]
    second_level = evidence_level[ranked[1][0]] if len(ranked) > 1 else -1
    margin = (top_level - second_level) * 100 + top_score - second_score

    # REV se reserva para conflictos entre evidencias comparables. Cuando hay
    # evidencia fuerte se emplea una diferencia relativa del 5 %. Si solo
    # existen apoyos generales, se exige un empate exacto.
    comparable_strong = (
        top_level == second_level
        and top_level >= 2
        and second_score >= 5
        and (top_score - second_score) / top_score <= 0.05
    )
    support_tie = (
        top_level == second_level and top_level <= 1 and top_score == second_score
    )
    # La activación de dos reglas asociadas al mismo disparador no constituye
    # por sí sola una ambigüedad. Solo es relevante si los dominios activados
    # son los dos primeros del ranking y sus evidencias resultan comparables.
    # Así se evitan falsos REV causados por vocabulario incidental en la
    # definición cuando el lema aporta una clasificación inequívoca.
    context_conflicts = {
        trigger: domains
        for trigger, domains in context_activation.items()
        if len(domains) > 1
    }
    runner_up = ranked[1][0] if len(ranked) > 1 else None
    decisive_context_conflicts = {
        trigger: domains
        for trigger, domains in context_conflicts.items()
        if runner_up is not None
        and winner in domains
        and runner_up in domains
        and (comparable_strong or support_tie)
    }
    if decisive_context_conflicts:
        detail = "; ".join(
            f"C:{trigger}={'|'.join(sorted(domains))}"
            for trigger, domains in decisive_context_conflicts.items()
        )
        return "REV", "contexto_multiple", top_score, margin, (
            "; ".join(evidence[winner] + [detail])
        )
    if comparable_strong or support_tie:
        return "REV", "empate", top_score, margin, "; ".join(evidence[winner])

    method = "lema" if any(x.startswith("L:") for x in evidence[winner]) else "definicion"
    return winner, method, top_score, margin, "; ".join(evidence[winner])


def main(input_file=DEFAULT_INPUT_FILE, output_file=DEFAULT_OUTPUT_FILE,
         summary_file=DEFAULT_SUMMARY_FILE):
    wb = load_workbook(input_file, read_only=True, data_only=True)
    rows = []
    corpus_counts = defaultdict(int)
    country_counts = defaultdict(lambda: defaultdict(int))

    for sheet_name in wb.sheetnames:
        country = sheet_name[:2]
        for row in wb[sheet_name].iter_rows(values_only=True):
            if not row or not row[0] or not str(row[0]).strip():
                continue
            lemma = str(row[0]).strip()
            definition = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            typology = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            loan = str(row[3]).strip() if len(row) > 3 and row[3] else ""
            domain, method, score, margin, evidence = classify(lemma, definition)
            corpus_counts[domain] += 1
            country_counts[country][domain] += 1
            rows.append({
                "Glosario": sheet_name,
                "País": COUNTRY_LABELS.get(country, country),
                "Lema": lemma,
                "Definición": definition,
                "Tipología": typology,
                "Préstamo": loan,
                "Subdominio": domain,
                "Método": method,
                "Puntuación": score,
                "Margen": margin,
                "Evidencia": evidence,
            })

    pd.DataFrame(rows).to_excel(output_file, index=False)

    domains = ["SisFin", "CrePre", "AhoDep", "InvMer", "SegPre",
               "ConEco", "MedDig", "ConJur", "REV", "NC"]
    summary = []
    for country in dict.fromkeys(sheet[:2] for sheet in wb.sheetnames):
        counts = country_counts[country]
        total = sum(counts.values())
        item = {"País": COUNTRY_LABELS.get(country, country), "Total": total}
        for domain in domains:
            item[domain] = counts.get(domain, 0)
            item[f"{domain}_%"] = round(100 * counts.get(domain, 0) / total, 1)
        summary.append(item)
    pd.DataFrame(summary).to_excel(summary_file, index=False)

    total = len(rows)
    print(f"Total: {total:,}")
    for domain in domains:
        count = corpus_counts.get(domain, 0)
        print(f"{domain:7s} {count:6,d} {100 * count / total:5.1f}%")
    print(f"Salida: {output_file}")
    print(f"Resumen: {summary_file}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Clasifica las entradas del corpus TPEF en ocho subdominios."
    )
    parser.add_argument("input", nargs="?", type=Path, default=DEFAULT_INPUT_FILE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_FILE)
    parser.add_argument("--summary", type=Path, default=DEFAULT_SUMMARY_FILE)
    args = parser.parse_args()
    main(args.input, args.output, args.summary)
