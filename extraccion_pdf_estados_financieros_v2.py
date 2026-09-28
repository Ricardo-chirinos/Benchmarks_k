import os
import re
import unicodedata

import pdfplumber
import openpyxl

# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA_PROYECTO = os.path.dirname(os.path.abspath(__file__))
CARPETA_PDFS = os.path.join(CARPETA_PROYECTO, "pdfs")
ARCHIVO_SALIDA = os.path.join(CARPETA_PROYECTO, "resultado_tablas.xlsx")

LINEAS_CONTEXTO = 12
VENTANA_LINEAS_TITULO = 4
MAX_LINEAS_BLOQUE_SIGUIENTE = 10
MAX_RETROCESO_ENCABEZADO = 15
INDICE_RIESGO_MINIMO = 0.02
INDICE_RIESGO_MAXIMO = 0.40

SECCION_NO_CLASIFICADA = "seccion_no_clasificada"


# ============================================================
# UTILIDADES GENERALES
# ============================================================

def obtener_archivos_pdf(carpeta):
    """
    Recorre `carpeta` de forma recursiva (`os.walk`), así que soporta tanto
    poner los PDF sueltos directamente en `pdfs/` como organizarlos en
    subcarpetas por institución (ej. `pdfs/CMR/marzo_2025.pdf`,
    `pdfs/CAT/marzo_2025.pdf`), a cualquier profundidad. El nombre de la
    subcarpeta inmediata además se usa como pista de entidad en
    `identificar_entidad_y_periodo` si el nombre del archivo no alcanza.
    """
    archivos_pdf = []
    if not os.path.isdir(carpeta):
        return archivos_pdf
    for raiz, _subcarpetas, nombres_archivo in os.walk(carpeta):
        for nombre_archivo in nombres_archivo:
            if nombre_archivo.lower().endswith(".pdf"):
                ruta_pdf = os.path.join(raiz, nombre_archivo)
                if os.path.isfile(ruta_pdf):
                    archivos_pdf.append(ruta_pdf)
    archivos_pdf.sort()
    return archivos_pdf


def leer_pdf(ruta_pdf):
    paginas = []
    with pdfplumber.open(ruta_pdf) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text()
            paginas.append({"pagina": numero_pagina, "texto": texto})
    return paginas


def normalizar_texto(texto):
    if not texto:
        return ""
    t = unicodedata.normalize("NFKD", texto)
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.lower()
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def es_probable_encabezado(linea):
    linea_limpia = linea.strip()
    if len(linea_limpia) < 4 or len(linea_limpia) > 140:
        return False
    letras = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", linea_limpia)
    if len(letras) < 4:
        return False
    digitos = re.findall(r"\d", linea_limpia)
    if len(digitos) > len(letras):
        return False
    valores = re.findall(
        r"(?<!\S)\(?[+-]?\d+(?:\.\d+)*(?:,\d+)?%?\)?(?!\S)", linea_limpia
    )
    if len(valores) >= 2:
        return False
    return True


# ============================================================
# PASO 1: IDENTIFICACIÓN DE ENTIDAD, AÑO Y PERÍODO
# ============================================================
#
# Antes estos tres valores estaban fijos en main() ("CMR", 2025, "202503"),
# por lo que las 22 filas del Excel quedaban etiquetadas como marzo 2025.
# Ahora se derivan del nombre del archivo, con la portada del PDF como
# respaldo cuando el nombre no trae mes o año.

MESES_A_PERIODO = {
    "enero": "01",
    "febrero": "02",
    "marzo": "03",
    "abril": "04",
    "mayo": "05",
    "junio": "06",
    "julio": "07",
    "agosto": "08",
    "septiembre": "09",
    "setiembre": "09",
    "octubre": "10",
    "noviembre": "11",
    "diciembre": "12",
}

# Diccionario de entidades. Se busca por token normalizado dentro del nombre
# del archivo. Ampliar a medida que se incorporen nuevas compañías.
ENTIDADES_CONOCIDAS = {
    "cmr": "CMR",
    "falabella": "CMR",
    "hites": "Hites",
    "ripley": "Ripley",
    "abcdin": "ABCDIN",
    "la polar": "La Polar",
    "polar": "La Polar",
    "cat administradora": "CAT",
    "solventa": "Solventa",
    "matic kard": "SB Pay",
    "sb pay": "SB Pay",
    "sbpay": "SB Pay",
    # Identificado 2026-09-27 cruzando contra el Excel de benchmark (hoja
    # "1. SSFF"): razón social "Servicios Financieros y Administración de
    # Créditos Comerciales S.A." — sólo se agrega el reconocimiento de
    # nombre/período; la extracción de campos para esta entidad NO está
    # validada todavía (su hoja del Excel tiene un layout de columnas
    # distinto, sin columnas "Comprobación").
    "servicios financieros y administracion de creditos comerciales": "SSFF",
    "servicios financieros": "SSFF",
    # Identificado 2026-09-27 (Ricardo compartió "05. EEFF TRICOT marzo
    # 2025.pdf" / "04. EEFF TRICOT diciembre 2024.pdf" y salían como
    # DESCONOCIDA). Razón social "Empresas Tricot S.A. y Subsidiarias".
    "empresas tricot": "Tricot",
    "tricot": "Tricot",
}

ENTIDAD_POR_DEFECTO = "DESCONOCIDA"


def _buscar_anio(texto_normalizado):
    """Devuelve el año más alto encontrado (evita quedarse con el año
    comparativo, que siempre es menor que el del período reportado).

    Usa lookaround en vez de \\b: en nombres de archivo como
    "22._EEFF_CMR_marzo_2025.pdf" el año queda pegado a un guion bajo, que
    \\b trata como parte de la palabra, y el año nunca se encontraba."""
    anios = [
        int(a) for a in re.findall(r"(?<!\d)(?:19|20)\d{2}(?!\d)", texto_normalizado)
    ]
    return max(anios) if anios else None


def _buscar_anio_para_mes(texto_normalizado, mes_numero):
    """Busca el año que aparece junto al mes ya identificado ("de marzo de
    2025"), en vez de tomar el año más alto de todo el texto. Evita que una
    fecha de firma del informe del auditor (que suele quedar en el año
    siguiente al período reportado y aparecer en las primeras páginas) se
    confunda con el año del período."""
    if mes_numero is None:
        return None
    nombre_mes = _NUMERO_A_MES.get(mes_numero)
    if nombre_mes is None:
        return None
    coincidencia = re.search(
        rf"de\s+{nombre_mes}\s+(?:de\s+)?(\d{{4}})", texto_normalizado
    )
    return int(coincidencia.group(1)) if coincidencia else None


def _buscar_mes(texto_normalizado):
    for nombre_mes, numero in MESES_A_PERIODO.items():
        # Con "in" a secas, "mayo" matcheaba como substring de "mayor" (bug
        # detectado 2026-09-27 con SSFF: su informe de auditores dice "el
        # riesgo de fraude es mayor que..." antes de mencionar "diciembre",
        # así que el período quedaba en mayo/05 en vez de diciembre/12).
        # Lookaround en vez de \b por la misma razón que en `_buscar_anio`:
        # un mes puede quedar pegado a un guion bajo u otro caracter no
        # alfabético que \b trataría como límite de palabra.
        if re.search(rf"(?<!\w){re.escape(nombre_mes)}(?!\w)", texto_normalizado):
            return numero
    return None


_NUMERO_A_MES = {v: k for k, v in MESES_A_PERIODO.items() if k not in ("setiembre",)}


def _buscar_entidad(texto_normalizado):
    for token, nombre in ENTIDADES_CONOCIDAS.items():
        if token in texto_normalizado:
            return nombre
    return None


def identificar_entidad_y_periodo(ruta_o_nombre_pdf, paginas):
    """
    Devuelve (entidad, anio, periodo) donde `periodo` tiene formato AAAAMM.

    `ruta_o_nombre_pdf` puede ser una ruta completa o sólo el nombre del
    archivo. Orden de prioridad (son los datos que el usuario controla,
    antes de recurrir al texto del PDF):
        1) nombre del archivo
        2) nombre de la carpeta contenedora inmediata (soporta organizar
           los PDF como `pdfs/<INSTITUCION>/<archivo>.pdf`)
        3) texto de las primeras páginas del PDF
    """
    nombre_archivo = os.path.basename(ruta_o_nombre_pdf)
    nombre_carpeta = os.path.basename(os.path.dirname(ruta_o_nombre_pdf))

    base = normalizar_texto(nombre_archivo)
    carpeta_normalizada = normalizar_texto(nombre_carpeta)

    anio = _buscar_anio(base)
    mes = _buscar_mes(base)
    entidad = _buscar_entidad(base) or _buscar_entidad(carpeta_normalizada)

    if anio is None or mes is None or entidad is None:
        cabecera = normalizar_texto(
            " ".join((p["texto"] or "") for p in paginas[:3])
        )
        if mes is None:
            mes = _buscar_mes(cabecera)
        if anio is None:
            # Preferir el año pegado al mes del período ("de marzo de 2025")
            # sobre el año más alto de las primeras páginas, que puede ser
            # la fecha de firma del informe de auditoría (típicamente el
            # año siguiente al período reportado).
            anio = _buscar_anio_para_mes(cabecera, mes) or _buscar_anio(cabecera)
        if entidad is None:
            entidad = _buscar_entidad(cabecera)

    if entidad is None:
        entidad = ENTIDAD_POR_DEFECTO

    periodo = f"{anio}{mes}" if (anio is not None and mes is not None) else None
    return entidad, anio, periodo


# ============================================================
# PASO 3: PATRONES DE TABLA (formato nuevo + formato antiguo)
# ============================================================
#
# Cada entrada es una lista de "grupos"; un grupo matchea si TODOS sus
# tokens aparecen en la ventana de líneas. Las variantes antiguas provienen
# del EEFF diciembre 2019, donde los títulos no llevan la palabra
# "intermedios" y las notas se llaman distinto.

TABLA_PATRONES = {
    "movimiento_provisiones": [
        ["resumen del movimiento en provisiones constituidas"],
        # Formato antiguo - Nota 6 c) / Nota 13 b)
        ["movimiento de las provisiones"],
        ["movimiento que se ha producido en las provisiones"],
        ["provisiones sobre cartera vencida"],
        # SSFF - Nota 9: título repartido en 3 líneas con una fila "Evaluación
        # Grupal" intercalada ("RESUMEN DEL MOVIMIENTO EN PROVISIONES" / ... /
        # "CONSTITUIDAS POR" / "Evaluación Grupal" / "CARTERA DE RIESGO DE
        # CRÉDITO..."), y además "PROVISIONES" y "CONSTITUIDAS" quedan
        # pegados sin espacio por un artefacto de pdfplumber. Los tokens no
        # necesitan ser contiguos entre sí dentro de la ventana de líneas,
        # así que se separan en 3 para no depender de ese espacio ni de la
        # línea intercalada.
        ["resumen del movimiento en provisiones", "constituidas por", "cartera de riesgo de credito"],
    ],
    "gasto_perdidas_totales": [
        ["resumen del gasto por perdidas crediticias"],
        # Formato antiguo - Nota 21
        ["provisiones y deterioro por riesgo de credito"],
        ["resultado neto provision por riesgo de credito"],
        # SB Pay (Matic Kard) - Nota 28 "Provisiones por riesgo de crédito
        # neto": mismo concepto (castigos + recuperos + provisiones netas
        # del período), título propio de esta entidad.
        ["provisiones por riesgo de credito neto"],
    ],
    "gasto_perdidas_detalle": [
        ["resumen del gasto de provisiones"],
        ["castigos directos y recuperaciones"],
    ],
    # Tricot - Nota 6 a.5) "Castigos": tabla "Total castigos" / "Total
    # recuperos deudas castigadas" / "Total castigo neto", con 4 columnas
    # (actual, comparativo, var $, var %) — no la estructura de
    # `resolver_fila_movimiento` (que asume "total del primer bloque" en
    # un layout de 2 bloques actual+comparativo sin columnas de variación).
    "castigos_variacion_tricot": [
        ["a.5) castigos"],
    ],
    "concentracion_morosidad": [
        ["concentracion", "morosidad"],
        ["riesgo de credito por dias de morosidad"],
    ],
    "estado_resultados": [
        ["estados intermedios de resultados"],
        # Formato antiguo
        ["estados del resultado"],
        ["estado de resultados"],
        # CAT usa el plural ("Estados de Resultados Integrales"), que no
        # matcheaba contra "estado de resultados" (singular). Sin esto la
        # sección quedaba pegada a la nota anterior (ej. composición de
        # cartera) durante toda la página del estado de resultados, y
        # `recupero_acum` no reconocía la fila correcta como válida.
        ["estados de resultados"],
    ],
    "estado_situacion_financiera": [
        ["estados intermedios de situacion financiera"],
        # Formato antiguo
        ["estados de situacion financiera"],
    ],
    # Formato antiguo: la composición de cartera vive en la Nota 6 a),
    # que no tiene equivalente directo en el formato nuevo.
    "composicion_cartera": [
        ["composicion de la cartera de colocaciones"],
        # CAT (2026-09-27): esta frase también aparece, partida en dos líneas,
        # dentro del Estado de Resultados ("Provisiones por riesgo de
        # créditos y cuentas por cobrar / a clientes", nota de gasto por
        # pérdidas crediticias). Sin la exclusión, ese falso positivo hace
        # que la sección salte de vuelta a `composicion_cartera` en medio del
        # Estado de Resultados y `recupero_acum` pierde la fila correcta
        # ("Recuperación de créditos castigados"). El título real de la nota
        # de cartera (formato nuevo y antiguo) nunca trae "provisiones por
        # riesgo" en la misma ventana de líneas, así que excluirlo no afecta
        # la detección legítima.
        {
            "incluye": ["creditos y cuentas por cobrar a clientes"],
            "excluye": ["provisiones por riesgo"],
        },
    ],
    # SB Pay (Matic Kard) - Nota 9.3 "Estratificación e información de la
    # cartera": tabla por segmento de producto (Antiguo/Nuevo con
    # uso/Refinanciado/Renegociado + fila TOTAL), no la estructura 2N+1 de
    # CMR/CAT. Su fila TOTAL trae 5 columnas: Saldo(>0), Saldo contingente,
    # Provisión contingente, Provisión cartera normal, Provisión cartera en
    # incumplimiento — sin columna neta, así que no se valida con
    # `resolver_fila_cartera` sino con `resolver_fila_cartera_segmentada`.
    "cartera_segmentada_por_producto": [
        ["segmentos", "saldos", "provisiones"],
    ],
    # Tricot - Nota 6 "II. Cartera total": fila TOTAL con 2 períodos lado a
    # lado, cada uno de 5 columnas (N° clientes, cartera bruta, provisión,
    # cartera neta, factor de provisión) — no es la estructura simétrica
    # 2N+1 de CMR/CAT, se valida con `resolver_fila_cartera_dos_periodos`.
    "cartera_total_dos_periodos": [
        # Excluye "% Castigo/cartera total" y "% Castigo neto/cartera
        # total" (indicadores de castigo de Tricot): esa frase contiene
        # "cartera total" como substring pero es una tabla completamente
        # distinta (porcentajes, no la fila TOTAL bruto/provisión/neto) que
        # aparece ANTES del título real "II. Cartera total" en el documento
        # — sin la exclusión, el título salta demasiado pronto y arrastra
        # varias tablas de subtotales por producto que no son el total
        # consolidado que se busca.
        {"incluye": ["cartera total"], "excluye": ["castigo/cartera", "castigos/cartera", "castigo neto/cartera"]},
    ],
    # Hites - Nota 10 "Indicadores de riesgo y castigo": fila TOTAL de 6
    # columnas (Provisión, Cartera bruta, Índice de riesgo %, Castigos, %
    # castigos, % mensual castigos) por período, en tablas separadas por
    # fecha (no interleaved). Se valida con `resolver_fila_indicadores_riesgo`.
    "indicadores_riesgo_castigo": [
        ["indicadores de riesgo y castigo"],
    ],
    # Hites - subtítulo "Provisión asociada a cada tipo de cartera:", justo
    # antes de las filas "Total castigos del período" / "Total recuperos
    # del período" (2 columnas: actual, comparativo, sin identidad contable
    # que autovalidar).
    "provision_por_cartera_hites": [
        ["provision asociada a cada tipo de cartera"],
    ],
    # SSFF - "Estratificación de la cartera" (verificado 2026-09-28 contra
    # PDF real dic-2024): tabla por tramos de mora con la misma estructura
    # 2N+1 que `composicion_cartera`/`concentracion_morosidad` (fila ">=90"
    # con 7 columnas: normal/incumplimiento/total para bruto, ídem para
    # provisión, y neto), pero el título de la nota nunca dice "morosidad"
    # ni "concentración" en la misma línea que su propio nombre — es un
    # título propio ("Estratificación de la cartera") que antes no estaba
    # mapeado a ninguna sección, así que la tabla quedaba con el
    # `tabla_activa` de la nota anterior (rollforward de provisiones) y
    # `consumo_90` nunca la encontraba.
    "estratificacion_cartera_ssff": [
        ["estratificacion de la cartera"],
    ],
}

CARTERA_PATRONES = {
    "colocaciones_consumo": ["colocaciones de consumo"],
    "creditos_contingentes": [
        "exposicion por creditos contingentes",
        "creditos contingentes",
    ],
    "adeudado_bancos": ["adeudado por bancos"],
    "colocaciones_comerciales": ["colocaciones comerciales"],
    "colocaciones_vivienda": [
        "colocaciones para vivienda",
        "colocaciones para la vivienda",
    ],
}


def combinar_seccion(tabla_activa, cartera_activa):
    if tabla_activa == "movimiento_provisiones":
        if cartera_activa in (None, "colocaciones_consumo"):
            # En el formato antiguo la única cartera es consumo y muchas
            # veces el encabezado de cartera no se repite dentro de la nota.
            return "movimiento_provisiones_consumo"
        return SECCION_NO_CLASIFICADA
    if tabla_activa == "gasto_perdidas_totales":
        return "gasto_perdidas_crediticias"
    if tabla_activa == "gasto_perdidas_detalle":
        return "gasto_perdidas_detalle_por_cartera"
    if tabla_activa == "concentracion_morosidad":
        return "morosidad_colocaciones"
    if tabla_activa == "estado_resultados":
        return "estado_resultados"
    if tabla_activa == "estado_situacion_financiera":
        return "estado_situacion_financiera"
    if tabla_activa == "composicion_cartera":
        return "composicion_cartera"
    if tabla_activa == "cartera_segmentada_por_producto":
        return "cartera_segmentada_por_producto"
    if tabla_activa == "cartera_total_dos_periodos":
        return "cartera_total_dos_periodos"
    if tabla_activa == "indicadores_riesgo_castigo":
        return "indicadores_riesgo_castigo"
    if tabla_activa == "castigos_variacion_tricot":
        return "castigos_recuperos_variacion"
    if tabla_activa == "provision_por_cartera_hites":
        return "recupero_castigo_simple_hites"
    if tabla_activa == "estratificacion_cartera_ssff":
        # Misma estructura 2N+1 que `morosidad_colocaciones`; se reutiliza
        # la sección existente para no duplicar el resolutor ni las
        # SECCIONES_ESPERADAS de consumo_90.
        return "morosidad_colocaciones"
    return SECCION_NO_CLASIFICADA


# ============================================================
# PASO 2: SEGUIMIENTO DE PERÍODO DINÁMICO
# ============================================================
#
# Antes PERIODO_PATRONES traía hardcodeado "al 31 de marzo de 2025". Con
# cualquier otro PDF el período quedaba siempre en None y se perdía la
# desambiguación entre columna actual y columna comparativa.

_MESES_TEXTO = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


def _variantes_fecha(anio):
    """Genera las formas en que un año puede aparecer en un encabezado de
    fecha, incluyendo variantes con espacio anómalo dentro del nombre del
    mes (artefacto frecuente de la extracción de texto de pdfplumber)."""
    variantes = []
    for mes in _MESES_TEXTO:
        variantes.append(f"de {mes} de {anio}")
        variantes.append(f"de {mes} {anio}")
        # "ma rzo", "septie mbre", etc.
        for corte in range(2, len(mes) - 1):
            mes_partido = f"{mes[:corte]} {mes[corte:]}"
            variantes.append(f"de {mes_partido} de {anio}")
    variantes.append(f"al 31 de diciembre de {anio}")
    variantes.append(f"al 31 de diciembre {anio}")
    return variantes


def construir_periodo_patrones(anio_actual):
    """
    Devuelve el diccionario {periodo: [patrones]} para el año del reporte.

    - "actual"      -> el año del período que se está extrayendo
    - "comparativo" -> el año anterior (columna de comparación)

    Si no se conoce el año, devuelve un diccionario vacío: todas las líneas
    quedarán con periodo=None, que `prioridad_periodo` trata como neutral.
    """
    if anio_actual is None:
        return {}
    return {
        "actual": _variantes_fecha(anio_actual) + [str(anio_actual)],
        "comparativo": (
            _variantes_fecha(anio_actual - 1) + [str(anio_actual - 1)]
        ),
    }


def calcular_secciones_por_linea(lineas, anio_actual=None, estado_inicial=None):
    tabla_activa, cartera_activa, periodo_activo = estado_inicial or (None, None, None)
    secciones_por_linea = []

    periodo_patrones = construir_periodo_patrones(anio_actual)
    patron_nota = re.compile(r"^nota\s+\d+\s*[-–.:)]?\s", flags=re.IGNORECASE)
    lineas_normalizadas = [normalizar_texto(l) for l in lineas]

    for indice, linea_normalizada in enumerate(lineas_normalizadas):
        if patron_nota.match(linea_normalizada):
            tabla_activa = None
            cartera_activa = None
            periodo_activo = None

        inicio_ventana = max(0, indice - VENTANA_LINEAS_TITULO + 1)
        ventana_texto = " ".join(lineas_normalizadas[inicio_ventana:indice + 1])

        tabla_encontrada = None
        for nombre_tabla, grupos in TABLA_PATRONES.items():
            for grupo in grupos:
                if isinstance(grupo, dict):
                    incluye = grupo["incluye"]
                    excluye = grupo.get("excluye", [])
                else:
                    incluye = grupo
                    excluye = []
                if all(token in ventana_texto for token in incluye) and not any(
                    token in ventana_texto for token in excluye
                ):
                    tabla_encontrada = nombre_tabla
                    break
            if tabla_encontrada:
                break

        if tabla_encontrada and tabla_encontrada != tabla_activa:
            tabla_activa = tabla_encontrada
            cartera_activa = None
            periodo_activo = None

        cartera_encontrada = None
        for nombre_cartera, patrones in CARTERA_PATRONES.items():
            for patron in patrones:
                if patron in linea_normalizada:
                    cartera_encontrada = nombre_cartera
                    break
            if cartera_encontrada:
                break

        if cartera_encontrada:
            cartera_activa = cartera_encontrada

        periodo_encontrado = None
        for nombre_periodo, patrones in periodo_patrones.items():
            for patron in patrones:
                if patron in ventana_texto:
                    periodo_encontrado = nombre_periodo
                    break
            if periodo_encontrado:
                break

        if periodo_encontrado:
            periodo_activo = periodo_encontrado

        seccion = combinar_seccion(tabla_activa, cartera_activa)
        secciones_por_linea.append({
            "seccion": seccion,
            "tabla_activa": tabla_activa,
            "cartera_activa": cartera_activa,
            "periodo": periodo_activo,
        })

    return secciones_por_linea, (tabla_activa, cartera_activa, periodo_activo)


# ============================================================
# SECCIONES ESPERADAS Y EXCLUSIONES POR CAMPO
# ============================================================
#
# El orden importa: `seleccionar_resultado_por_campo` usa el índice de la
# sección dentro de la lista como primer criterio de desempate.

SECCIONES_ESPERADAS = {
    "saldo_efectivo": [
        "morosidad_colocaciones",
        "composicion_cartera",
        # SB Pay: ver `cartera_segmentada_por_producto` más abajo.
        "cartera_segmentada_por_producto",
        # Tricot / Hites: ver `cartera_total_dos_periodos` e
        # `indicadores_riesgo_castigo` más abajo. Antes de
        # `movimiento_provisiones_consumo`: ese fallback genérico
        # (`resolver_fila_movimiento`, sin identidad contable) puede
        # engancharse con una fila de rollforward no relacionada en
        # documentos con títulos de nota menos estandarizados (ej. Hites),
        # y estas 2 secciones sí autovalidan bruto-provisión=neto. CMR/CAT
        # nunca generan estas 2 secciones, así que el orden no les afecta.
        "cartera_total_dos_periodos",
        "indicadores_riesgo_castigo",
        "movimiento_provisiones_consumo",
    ],
    "provision_efectiva": [
        "composicion_cartera",
        "cartera_segmentada_por_producto",
        "cartera_total_dos_periodos",
        "indicadores_riesgo_castigo",
        "movimiento_provisiones_consumo",
    ],
    "castigo_acum": [
        "movimiento_provisiones_consumo",
        "gasto_perdidas_detalle_por_cartera",
        # SB Pay: "Castigos" vive en la Nota 28 (mapeada a
        # gasto_perdidas_crediticias, ver TABLA_PATRONES), no en el
        # rollforward de provisiones. Al final para no interferir con
        # CMR/CAT, que sí resuelven por movimiento_provisiones_consumo.
        "gasto_perdidas_crediticias",
        # Hites: fila TOTAL de "Indicadores de riesgo y castigo" (Nota 10)
        # trae provisión, cartera bruta Y castigos en la misma fila.
        "indicadores_riesgo_castigo",
        # Tricot: tabla "a.5) Castigos" (4 columnas: actual, comparativo,
        # var $, var %).
        "castigos_recuperos_variacion",
        # Hites: fila simple de 2 columnas "Total castigos del período".
        "recupero_castigo_simple_hites",
    ],
    # "movimiento_provisiones_consumo" queda al final a propósito: la fila
    # "Recuperación de créditos castigados" del rollforward NIIF9 (Nota 7c
    # en CAT) es un concepto distinto — recupero dentro del movimiento de
    # provisiones, no el recupero de cartera castigada que reporta el
    # Estado de Resultados y que es el valor verificado en el Excel (ver
    # CAT mar-2025: rollforward=124, Estado de Resultados=5.496, Excel=
    # 5.496). CMR no tiene este conflicto porque esa fila viene vacía en su
    # rollforward, pero si algún día trae dato real, gana igual el Estado de
    # Resultados/gasto_perdidas_crediticias, que es la fuente consistente
    # con el proceso manual.
    "recupero_acum": [
        "gasto_perdidas_crediticias",
        "gasto_perdidas_detalle_por_cartera",
        "estado_resultados",
        "movimiento_provisiones_consumo",
        # Tricot: misma tabla "a.5) Castigos" que castigo_acum.
        "castigos_recuperos_variacion",
        # Hites: tabla simple de 2 columnas (actual/comparativo) "Total
        # recuperos del período", bajo el subtítulo "Provisión asociada a
        # cada tipo de cartera".
        "recupero_castigo_simple_hites",
    ],
    "consumo_90": ["morosidad_colocaciones"],
}

EXCLUSIONES_POR_CAMPO = {}


# ============================================================
# PALABRAS CLAVE
# ============================================================

PALABRAS_CLAVE = {
    "saldo_efectivo": [
        "subtotal",
        "totales colocaciones",
        "colocaciones de consumo",
        "activos financieros antes de provisiones",
        "deudores por tarjetas de credito",
        # SB Pay: fila "TOTAL" de la tabla de cartera segmentada por
        # producto (Nota 9.3). Palabra genérica, pero sólo se acepta dentro
        # de `cartera_segmentada_por_producto` (ver SECCIONES_ESPERADAS),
        # así que no interfiere con los "TOTAL X" de otras secciones.
        "total",
    ],
    "provision_efectiva": [
        "provisión efectiva", "provisiones efectivas",
        "saldos al 31 de marzo", "saldos al 30 de junio",
        "saldos al 30 de septiembre", "saldos al 31 de diciembre",
        "resumen del movimiento en provisiones constituidas",
        "provisiones constituidas", "total provisiones",
        # La fila "Subtotal" de composición de cartera trae saldo, provisión
        # y neto en la misma línea (ver resolver_fila_cartera). Sin esta
        # palabra clave, provision_efectiva nunca encuentra esa fila por sí
        # sola aunque saldo_efectivo sí la encuentre.
        "subtotal",
        # SB Pay: misma fila "TOTAL" de la cartera segmentada (ver arriba).
        "total",
    ],
    "castigo_acum": [
        "castigos acumulados", "castigo acumulado",
        "aplicación de provisiones por castigos",
        "castigos de cartera deteriorada",
        "total de castigos",
        "castigos", "castigo",
        "castigos del período", "castigos del periodo", "write-off",
        # SSFF: su fila de castigos se titula distinto ("Baja de activos
        # financieros deteriorados del ejercicio"), sin usar la palabra
        # "castigo" en absoluto. Se incluye "del ejercicio" en la palabra
        # clave para que no quede texto no numérico después del término
        # (extraer_valores_tabulares descarta la línea si sobra texto con
        # letras, como pasaría dejando sólo "deteriorados").
        "baja de activos financieros deteriorados del ejercicio",
    ],
    "recupero_acum": [
        "recuperación de créditos castigados",
        "recuperacion de creditos castigados",
        "recuperación de activos castigados",
        "recuperacion de activos castigados",
        "recuperación de castigos", "recuperacion de castigos",
        "recuperos", "recuperaciones",
        # SB Pay (Nota 28): "Recuperación de créditos por deudores tarjeta
        # crédito castigados" — la frase completa nunca queda contigua
        # (pdfplumber la parte en 2 líneas con el valor intercalado en
        # medio, ver buscar_valor_en_linea_pura_siguiente), así que se
        # matchea el inicio de la frase.
        "recuperación de créditos por", "recuperacion de creditos por",
        # Tricot: "Total recuperos deudas castigadas(1)" — la palabra clave
        # genérica "recuperos" deja "deudas castigadas..." como texto
        # sobrante después del término, que `extraer_valores_tabulares`
        # descarta por tener letras; se matchea la frase completa.
        "total recuperos deudas castigadas",
        # Hites: "Total recuperos del período" — mismo problema ("del
        # período" queda como texto sobrante con letras).
        "total recuperos del período", "total recuperos del periodo",
    ],
    "consumo_90": [
        ">= 90 días", ">= 90 dias", "90 días y más", "90 días o más",
        "90 dias y mas", "90 dias o mas", "cartera en incumplimiento",
        "riesgo de crédito por días de morosidad",
        # SSFF - tabla "Estratificación de la cartera" (Nota 9): la fila se
        # titula literalmente ">=90" sin espacio y sin la palabra "días" (a
        # diferencia de CMR/CAT, que sí usan "≥ 90 días" con espacio).
        # Verificado 2026-09-28 contra el PDF real dic-2024 de SSFF.
        ">=90",
    ],
    # NOTA (2026-09-27): `gasto_perdidas_crediticias` se eliminó — no tiene
    # destino en el Excel de benchmark, confirmado como error menor.
}


# ============================================================
# PASO 4: MATCH DE ETIQUETAS TOLERANTE A PLURALES
# ============================================================
#
# El patrón original `(?<!\w)subtotal(?!\w)` fallaba contra "Subtotales"
# (formato 2019) y funcionaba contra "Subtotal" (formato 2025). El sufijo
# opcional resuelve ambos casos sin abrir la puerta a falsos positivos.

SUFIJO_PLURAL = r"(?:es|s)?"


def construir_patron_termino(palabra):
    return rf"(?<!\w){re.escape(palabra)}{SUFIJO_PLURAL}(?!\w)"


# ============================================================
# EXTRACCIÓN DE VALORES TABULARES (crudos, como texto)
# ============================================================

def extraer_valores_tabulares(linea, termino):
    coincidencia = re.search(
        construir_patron_termino(termino), linea, flags=re.IGNORECASE
    )
    if not coincidencia:
        return []

    texto_despues = linea[coincidencia.end():].strip()
    if not texto_despues:
        return []

    texto_despues = re.sub(r"^[\s:;|=]+", "", texto_despues).strip()
    if not texto_despues:
        return []

    texto_despues = re.sub(
        r"^\s*de\s+\d{4}\b", "", texto_despues, flags=re.IGNORECASE
    ).strip()
    if not texto_despues:
        return []

    if re.search(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", texto_despues):
        return []

    patron_caracteres_permitidos = r"^[0-9\s\.,\-\+\(\)%:;\|/=]+$"
    if not re.fullmatch(patron_caracteres_permitidos, texto_despues):
        return []

    patron_valor = r"""
        \(\s*[+-]?\d+(?:\.\d+)*(?:,\d+)?%?\s*\)
        |
        [+-]?\d+(?:\.\d+)*(?:,\d+)?%?
        |
        (?<!\S)-(?!\S)
    """
    valores = re.findall(patron_valor, texto_despues, flags=re.VERBOSE)
    return [re.sub(r"\s+", "", v) for v in valores if v.strip()]


# ============================================================
# BÚSQUEDA DE VALOR EN BLOQUE SIGUIENTE (encabezados "puros")
# ============================================================

ETIQUETAS_BLOQUE_PRIORITARIAS = [
    "colocaciones de consumo",
    "deudores por tarjetas de credito",
]
ETIQUETAS_BLOQUE_RESPALDO = ["subtotal", "total"]


def buscar_valor_en_bloque_siguiente(lineas, indice_header):
    resultado_respaldo = None
    limite = min(len(lineas), indice_header + 1 + MAX_LINEAS_BLOQUE_SIGUIENTE)

    for indice in range(indice_header + 1, limite):
        linea = lineas[indice].strip()
        if not linea:
            continue
        if linea.isupper() and es_probable_encabezado(linea):
            break

        for etiqueta in ETIQUETAS_BLOQUE_PRIORITARIAS:
            valores = extraer_valores_tabulares(linea, etiqueta)
            if valores:
                return indice, valores

        if resultado_respaldo is None:
            for etiqueta in ETIQUETAS_BLOQUE_RESPALDO:
                valores_respaldo = extraer_valores_tabulares(linea, etiqueta)
                if valores_respaldo:
                    resultado_respaldo = (indice, valores_respaldo)
                    break

    if resultado_respaldo:
        return resultado_respaldo
    return None, []


def buscar_valor_en_linea_pura_siguiente(lineas, indice_header):
    """A veces pdfplumber reconstruye una etiqueta partida en dos líneas con
    la fila de valores intercalada en medio (ej. SB Pay Nota 28:
    'Recuperación de créditos por deudores tarjeta crédito' / '887 844' /
    'castigados' — el cierre de la etiqueta queda DESPUÉS del valor). Ni el
    resto de la línea del término ni `buscar_valor_en_bloque_siguiente`
    (que exige una etiqueta conocida) cubren este caso. Si la línea
    inmediatamente siguiente al término es puramente numérica, se usa
    directamente como fila de valores."""
    if indice_header + 1 >= len(lineas):
        return None, []

    linea_siguiente = lineas[indice_header + 1].strip()
    if not linea_siguiente:
        return None, []
    if re.search(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", linea_siguiente):
        return None, []

    patron_caracteres_permitidos = r"^[0-9\s\.,\-\+\(\)%:;\|/=]+$"
    if not re.fullmatch(patron_caracteres_permitidos, linea_siguiente):
        return None, []

    patron_valor = r"""
        \(\s*[+-]?\d+(?:\.\d+)*(?:,\d+)?%?\s*\)
        |
        [+-]?\d+(?:\.\d+)*(?:,\d+)?%?
        |
        (?<!\S)-(?!\S)
    """
    valores = [
        re.sub(r"\s+", "", v)
        for v in re.findall(patron_valor, linea_siguiente, flags=re.VERBOSE)
        if v.strip()
    ]
    if not valores:
        return None, []
    return indice_header + 1, valores


# ============================================================
# BÚSQUEDA DE PALABRAS CLAVE
# ============================================================

def buscar_palabras_clave(texto, palabras, secciones_por_linea, lineas, campo=""):
    coincidencias = []
    if not texto:
        return coincidencias

    palabras_ordenadas = sorted(set(palabras), key=len, reverse=True)
    exclusiones = EXCLUSIONES_POR_CAMPO.get(campo, [])

    for indice_linea, linea in enumerate(lineas):
        linea_limpia = linea.strip()
        if not linea_limpia:
            continue

        linea_normalizada = normalizar_texto(linea_limpia)
        if any(excl in linea_normalizada for excl in exclusiones):
            continue

        for palabra in palabras_ordenadas:
            coincidencia = re.search(
                construir_patron_termino(palabra),
                linea_limpia,
                flags=re.IGNORECASE,
            )
            if not coincidencia:
                continue

            termino_encontrado = coincidencia.group()
            valores_tabulares = extraer_valores_tabulares(
                linea=linea_limpia, termino=termino_encontrado
            )
            indice_linea_valor = indice_linea
            linea_valor = linea_limpia
            linea_encabezado = ""

            if not valores_tabulares:
                resto_linea = linea_limpia[coincidencia.end():].strip()
                if not resto_linea:
                    indice_valor, valores_bloque = buscar_valor_en_bloque_siguiente(
                        lineas, indice_linea
                    )
                    if valores_bloque:
                        valores_tabulares = valores_bloque
                        indice_linea_valor = indice_valor
                        linea_valor = lineas[indice_valor].strip()
                        linea_encabezado = linea_limpia

            if not valores_tabulares:
                # Etiqueta partida con el valor intercalado en medio (ver
                # buscar_valor_en_linea_pura_siguiente): se intenta tanto si
                # sobraba texto no numérico después del término como si no.
                indice_valor, valores_directos = buscar_valor_en_linea_pura_siguiente(
                    lineas, indice_linea
                )
                if valores_directos:
                    valores_tabulares = valores_directos
                    indice_linea_valor = indice_valor
                    linea_valor = lineas[indice_valor].strip()
                    linea_encabezado = linea_limpia

            if not valores_tabulares:
                continue

            valores_numericos = [v for v in valores_tabulares if re.search(r"\d", v)]
            info_seccion = secciones_por_linea[indice_linea]

            coincidencias.append({
                "termino": palabra,
                "numero_linea": indice_linea_valor + 1,
                "linea": linea_valor,
                "linea_encabezado": linea_encabezado,
                "valores_tabulares": valores_tabulares,
                "valores_numericos": valores_numericos,
                "cantidad_columnas": len(valores_tabulares),
                "tiene_cifras": bool(valores_numericos),
                "solo_guiones": all(v == "-" for v in valores_tabulares),
                "seccion": info_seccion["seccion"],
                "periodo": info_seccion["periodo"],
                "indice_linea_valor": indice_linea_valor,
                "lineas_pagina": lineas,
            })
            break

    return coincidencias


def buscar_campos(texto, secciones_por_linea, lineas):
    resultados = {}
    for campo, palabras in PALABRAS_CLAVE.items():
        resultados[campo] = buscar_palabras_clave(
            texto, palabras, secciones_por_linea, lineas, campo=campo
        )
    return resultados


def buscar_campos_en_pdf(paginas, anio_actual=None):
    """El estado de sección se arrastra de una página a la siguiente."""
    resultados_pdf = []
    estado = None  # (tabla_activa, cartera_activa, periodo_activo)

    for pagina in paginas:
        numero_pagina = pagina["pagina"]
        texto = pagina["texto"]
        if not texto:
            # Página sin texto extraíble: no se resetea el estado a propósito,
            # una tabla puede continuar después de una página de imagen.
            continue

        lineas = texto.splitlines()

        secciones_por_linea, estado = calcular_secciones_por_linea(
            lineas, anio_actual, estado
        )

        resultados_pagina = buscar_campos(texto, secciones_por_linea, lineas)

        for campo, coincidencias in resultados_pagina.items():
            for coincidencia in coincidencias:
                resultado = dict(coincidencia)
                resultado["campo"] = campo
                resultado["pagina"] = numero_pagina
                resultados_pdf.append(resultado)

    return resultados_pdf


# ============================================================
# SELECCIÓN DEL MEJOR RESULTADO POR CAMPO
# ============================================================

def filtrar_por_seccion_esperada(resultados_pdf):
    resultados_filtrados = []
    for resultado in resultados_pdf:
        secciones_esperadas = SECCIONES_ESPERADAS.get(resultado["campo"], [])
        if resultado["seccion"] not in secciones_esperadas:
            continue
        if resultado["seccion"] in SECCIONES_TIPO_CARTERA:
            valores = quitar_columna_nota(resultado["valores_tabulares"])
            if resolver_fila_cartera(valores) is None:
                continue
        if resultado["seccion"] in SECCIONES_CARTERA_SEGMENTADA:
            if resolver_fila_cartera_segmentada(resultado["valores_tabulares"]) is None:
                continue
        if resultado["seccion"] in SECCIONES_CARTERA_DOS_PERIODOS:
            if resolver_fila_cartera_dos_periodos(resultado["valores_tabulares"]) is None:
                continue
        if resultado["seccion"] in SECCIONES_INDICADORES_RIESGO:
            if resolver_fila_indicadores_riesgo(resultado["valores_tabulares"]) is None:
                continue
        if resultado["seccion"] in SECCIONES_CASTIGO_RECUPERO_VARIACION:
            if resolver_fila_castigo_recupero_variacion(resultado["valores_tabulares"]) is None:
                continue
        if resultado["seccion"] in SECCIONES_DOS_COLUMNAS_SIMPLE:
            if resolver_fila_dos_columnas_simple(resultado["valores_tabulares"]) is None:
                continue

        resultados_filtrados.append(resultado)
    return resultados_filtrados


def prioridad_periodo(periodo):
    """
    "actual" siempre le gana a "comparativo", incluso si el período actual
    trae sólo guiones (cero). None se trata como neutral, para no penalizar
    campos donde ambos años conviven en la misma línea y se resuelven por
    índice de columna.
    """
    if periodo == "actual":
        return 0
    if periodo is None:
        return 1
    return 2


def seleccionar_resultado_por_campo(resultados_pdf):
    """
    Orden de prioridad:
        1) sección esperada más específica para el campo
        2) período "actual" antes que "comparativo"
        3) tiene cifras (evita filas totalmente vacías si hay opción)
        4) no es sólo guiones
        5) página / número de línea como desempate final
    """
    resultados_validos = filtrar_por_seccion_esperada(resultados_pdf)
    seleccion = {}

    for resultado in resultados_validos:
        campo = resultado["campo"]
        secciones_esperadas = SECCIONES_ESPERADAS.get(campo, [])
        prioridad_seccion = secciones_esperadas.index(resultado["seccion"])
        clave_orden = (
            prioridad_seccion,
            prioridad_periodo(resultado.get("periodo")),
            0 if resultado["tiene_cifras"] else 1,
            0 if not resultado["solo_guiones"] else 1,
            resultado["pagina"],
            resultado["numero_linea"],
        )
        actual = seleccion.get(campo)
        if actual is None or clave_orden < actual["clave_orden"]:
            seleccion[campo] = {"clave_orden": clave_orden, "resultado": resultado}

    return {campo: datos["resultado"] for campo, datos in seleccion.items()}


# ============================================================
# CONVERSIÓN NUMÉRICA Y REGLA DE VALOR FINAL
# ============================================================

def convertir_a_numero(valor_texto):
    """Formato chileno: '.' separador de miles, ',' decimal, paréntesis
    para negativos, '-' como cero."""
    if valor_texto is None:
        return None

    valor_texto = valor_texto.strip()
    if valor_texto == "-" or valor_texto == "":
        return 0

    es_negativo = valor_texto.startswith("(") and valor_texto.endswith(")")
    valor_limpio = valor_texto.strip("()")

    if "," in valor_limpio and "." in valor_limpio:
        valor_limpio = valor_limpio.replace(".", "").replace(",", ".")
    elif "," in valor_limpio:
        valor_limpio = valor_limpio.replace(",", ".")
    else:
        valor_limpio = valor_limpio.replace(".", "")

    valor_limpio = valor_limpio.replace("%", "").strip()
    if not valor_limpio:
        return None

    try:
        numero = float(valor_limpio)
    except ValueError:
        return None

    if es_negativo:
        numero = -numero
    if numero == int(numero):
        return int(numero)
    return numero


# Reglas de POSICIÓN para extraer el valor final de cada campo.
#
# ⚠ ADVERTENCIA (pendiente de resolver): estos índices están calibrados
# contra la geometría del formato 2025. La tabla de morosidad (Nota 13 n)
# tiene 13 columnas y el total antes de provisiones queda en el índice 5.
# En el formato antiguo (2019) la tabla de cartera tiene 7 columnas, por lo
# que el índice 5 NO apunta al mismo concepto. Cuando `columnas_esperadas`
# no calza, el script lo informa como advertencia en lugar de silenciarlo.
REGLAS_VALOR_FINAL = {
    "saldo_efectivo": {"indice": 5, "columnas_esperadas": 13},
    "consumo_90": {"indice": 5, "columnas_esperadas": 13},
    "provision_efectiva": {"indice": -1, "columnas_esperadas": 3},
    "castigo_acum": {"indice": -1, "columnas_esperadas": 3},
    "recupero_acum": {"indice": 0, "columnas_esperadas": None},
}

# Campos cuyo valor debe ser positivo por naturaleza contable. En las tablas
# de movimiento de provisiones los castigos aparecen entre paréntesis; el
# benchmark los registra en positivo.
CAMPOS_SIEMPRE_POSITIVOS = {
    "saldo_efectivo",
    "provision_efectiva",
    "castigo_acum",
    "recupero_acum",
    "consumo_90",
}


def mapear_columnas_por_anio(lineas, indice_fila, max_retroceso=MAX_RETROCESO_ENCABEZADO):
    """
    Busca hacia arriba, desde `indice_fila`, la primera línea que contenga
    dos o más años (encabezado comparativo de la tabla).

    Devuelve {posicion_columna: anio} o {} si no encuentra encabezado.
    """
    if indice_fila is None or not lineas:
        return {}

    limite_inferior = max(-1, indice_fila - max_retroceso)
    for i in range(indice_fila - 1, limite_inferior, -1):
        if i < 0 or i >= len(lineas):
            continue
        anios = re.findall(r"\b(?:19|20)\d{2}\b", lineas[i])
        if len(anios) >= 2:
            return {pos: int(a) for pos, a in enumerate(anios)}
    return {}


def _normalizar_signo(campo, valor):
    """En las tablas de movimiento los castigos y aplicaciones vienen entre
    paréntesis (negativos). El benchmark los registra en positivo."""
    if valor is None:
        return None
    if campo in CAMPOS_SIEMPRE_POSITIVOS:
        return abs(valor)
    return valor

def _concatenar_advertencia(existente, nueva):
    if not existente:
        return nueva
    if nueva in existente:
        return existente
    return f"{existente} | {nueva}"


def aplicar_validaciones_cruzadas(filas_detalle):
    """
    a) ANTI-DUPLICADO: si dos campos distintos resuelven al mismo valor, es
       casi seguro que una sección degradó hacia la de otro campo. Éste es
       el chequeo que habría detectado CMR 201912, donde `saldo_efectivo` y
       `castigo_acum` devolvían ambos -157.935.
    b) MAGNITUD: el saldo de cartera nunca es negativo, y la provisión debe
       ser una fracción plausible del saldo.
    """
    valores_por_campo = {f["campo"]: f["valor_final"] for f in filas_detalle}

    # --- a) anti-duplicado ---
    vistos = {}
    for fila in filas_detalle:
        valor = fila["valor_final"]
        if valor is None or valor == 0:
            continue
        if valor in vistos:
            campo_previo = vistos[valor]
            # Cada campo debe nombrar al OTRO, no a sí mismo.
            fila["advertencia"] = _concatenar_advertencia(
                fila.get("advertencia"),
                f"Valor idéntico al de '{campo_previo}' — probable colisión "
                f"de sección: ambos campos leyeron la misma tabla.",
            )
            for otra in filas_detalle:
                if otra["campo"] == campo_previo:
                    otra["advertencia"] = _concatenar_advertencia(
                        otra.get("advertencia"),
                        f"Valor idéntico al de '{fila['campo']}' — probable "
                        f"colisión de sección: ambos campos leyeron la misma "
                        f"tabla.",
                    )
        else:
            vistos[valor] = fila["campo"]

    # --- b) magnitud ---
    saldo = valores_por_campo.get("saldo_efectivo")
    provision = valores_por_campo.get("provision_efectiva")

    for fila in filas_detalle:
        if fila["campo"] in CAMPOS_SIEMPRE_POSITIVOS and fila["valor_final"] == 0:
            fila["advertencia"] = _concatenar_advertencia(
                fila.get("advertencia"),
                "Valor 0: probable columna con guion o posición equivocada.",
            )

    if saldo and provision and saldo > 0:
        indice_riesgo = provision / saldo
        if not (INDICE_RIESGO_MINIMO <= indice_riesgo <= INDICE_RIESGO_MAXIMO):
            mensaje = (
                f"Índice de riesgo implícito {indice_riesgo:.1%} fuera del "
                f"rango plausible ({INDICE_RIESGO_MINIMO:.0%}-"
                f"{INDICE_RIESGO_MAXIMO:.0%}): revisar saldo y/o provisión."
            )
            for fila in filas_detalle:
                if fila["campo"] in ("saldo_efectivo", "provision_efectiva"):
                    fila["advertencia"] = _concatenar_advertencia(
                        fila.get("advertencia"), mensaje
                    )

    return filas_detalle

def quitar_columna_nota(valores):
    """En el estado de resultados la primera columna es la referencia a la
    nota ('Provisiones por riesgo de crédito 21 (163.687) ...'). Un entero
    de 1-2 dígitos sin separador de miles al inicio es esa referencia."""
    if len(valores) >= 3 and re.fullmatch(r"\d{1,2}", valores[0]):
        return valores[1:]
    return valores


def resolver_fila_cartera(valores):
    """
    Fila tipo 'Subtotal' de composición de cartera o de morosidad.
    Estructura: N columnas de activos + N de provisiones + 1 neta = 2N+1.
    Se valida con la identidad saldo - provisión = neto; si no cuadra,
    la fila no corresponde y se devuelve None.
    Devuelve (saldo, provision, neto) en positivo.
    """
    total = len(valores)
    if total < 5 or total % 2 == 0:
        return None
    n = (total - 1) // 2
    saldo = convertir_a_numero(valores[n - 1])
    provision = convertir_a_numero(valores[2 * n - 1])
    neto = convertir_a_numero(valores[2 * n])
    if saldo is None or provision is None or neto is None:
        return None
    if saldo == 0 and provision == 0 and neto == 0:
        # Fila real pero sin actividad en esa categoría de cartera (ej. el
        # "Subtotal" de "Colocaciones comerciales" en un emisor que solo
        # opera consumo). saldo=provision=neto=0 siempre cumple la identidad
        # de forma trivial, así que sin este descarte una categoría vacía le
        # gana por prioridad de sección a la fila real que sí tiene datos
        # (bug detectado 2026-09-27 con CMR mar-2025: el "Subtotal" vacío de
        # comerciales ganaba sobre el de consumo, que sí tenía cifras).
        return None
    if abs(abs(saldo) - abs(provision) - neto) > 2:   # tolerancia de redondeo
        if total == 9:
            # SSFF (Nota 9): layout real es 3 sub-grupos de 3 columnas
            # (bruto normal/incumplimiento/total, provisión ídem, neto
            # ídem) en vez del genérico 2N+1 (N=4 bruto+prov, 1 neto) que
            # asume esta función. Los valores "totales" verdaderos están en
            # los índices 2, 5 y 8, no 3, 7, 8. Sólo se intenta como
            # fallback cuando la interpretación genérica no cuadra, para no
            # afectar a un eventual N=4 real de 9 columnas en otra entidad.
            saldo_alt = convertir_a_numero(valores[2])
            provision_alt = convertir_a_numero(valores[5])
            neto_alt = convertir_a_numero(valores[8])
            if saldo_alt is None or provision_alt is None or neto_alt is None:
                return None
            if saldo_alt == 0 and provision_alt == 0 and neto_alt == 0:
                return None
            if abs(abs(saldo_alt) - abs(provision_alt) - neto_alt) > 2:
                return None
            return abs(saldo_alt), abs(provision_alt), neto_alt
        return None
    return abs(saldo), abs(provision), neto


def resolver_fila_cartera_segmentada(valores):
    """
    Fila 'TOTAL' de la cartera segmentada por producto (SB Pay, Nota 9.3):
    5 columnas = Saldo(>0) + Saldo contingente + Provisión contingente +
    Provisión cartera normal + Provisión cartera en incumplimiento. A
    diferencia de `resolver_fila_cartera` no trae columna neta, así que no
    hay identidad contable con la que autovalidar la fila — sólo se exige
    que las 5 columnas sean numéricas y el saldo sea positivo.
    Devuelve (saldo, provision_normal_mas_incumplimiento).
    """
    if len(valores) != 5:
        return None
    numeros = [convertir_a_numero(v) for v in valores]
    if any(n is None for n in numeros):
        return None
    saldo = numeros[0]
    if saldo <= 0:
        return None
    provision = numeros[3] + numeros[4]
    return saldo, provision


def resolver_fila_movimiento(valores):
    """
    Tablas de movimiento (castigos, recuperos, provisiones constituidas).
    - Par   -> dos bloques (actual | comparativo): total del PRIMER bloque.
    - Impar -> un solo bloque: el total es la última columna.
    """
    total = len(valores)
    if total == 0:
        return None
    if total % 2 == 0:
        return convertir_a_numero(valores[total // 2 - 1])
    return convertir_a_numero(valores[-1])


def resolver_fila_cartera_dos_periodos(valores):
    """
    Fila TOTAL con 2 períodos lado a lado (actual + comparativo), cada uno
    con las columnas [N° clientes, Cartera bruta, Provisión, Cartera neta] y,
    en algunas entidades (Tricot), un 5to valor de factor de provisión al
    final del bloque. Distinto del genérico `resolver_fila_cartera`: aquí los
    2 períodos completos conviven en la misma fila en vez de compartir sólo
    la columna neta.
    - Tricot (Nota 6, "II. Cartera total"): 10 valores (bloques de 5).
    - Hites (Nota 10, "Cartera Total"): 8 valores (bloques de 4).
    Se toma el primer bloque (período actual) y se valida con la identidad
    bruto - provisión = neto. Devuelve (bruto, provision) o None.
    """
    total = len(valores)
    if total == 10:
        tamano_bloque = 5
    elif total == 8:
        tamano_bloque = 4
    else:
        return None
    bloque = valores[:tamano_bloque]
    bruto = convertir_a_numero(bloque[1])
    provision = convertir_a_numero(bloque[2])
    neto = convertir_a_numero(bloque[3])
    if bruto is None or provision is None or neto is None:
        return None
    if abs(abs(bruto) - abs(provision) - neto) > 2:
        return None
    # Corregido 2026-09-28: se vio en el PDF real de Tricot (no en el .md de
    # prueba) que un match falso de título puede caer en una fila donde
    # bruto == provisión (con neto ~0), lo cual es contablemente imposible
    # para una cartera real (la provisión nunca iguala ni supera el bruto).
    # Ver [[eeff-tricot-hites-ripley]] en memoria: pendiente de diagnóstico
    # exacto contra el PDF real, esta validación evita reportar un valor
    # incorrecto mientras tanto.
    if abs(bruto) <= abs(provision):
        return None
    return abs(bruto), abs(provision)


def resolver_fila_indicadores_riesgo(valores):
    """
    Fila TOTAL de "Indicadores de riesgo y castigo" (Hites, Nota 10): 6
    columnas de UN solo período (las tablas de cada fecha están separadas,
    no interleaved) — Provisión, Cartera bruta, Índice de riesgo %,
    Castigos, % castigos, % mensual castigos. Se exige provisión < cartera
    bruta como chequeo de sensatez (no hay columna neta con la que validar
    una identidad contable exacta). Devuelve (bruto, provision, castigo).
    """
    if len(valores) != 6:
        return None
    provision = convertir_a_numero(valores[0])
    bruto = convertir_a_numero(valores[1])
    castigo = convertir_a_numero(valores[3])
    if provision is None or bruto is None or castigo is None:
        return None
    if bruto <= 0 or provision < 0 or provision >= bruto:
        return None
    return abs(bruto), abs(provision), abs(castigo)


def resolver_fila_castigo_recupero_variacion(valores):
    """
    Tabla "a.5) Castigos" de Tricot: 4 columnas (actual, comparativo, var $,
    var %) — a diferencia de `resolver_fila_movimiento` (que asume 2 bloques
    actual+comparativo sin columnas de variación), aquí el total del período
    actual es siempre la PRIMERA columna, no la del medio.

    Algunas filas (ej. "Total recuperos deudas castigadas(1)") traen pegada
    una referencia de nota al pie entre paréntesis justo después de la
    etiqueta, que `extraer_valores_tabulares` no distingue de un valor real
    y agrega como una 5ta columna al principio — se descarta si aparece.
    """
    if len(valores) == 5 and re.fullmatch(r"\(\d{1,2}\)", valores[0]):
        valores = valores[1:]
    if len(valores) != 4:
        return None
    return convertir_a_numero(valores[0])


def resolver_fila_dos_columnas_simple(valores):
    """
    Fila de 2 columnas (actual, comparativo) sin identidad contable que
    autovalidar, ej. "Total recuperos del período" de Hites. Sólo exige que
    ambas columnas sean numéricas; se toma la primera (período actual).
    """
    if len(valores) != 2:
        return None
    numeros = [convertir_a_numero(v) for v in valores]
    if any(n is None for n in numeros):
        return None
    return numeros[0]


SECCIONES_TIPO_CARTERA = {"composicion_cartera", "morosidad_colocaciones"}

# SB Pay: fila "TOTAL" de la cartera segmentada por producto (Nota 9.3),
# ver `resolver_fila_cartera_segmentada`. Set aparte de SECCIONES_TIPO_
# CARTERA porque no comparte la estructura 2N+1 autovalidable.
SECCIONES_CARTERA_SEGMENTADA = {"cartera_segmentada_por_producto"}

# Tricot/Hites: fila TOTAL con 2 períodos lado a lado, ver
# `resolver_fila_cartera_dos_periodos`.
SECCIONES_CARTERA_DOS_PERIODOS = {"cartera_total_dos_periodos"}

# Hites: fila TOTAL de "Indicadores de riesgo y castigo", ver
# `resolver_fila_indicadores_riesgo`.
SECCIONES_INDICADORES_RIESGO = {"indicadores_riesgo_castigo"}

# Tricot: tabla "a.5) Castigos" (4 columnas con variación), ver
# `resolver_fila_castigo_recupero_variacion`.
SECCIONES_CASTIGO_RECUPERO_VARIACION = {"castigos_recuperos_variacion"}

# Hites: fila simple de 2 columnas "Total recuperos del período", ver
# `resolver_fila_dos_columnas_simple`.
SECCIONES_DOS_COLUMNAS_SIMPLE = {"recupero_castigo_simple_hites"}

# ⚠ Confirmado con Ricardo (2026-09-27) cruzando contra el Excel verificado:
# para la mayoría de las entidades (CMR) "Colocaciones consumo >=90dm" es el
# monto BRUTO de la fila "≥ 90 días" (saldo, antes de provisión). Para CAT en
# cambio el valor histórico corresponde a la columna de PROVISIÓN de esa
# misma fila, no al bruto — no se sabe si es una decisión deliberada o un
# artefacto del proceso manual, pero hay que replicarlo para no romper la
# serie histórica de CAT. Ampliar este set si aparece el mismo patrón en
# otra entidad nueva.
ENTIDADES_CONSUMO90_USA_PROVISION = {"CAT"}

# Entidades donde `consumo_90` ("Colocaciones consumo >=90dm") no viene en el
# PDF por ningún camino y se completa manualmente fuera del proceso
# automático (dato externo o suma manual de varios campos). Confirmado con
# Ricardo: SB Pay y Solventa (2026-09-27), Tricot (2026-09-28, la cifra que
# usa es una suma manual de varios campos, no un valor único en el PDF). Sirve
# sólo para dar un mensaje más claro en el resumen; no cambia el resultado
# (sigue siendo None, como antes).
ENTIDADES_CONSUMO90_MANUAL = {"SB Pay", "Solventa", "Tricot"}

# SSFF reporta sus EEFF en M$ (miles) mientras que la taxonomía canónica y el
# Excel de benchmark de Ricardo están en MM$ (millones) — mismo problema ya
# detectado para Solventa (ver [[eeff-sbpay-solventa-hallazgo]]), pero SSFF sí
# tiene reglas de extracción funcionando, así que aquí basta con convertir el
# valor final. Se aplica en `construir_tabla_final`, no en los resolutores,
# para no afectar la identidad contable (saldo-provisión=neto), que se valida
# en la unidad original del PDF.
ENTIDADES_EN_MILES = {"SSFF"}


def calcular_valor_final(campo, resultado, anio_actual=None, entidad=None):
    """
    Devuelve (valor_numerico, advertencia).

    Orden de resolución:
        1) Tabla de cartera   -> estructura simétrica 2N+1, autoverificada
        2) Tabla de movimiento -> bloques año actual / comparativo
        3) Índice fijo         -> fallback heredado, con advertencia

    `anio_actual` ya no se usa (el anclaje por encabezado quedó obsoleto),
    pero se mantiene en la firma porque `construir_tabla_final` lo pasa
    posicionalmente.
    """
    valores = resultado["valores_tabulares"]
    if not valores:
        return None, "La fila seleccionada no contiene valores."

    # Antes una fila de puros guiones se descartaba como "sin dato" (None).
    # Confirmado con Ricardo (2026-09-27): en el proceso manual, si la
    # sección/fila buscada SÍ se encuentra pero el valor viene en guion, se
    # registra como 0, no como dato faltante. Se deja que los resolutores de
    # abajo (que ya tratan "-" como 0 vía `convertir_a_numero`) lo resuelvan;
    # sólo se sigue devolviendo None cuando no hay ninguna fila candidata
    # (ver más arriba, "resultado is None" en `construir_tabla_final`).

    # En el estado de resultados la primera columna es la referencia a la
    # nota; hay que descartarla antes de contar posiciones.
    valores = quitar_columna_nota(valores)
    seccion = resultado.get("seccion")

    # --- 1) Tablas de cartera --------------------------------------------
    # Composición de cartera y morosidad comparten estructura:
    #   N columnas de activos + N de provisiones + 1 columna neta = 2N+1
    # La identidad saldo - provisión = neto valida que la fila sea la
    # correcta, así que no hace falta conocer el formato de antemano.
    if seccion in SECCIONES_TIPO_CARTERA:
        resuelto = resolver_fila_cartera(valores)
        if resuelto is not None:
            saldo, provision, _neto = resuelto
            if campo == "consumo_90" and entidad in ENTIDADES_CONSUMO90_USA_PROVISION:
                return provision, None
            if campo in ("saldo_efectivo", "consumo_90"):
                return saldo, None
            if campo == "provision_efectiva":
                return provision, None
        return None, (
            "Fila de cartera inválida: no cumple saldo - provisión = neto. "
            "Revisar manualmente."

)

    # --- 1b) Cartera segmentada por producto (SB Pay, Nota 9.3) ----------
    # Sin columna neta que autovalide la fila (ver
    # resolver_fila_cartera_segmentada); la provisión ya excluye la
    # provisión adicional CMF porque esta tabla sólo trae cartera normal +
    # incumplimiento, no la contingente (esa vive aparte en la Nota 22).
    if seccion in SECCIONES_CARTERA_SEGMENTADA:
        resuelto = resolver_fila_cartera_segmentada(valores)
        if resuelto is not None:
            saldo, provision = resuelto
            if campo == "saldo_efectivo":
                return saldo, None
            if campo == "provision_efectiva":
                return provision, None
        return None, (
            "Fila de cartera segmentada inválida: se esperaban 5 columnas "
            "numéricas con saldo positivo. Revisar manualmente."
        )

    # --- 1c) Cartera total con 2 períodos lado a lado (Tricot, Hites) -----
    if seccion in SECCIONES_CARTERA_DOS_PERIODOS:
        resuelto = resolver_fila_cartera_dos_periodos(valores)
        if resuelto is not None:
            bruto, provision = resuelto
            if campo == "saldo_efectivo":
                return bruto, None
            if campo == "provision_efectiva":
                return provision, None
        return None, (
            "Fila de cartera con 2 períodos inválida: se esperaban 8 o 10 "
            "columnas numéricas con bruto - provisión = neto. Revisar "
            "manualmente."
        )

    # --- 1d) Indicadores de riesgo y castigo (Hites, Nota 10) -------------
    if seccion in SECCIONES_INDICADORES_RIESGO:
        resuelto = resolver_fila_indicadores_riesgo(valores)
        if resuelto is not None:
            bruto, provision, castigo = resuelto
            if campo == "saldo_efectivo":
                return bruto, None
            if campo == "provision_efectiva":
                return provision, None
            if campo == "castigo_acum":
                return castigo, None
        return None, (
            "Fila de indicadores de riesgo inválida: se esperaban 6 "
            "columnas numéricas con provisión < cartera bruta. Revisar "
            "manualmente."
        )

    # --- 1e) Castigos/recuperos con columna de variación (Tricot) ---------
    if seccion in SECCIONES_CASTIGO_RECUPERO_VARIACION:
        valor = resolver_fila_castigo_recupero_variacion(valores)
        if valor is not None:
            return _normalizar_signo(campo, valor), None
        return None, (
            "Fila de castigos/recuperos con variación inválida: se "
            "esperaban 4 columnas numéricas. Revisar manualmente."
        )

    # --- 1f) Fila simple de 2 columnas sin identidad (Hites) --------------
    if seccion in SECCIONES_DOS_COLUMNAS_SIMPLE:
        valor = resolver_fila_dos_columnas_simple(valores)
        if valor is not None:
            return _normalizar_signo(campo, valor), None
        return None, (
            "Fila de 2 columnas inválida: se esperaban exactamente 2 "
            "columnas numéricas. Revisar manualmente."
        )

    # --- 2) Tablas de movimiento -----------------------------------------
    # Castigos, recuperos y provisiones constituidas. Si el número de
    # columnas es par, son dos bloques (actual | comparativo) y se toma el
    # total del primero; si es impar, hay un solo bloque y el total es la
    # última columna.
    if seccion in (
        "movimiento_provisiones_consumo",
        "gasto_perdidas_crediticias",
        "gasto_perdidas_detalle_por_cartera",
        "estado_resultados",
    ):
        valor = resolver_fila_movimiento(valores)
        if valor is not None:
            return _normalizar_signo(campo, valor), None

    # --- 3) Fallback: índice fijo ----------------------------------------
    regla = REGLAS_VALOR_FINAL.get(campo)
    if regla is None:
        return None, f"No hay regla de posición definida para '{campo}'."

    indice = regla["indice"]
    columnas_esperadas = regla["columnas_esperadas"]

    advertencia = (
        f"Sin regla estructural para la sección '{seccion}'; se usó el "
        f"índice fijo {indice} (calibrado para el formato 2025)."
    )
    if columnas_esperadas is not None and len(valores) != columnas_esperadas:
        advertencia += (
            f" Además se esperaban {columnas_esperadas} columnas y se "
            f"encontraron {len(valores)}: verificar manualmente."
        )

    try:
        valor_texto = valores[indice]
    except IndexError:
        return None, (
            f"Índice {indice} fuera de rango: la fila sólo tiene "
            f"{len(valores)} valores."
        )

    valor_numerico = convertir_a_numero(valor_texto)
    return _normalizar_signo(campo, valor_numerico), advertencia


def construir_tabla_final(resultados_pdf, anio_actual=None, entidad=None):
    seleccion = seleccionar_resultado_por_campo(resultados_pdf)
    filas_detalle = []

    for campo in PALABRAS_CLAVE.keys():
        resultado = seleccion.get(campo)
        if resultado is None:
            if campo == "consumo_90" and entidad in ENTIDADES_CONSUMO90_MANUAL:
                advertencia = (
                    "Campo manual/externo para esta entidad: no viene en el "
                    "PDF por ningún camino, se completa fuera del proceso "
                    "automático (no es un error de extracción)."
                )
            else:
                advertencia = "No se encontró ninguna coincidencia válida."
            filas_detalle.append({
                "campo": campo,
                "pagina": None,
                "linea": None,
                "valores_tabulares": [],
                "valor_final": None,
                "advertencia": advertencia,
            })
            continue

        valor_final, advertencia = calcular_valor_final(campo, resultado, anio_actual, entidad)
        if valor_final is not None and entidad in ENTIDADES_EN_MILES:
            # PDF en M$ (miles) -> taxonomía canónica en MM$ (millones).
            valor_final = round(valor_final / 1000, 3)
        filas_detalle.append({
            "campo": campo,
            "pagina": resultado["pagina"],
            "linea": resultado["linea"],
            "valores_tabulares": resultado["valores_tabulares"],
            "valor_final": valor_final,
            "advertencia": advertencia,
        })

    return aplicar_validaciones_cruzadas(filas_detalle)


# ============================================================
# PRESENTACIÓN
# ============================================================

def imprimir_tabla_final(filas_detalle, nombre, periodo, detalle=True):
    print()
    print("=" * 70)
    print("TABLA FINAL (formato compatible con Excel de comprobación)")
    print("=" * 70)

    valores_por_campo = {f["campo"]: f["valor_final"] for f in filas_detalle}

    print(f"{'Nombre':<28}: {nombre}")
    print(f"{'Periodo':<28}: {periodo}")
    print(f"{'Saldo Efectivo (MM$)':<28}: {valores_por_campo.get('saldo_efectivo')}")
    print(f"{'Provisión Efectiva (MM$)':<28}: {valores_por_campo.get('provision_efectiva')}")
    print(f"{'Castigo_Acum':<28}: {valores_por_campo.get('castigo_acum')}")
    print(f"{'Recupero_Acum':<28}: {valores_por_campo.get('recupero_acum')}")
    print(f"{'Colocaciones consumo >=90dm':<28}: {valores_por_campo.get('consumo_90')}")

    if not detalle:
        return

    print()
    print("-" * 70)
    print("DETALLE POR CAMPO (trazabilidad)")
    print("-" * 70)
    for fila in filas_detalle:
        print()
        print(f"Campo:          {fila['campo']}")
        print(f"Página:         {fila['pagina']}")
        print(f"Línea:          {fila['linea']}")
        print(f"Valores crudos: {fila['valores_tabulares']}")
        print(f"Valor final:    {fila['valor_final']}")
        if fila["advertencia"]:
            print(f"⚠ Advertencia: {fila['advertencia']}")


def imprimir_resumen_lote(resumen):
    print()
    print("=" * 78)
    print("RESUMEN DEL LOTE")
    print("=" * 78)
    print(f"{'Archivo':<42} {'Entidad':<12} {'Periodo':<9} {'Estado'}")
    print("-" * 78)
    for fila in resumen:
        print(
            f"{fila['archivo'][:41]:<42} "
            f"{str(fila['entidad'])[:11]:<12} "
            f"{str(fila['periodo']):<9} "
            f"{fila['estado']}"
        )

    total = len(resumen)
    completos = sum(1 for f in resumen if f["campos_ok"] == len(PALABRAS_CLAVE))
    parciales = sum(1 for f in resumen if 0 < f["campos_ok"] < len(PALABRAS_CLAVE))
    vacios = sum(1 for f in resumen if f["campos_ok"] == 0 and not f["error"])
    errores = sum(1 for f in resumen if f["error"])

    print("-" * 78)
    print(f"Total procesados : {total}")
    print(f"Completos        : {completos}")
    print(f"Parciales        : {parciales}")
    print(f"Sin extracción   : {vacios}")
    print(f"Con error        : {errores}")

    # --- Diagnóstico por campo -------------------------------------------
    # Permite distinguir un campo que falla SIEMPRE (problema de lógica o
    # ausencia estructural en el reporte) de uno que falla sólo en ciertos
    # períodos (problema de formato acotado a una cohorte).
    print()
    print("-" * 78)
    print("DIAGNÓSTICO POR CAMPO")
    print("-" * 78)
    print(f"{'Campo':<30} {'Faltante en':<13} {'Con advertencia':<16} Períodos que fallan")
    print("-" * 78)

    for campo in PALABRAS_CLAVE.keys():
        faltan = [
            f for f in resumen
            if not f["error"] and campo in f.get("campos_faltantes", [])
        ]
        advertidos = [
            f for f in resumen
            if not f["error"] and campo in f.get("campos_advertidos", [])
        ]
        periodos = [str(f["periodo"]) for f in faltan]
        if len(periodos) > 4:
            muestra = ", ".join(periodos[:4]) + f" … (+{len(periodos) - 4})"
        else:
            muestra = ", ".join(periodos) if periodos else "—"

        print(
            f"{campo:<30} "
            f"{len(faltan)}/{total:<11} "
            f"{len(advertidos)}/{total:<14} "
            f"{muestra}"
        )


# ============================================================
# PASO 5: EJECUCIÓN POR LOTE (MAIN)
# ============================================================
#
# El bug original estaba aquí: `ruta_pdf = archivos_pdf[0]` procesaba
# siempre el primer archivo alfabético y nunca iteraba. Ahora se recorre
# la lista completa y cada archivo se aísla en su propio try/except para
# que un PDF problemático no aborte el lote.

# Poner en True para ver la trazabilidad campo por campo de cada PDF.
# Con 22 archivos conviene dejarlo en False y usar el resumen final.
MOSTRAR_DETALLE_POR_ARCHIVO = False


def procesar_pdf(ruta_pdf):
    """Procesa un único PDF y devuelve un diccionario con el resultado.
    No captura excepciones: eso es responsabilidad de main()."""
    nombre_archivo = os.path.basename(ruta_pdf)
    paginas = leer_pdf(ruta_pdf)
    entidad, anio, periodo = identificar_entidad_y_periodo(ruta_pdf, paginas)

    if periodo is None:
        return {
            "archivo": nombre_archivo,
            "entidad": entidad,
            "anio": anio,
            "periodo": None,
            "paginas": len(paginas),
            "filas_detalle": [],
            "valores": {},
            "campos_ok": 0,
            "error": None,
            "estado": "⚠ No se pudo determinar el período",
            "campos_faltantes": list(PALABRAS_CLAVE.keys()),
            "campos_advertidos": [],
        }

    resultados_pdf = buscar_campos_en_pdf(paginas, anio)
    filas_detalle = construir_tabla_final(resultados_pdf, anio, entidad)
    valores = {f["campo"]: f["valor_final"] for f in filas_detalle}
    campos_ok = sum(1 for v in valores.values() if v is not None)

    campos_faltantes = [c for c, v in valores.items() if v is None]
    campos_advertidos = [
        f["campo"] for f in filas_detalle
        if f["advertencia"] and f["valor_final"] is not None
    ]
    estado = f"{campos_ok}/{len(PALABRAS_CLAVE)} campos"
    if campos_advertidos:
            estado += f"  | ⚠ {len(campos_advertidos)}"

    estado = f"{campos_ok}/{len(PALABRAS_CLAVE)} campos"
    if campos_faltantes:
        estado += f"  | falta: {', '.join(campos_faltantes)}"

    return {
        "archivo": nombre_archivo,
        "entidad": entidad,
        "anio": anio,
        "periodo": periodo,
        "paginas": len(paginas),
        "filas_detalle": filas_detalle,
        "valores": valores,
        "campos_ok": campos_ok,
        "campos_faltantes": campos_faltantes,
        "campos_advertidos": campos_advertidos,
        "error": None,
        "estado": estado,
    }


# ============================================================
# EXPORTACIÓN A EXCEL
# ============================================================
#
# Replica la estructura del Excel de benchmark de Ricardo: una hoja por
# entidad ("1. <Entidad>"), filas en orden cronológico ascendente (más
# antiguo arriba), columnas B-J con los datos extraídos y columnas A, G,
# I, K-R con las mismas fórmulas que ya usa su plantilla (Comprobación,
# Mora +90, Castigo_Marg, acumulados 12 meses, índices).
#
# Verificado contra `Benchmark - Operadoras de tarjetas - 202503_v2.xlsx`
# (hojas "1. CMR", "1. CAT", "1. SB Pay") — mismo layout de 17 columnas en
# las 3. Dos ajustes deliberados respecto al original, documentados para
# Ricardo (ver mensaje del hilo):
#   - Primera fila de cada entidad: sin fila previa no hay "marginal" que
#     calcular, así que Castigo_Marg/Recupero_Marg usan el valor acumulado
#     tal cual (=F/H), en vez de una corrección manual puntual como la que
#     tenía CMR en su primera fila.
#   - Acumulado 12 meses (columnas M/O): se deja en blanco hasta la 4ta
#     fila de la entidad (ventana completa), igual que hace el original en
#     la hoja CMR. La hoja CAT del original en cambio muestra sumas
#     parciales desde la 2da fila — no se replica esa variante para no
#     mostrar un "acum. 12 meses" que en realidad es de 1-2 trimestres.

ENCABEZADOS_HOJA_ENTIDAD = [
    None, "Nombre", "Periodo", "Saldo Efectivo (MM$)", "Provisión Efectiva (MM$)",
    "Castigo_Acum", "Comprobación", "Recupero_Acum", "Comprobación",
    "Colocaciones consumo >=90dm", "Mora +90", "Castigo_Marg",
    "Castigos acum 12 meses", "Recupero_Marg", "Recuperos acum 12 meses",
    "Castigo_Neto_acum_12m", "Indice_Riesgo", "Indice_Castigo_Acum",
]  # columnas A..R; A no lleva encabezado en la plantilla original.

# Orden de hojas del Excel de benchmark original de Ricardo (no alfabético).
# Una entidad nueva que no aparezca aquí se agrega al final, en el orden en
# que se detecte.
ORDEN_HOJAS_ORIGINAL = [
    "CMR", "CAT", "SSFF", "CAR", "Hites", "ILP", "AD retail",
    "abc (ILP+AD)", "Tricot", "Solventa", "Corona", "SB Pay",
]


def exportar_a_excel(resumen, ruta_salida):
    por_entidad = {}
    for fila in resumen:
        entidad = fila.get("entidad")
        periodo = fila.get("periodo")
        if not entidad or periodo is None or fila.get("error"):
            continue
        por_entidad.setdefault(entidad, []).append(fila)

    libro = openpyxl.Workbook()
    libro.remove(libro.active)

    def _orden_entidad(entidad):
        try:
            return (0, ORDEN_HOJAS_ORIGINAL.index(entidad))
        except ValueError:
            return (1, entidad)  # entidad desconocida: al final, alfabético entre sí

    for entidad in sorted(por_entidad, key=_orden_entidad):
        filas_entidad = sorted(por_entidad[entidad], key=lambda f: f["periodo"])
        nombre_hoja = f"1. {entidad}"[:31]  # Excel limita el título a 31 caracteres
        hoja = libro.create_sheet(nombre_hoja)

        for columna, encabezado in enumerate(ENCABEZADOS_HOJA_ENTIDAD, start=1):
            if encabezado is not None:
                hoja.cell(row=1, column=columna, value=encabezado)

        total_filas = len(filas_entidad)
        for indice, fila in enumerate(filas_entidad):
            r = indice + 2
            valores = fila.get("valores", {})
            es_primera = indice == 0
            es_ultima = indice == total_filas - 1
            tiene_ventana_12m = indice >= 3  # 4ta fila de la entidad en adelante

            hoja.cell(row=r, column=1, value=f"=ROUND(C{r}/100,0)")           # A Año
            hoja.cell(row=r, column=2, value=entidad)                        # B Nombre
            hoja.cell(row=r, column=3, value=int(fila["periodo"]))           # C Periodo

            if valores.get("saldo_efectivo") is not None:
                hoja.cell(row=r, column=4, value=valores["saldo_efectivo"])
            if valores.get("provision_efectiva") is not None:
                hoja.cell(row=r, column=5, value=valores["provision_efectiva"])
            if valores.get("castigo_acum") is not None:
                hoja.cell(row=r, column=6, value=valores["castigo_acum"])
            if not es_ultima:
                hoja.cell(row=r, column=7, value=f"=F{r}<F{r + 1}")           # G
            if valores.get("recupero_acum") is not None:
                hoja.cell(row=r, column=8, value=valores["recupero_acum"])
            if not es_ultima:
                hoja.cell(row=r, column=9, value=f"=H{r}<H{r + 1}")           # I
            if valores.get("consumo_90") is not None:
                hoja.cell(row=r, column=10, value=valores["consumo_90"])
                hoja.cell(row=r, column=11, value=f"=J{r}/D{r}")              # K Mora +90

            if es_primera:
                hoja.cell(row=r, column=12, value=f"=F{r}")                  # L
                hoja.cell(row=r, column=14, value=f"=H{r}")                  # N
            else:
                hoja.cell(
                    row=r, column=12,
                    value=f"=IF(A{r}<>A{r - 1},F{r},F{r}-F{r - 1})",
                )
                hoja.cell(
                    row=r, column=14,
                    value=f"=IF(A{r}<>A{r - 1},H{r},H{r}-H{r - 1})",
                )

            if tiene_ventana_12m:
                hoja.cell(row=r, column=13, value=f"=SUM(L{r - 3}:L{r})")     # M
                hoja.cell(row=r, column=15, value=f"=SUM(N{r - 3}:N{r})")     # O
                hoja.cell(row=r, column=16, value=f"=M{r}-O{r}")              # P

            hoja.cell(row=r, column=17, value=f"=E{r}/D{r}")                 # Q
            if tiene_ventana_12m:
                hoja.cell(row=r, column=18, value=f"=P{r}/D{r}")              # R

        for columna in hoja.columns:
            ancho = max(
                (len(str(c.value)) for c in columna if c.value is not None),
                default=0,
            )
            hoja.column_dimensions[columna[0].column_letter].width = min(
                max(ancho + 2, 10), 32
            )

    if not libro.sheetnames:
        hoja = libro.create_sheet("Resumen")
        hoja.append(["Sin resultados: revisar errores en el procesamiento por lote."])

    libro.save(ruta_salida)


def main():
    print("=" * 78)
    print("EXTRACTOR DE ESTADOS FINANCIEROS - PROCESAMIENTO POR LOTE")
    print("=" * 78)

    archivos_pdf = obtener_archivos_pdf(CARPETA_PDFS)
    print(f"\nCarpeta       : {CARPETA_PDFS}")
    print(f"PDF encontrados: {len(archivos_pdf)}")

    if not archivos_pdf:
        print("No se encontraron archivos PDF en la carpeta.")
        return

    resumen = []

    for indice, ruta_pdf in enumerate(archivos_pdf, start=1):
        nombre_archivo = os.path.basename(ruta_pdf)
        print(f"\n[{indice}/{len(archivos_pdf)}] {nombre_archivo}")

        try:
            resultado = procesar_pdf(ruta_pdf)
        except Exception as error:  # aislamiento por archivo
            print(f"  ✖ {type(error).__name__}: {error}")
            resumen.append({
                "archivo": nombre_archivo,
                "entidad": None,
                "anio": None,
                "periodo": None,
                "campos_ok": 0,
                "campos_faltantes": list(PALABRAS_CLAVE.keys()),
                "campos_advertidos": [],
                "error": error,
                "estado": f"ERROR: {type(error).__name__}",
            })
            continue

        print(
            f"  Entidad: {resultado['entidad']} | "
            f"Periodo: {resultado['periodo']} | "
            f"Páginas: {resultado['paginas']} | "
            f"{resultado['estado']}"
        )

        if resultado["filas_detalle"]:
            imprimir_tabla_final(
                resultado["filas_detalle"],
                resultado["entidad"],
                resultado["periodo"],
                detalle=MOSTRAR_DETALLE_POR_ARCHIVO,
            )

        resumen.append(resultado)

    imprimir_resumen_lote(resumen)

    print()
    try:
        exportar_a_excel(resumen, ARCHIVO_SALIDA)
        print(f"Exportado a Excel: {ARCHIVO_SALIDA}")
    except Exception as error:
        print(f"✖ No se pudo exportar a Excel: {type(error).__name__}: {error}")


if __name__ == "__main__":
    main()
