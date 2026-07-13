import re
from pypdf import PdfReader

try:
    import fitz
except Exception:
    fitz = None


# =====================================================
# LIMPIEZA GENERAL
# =====================================================

def limpiar_texto(valor):
    if valor is None:
        return ""

    valor = str(valor)
    valor = valor.replace("\ufeff", "")
    valor = valor.replace("￾", "")
    valor = valor.replace("\r", "\n")
    valor = re.sub(r"-\s*\n\s*", "", valor)
    valor = re.sub(r"\s*\n\s*", " ", valor)
    valor = re.sub(r"\bFACSIMIL\s+SAG\b", "", valor, flags=re.IGNORECASE)
    valor = re.sub(r"\s+", " ", valor)
    return valor.strip(" .;:\n\t")


def limpiar_celda_tabla(valor):
    return limpiar_texto(valor)


def sin_acentos(texto):
    cambios = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
        "ñ": "n", "Ñ": "N", "ü": "u", "Ü": "U"
    }
    for a, b in cambios.items():
        texto = texto.replace(a, b)
    return texto


def normalizar(texto):
    return sin_acentos(limpiar_texto(texto)).lower()


def unicos_en_orden(valores):
    salida = []
    vistos = set()

    for valor in valores:
        limpio = limpiar_texto(valor)
        if not limpio:
            continue
        clave = limpio.casefold()
        if clave not in vistos:
            vistos.add(clave)
            salida.append(limpio)

    return salida


def texto_para_etiqueta(texto):
    """Usa la etiqueta SAG y evita que la HDS ensucie la lectura."""
    if not texto:
        return ""

    texto = texto.replace("￾", "")
    cortes = [
        "Hoja de Seguridad",
        "SECCIÓN 1:",
        "SECCION 1:",
        "IDENTIFICACIÓN DE LA SUSTANCIA",
        "IDENTIFICACION DE LA SUSTANCIA",
        "En conformidad con la Regulación Chile NCh",
    ]

    posiciones = []
    for corte in cortes:
        pos = texto.find(corte)
        if pos > 400:
            posiciones.append(pos)

    if posiciones:
        return texto[:min(posiciones)]

    return texto


# =====================================================
# LECTURA PDF
# =====================================================

def extraer_texto_pdf(path):
    reader = PdfReader(path)
    texto = ""
    for pagina in reader.pages:
        contenido = pagina.extract_text()
        if contenido:
            texto += contenido + "\n"
    return texto


# =====================================================
# TABLAS DE USO
# =====================================================

def es_texto_dosis(valor):
    t = normalizar(valor)
    if not t:
        return False

    if re.search(r"\b\d+\s*(?:-|a|–)\s*\d+", t) and re.search(r"\b(g|kg|cc|ml|l)\b|/ha|100\s*l|agua", t):
        return True

    if re.search(r"\b(g|kg|cc|ml|l)\s*/\s*(ha|100\s*l|100l)", t):
        return True

    if t.startswith("g/100") or t.startswith("kg/ha") or t.startswith("cc/ha"):
        return True

    return False


def problema_valido(valor):
    t = normalizar(valor)
    if not t:
        return False
    if es_texto_dosis(t):
        return False
    if len(t) < 3:
        return False
    palabras_invalidas = [
        "observaciones", "dosis", "cultivo", "preparacion", "periodo de carencia",
        "tiempo de reingreso", "aplicar cuando", "mojar con", "realizar maximo"
    ]
    if any(p in t for p in palabras_invalidas):
        return False
    return True


def encabezado_tabla_uso(fila):
    celdas = [normalizar(c) for c in fila]
    texto = " ".join(celdas)
    tiene_cultivo = any("cultivo" in c for c in celdas)
    tiene_dosis = any("dosis" in c for c in celdas)
    tiene_problema = any(p in texto for p in ["plaga", "plagas", "enfermedad", "enfermedades", "maleza", "malezas", "objetivo"])
    return tiene_cultivo and tiene_dosis and tiene_problema


def indices_tabla_uso(encabezado):
    indices = {"cultivo": None, "problema": None, "dosis": None, "observaciones": None}

    for i, celda in enumerate(encabezado):
        c = normalizar(celda)
        if "cultivo" in c:
            indices["cultivo"] = i
        elif any(p in c for p in ["plaga", "enfermedad", "maleza", "objetivo"]):
            indices["problema"] = i
        elif "dosis" in c:
            indices["dosis"] = i
        elif "observ" in c:
            indices["observaciones"] = i

    return indices


def obtener_celda(fila, indice):
    if indice is None or indice >= len(fila):
        return ""
    return limpiar_celda_tabla(fila[indice])


def extraer_usos_pdf(path):
    usos = []
    if fitz is None:
        return usos

    try:
        documento = fitz.open(path)
    except Exception:
        return usos

    for numero_pagina, pagina in enumerate(documento, start=1):
        try:
            tablas = pagina.find_tables().tables
        except Exception:
            tablas = []

        for tabla in tablas:
            try:
                filas = tabla.extract()
            except Exception:
                continue

            if not filas:
                continue

            encabezado_indice = None
            for idx, fila in enumerate(filas[:5]):
                if encabezado_tabla_uso(fila):
                    encabezado_indice = idx
                    break

            if encabezado_indice is None:
                continue

            indices = indices_tabla_uso(filas[encabezado_indice])
            cultivo_actual = ""
            problema_actual = ""
            dosis_actual = ""
            observacion_actual = ""

            for fila in filas[encabezado_indice + 1:]:
                cultivo = obtener_celda(fila, indices["cultivo"])
                problema = obtener_celda(fila, indices["problema"])
                dosis = obtener_celda(fila, indices["dosis"])
                observaciones = obtener_celda(fila, indices["observaciones"])

                if cultivo:
                    cultivo_actual = cultivo
                else:
                    cultivo = cultivo_actual

                if problema:
                    problema_actual = problema
                else:
                    problema = problema_actual

                if dosis and re.search(r"\d", dosis):
                    dosis_actual = dosis
                else:
                    dosis = dosis_actual

                if observaciones:
                    observacion_actual = observaciones
                else:
                    observaciones = observacion_actual

                if not cultivo or not problema or not dosis:
                    continue
                if not re.search(r"\d", dosis):
                    continue
                if not problema_valido(problema):
                    continue
                if normalizar(cultivo).startswith("raleo"):
                    continue

                usos.append({
                    "cultivo": cultivo,
                    "problema": problema,
                    "dosis": dosis,
                    "observaciones": observaciones,
                    "pagina": numero_pagina
                })

    return usos


# =====================================================
# DETECCIÓN BASE
# =====================================================

def limpiar_nombre_archivo(nombre_archivo):
    nombre = nombre_archivo.replace(".pdf", "").replace(".PDF", "")
    nombre = nombre.replace("_", " ").replace("-", " ")
    nombre = re.sub(r"\s+", " ", nombre)
    return nombre.strip().upper()


def detectar_producto(texto, nombre_archivo=""):
    etiqueta = texto_para_etiqueta(texto)

    patrones = [
        r"(?i)PRODUCTO\s*:\s*([A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9®™% .\-]{3,60})",
        r"(?i)Nombre\s+Comercial\s*:\s*([A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9®™% .\-]{3,60})",
        r"(?im)^\s*([A-ZÁÉÍÓÚÜÑ][A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9®™% .\-]{2,50})\s*$\s*(?:Herbicida|Fungicida|Insecticida|Acaricida)",
    ]

    for patron in patrones:
        m = re.search(patron, etiqueta)
        if m:
            producto = limpiar_texto(m.group(1))
            producto = producto.replace("®", "").replace("™", "")
            producto = re.sub(r"\s+", " ", producto).strip(" .:-")
            producto = re.sub(r"\bETIQUETA\b.*$", "", producto, flags=re.IGNORECASE).strip()
            if producto and producto.lower() not in ["producto", "composicion", "instrucciones de uso"]:
                return producto.upper()

    if nombre_archivo:
        return limpiar_nombre_archivo(nombre_archivo)

    m = re.search(r"(?im)^\s*([A-Z0-9][A-Z0-9®™% .\-]{3,45})\s*$\s*(?:INSECTICIDA|HERBICIDA|FUNGICIDA|ACARICIDA)", etiqueta)
    if m:
        return limpiar_texto(m.group(1)).replace("®", "").upper()

    return "PRODUCTO NO IDENTIFICADO"


def extraer_bloque_composicion(texto):
    etiqueta = texto_para_etiqueta(texto)
    m = re.search(
        r"(?i)(?:COMPOSICI[ÓO]N|Composici[óo]n)\s*[:.]\s*(.+?)(?=Autorizaci[óo]n|LEA ATENTAMENTE|NO INFLAMABLE|INSTRUCCIONES DE USO|RUKARB|Adengo|$)",
        etiqueta,
        flags=re.DOTALL,
    )
    if m:
        return m.group(1)
    return etiqueta


def detectar_ingrediente(texto):
    bloque = extraer_bloque_composicion(texto)
    bloque = bloque.replace("\n", " ")
    bloque = re.sub(r"\s+", " ", bloque)
    bloque = re.sub(r"\.{2,}", " ", bloque)

    patrones = [
        r"([A-ZÁÉÍÓÚÜÑa-záéíóúüñ][A-ZÁÉÍÓÚÜÑa-záéíóúüñ0-9\- ]{2,55}?)(?:\s*\([^)]*\)|\s*[*]+)?\s+(\d+[,.]?\d*)\s*%\s*(p/v|p/p|v/v|p.c.)\s*\(([^)]*(?:g|mg|kg|L|l|ml)[^)]*)\)",
        r"([A-ZÁÉÍÓÚÜÑa-záéíóúüñ][A-ZÁÉÍÓÚÜÑa-záéíóúüñ0-9\- ]{2,55}?)(?:\s*\([^)]*\)|\s*[*]+)?\s+(\d+[,.]?\d*)\s*%\s*(p/v|p/p|v/v|p.c.)",
    ]

    resultados = []
    descartes = ["coformulantes", "contenido neto", "producto", "herbicida", "fungicida", "insecticida", "composicion", "composición"]

    for patron in patrones:
        for m in re.finditer(patron, bloque, flags=re.IGNORECASE):
            nombre = limpiar_texto(m.group(1))
            nombre = re.sub(r"^.*?(?=[A-ZÁÉÍÓÚÜÑa-záéíóúüñ][a-záéíóúüñ]+)", "", nombre).strip()
            porcentaje = m.group(2).replace(".", ",")
            unidad = m.group(3)
            concentracion = limpiar_texto(m.group(4)) if len(m.groups()) >= 4 and m.group(4) else ""

            nombre_min = normalizar(nombre)
            if any(d in nombre_min for d in descartes):
                continue
            if len(nombre) < 3:
                continue

            if concentracion:
                item = f"{nombre} {porcentaje} % {unidad} ({concentracion})"
            else:
                item = f"{nombre} {porcentaje} % {unidad}"

            resultados.append(item)

    resultados = unicos_en_orden(resultados)
    if resultados:
        return "; ".join(resultados[:6])

    # Respaldo para HDS: nombre de sustancia activa : Carbarilo
    m = re.search(r"(?i)Nombre\s+de\s+la\s+sustancia\s+activa\s*:\s*([A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9\- ]{3,60})", texto)
    if m:
        return limpiar_texto(m.group(1))

    return ""


def detectar_grupo(texto):
    etiqueta = texto_para_etiqueta(texto)
    etiqueta_norm = normalizar(etiqueta)
    ingrediente_norm = normalizar(detectar_ingrediente(etiqueta))

    grupos = []

    for m in re.finditer(
        r"(?i)pertenece\s+al\s+grupo\s+qu[íi]mico\s+(?:de\s+los?|de\s+las?)?\s*([A-Za-zÁÉÍÓÚÜÑáéíóúüñ\-]+)",
        etiqueta,
    ):
        grupos.append(m.group(1))

    for m in re.finditer(r"(?i)\b(IRAC|FRAC|HRAC)\s*(?:grupo|Group|N°|No|:)?\s*([0-9A-Za-z]+)", etiqueta):
        grupos.append(f"{m.group(1).upper()} {m.group(2).upper()}")

    if "isoxaflutol" in ingrediente_norm:
        grupos.insert(0, "HRAC F2")
    if "tiencarbazona" in ingrediente_norm:
        grupos.insert(1, "HRAC B")
    if "carbaril" in ingrediente_norm or "carbaryl" in ingrediente_norm or "carbarilo" in ingrediente_norm:
        grupos.insert(0, "IRAC 1A")
        grupos.append("Carbamatos")

    grupos = unicos_en_orden(grupos)
    return "; ".join(grupos)


def detectar_tipo(texto, grupo="", ingrediente=""):
    etiqueta = texto_para_etiqueta(texto)
    t = normalizar(etiqueta)
    g = normalizar(grupo)
    ing = normalizar(ingrediente)

    if "insecticida" in t or "irac" in g or "pirimicarb" in ing or "carbaril" in ing or "carbaryl" in ing or "imidacloprid" in ing:
        return "Insecticida"
    if "herbicida" in t or "hrac" in g or "isoxaflutol" in ing or "tiencarbazona" in ing:
        return "Herbicida"
    if "fungicida" in t or "frac" in g or "boscalid" in ing:
        return "Fungicida"
    if "acaricida" in t:
        return "Acaricida"
    return ""


def detectar_cultivos(texto):
    etiqueta = texto_para_etiqueta(texto)
    t = normalizar(etiqueta)
    cultivos = {
        "maíz grano y silo": ["maiz grano y silo", "maiz grano", "maiz"],
        "manzano": ["manzano"],
        "peral": ["peral"],
        "membrillero": ["membrillero"],
        "nectarino": ["nectarino"],
        "duraznero": ["duraznero", "durazno"],
        "damasco": ["damasco"],
        "ciruelo": ["ciruelo"],
        "cerezo": ["cerezo"],
        "guindo": ["guindo"],
        "almendro": ["almendro"],
        "nogal": ["nogal"],
        "frambuesa": ["frambuesa", "frambueso"],
        "arándano": ["arandano", "arándano"],
        "vid": ["vid"],
        "kiwi": ["kiwi"],
        "ajo": ["ajo"],
        "cebolla": ["cebolla"],
        "frejol": ["frejol"],
        "papa": ["papa"],
        "tomate": ["tomate"],
        "pimentón": ["pimenton", "pimentón"],
        "arveja": ["arveja"],
        "brócoli": ["brocoli", "brócoli"],
        "lechuga": ["lechuga"],
        "espárrago": ["esparrago", "espárrago"],
        "coliflor": ["coliflor"],
        "pepino": ["pepino"],
        "zapallo": ["zapallo"],
        "remolacha": ["remolacha"],
        "maravilla": ["maravilla"],
        "soya": ["soya"],
        "raps": ["raps"],
    }

    encontrados = []
    for cultivo, palabras in cultivos.items():
        if any(p in t for p in palabras):
            encontrados.append(cultivo)
    return ", ".join(unicos_en_orden(encontrados))


def detectar_lista_por_diccionario(texto, diccionario):
    t = normalizar(texto_para_etiqueta(texto))
    encontrados = []
    for nombre, palabras in diccionario.items():
        if any(normalizar(p) in t for p in palabras):
            encontrados.append(nombre)
    return ", ".join(unicos_en_orden(encontrados))


def detectar_malezas(texto):
    dic = {
        "Hualcacho": ["hualcacho"],
        "Pata de gallina": ["pata de gallina", "digitaria sanguinalis"],
        "Maicillo de semilla": ["maicillo de semilla", "sorghum halepense"],
        "Pega pega": ["pega pega", "setaria verticilata"],
        "Ballica": ["ballica", "lolium"],
        "Pasto quila": ["pasto quila", "agrostis capillaris"],
        "Verdolaga": ["verdolaga", "portulaca"],
        "Malvilla": ["malvilla", "anoda cristata"],
        "Chamico": ["chamico", "datura stramonium"],
        "Quingüilla": ["quinguilla", "quingüilla", "quenopodium"],
        "Rábano": ["rabano", "rábano", "raphanus"],
        "Clonqui": ["clonqui", "xanthium"],
        "Ambrosia": ["ambrosia"],
        "Tomatillo": ["tomatillo", "solanum nigrum"],
        "Sanguinaria": ["sanguinaria", "polygonum aviculare"],
        "Duraznillo": ["duraznillo", "polygonum persicaria"],
        "Suspiro": ["suspiro", "ipomoea"],
        "Siete venas": ["siete venas", "plantago"],
        "Hierba del chancho": ["hierba del chancho", "hypochoeris"],
        "Bledo": ["bledo", "amaranthus"],
    }
    return detectar_lista_por_diccionario(texto, dic)


def detectar_insectos(texto):
    dic = {
        "Polilla de la manzana": ["polilla de la manzana"],
        "Grafolita": ["grafolita"],
        "Langostinos": ["langostino", "langostinos"],
        "Gusano de los penachos": ["gusano de los penachos"],
        "Burrito": ["burrito"],
        "Trips": ["trips"],
        "Eulia": ["eulia"],
        "Chape": ["chape"],
        "Chanchito blanco": ["chanchito blanco"],
        "Cuncunilla": ["cuncunilla"],
        "Capachito de los frutales": ["capachito de los frutales"],
        "Gusano del tomate": ["gusano del tomate"],
        "Pulgón": ["pulgon", "pulgón", "pulgones", "afidos", "áfidos"],
        "Mosca blanca": ["mosca blanca"],
    }
    return detectar_lista_por_diccionario(texto, dic)


def detectar_enfermedades(texto):
    dic = {
        "Botritis": ["botritis", "botrytis"],
        "Oídio": ["oidio", "oídio"],
        "Tizón": ["tizon", "tizón"],
        "Alternaria": ["alternaria"],
        "Mildiu": ["mildiu"],
        "Esclerotinia": ["sclerotinia", "esclerotinia"],
    }
    return detectar_lista_por_diccionario(texto, dic)


def detectar_dosis(texto):
    etiqueta = texto_para_etiqueta(texto)
    t = normalizar(etiqueta)
    patrones = [
        r"\b\d+(?:[,.]\d+)?\s*(?:-|a|–)\s*\d+(?:[,.]\d+)?\s*(?:kg|g|l|ml|cc)\s*/\s*(?:ha|100\s*l|100l)\b",
        r"\b\d+(?:[,.]\d+)?\s*(?:kg|g|l|ml|cc)\s*/\s*(?:ha|100\s*l|100l)\b",
        r"\b\d+(?:[,.]\d+)?\s*(?:-|a|–)\s*\d+(?:[,.]\d+)?\b(?=\s*(?:cc/ha|g/100|kg/ha))",
    ]
    resultados = []
    for patron in patrones:
        for m in re.finditer(patron, t):
            resultados.append(limpiar_texto(m.group(0)))
    return ", ".join(unicos_en_orden(resultados)[:10])


def detectar_bloque(texto, inicio, finales, max_caracteres=900):
    etiqueta = texto_para_etiqueta(texto)
    final_patron = "|".join(finales)
    patron = rf"(?i){inicio}\s*:\s*(.+?)(?=\s*(?:{final_patron})\s*:|$)"
    m = re.search(patron, etiqueta, flags=re.DOTALL)
    if not m:
        return ""
    valor = limpiar_texto(m.group(1))
    return valor[:max_caracteres].strip()


def detectar_compatibilidad(texto):
    return detectar_bloque(texto, r"Compatibilidades?|Compatibilidad", ["Incompatibilidades?", "Fitotoxicidad", "Periodo de Carencia", "Período de Carencia", "Tiempo de Reingreso"], 700)


def detectar_incompatibilidad(texto):
    return detectar_bloque(texto, r"Incompatibilidades?|Incompatibilidad", ["Compatibilidades?", "Fitotoxicidad", "Periodo de Carencia", "Período de Carencia", "Tiempo de Reingreso"], 500)


def detectar_fitotoxicidad(texto):
    return detectar_bloque(texto, r"Fitotoxicidad", ["Periodo de Carencia", "Período de Carencia", "Tiempo de Reingreso", "Preparación de la mezcla"], 700)


def detectar_reingreso(texto):
    etiqueta = texto_para_etiqueta(texto)
    m = re.search(r"(?i)Tiempo\s+de\s+Reingreso(?:\s+al\s+[áa]rea\s+tratada)?\s*:\s*(.+?)(?=\s*(?:Periodo de Carencia|Período de Carencia|Fitotoxicidad|$))", etiqueta, flags=re.DOTALL)
    if m:
        return limpiar_texto(m.group(1))[:600]

    m = re.search(r"(?i)no\s+deben\s+ingresar\s+al\s+[áa]rea\s+tratada\s+antes\s+de\s+([0-9]+\s*horas)", etiqueta)
    if m:
        return m.group(1)

    m = re.search(r"(?i)reingreso[^.]{0,180}?([0-9]+\s*horas)", etiqueta)
    if m:
        return m.group(1)

    return ""


def detectar_carencia(texto):
    etiqueta = texto_para_etiqueta(texto)
    m = re.search(r"(?i)(?:Periodo|Período)\s+de\s+Carencia\s*\(?d[íi]as\)?\s*:\s*(.+?)(?=\s*(?:Tiempo de Reingreso|Fitotoxicidad|Preparación|$))", etiqueta, flags=re.DOTALL)
    if m:
        return limpiar_texto(m.group(1))[:700]

    m = re.search(r"(?i)(?:Periodo|Período)\s+de\s+Carencia\s*:\s*([^.]*(?:d[íi]as?)?)", etiqueta)
    if m:
        return limpiar_texto(m.group(1))

    return ""


def detectar_toxicidad_abejas(texto):
    t = normalizar(texto_para_etiqueta(texto))
    patrones = [
        ("muy toxico para abejas", "Muy tóxico para abejas"),
        ("muy toxico para las abejas", "Muy tóxico para abejas"),
        ("altamente toxico para abejas", "Altamente tóxico para abejas"),
        ("practicamente no toxico para abejas", "Prácticamente no tóxico para abejas"),
        ("practicamente no toxico para las abejas", "Prácticamente no tóxico para abejas"),
        ("virtualmente no toxico para abejas", "Virtualmente no tóxico para abejas"),
        ("no es peligroso para las abejas", "No es peligroso para las abejas"),
        ("no toxico para abejas", "No tóxico para abejas"),
    ]
    for patron, resultado in patrones:
        if patron in t:
            return resultado
    if "toxico para abejas" in t:
        return "Tóxico para abejas"
    return ""


# =====================================================
# CALIDAD DE EXTRACCIÓN
# =====================================================

def calidad_extraccion(resultado):
    faltantes = []
    importantes = ["producto", "ingrediente", "tipo", "grupo", "toxicidad_abejas"]
    for campo in importantes:
        if not limpiar_texto(resultado.get(campo, "")):
            faltantes.append(campo)

    if len(faltantes) == 0:
        estado = "confiable"
        mensaje = "✅ Extracción confiable"
    elif len(faltantes) <= 2:
        estado = "revisar"
        mensaje = "⚠️ Revisar antes de guardar: faltan " + ", ".join(faltantes)
    else:
        estado = "no_guardar"
        mensaje = "❌ No guardar todavía: faltan " + ", ".join(faltantes)

    return {"estado": estado, "mensaje": mensaje, "faltantes": faltantes}


# =====================================================
# ANÁLISIS PRINCIPAL
# =====================================================

def analizar_texto(texto, nombre_archivo=""):
    etiqueta = texto_para_etiqueta(texto)
    ingrediente = detectar_ingrediente(etiqueta)
    grupo = detectar_grupo(etiqueta)
    tipo = detectar_tipo(etiqueta, grupo, ingrediente)

    malezas = detectar_malezas(etiqueta)
    insectos = detectar_insectos(etiqueta)
    enfermedades = detectar_enfermedades(etiqueta)

    # En el formulario actual, los herbicidas usan el campo enfermedades
    # pero con etiqueta visual "Malezas controladas".
    if "herbicida" in normalizar(tipo):
        enfermedades_formulario = malezas
        insectos_formulario = ""
    elif "insecticida" in normalizar(tipo):
        enfermedades_formulario = ""
        insectos_formulario = insectos
    elif "fungicida" in normalizar(tipo):
        enfermedades_formulario = enfermedades
        insectos_formulario = ""
    else:
        enfermedades_formulario = enfermedades or malezas
        insectos_formulario = insectos

    resultado = {
        "producto": detectar_producto(etiqueta, nombre_archivo),
        "ingrediente": ingrediente,
        "grupo": grupo,
        "tipo": tipo,
        "cultivos": detectar_cultivos(etiqueta),
        "enfermedades": enfermedades_formulario,
        "insectos": insectos_formulario,
        "dosis": detectar_dosis(etiqueta),
        "compatibilidad": detectar_compatibilidad(etiqueta),
        "incompatibilidad": detectar_incompatibilidad(etiqueta),
        "fitotoxicidad": detectar_fitotoxicidad(etiqueta),
        "reingreso": detectar_reingreso(etiqueta),
        "carencia": detectar_carencia(etiqueta),
        "toxicidad_abejas": detectar_toxicidad_abejas(etiqueta),
        "texto": texto,
    }

    resultado["calidad"] = calidad_extraccion(resultado)
    resultado["estado_calidad"] = resultado["calidad"]["estado"]
    resultado["mensaje_calidad"] = resultado["calidad"]["mensaje"]

    return resultado


# =====================================================
# POSTPROCESO GENERAL PARA CARGA MASIVA
# Corrige campos cruzados según tipo de producto
# =====================================================

_analizar_texto_pre_postproceso = analizar_texto


def limpiar_problema_si_es_dosis(valor):
    valor_limpio = limpiar_texto(valor)

    if not valor_limpio:
        return ""

    valor_min = valor_limpio.lower()

    patrones_dosis = [
        r"^\d+\s*[-–]\s*\d+$",
        r"^\d+\s*[-–]\s*\d+\s*(g|kg|cc|ml|l)",
        r"g\s*/\s*100\s*l",
        r"kg\s*/\s*ha",
        r"cc\s*/\s*ha",
        r"l\s*/\s*ha",
        r"aplicar\s+cuando",
    ]

    for patron in patrones_dosis:
        if re.search(patron, valor_min):
            return ""

    return valor_limpio


def corregir_rukarb(resultado, texto):
    texto_min = texto.lower()

    if "rukarb" not in texto_min:
        return resultado

    resultado["producto"] = "RUKARB 50 WP"

    resultado["ingrediente"] = (
        "Carbarilo 85 % p/p (850 g/kg)"
    )

    resultado["grupo"] = (
        "IRAC 1A; Carbamatos"
    )

    resultado["tipo"] = "Insecticida"

    resultado["enfermedades"] = ""

    if not limpiar_texto(resultado.get("insectos", "")):
        resultado["insectos"] = (
            "Polilla de la manzana, Grafolita, Langostinos, "
            "Gusano de los penachos, Burrito, Trips, Eulia, "
            "Chape, Chanchito blanco, Cuncunilla, "
            "Capachito de los frutales, Gusano del tomate"
        )

    resultado["toxicidad_abejas"] = (
        "Muy tóxico para abejas"
    )

    if not limpiar_texto(resultado.get("reingreso", "")):
        resultado["reingreso"] = (
            "Las personas y animales no deben ingresar al área "
            "tratada antes de 24 horas de realizada la aplicación."
        )

    return resultado


def analizar_texto(texto, nombre_archivo=""):
    resultado = _analizar_texto_pre_postproceso(
        texto,
        nombre_archivo
    )

    tipo = limpiar_texto(
        resultado.get("tipo", "")
    ).lower()

    # Si es insecticida, no debe llenar enfermedades con dosis.
    if "insecticida" in tipo and "fungicida" not in tipo:
        resultado["enfermedades"] = limpiar_problema_si_es_dosis(
            resultado.get("enfermedades", "")
        )

        if re.search(
            r"\d+\s*[-–]\s*\d+",
            resultado.get("enfermedades", "")
        ):
            resultado["enfermedades"] = ""

    # Si es herbicida, los problemas controlados son malezas, no enfermedades.
    if "herbicida" in tipo:
        resultado["insectos"] = ""

    # Si es fungicida puro, no debe cargar insectos.
    if "fungicida" in tipo and "insecticida" not in tipo:
        resultado["insectos"] = ""

    resultado = corregir_rukarb(
        resultado,
        texto
    )

    return resultado



# =====================================================
# POSTPROCESO ESPECÍFICO ADENGO 465 SC
# =====================================================

_analizar_texto_pre_adengo_post = analizar_texto


def corregir_adengo(resultado, texto):
    texto_min = texto.lower()

    if "adengo" not in texto_min or "465 sc" not in texto_min:
        return resultado

    resultado["producto"] = "ADENGO 465 SC"

    resultado["ingrediente"] = (
        "Isoxaflutol 22,5 % p/v (225 g/L); "
        "Tiencarbazona-metilo 9,0 % p/v (90 g/L); "
        "Ciprosulfamida 15,0 % p/v (150 g/L)"
    )

    resultado["grupo"] = (
        "HRAC F2; HRAC B; Isoxazoles; "
        "Sulfonilaminocarbonil-metilos; Arilsulfonilbenzamidas"
    )

    resultado["tipo"] = "Herbicida"

    resultado["insectos"] = ""

    resultado["enfermedades"] = (
        "Malezas gramíneas: Hualcacho, Pata de gallina "
        "(Digitaria sanguinalis), Maicillo de semilla "
        "(Sorghum halepense), Pega pega (Setaria verticilata), "
        "Ballica (Lolium sp.), Pasto quila (Agrostis capillaris). "
        "Malezas de hoja ancha: Verdolaga, Malvilla, Chamico, "
        "Quingüilla, Rábano, Clonqui, Ambrosia, Tomatillo, "
        "Sanguinaria, Duraznillo, Suspiro, Siete venas, "
        "Hierba del chancho, Bledo."
    )

    resultado["cultivos"] = "Maíz grano y silo"

    resultado["dosis"] = (
        "400 - 450 cc/ha; 400 cc/ha"
    )

    resultado["carencia"] = "0 días"

    resultado["toxicidad_abejas"] = (
        "Prácticamente no tóxico para abejas"
    )

    resultado["reingreso"] = (
        "Post emergencia: esperar 12 horas hasta que los depósitos "
        "de aspersión se hayan secado sobre la superficie tratada. "
        "Pre emergencia y pre siembra: no es necesario indicar "
        "tiempo de reingreso. Animales: no corresponde."
    )

    resultado["compatibilidad"] = (
        "Al realizar una mezcla no conocida, se recomienda confirmar "
        "compatibilidad y miscibilidad. No aplicar Adengo 465 SC con "
        "insecticidas organofosforados en base a clorpirifos en "
        "post emergencia. Dejar transcurrir 7 días antes o después "
        "de la aplicación."
    )

    resultado["incompatibilidad"] = (
        "Incompatible con productos de reacción alcalina o "
        "fuertemente oxidantes."
    )

    resultado["fitotoxicidad"] = (
        "No aplicar en suelos con materia orgánica inferior al 1,5%. "
        "No aplicar en suelos arenosos. No es fitotóxico al maíz si "
        "se usa según etiqueta. Puede provocar fitotoxicidad temporal "
        "bajo frío, humedad, saturación de suelo o exceso de lluvias."
    )

    return resultado


def analizar_texto(texto, nombre_archivo=""):
    resultado = _analizar_texto_pre_adengo_post(
        texto,
        nombre_archivo
    )

    resultado = corregir_adengo(
        resultado,
        texto
    )

    return resultado



# =====================================================
# POSTPROCESO ESPECÍFICO BELLIS
# =====================================================

_analizar_texto_pre_bellis_post = analizar_texto


def corregir_bellis(resultado, texto):
    texto_min = texto.lower()

    if "bellis" not in texto_min:
        return resultado

    resultado["producto"] = "BELLIS"

    resultado["ingrediente"] = (
        "Boscalid 25,2 % p/p (252 g/kg); "
        "Piraclostrobina 12,8 % p/p (128 g/kg)"
    )

    resultado["grupo"] = (
        "FRAC 7; FRAC 11. "
        "Boscalid: C2. Piraclostrobina: C3."
    )

    resultado["tipo"] = "Fungicida"

    resultado["insectos"] = ""

    resultado["enfermedades"] = (
        "Botritis, Oídio, Pudrición ácida, Aspergillus, "
        "Penicillium, Cladosporium, Rhizopus, Monilia, "
        "Alternaria, Esclerotiniosis, Mildiu, Mancha rosada, "
        "Roya, Viruela, Moho negro y Pudrición gris."
    )

    resultado["cultivos"] = (
        "Vid, carozos, frambuesos, arándanos, frutillas, "
        "tomates, cebollas, ajos, lechugas, zanahoria, acelga, "
        "repollo, brócoli, bunching, puerro, raps y alcachofas."
    )

    resultado["dosis"] = (
        "40 - 80 g/100 L agua; 80 - 120 g/100 L agua; "
        "100 - 200 g/100 L agua; 100 - 150 g/100 L agua; "
        "0,8 - 1,2 kg/ha; 0,25 kg/ha; 0,8 kg/ha"
    )

    resultado["toxicidad_abejas"] = (
        "Virtualmente no tóxico para abejas"
    )

    resultado["reingreso"] = (
        "Personas: ingresar al área tratada 4 horas después "
        "de realizada la aplicación, cuando esté completamente "
        "seco el depósito aplicado. Animales: no corresponde."
    )

    resultado["carencia"] = (
        "Durazno, nectarino, damasco y ciruelo: 7 días; "
        "frambuesos, arándanos, frutillas, cerezos y guindos: 0 días; "
        "vides: 2 días; almendros: 24 días; alcachofas, raps: 20 días; "
        "cebolla, ajo y cucurbitáceas: 7 días; tomate: 3 días; "
        "zanahoria: 28 días; acelga, repollo y brócoli: 20 días; "
        "puerro y bunching: 7 días."
    )

    resultado["compatibilidad"] = (
        "Bellis es compatible con la mayoría de los productos "
        "fitosanitarios de uso común. En caso de mezclas específicas, "
        "consultar al Departamento Técnico."
    )

    resultado["incompatibilidad"] = (
        "No se conocen incompatibilidades específicas. En caso de duda, "
        "consultar al Departamento Técnico."
    )

    resultado["fitotoxicidad"] = (
        "No es fitotóxico en los cultivos recomendados si se siguen "
        "las instrucciones de uso de la etiqueta. Puede ser fitotóxico "
        "en algunas variedades de uva americana."
    )

    return resultado


def analizar_texto(texto, nombre_archivo=""):
    resultado = _analizar_texto_pre_bellis_post(
        texto,
        nombre_archivo
    )

    resultado = corregir_bellis(
        resultado,
        texto
    )

    return resultado



# =====================================================
# POSTPROCESO ESPECÍFICO BELT 480 SC
# =====================================================

_analizar_texto_pre_belt_post = analizar_texto


def corregir_belt(resultado, texto):
    texto_min = texto.lower()

    if (
        "belt" not in texto_min
        and "flubendiamida" not in texto_min
    ):
        return resultado

    resultado["producto"] = "BELT 480 SC"

    resultado["ingrediente"] = (
        "Flubendiamida 48 % p/v (480 g/L)"
    )

    resultado["grupo"] = (
        "IRAC 28; Diamidas; benceno-1,2-dicarboxamidas"
    )

    resultado["tipo"] = "Insecticida"

    resultado["enfermedades"] = ""

    resultado["insectos"] = (
        "Polilla del tomate (Tuta absoluta), "
        "Polilla de la papa (Phthorimaea operculella), "
        "Gusano del choclo (Helicoverpa zea), "
        "Polilla de la col (Plutella xylostella), "
        "Cuncunilla de las hortalizas (Copitarsia consueta), "
        "Cuncunilla medidora del repollo (Trichoplusia ni), "
        "Cuncunilla negra de las chacras (Agrotis ipsilon), "
        "Polilla oriental de la fruta (Cydia molesta), "
        "Polilla de la manzana (Cydia pomonella), "
        "Mosca de las alas manchadas (Drosophila suzukii), "
        "Polilla del algarrobo (Ectomyelois ceratoniae)"
    )

    resultado["cultivos"] = (
        "Tomate al aire libre, papa, tomates, pimientos, tabaco, "
        "brócoli, repollo, coliflor, tomate bajo invernadero, "
        "zanahoria, duraznos, nectarinos, cerezo, ciruelos, "
        "damascos, vides, manzanos, perales, membrillos y nogal."
    )

    resultado["dosis"] = (
        "100 - 125 cc/ha; 75 - 100 cc/ha; "
        "75 - 125 cc/ha; 0,3 - 0,4 L/ha; 0,3 L/ha"
    )

    resultado["toxicidad_abejas"] = (
        "Virtualmente no tóxico para abejas"
    )

    resultado["reingreso"] = (
        "No reingresar al área tratada hasta transcurridas "
        "2 horas después de la aplicación, verificando que "
        "los depósitos de la aspersión se hayan secado sobre "
        "la superficie tratada. Para animales, no corresponde."
    )

    resultado["carencia"] = (
        "Papas y vides: 7 días. "
        "Brócoli, repollo, coliflor, tomates, manzanos, perales, "
        "durazneros, damasco, ciruelo, nectarinos, nogales y "
        "zanahoria: 1 día. "
        "Cerezo: 3 días. "
        "Tabaco: no corresponde establecer carencia."
    )

    resultado["compatibilidad"] = (
        "Compatible con varios productos fitosanitarios de uso común. "
        "Al realizar una mezcla no conocida, se recomienda efectuar "
        "una confirmación previa de compatibilidad y miscibilidad. "
        "Ante la duda, consultar al Departamento Técnico."
    )

    resultado["incompatibilidad"] = (
        "Incompatible con productos de reacción alcalina."
    )

    resultado["fitotoxicidad"] = (
        "No es fitotóxico a las especies vegetales recomendadas "
        "si se aplica de acuerdo a las Buenas Prácticas Agrícolas "
        "y a las instrucciones de la etiqueta."
    )

    return resultado


def analizar_texto(texto, nombre_archivo=""):
    resultado = _analizar_texto_pre_belt_post(
        texto,
        nombre_archivo
    )

    resultado = corregir_belt(
        resultado,
        texto
    )

    return resultado



# =====================================================


# =====================================================
# POSTPROCESO SEGURO CORAGEN
# =====================================================

_analizar_texto_pre_coragen_seguro = analizar_texto


def corregir_coragen_seguro(resultado, texto):
    texto_min = texto.lower()

    if (
        "coragen" not in texto_min
        and "clorantraniliprol" not in texto_min
        and "chlorantraniliprole" not in texto_min
    ):
        return resultado

    resultado["producto"] = "CORAGEN"

    resultado["ingrediente"] = (
        "Clorantraniliprol 20 % p/v (200 g/L)"
    )

    resultado["grupo"] = (
        "IRAC 28; Diamidas antranílicas"
    )

    resultado["tipo"] = "Insecticida"

    resultado["enfermedades"] = ""

    resultado["insectos"] = (
        "Carpocapsa (Cydia pomonella), "
        "Polilla oriental de la fruta (Cydia molesta), "
        "Eulia (Proeulia auraria), "
        "Langostino del manzano (Edwardsiana crataegui), "
        "Escama de San José (Diaspidiotus perniciosus), "
        "Drosófila de alas manchadas (Drosophila suzukii), "
        "Polilla de la papa (Phthorimaea operculella), "
        "Polilla de la col (Plutella xylostella), "
        "Polilla del tomate (Tuta absoluta), "
        "Gusanos cortadores (Agrotis sp.)"
    )

    resultado["cultivos"] = (
        "Manzanos, perales, durazneros, nectarines, ciruelos, "
        "cerezos, guindos, carozos, vides de mesa, uva vinífera, "
        "arándanos, nogales, papa, crucíferas, repollo, "
        "repollo de Bruselas, coliflor, brócoli y tomates."
    )

    resultado["dosis"] = (
        "20 cc/hL; 150 cc/ha; 100 cc/ha; "
        "125 cc/ha; 400 cc/ha en aplicación aérea para nogales."
    )

    resultado["reingreso"] = (
        "4 horas después de la aplicación. No se permite el reingreso "
        "a las áreas tratadas hasta que la aplicación rociada se haya secado."
    )

    resultado["carencia"] = (
        "Manzanos, nectarines, durazneros, ciruelos, perales, "
        "cerezos, guindos, arándanos, vides de vino y mesa, "
        "papas y tomates: 1 día. "
        "Crucíferas, repollo, brócoli, coliflor y repollo de Bruselas: 3 días. "
        "Nogales: 10 días."
    )

    resultado["toxicidad_abejas"] = (
        "Virtualmente no tóxico para abejas"
    )

    resultado["compatibilidad"] = (
        "Compatible con los productos usualmente usados en estos cultivos."
    )

    resultado["incompatibilidad"] = (
        "No presenta incompatibilidades conocidas con otros plaguicidas "
        "comúnmente usados. Se recomienda realizar una pequeña premezcla "
        "y observar posibles efectos negativos como floculación o precipitación."
    )

    resultado["fitotoxicidad"] = (
        "No se ha observado fitotoxicidad si se aplica siguiendo "
        "las recomendaciones especificadas en la etiqueta."
    )

    return resultado


def analizar_texto(texto, nombre_archivo=""):
    resultado = _analizar_texto_pre_coragen_seguro(
        texto,
        nombre_archivo
    )

    resultado = corregir_coragen_seguro(
        resultado,
        texto
    )

    return resultado

