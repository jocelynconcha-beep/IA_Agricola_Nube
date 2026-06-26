import re
from pypdf import PdfReader

try:
    import fitz
except Exception:
    fitz = None


def extraer_texto_pdf(path):
    reader = PdfReader(path)
    texto = ""

    for pagina in reader.pages:
        contenido = pagina.extract_text()
        if contenido:
            texto += contenido + "\n"

    return texto


def limpiar_celda_tabla(valor):
    if valor is None:
        return ""

    texto = str(valor)
    texto = texto.replace("￾", "")
    texto = texto.replace("\r", "\n")
    texto = re.sub(r"-\s*\n\s*", "", texto)
    texto = re.sub(r"\s*\n\s*", " ", texto)
    texto = re.sub(r"\bFACSIMIL\s+SAG\b", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\bG\s*A\s*S\s*L\s*I\s*M\s*I\s*S\s*C\b", "", texto, flags=re.IGNORECASE)
    texto = texto.replace("fol ración", "floración")
    texto = texto.replace("fol r", "flor")
    texto = texto.replace("Defni ir", "Definir")
    texto = texto.replace("defni ir", "definir")
    texto = re.sub(r"\s+", " ", texto)
    texto = re.sub(r"^(?:[A-Z]\s+){1,8}(?=[A-ZÁÉÍÓÚÑa-záéíóúñ0-9])", "", texto)

    return texto.strip(" .;")


def encabezado_tabla_uso(fila):
    celdas = [limpiar_celda_tabla(celda).lower() for celda in fila]
    texto = " ".join(celdas)

    tiene_cultivo = any("cultivo" in celda for celda in celdas)
    tiene_dosis = any("dosis" in celda for celda in celdas)
    tiene_problema = any(
        palabra in texto
        for palabra in ["plaga", "enfermedad", "maleza"]
    )

    return tiene_cultivo and tiene_dosis and tiene_problema


def indices_tabla_uso(encabezado):
    indices = {
        "cultivo": None,
        "problema": None,
        "dosis": None,
        "observaciones": None
    }

    for indice, celda in enumerate(encabezado):
        celda_limpia = limpiar_celda_tabla(celda).lower()

        if "cultivo" in celda_limpia:
            indices["cultivo"] = indice
        elif any(palabra in celda_limpia for palabra in ["plaga", "enfermedad", "maleza"]):
            indices["problema"] = indice
        elif "dosis" in celda_limpia:
            indices["dosis"] = indice
        elif "observ" in celda_limpia:
            indices["observaciones"] = indice

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

            for indice_fila, fila in enumerate(filas[:3]):
                if encabezado_tabla_uso(fila):
                    encabezado_indice = indice_fila
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

                if cultivo == "" and problema == "" and dosis == "":
                    continue

                if dosis == "" or not re.search(r"\d", dosis):
                    continue

                usos.append({
                    "cultivo": cultivo,
                    "problema": problema,
                    "dosis": dosis,
                    "observaciones": observaciones,
                    "pagina": numero_pagina
                })

    return usos


def limpiar_nombre_archivo(nombre_archivo):
    nombre = nombre_archivo.replace(".pdf", "").replace(".PDF", "")
    nombre = nombre.replace("_", " ").replace("-", " ")
    nombre = re.sub(r"\s+", " ", nombre)
    return nombre.strip().upper()


def detectar_producto(texto, nombre_archivo=""):
    texto_upper = texto.upper()

    nombre_archivo_limpio = limpiar_nombre_archivo(nombre_archivo)

    if nombre_archivo_limpio:
        return nombre_archivo_limpio

    productos_conocidos = {
        "PIRIMOR": ["PIRIMOR", "PIRIMOR®"],
        "BOSCALID 50% WG": ["BOSCALID 50% WG", "BOSCALID"],
        "CONFIDOR 350 SC": ["CONFIDOR 350 SC", "CONFIDOR"],
        "GESATOP": ["GESATOP"],
        "GRANITE": ["GRANITE", "GRANITE®", "GRANITE™"],
        "CONSENTO 450 SC": ["CONSENTO 450 SC", "CONSENTO"],
        "SELECRON 720 EC": ["SELECRON 720 EC", "SELECRON"],
        "BELLIS": ["BELLIS"],
        "FUNGISEI": ["FUNGISEI"],
        "CAYENNE 500 SC": ["CAYENNE 500 SC", "CAYENNE"]
    }

    for nombre_limpio, variantes in productos_conocidos.items():
        for variante in variantes:
            if variante in texto_upper:
                return nombre_limpio

    return "PRODUCTO NO IDENTIFICADO"


def detectar_ingrediente(texto):
    texto_lower = texto.lower()

    ingredientes = [
        "boscalid",
        "imidacloprid",
        "simazina",
        "fenamidona",
        "propamocarb",
        "profenofos",
        "pirimicarb",
        "acetamiprid",
        "thiamethoxam",
        "azoxystrobin",
        "difenoconazole",
        "tebuconazole",
        "captan",
        "mancozeb",
        "bacillus subtilis cepa iab/bs03",
        "bacillus subtilis",
        "cobre",
        "oxifluorfen",
        "penoxsulam"
    ]

    encontrados = []

    for ing in ingredientes:
        if ing in texto_lower:
            encontrados.append(ing.title())

    return ", ".join(sorted(set(encontrados)))


def detectar_grupo(texto):
    texto_lower = texto.lower()

    if "oxifluorfen" in texto_lower and "penoxsulam" in texto_lower:
        return "HRAC 14, HRAC 2"

    if "boscalid" in texto_lower:
        return "FRAC 7"

    if "imidacloprid" in texto_lower:
        return "IRAC 4"

    if "profenofos" in texto_lower:
        return "IRAC 1B"

    if "pirimicarb" in texto_lower:
        return "IRAC 1A"

    if "simazina" in texto_lower:
        return "HRAC 5"

    if "bacillus subtilis" in texto_lower:
        return "BIOLOGICO"

    patrones = [
        r"grupo\s+(\d+[a-z]?)",
        r"irac\s+(\d+[a-z]?)",
        r"frac\s+(\d+[a-z]?)",
        r"hrac\s+(\d+[a-z]?)"
    ]

    grupos = []

    for patron in patrones:
        for match in re.finditer(patron, texto_lower):
            numero = match.group(1).upper()

            if "herbicida" in texto_lower or "hrac" in texto_lower:
                grupos.append("HRAC " + numero)
            elif "insecticida" in texto_lower or "irac" in texto_lower:
                grupos.append("IRAC " + numero)
            else:
                grupos.append("FRAC " + numero)

    grupos = sorted(set(grupos))

    return ", ".join(grupos)


def detectar_tipo(texto, grupo, ingrediente):
    texto_lower = texto.lower()
    grupo_lower = grupo.lower()
    ingrediente_lower = ingrediente.lower()

    if "biologico" in grupo_lower or "biológico" in grupo_lower:
        return "Biológico"

    if "bacillus" in ingrediente_lower:
        return "Biológico"

    if "irac" in grupo_lower:
        return "Insecticida"

    if "pirimicarb" in ingrediente_lower:
        return "Insecticida"

    if "imidacloprid" in ingrediente_lower:
        return "Insecticida"

    if "profenofos" in ingrediente_lower:
        return "Insecticida"

    if "hrac" in grupo_lower:
        return "Herbicida"

    if "oxifluorfen" in ingrediente_lower:
        return "Herbicida"

    if "penoxsulam" in ingrediente_lower:
        return "Herbicida"

    if "simazina" in ingrediente_lower:
        return "Herbicida"

    if "frac" in grupo_lower:
        return "Fungicida"

    if "insecticida" in texto_lower:
        return "Insecticida"

    if "herbicida" in texto_lower:
        return "Herbicida"

    if "fungicida" in texto_lower:
        return "Fungicida"

    return ""


def detectar_cultivos(texto):
    texto_lower = texto.lower()

    cultivos = {
        "vid": ["vid", "vides", "uva"],
        "tomate": ["tomate", "tomates"],
        "papa": ["papa", "papas"],
        "cerezo": ["cerezo", "cerezos"],
        "manzano": ["manzano", "manzanos"],
        "nogal": ["nogal", "nogales"],
        "berries": ["berries", "frutilla", "frutillas", "arándano", "arandano", "frambueso", "frambuesa"],
        "carozos": ["carozos", "durazno", "nectarino", "damasco", "ciruelo"],
        "almendro": ["almendro", "almendros"],
        "cebolla": ["cebolla", "cebollas"],
        "pimentón": ["pimenton", "pimentón", "pimentones"],
        "tabaco": ["tabaco"]
    }

    encontrados = []

    for cultivo, palabras in cultivos.items():
        if any(p in texto_lower for p in palabras):
            encontrados.append(cultivo)

    return ", ".join(sorted(set(encontrados)))


def detectar_enfermedades(texto):
    texto_lower = texto.lower()

    enfermedades = {
        "botritis": ["botritis", "botrytis"],
        "oidio": ["oidio", "oídio"],
        "tizón": ["tizon", "tizón"],
        "alternaria": ["alternaria"],
        "mildiu": ["mildiu"],
        "esclerotinia": ["sclerotinia", "esclerotinia"]
    }

    encontrados = []

    for enfermedad, palabras in enfermedades.items():
        if any(p in texto_lower for p in palabras):
            encontrados.append(enfermedad)

    return ", ".join(sorted(set(encontrados)))


def detectar_insectos(texto):
    texto_lower = texto.lower()

    insectos = {
        "trips": ["trips"],
        "pulgón": ["pulgon", "pulgón", "pulgones", "áfidos", "afidos"],
        "mosca blanca": ["mosca blanca"],
        "tuta": ["tuta"],
        "chanchito blanco": ["chanchito blanco"],
        "polilla": ["polilla"],
        "escama": ["escama"]
    }

    encontrados = []

    for insecto, palabras in insectos.items():
        if any(p in texto_lower for p in palabras):
            encontrados.append(insecto)

    return ", ".join(sorted(set(encontrados)))


def detectar_dosis(texto):
    texto_limpio = texto.lower()

    texto_limpio = texto_limpio.replace("￾", "")
    texto_limpio = texto_limpio.replace("-\n", "-")
    texto_limpio = texto_limpio.replace("\n", " ")
    texto_limpio = texto_limpio.replace("lts", "l")
    texto_limpio = texto_limpio.replace("litros", "l")
    texto_limpio = texto_limpio.replace("litro", "l")
    texto_limpio = texto_limpio.replace("hectárea", "ha")
    texto_limpio = texto_limpio.replace("hectarea", "ha")
    texto_limpio = re.sub(r"\s+", " ", texto_limpio)

    texto_limpio = re.sub(
        r"\b(\d{2,3})(1[.,]000)\s*(g|kg|l|ml|cc)\s*/\s*ha\b",
        r"\1-\2 \3/ha",
        texto_limpio
    )

    patrones = [
        r"\b\d+(?:[,.]\d+)?\s*(?:-|a|–)\s*\d+(?:[,.]\d+)?\s*(?:kg|g|l|ml|cc)\s*/\s*(?:ha|há|hás|100\s*l|100l)\b",
        r"\b\d+(?:[,.]\d+)?\s*(?:kg|g|l|ml|cc)\s*/\s*(?:ha|há|hás|100\s*l|100l)\b",
        r"\b\d+(?:[,.]\d+)?\s*(?:-|a|–)\s*\d+(?:[,.]\d+)?\s*(?:kg|g|l|ml|cc)\s+por\s+(?:ha|100\s*l)\b",
        r"\b\d+(?:[,.]\d+)?\s*(?:kg|g|l|ml|cc)\s+por\s+(?:ha|100\s*l)\b",
        r"\b\d+(?:[,.]\d+)?\s*(?:-|a|–)\s*\d+(?:[,.]\d+)?\s*(?:cc|ml|g)\s*/\s*100\s*(?:l|litros)\b",
        r"\b\d+(?:[,.]\d+)?\s*(?:cc|ml|g)\s*/\s*100\s*(?:l|litros)\b"
    ]

    dosis_encontradas = []

    for patron in patrones:
        resultados = re.findall(patron, texto_limpio)

        for resultado in resultados:
            resultado = resultado.strip()
            resultado = resultado.replace(" a ", "-")
            resultado = re.sub(r"\s+", " ", resultado)

            if resultado not in dosis_encontradas:
                dosis_encontradas.append(resultado)

    if dosis_encontradas:
        return ", ".join(dosis_encontradas[:8])

    match = re.search(r"dosis[^.]{0,250}", texto_limpio)

    if match:
        return match.group(0).strip()

    return ""



def detectar_compatibilidad(texto):
    texto_limpio = texto.lower()
    texto_limpio = texto_limpio.replace("\n", " ")
    texto_limpio = re.sub(r"\s+", " ", texto_limpio)

    match = re.search(
        r"compatibilidades[^.]{0,500}",
        texto_limpio
    )

    if match:
        resultado = match.group(0).strip()
        resultado = resultado.replace("compatibilidades", "Compatible con")
        return resultado

    match = re.search(
        r"compatible con[^.]{0,500}",
        texto_limpio
    )

    if match:
        return match.group(0).strip()

    return ""


def detectar_incompatibilidad(texto):
    texto_limpio = texto.lower()
    texto_limpio = texto_limpio.replace("\n", " ")
    texto_limpio = re.sub(r"\s+", " ", texto_limpio)

    match = re.search(
        r"incompatibilidades[^.]{0,400}",
        texto_limpio
    )

    if match:
        resultado = match.group(0).strip()
        resultado = resultado.replace("incompatibilidades", "Incompatible con")
        return resultado

    match = re.search(
        r"no es compatible[^.]{0,400}",
        texto_limpio
    )

    if match:
        return match.group(0).strip()

    return ""


def detectar_fitotoxicidad(texto):
    texto_limpio = texto.lower()
    texto_limpio = texto_limpio.replace("\n", " ")
    texto_limpio = re.sub(r"\s+", " ", texto_limpio)

    match = re.search(
        r"fitotoxicidad[^.]{0,400}",
        texto_limpio
    )

    if match:
        return match.group(0).strip()

    match = re.search(
        r"no es fitot[oó]xico[^.]{0,400}",
        texto_limpio
    )

    if match:
        return match.group(0).strip()

    return ""



def detectar_reingreso(texto):
    texto_lower = texto.lower()

    patrones = [
        r"reingreso[^.]{0,120}?(\d+\s*horas)",
        r"no reingresar[^.]{0,120}?(\d+\s*horas)",
        r"transcurrir\s+(\d+\s*horas)",
        r"transcurridas\s+(\d+\s*horas)"
    ]

    for patron in patrones:
        match = re.search(patron, texto_lower)
        if match:
            return match.group(1)

    return ""


def detectar_carencia(texto):
    texto_lower = texto.lower()

    match = re.search(r"(periodo|período)\s+de\s+carencia[^.]{0,220}", texto_lower)
    if match:
        return match.group(0).strip()

    match = re.search(r"carencia[^.]{0,160}", texto_lower)
    if match:
        return match.group(0).strip()

    return ""


def detectar_toxicidad_abejas(texto):
    texto_lower = texto.lower()

    texto_lower = texto_lower.replace("￾", "")
    texto_lower = texto_lower.replace("-\n", "")
    texto_lower = texto_lower.replace("- ", "")
    texto_lower = texto_lower.replace("\n", " ")
    texto_lower = re.sub(r"\s+", " ", texto_lower)

    texto_lower = texto_lower.replace("tóxi co", "tóxico")
    texto_lower = texto_lower.replace("toxi co", "toxico")
    texto_lower = texto_lower.replace("tóxi-co", "tóxico")
    texto_lower = texto_lower.replace("toxi-co", "toxico")
    texto_lower = texto_lower.replace("tóxi co", "tóxico")

    texto_simple = texto_lower
    texto_simple = texto_simple.replace("á", "a")
    texto_simple = texto_simple.replace("é", "e")
    texto_simple = texto_simple.replace("í", "i")
    texto_simple = texto_simple.replace("ó", "o")
    texto_simple = texto_simple.replace("ú", "u")

    patrones = [
        ("virtualmente no toxico para abejas", "Virtualmente no tóxico para abejas"),
        ("virtualmente no toxico para las abejas", "Virtualmente no tóxico para abejas"),
        ("practicamente no toxico para abejas", "Prácticamente no tóxico para abejas"),
        ("practicamente no toxico para las abejas", "Prácticamente no tóxico para abejas"),
        ("no toxico para abejas", "No tóxico para abejas"),
        ("no toxico para las abejas", "No tóxico para abejas"),
        ("ligeramente toxico para abejas", "Ligeramente tóxico para abejas"),
        ("ligeramente toxico para las abejas", "Ligeramente tóxico para abejas"),
        ("moderadamente toxico para abejas", "Moderadamente tóxico para abejas"),
        ("moderadamente toxico para las abejas", "Moderadamente tóxico para abejas"),
        ("altamente toxico para abejas", "Altamente tóxico para abejas"),
        ("altamente toxico para las abejas", "Altamente tóxico para abejas"),
        ("muy toxico para abejas", "Muy tóxico para abejas"),
        ("muy toxico para las abejas", "Muy tóxico para abejas")
    ]

    for patron, resultado in patrones:
        if patron in texto_simple:
            return resultado

    return ""


def analizar_texto(texto, nombre_archivo=""):
    ingrediente = detectar_ingrediente(texto)
    grupo = detectar_grupo(texto)
    tipo = detectar_tipo(texto, grupo, ingrediente)

    return {
        "producto": detectar_producto(texto, nombre_archivo),
        "ingrediente": ingrediente,
        "grupo": grupo,
        "tipo": tipo,
        "cultivos": detectar_cultivos(texto),
        "enfermedades": detectar_enfermedades(texto),
        "insectos": detectar_insectos(texto),
        "dosis": detectar_dosis(texto),
        "compatibilidad": detectar_compatibilidad(texto),
        "incompatibilidad": detectar_incompatibilidad(texto),
        "fitotoxicidad": detectar_fitotoxicidad(texto),
        "reingreso": detectar_reingreso(texto),
        "carencia": detectar_carencia(texto),
        "toxicidad_abejas": detectar_toxicidad_abejas(texto),
        "texto": texto
    }


# ==============================================================
# EXTRACCIÓN MEJORADA DE DATOS DE ETIQUETA
# ==============================================================

_detectar_ingrediente_anterior = detectar_ingrediente
_detectar_grupo_anterior = detectar_grupo
_detectar_tipo_anterior = detectar_tipo
_detectar_reingreso_anterior = detectar_reingreso
_detectar_compatibilidad_anterior = detectar_compatibilidad
_detectar_incompatibilidad_anterior = detectar_incompatibilidad


def _limpiar_texto_extraido(valor):
    if valor is None:
        return ""

    valor = str(valor)
    valor = valor.replace("\r", " ")
    valor = re.sub(r"\s+", " ", valor)
    return valor.strip(" \n\t:;.-")


def _unicos_en_orden(valores):
    resultado = []
    vistos = set()

    for valor in valores:
        valor_limpio = _limpiar_texto_extraido(valor)

        if not valor_limpio:
            continue

        clave = valor_limpio.casefold()

        if clave not in vistos:
            vistos.add(clave)
            resultado.append(valor_limpio)

    return resultado


def detectar_ingrediente(texto):
    bloque = re.search(
        r"COMPOSICI[ÓO]N\s*:\s*(.+?)"
        r"(?=\bCoformulantes\b|\bAutorizaci[óo]n\b)",
        texto,
        flags=re.IGNORECASE | re.DOTALL
    )

    ingredientes = []

    if bloque:
        contenido = bloque.group(1)

        patron = re.compile(
            r"([A-ZÁÉÍÓÚÜÑ][A-Za-zÁÉÍÓÚÜÑáéíóúüñ"
            r"\s\-]+?)"
            r"[*\"']*\.{2,}\s*"
            r"([0-9]+(?:[.,][0-9]+)?\s*%\s*p/v"
            r"(?:\s*\([^)]*\))?)",
            flags=re.IGNORECASE
        )

        for nombre, concentracion in patron.findall(contenido):
            nombre = _limpiar_texto_extraido(nombre)
            concentracion = _limpiar_texto_extraido(concentracion)

            if nombre and concentracion:
                ingredientes.append(
                    f"{nombre} {concentracion}"
                )

    ingredientes = _unicos_en_orden(ingredientes)

    if ingredientes:
        return "; ".join(ingredientes)

    return _detectar_ingrediente_anterior(texto)


def detectar_grupo(texto):
    coincidencias = re.findall(
        r"pertenece\s+al\s+grupo\s+qu[íi]mico\s+"
        r"de\s+los?\s+([A-Za-zÁÉÍÓÚÜÑáéíóúüñ\-]+)",
        texto,
        flags=re.IGNORECASE
    )

    grupos = _unicos_en_orden(coincidencias)

    if grupos:
        return "; ".join(grupos)

    bloque = re.search(
        r"Grupo\s+Qu[íi]mico\s*:\s*(.+?)(?=\n[A-ZÁÉÍÓÚÑ]{3,}|$)",
        texto,
        flags=re.IGNORECASE | re.DOTALL
    )

    if bloque:
        valor = _limpiar_texto_extraido(bloque.group(1))

        if valor:
            return valor

    return _detectar_grupo_anterior(texto)


def detectar_tipo(texto, grupo="", ingrediente=""):
    tipos_posibles = (
        "FUNGICIDA|INSECTICIDA|HERBICIDA|ACARICIDA|"
        "NEMATICIDA|BACTERICIDA|MOLUSQUICIDA|"
        "REGULADOR DE CRECIMIENTO"
    )

    encabezado = re.search(
        rf"(?im)^\s*((?:{tipos_posibles})"
        rf"(?:\s*[-/]\s*(?:{tipos_posibles}))*)\s*$",
        texto
    )

    if encabezado:
        valor = encabezado.group(1)
        partes = re.split(r"\s*[-/]\s*", valor)
        partes = [
            parte.strip().title()
            for parte in partes
            if parte.strip()
        ]

        return " - ".join(partes)

    descripcion = re.search(
        r"\bes\s+un(?:a)?\s+"
        r"((?:fungicida|insecticida|herbicida|acaricida|"
        r"nematicida|bactericida)"
        r"(?:\s*[- ]\s*(?:fungicida|insecticida|herbicida|"
        r"acaricida|nematicida|bactericida))*)",
        texto,
        flags=re.IGNORECASE
    )

    if descripcion:
        valor = descripcion.group(1)
        palabras = re.findall(
            r"fungicida|insecticida|herbicida|acaricida|"
            r"nematicida|bactericida",
            valor,
            flags=re.IGNORECASE
        )

        palabras = _unicos_en_orden(
            palabra.title() for palabra in palabras
        )

        if palabras:
            return " - ".join(palabras)

    return _detectar_tipo_anterior(
        texto,
        grupo,
        ingrediente
    )


def detectar_reingreso(texto):
    coincidencia = re.search(
        r"Tiempo\s+de\s+reingreso"
        r"(?:\s+al\s+[áa]rea\s+tratada)?\s*:\s*"
        r"(.+?)"
        r"(?=\s*Carencia\s*:|\s*Fitotoxicidad\s*:|$)",
        texto,
        flags=re.IGNORECASE | re.DOTALL
    )

    if coincidencia:
        valor = _limpiar_texto_extraido(
            coincidencia.group(1)
        )

        if valor:
            return valor

    return _detectar_reingreso_anterior(texto)


def detectar_compatibilidad(texto):
    coincidencia = re.search(
        r"(?<!In)Compatibilidad\s*:\s*(.+?)"
        r"(?=\s*Incompatibilidad\s*:|"
        r"\s*Fitotoxicidad\s*:|"
        r"\s*Nota\s+del\s+Fabricante|$)",
        texto,
        flags=re.IGNORECASE | re.DOTALL
    )

    if coincidencia:
        valor = _limpiar_texto_extraido(
            coincidencia.group(1)
        )

        if valor:
            return valor

    valor_anterior = _limpiar_texto_extraido(
        _detectar_compatibilidad_anterior(texto)
    )

    frases_incompletas = {
        "compatible con",
        "compatibilidad",
        "compatible"
    }

    if valor_anterior.casefold() in frases_incompletas:
        return ""

    return valor_anterior


def detectar_incompatibilidad(texto):
    coincidencia = re.search(
        r"Incompatibilidad\s*:\s*(.+?)"
        r"(?=\s*Compatibilidad\s*:|"
        r"\s*Fitotoxicidad\s*:|"
        r"\s*Nota\s+del\s+Fabricante|$)",
        texto,
        flags=re.IGNORECASE | re.DOTALL
    )

    if coincidencia:
        valor = _limpiar_texto_extraido(
            coincidencia.group(1)
        )

        if valor:
            return valor

    valor_anterior = _limpiar_texto_extraido(
        _detectar_incompatibilidad_anterior(texto)
    )

    frases_incompletas = {
        "incompatible con",
        "incompatibilidad",
        "incompatible"
    }

    if valor_anterior.casefold() in frases_incompletas:
        return ""

    return valor_anterior


# ==============================================================
# MEJORAS PARA ETIQUETAS CON TABLAS COMBINADAS Y HERBICIDAS
# ==============================================================

_detectar_producto_previo = detectar_producto
_detectar_grupo_previo = detectar_grupo
_detectar_carencia_previa = detectar_carencia
_detectar_toxicidad_abejas_previa = detectar_toxicidad_abejas


def detectar_producto(texto, nombre_archivo=""):
    patrones = [
        r"(?im)^\s*([A-ZÁÉÍÓÚÜÑ][A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9®™' -]{2,40})\s*$"
        r"(?=\s*(?:Herbicida|Fungicida|Insecticida|Acaricida))",

        r"(?i)Nombre\s+comercial(?:\s+del\s+producto\s+químico)?\s*[:\-]?\s*"
        r"([A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9®™' -]{2,50})",
    ]

    for patron in patrones:
        coincidencia = re.search(patron, texto)

        if coincidencia:
            producto = _limpiar_texto_extraido(
                coincidencia.group(1)
            )

            producto = producto.replace("®", "").replace("™", "")
            producto = re.sub(r"\s+", " ", producto).strip()

            palabras_invalidas = {
                "precauciones y advertencias",
                "instrucciones de uso",
                "composición",
                "herbicida",
                "fungicida",
                "insecticida"
            }

            if producto.casefold() not in palabras_invalidas:
                return producto

    return _detectar_producto_previo(
        texto,
        nombre_archivo
    )


def detectar_grupo(texto):
    patrones = [
        r"pertenece\s+al\s+grupo\s+qu[íi]mico\s+"
        r"(?:de\s+los?|de\s+las?|)\s*"
        r"([A-Za-zÁÉÍÓÚÜÑáéíóúüñ\-]+)",

        r"clasificaci[óo]n\s+HRAC(?:/WSSA)?[^.]{0,80}?"
        r"grupo\s+([0-9A-Za-z]+)",

        r"\bHRAC\s*[:\-]?\s*([0-9A-Za-z]+)",
    ]

    resultados = []

    for patron in patrones:
        for valor in re.findall(
            patron,
            texto,
            flags=re.IGNORECASE
        ):
            valor = _limpiar_texto_extraido(valor)

            if valor:
                resultados.append(valor)

    resultados = _unicos_en_orden(resultados)

    grupo_quimico = ""
    grupo_hrac = ""

    for resultado in resultados:
        if re.fullmatch(r"[0-9A-Za-z]+", resultado):
            grupo_hrac = resultado
        else:
            grupo_quimico = resultado

    if grupo_hrac and grupo_quimico:
        return f"HRAC {grupo_hrac}; {grupo_quimico}"

    if grupo_hrac:
        return f"HRAC {grupo_hrac}"

    if grupo_quimico:
        return grupo_quimico

    return _detectar_grupo_previo(texto)


def detectar_carencia(texto):
    coincidencia = re.search(
        r"Per[íi]odos?\s+de\s+carencias?\s*:\s*(.+?)"
        r"(?=\s*Tiempo\s+de\s+reingreso\s*:|"
        r"\s*Reingreso\s*:|"
        r"\s*Nota\b|$)",
        texto,
        flags=re.IGNORECASE | re.DOTALL
    )

    if coincidencia:
        valor = _limpiar_texto_extraido(
            coincidencia.group(1)
        )

        if valor:
            return valor

    return _detectar_carencia_previa(texto)


def detectar_toxicidad_abejas(texto):
    texto_simple = _limpiar_texto_extraido(texto)

    patrones_baja = [
        r"no\s+es\s+peligroso\s+para\s+las\s+abejas",
        r"no\s+es\s+t[óo]xico\s+para\s+las\s+abejas",
        r"virtualmente\s+no\s+t[óo]xico\s+para\s+abejas",
    ]

    for patron in patrones_baja:
        if re.search(
            patron,
            texto_simple,
            flags=re.IGNORECASE
        ):
            return "No es peligroso para las abejas"

    patrones_alta = [
        r"muy\s+t[óo]xico\s+para\s+abejas",
        r"altamente\s+t[óo]xico\s+para\s+abejas",
        r"t[óo]xico\s+para\s+abejas",
    ]

    for patron in patrones_alta:
        coincidencia = re.search(
            patron,
            texto_simple,
            flags=re.IGNORECASE
        )

        if coincidencia:
            return coincidencia.group(0).capitalize()

    return _detectar_toxicidad_abejas_previa(texto)
