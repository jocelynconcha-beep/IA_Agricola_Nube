import os
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests
from dotenv import load_dotenv


CARPETA = Path(__file__).resolve().parent
load_dotenv(CARPETA / ".env")


def _obtener_configuracion():
    url = os.getenv("SUPABASE_URL", "").strip()
    key = os.getenv("SUPABASE_KEY", "").strip()

    # Cuando la app esté publicada, también podrá leer Streamlit Secrets.
    if not url or not key:
        try:
            import streamlit as st

            url = url or str(st.secrets.get("SUPABASE_URL", "")).strip()
            key = key or str(st.secrets.get("SUPABASE_KEY", "")).strip()
        except Exception:
            pass

    url = url.rstrip("/")

    # Evita formar /rest/v1/rest/v1.
    if url.endswith("/rest/v1"):
        url = url[:-8]

    if not url:
        raise RuntimeError("No se encontró SUPABASE_URL.")

    if not key:
        raise RuntimeError("No se encontró SUPABASE_KEY.")

    return url, key


def _headers(prefer=None):
    _, key = _obtener_configuracion()

    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }

    if prefer:
        headers["Prefer"] = prefer

    return headers


def _endpoint(tabla):
    url, _ = _obtener_configuracion()
    return f"{url}/rest/v1/{tabla}"


def _verificar_respuesta(respuesta, accion):
    if not respuesta.ok:
        raise RuntimeError(
            f"Error al {accion}: "
            f"{respuesta.status_code} {respuesta.text}"
        )


def _texto_o_none(valor):
    if valor is None:
        return None

    texto = str(valor).strip()

    if texto == "":
        return None

    return texto


def _dataframe(datos):
    if not datos:
        return pd.DataFrame()

    return pd.DataFrame(datos)


# -------------------------------------------------------------------
# Compatibilidad con la antigua versión SQLite
# -------------------------------------------------------------------

def conectar():
    raise RuntimeError(
        "La aplicación ahora usa Supabase y no una conexión SQLite directa."
    )


def crear_tablas():
    # Las tablas ya fueron creadas en Supabase.
    return None


def migrar_bd():
    # Las columnas ya fueron creadas en Supabase.
    return None


def crear_tabla_experiencias_campo():
    # La tabla ya fue creada en Supabase.
    return None


# -------------------------------------------------------------------
# SUPABASE STORAGE
# -------------------------------------------------------------------

def subir_pdf_storage(
    nombre_archivo,
    pdf_bytes,
    bucket="pdfs-productos"
):
    if not pdf_bytes:
        raise ValueError("El PDF está vacío.")

    nombre = str(nombre_archivo).strip()

    if not nombre:
        raise ValueError("El PDF no tiene nombre.")

    url, key = _obtener_configuracion()

    nombre_codificado = quote(nombre, safe="")
    endpoint = (
        f"{url}/storage/v1/object/"
        f"{bucket}/{nombre_codificado}"
    )

    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/pdf",
        "x-upsert": "true",
    }

    respuesta = requests.post(
        endpoint,
        headers=headers,
        data=pdf_bytes,
        timeout=120,
    )

    _verificar_respuesta(
        respuesta,
        "subir el PDF a Supabase Storage"
    )

    return (
        f"{url}/storage/v1/object/public/"
        f"{bucket}/{nombre_codificado}"
    )



# -------------------------------------------------------------------
# PRODUCTOS
# -------------------------------------------------------------------

def guardar_producto(
    nombre,
    ingrediente,
    grupo,
    tipo,
    cultivos,
    enfermedades,
    insectos,
    dosis="",
    compatibilidad="",
    incompatibilidad="",
    fitotoxicidad="",
    reingreso="",
    carencia="",
    toxicidad_abejas="",
    pdf=""
):
    datos = {
        "nombre": _texto_o_none(nombre),
        "ingrediente": _texto_o_none(ingrediente),
        "grupo": _texto_o_none(grupo),
        "tipo": _texto_o_none(tipo),
        "cultivos": _texto_o_none(cultivos),
        "enfermedades": _texto_o_none(enfermedades),
        "insectos": _texto_o_none(insectos),
        "dosis": _texto_o_none(dosis),
        "compatibilidad": _texto_o_none(compatibilidad),
        "incompatibilidad": _texto_o_none(incompatibilidad),
        "fitotoxicidad": _texto_o_none(fitotoxicidad),
        "reingreso": _texto_o_none(reingreso),
        "carencia": _texto_o_none(carencia),
        "toxicidad_abejas": _texto_o_none(toxicidad_abejas),
        "pdf": _texto_o_none(pdf),
    }

    respuesta = requests.post(
        _endpoint("productos"),
        headers=_headers("return=representation"),
        json=datos,
        timeout=60,
    )

    _verificar_respuesta(respuesta, "guardar el producto")

    filas = respuesta.json()

    if not filas:
        raise RuntimeError("Supabase no devolvió el ID del producto creado.")

    return int(filas[0]["id"])


def obtener_productos():
    respuesta = requests.get(
        _endpoint("productos"),
        headers=_headers(),
        params={
            "select": "*",
            "order": "id.desc",
        },
        timeout=60,
    )

    _verificar_respuesta(respuesta, "obtener los productos")

    return _dataframe(respuesta.json())


def eliminar_producto(id_producto):
    respuesta = requests.delete(
        _endpoint("productos"),
        headers=_headers("return=minimal"),
        params={
            "id": f"eq.{int(id_producto)}",
        },
        timeout=60,
    )

    _verificar_respuesta(respuesta, "eliminar el producto")


def eliminar_duplicados():
    df = obtener_productos()

    if df.empty or "pdf" not in df.columns:
        return

    df["pdf_limpio"] = df["pdf"].fillna("").astype(str).str.strip()

    duplicados = df[
        (df["pdf_limpio"] != "")
        & df.duplicated(subset=["pdf_limpio"], keep="first")
    ]

    for _, fila in duplicados.iterrows():
        eliminar_producto(int(fila["id"]))


# -------------------------------------------------------------------
# USOS DEL PRODUCTO
# -------------------------------------------------------------------

def guardar_usos_producto(producto_id, pdf, usos):
    # Eliminar usos anteriores del producto.
    respuesta = requests.delete(
        _endpoint("usos_producto"),
        headers=_headers("return=minimal"),
        params={
            "producto_id": f"eq.{int(producto_id)}",
        },
        timeout=60,
    )

    _verificar_respuesta(
        respuesta,
        "eliminar los usos anteriores del producto"
    )

    # Eliminar también registros anteriores asociados al mismo PDF.
    if _texto_o_none(pdf):
        respuesta = requests.delete(
            _endpoint("usos_producto"),
            headers=_headers("return=minimal"),
            params={
                "pdf": f"eq.{pdf}",
            },
            timeout=60,
        )

        _verificar_respuesta(
            respuesta,
            "eliminar los usos anteriores del PDF"
        )

    filas = []

    for uso in usos:
        filas.append({
            "producto_id": int(producto_id),
            "pdf": _texto_o_none(pdf),
            "cultivo": _texto_o_none(uso.get("cultivo", "")),
            "problema": _texto_o_none(uso.get("problema", "")),
            "dosis": _texto_o_none(uso.get("dosis", "")),
            "observaciones": _texto_o_none(
                uso.get("observaciones", "")
            ),
            "pagina": uso.get("pagina"),
        })

    if not filas:
        return

    respuesta = requests.post(
        _endpoint("usos_producto"),
        headers=_headers("return=minimal"),
        json=filas,
        timeout=60,
    )

    _verificar_respuesta(respuesta, "guardar los usos del producto")


def obtener_usos_producto(producto_id=None, pdf=None):
    parametros = {
        "select": "*",
        "order": "pagina.asc.nullslast,id.asc",
    }

    if producto_id is not None:
        parametros["producto_id"] = f"eq.{int(producto_id)}"
    elif pdf is not None:
        parametros["pdf"] = f"eq.{pdf}"

    respuesta = requests.get(
        _endpoint("usos_producto"),
        headers=_headers(),
        params=parametros,
        timeout=60,
    )

    _verificar_respuesta(respuesta, "obtener los usos del producto")

    return _dataframe(respuesta.json())


# -------------------------------------------------------------------
# EXPERIENCIAS DE CAMPO
# -------------------------------------------------------------------

def guardar_experiencia_campo(
    producto_id,
    producto_nombre,
    cultivo,
    problema,
    dosis_usada,
    horario_aplicacion,
    condicion_climatica,
    resultado_observado,
    comentario,
    fecha
):
    datos = {
        "producto_id": int(producto_id),
        "producto_nombre": _texto_o_none(producto_nombre),
        "cultivo": _texto_o_none(cultivo),
        "problema": _texto_o_none(problema),
        "dosis_usada": _texto_o_none(dosis_usada),
        "horario_aplicacion": _texto_o_none(horario_aplicacion),
        "condicion_climatica": _texto_o_none(condicion_climatica),
        "resultado_observado": _texto_o_none(resultado_observado),
        "comentario": _texto_o_none(comentario),
        "fecha": _texto_o_none(fecha),
        "creado_en": datetime.now(timezone.utc).isoformat(),
    }

    respuesta = requests.post(
        _endpoint("experiencias_campo"),
        headers=_headers("return=minimal"),
        json=datos,
        timeout=60,
    )

    _verificar_respuesta(
        respuesta,
        "guardar la experiencia de campo"
    )


def obtener_experiencias_producto(producto_id):
    respuesta = requests.get(
        _endpoint("experiencias_campo"),
        headers=_headers(),
        params={
            "select": "*",
            "producto_id": f"eq.{int(producto_id)}",
            "order": "id.desc",
        },
        timeout=60,
    )

    _verificar_respuesta(
        respuesta,
        "obtener las experiencias del producto"
    )

    return _dataframe(respuesta.json())


def obtener_todas_experiencias():
    respuesta = requests.get(
        _endpoint("experiencias_campo"),
        headers=_headers(),
        params={
            "select": "*",
            "order": "id.desc",
        },
        timeout=60,
    )

    _verificar_respuesta(
        respuesta,
        "obtener las experiencias de campo"
    )

    return _dataframe(respuesta.json())


def actualizar_experiencia_campo(
    experiencia_id,
    cultivo,
    problema,
    dosis_usada,
    horario_aplicacion,
    condicion_climatica,
    resultado_observado,
    comentario,
    fecha
):
    datos = {
        "cultivo": _texto_o_none(cultivo),
        "problema": _texto_o_none(problema),
        "dosis_usada": _texto_o_none(dosis_usada),
        "horario_aplicacion": _texto_o_none(horario_aplicacion),
        "condicion_climatica": _texto_o_none(condicion_climatica),
        "resultado_observado": _texto_o_none(resultado_observado),
        "comentario": _texto_o_none(comentario),
        "fecha": _texto_o_none(fecha),
    }

    respuesta = requests.patch(
        _endpoint("experiencias_campo"),
        headers=_headers("return=minimal"),
        params={
            "id": f"eq.{int(experiencia_id)}",
        },
        json=datos,
        timeout=60,
    )

    _verificar_respuesta(
        respuesta,
        "actualizar la experiencia de campo"
    )


def eliminar_experiencia_campo(experiencia_id):
    respuesta = requests.delete(
        _endpoint("experiencias_campo"),
        headers=_headers("return=minimal"),
        params={
            "id": f"eq.{int(experiencia_id)}",
        },
        timeout=60,
    )

    _verificar_respuesta(
        respuesta,
        "eliminar la experiencia de campo"
    )
