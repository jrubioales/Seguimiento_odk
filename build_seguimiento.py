"""Genera docs/index.html: seguimiento de levantamientos ODK en un solo archivo.

Lee de ODK Central tres formularios y sus listas de entidades (el universo de
lo que cada mercaderista tiene que levantar) y produce una pagina estatica con
los datos embebidos. No necesita servidor: GitHub Actions la regenera y GitHub
Pages la sirve.

    Formulario                               Universo (lista de entidades)      Periodo esperado
    seguimiento_muebles_alquilados_uiux_v4   muebles_por_local_lista            un registro por local cada mes
    acuerdos_comerciales                     acuerdos_por_local_lista           un registro por local y acuerdo cada 2 semanas en su vigencia
    promociones_mercaderista                 locales_promociones_lista          un registro por local cada semana (lun-dom)
    extravisibilidad_ganados_moderno         ganadas_moderno                    exhibiciones activas, ganadas y perdidas por semana y canal

Credenciales por variables de entorno (ODK_EMAIL, ODK_PASSWORD) o con
``--env ruta/.env``. Si SEGUIMIENTO_CLAVE esta definida, los datos se cifran
(AES-GCM) y la pagina pide esa clave para abrirse.
"""

from __future__ import annotations

import argparse
import base64
import csv
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests

AQUI = Path(__file__).resolve().parent
PLANTILLA = AQUI / "plantilla.html"
PLANTILLA_PPT = AQUI / "plantilla_ppt.pptx"    # portada de formato.pptx; la pagina agrega las fotos
SALIDA = AQUI / "docs" / "index.html"

ECUADOR = timezone(timedelta(hours=-5))  # sin horario de verano
TIMEOUT = 180

FORM_VIS = "seguimiento_muebles_alquilados_uiux_v4"
FORM_AC = "acuerdos_comerciales"
FORM_PR = "promociones_mercaderista"
FORM_EX = "extravisibilidad_ganados_moderno"
FORM_SOS = "sos_participacion_percha_moderno"
FORM_PL = "planes_comerciales"
FORM_CO = "check_outs"
FORM_KB = "kbds"

MESES_VIS = 2                      # meses que muestra visibilidad (el actual y el anterior)
SEMANAS_PROMO = 4                  # semanas hacia atras que muestra promociones
PERIODOS_AC = 8                    # periodos hacia atras que muestra acuerdos
INICIO_QUINCENA_AC = date(2026, 11, 9)  # desde este lunes acuerdos se mide cada 2 semanas (antes, semanal)
SEMANAS_EX = 12                    # semanas ISO hacia atras que muestra exhibiciones ganadas
MESES_MENS = 2                     # meses que muestran SOS, planes, check outs y KBDS
MAX_DETALLE = 6                    # novedades listadas por registro
FOTOS_RAMA = os.environ.get("FOTOS_RAMA", "fotos")  # rama del repo donde viven las fotos
MINIATURA = 400                    # px del lado largo de la miniatura de la galeria
ITERACIONES = 250_000              # PBKDF2 de la clave de la pagina

# Exhibiciones de prueba que siguen en ganadas_moderno: no se cuentan.
EXCLUIR_EX = {
    "f15f1235-79f2-47e0-bccc-1f975b9ac217",  # HIPERMARKET BUIJO · Cabecera jabón
    "941e6ea5-0335-43d1-bfe1-965c0cf67807",  # HIPERMARKET BUIJO · Ruma aceite
    "336f8cdd-f895-4cdd-af6f-51a1ea1098c1",  # LA ECUATORIANA · Ruma Skittles
}

MOTIVOS_AC = {
    "sin_stock": "sin stock",
    "sin_alcance": "sin alcance",
    "espacio_limitado": "espacio limitado",
    "administrador": "no autoriza el administrador",
}
MOTIVOS_VIS = {
    "otro": "otro",
    "retirado_tienda": "retirado por la tienda",
    "nunca_llego": "nunca llegó",
    "convenio_terminado": "convenio terminado",
    "retirado_por_dano": "retirado por daño",
    "movido": "movido de lugar",
}
DANIOS = {"roto": "roto", "falta_pieza": "falta pieza", "golpe": "golpe",
          "pintura": "pintura", "otro": "otro"}
ESTADOS_PR = {
    "inactivo": "promoción no activa",
    "agotado": "agotado",
    "activo_distinto": "precio/descuento distinto",
    "sin_alcance": "sin alcance",
}

# Negocio y categoria por palabra clave del producto (o del texto del mueble).
# Gana la primera regla que coincide, asi que lo especifico va antes que lo
# generico (ACHIOTE antes que ACEITE, TOCADOR antes que JABON). Fuentes: la
# hoja COMPLETO de ACUERDOS COMERCIALES.xlsx y los muebles de visibilidad.
# None en negocio = se desconoce; el texto del mueble de visibilidad ya trae
# su propio negocio.
CATALOGO: list[tuple[tuple[str, ...], str | None, str]] = [
    (("ACHIOTE",), "ALES", "ACHIOTE"),
    (("MARGARINA", "NATURELLA"), "ALES", "MARGARINAS"),
    (("MANTECA", "TRES CORONAS"), "ALES", "MANTECAS"),
    (("CALDO", "SAZON"), "ALES", "SAZONADORES"),
    (("PODEROX", "QUITAMANCHA"), "ALES", "QUITAMANCHAS"),
    (("ZK12", "TOCADOR"), None, "JABÓN DE TOCADOR"),
    (("JABON",), "ALES", "JABÓN DE LAVAR"),
    (("ALESOL", "GIRASOLI", "ACEITE", "ACEIRE"), "ALES", "ACEITES"),
    (("ZAGAZ", "LAVAPLATO", "LAVAVAJILLA", "LAVP"), "ILP", "LAVAVAJILLAS"),
    (("FAB", "DETERGENTE", "DET "), "ILP", "DETERGENTES"),
    (("PEDIGREE", "WHISKA", "MASCOTA"), "MARS", "MASCOTAS"),
    (("M&M", "SNICKERS", "SKITTLES", "ORBIT", "MILKY", "TWIX", "CHOCOLATE",
      "CARAMELO", "GOMA", "CHICLE", "CONFITERIA"), "MARS", "CONFITERÍA"),
    (("H&S", "H&G", "HEAD", "PANTENE", "HERBAL", "HAIR", "CAPILAR"), "P&G", "CUIDADO CAPILAR"),
    (("GILLETTE", "GILLETE", "PRESTOBARBA", "SHAVE", "AFEITADO"), "P&G", "AFEITADO"),
    (("ORAL",), "P&G", "CUIDADO ORAL"),
    (("OLD SPICE", "SECRET", "DESODORANTE", "APDO"), "P&G", "DESODORANTES"),
    (("GALLETA",), "SUPERIOR", "GALLETAS"),
]
SIN_NEGOCIO = "Sin clasificar"
_REGLAS = [(re.compile("|".join(rf"(?<![A-Z0-9]){re.escape(k)}" for k in claves)), neg, cat)
           for claves, neg, cat in CATALOGO]


# --- ODK ------------------------------------------------------------------------


class Odk:
    def __init__(self, base: str, proyecto: int, email: str, password: str):
        self.base = base.rstrip("/")
        self.proyecto = proyecto
        r = requests.post(f"{self.base}/v1/sessions",
                          json={"email": email, "password": password}, timeout=TIMEOUT)
        if r.status_code >= 400:
            raise SystemExit(f"Login ODK fallido ({r.status_code}): {r.text[:200]}")
        self.h = {"Authorization": f"Bearer {r.json()['token']}"}

    def _get(self, url: str, intentos: int = 3) -> requests.Response:
        """GET con reintentos: el servidor a veces corta descargas largas."""
        for i in range(intentos):
            try:
                r = requests.get(url, headers=self.h, timeout=TIMEOUT)
                if r.status_code < 500:
                    r.raise_for_status()
                    return r
            except (requests.ConnectionError, requests.Timeout, requests.exceptions.ChunkedEncodingError):
                if i == intentos - 1:
                    raise
            time.sleep(5 * (i + 1))
        r.raise_for_status()
        return r

    def _form(self, form: str) -> str:
        return f"{self.base}/v1/projects/{self.proyecto}/forms/{form}"

    def registros(self, form: str, desde: date | None = None, expandir: bool = False) -> list[dict]:
        params = ["%24top=500"]
        if desde:
            params.append(f"%24filter=__system/submissionDate%20ge%20{desde.isoformat()}T00:00:00-05:00")
        if expandir:
            params.append("%24expand=*")
        url = f"{self._form(form)}.svc/Submissions?{'&'.join(params)}"
        filas: list[dict] = []
        for _ in range(1000):  # tope ante un nextLink circular
            cuerpo = self._get(url).json()
            filas.extend(cuerpo.get("value", []))
            url = cuerpo.get("@odata.nextLink")
            if not url:
                break
        return filas

    def xml(self, form: str) -> str:
        return self._get(f"{self._form(form)}.xml").content.decode("utf-8")

    def lista(self, form: str, nombre: str) -> list[dict]:
        """Adjunto CSV del formulario (las listas de entidades se publican asi)."""
        r = self._get(f"{self._form(form)}/attachments/{nombre}")
        return list(csv.DictReader(io.StringIO(r.content.decode("utf-8-sig"))))


def leer_env(ruta: str) -> None:
    for linea in Path(ruta).read_text(encoding="utf-8-sig").splitlines():
        if "=" in linea and not linea.lstrip().startswith("#"):
            k, v = linea.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


# --- utilidades -----------------------------------------------------------------


def g(d: dict, ruta: str):
    """Lee ``a/b/c`` de un registro OData anidado."""
    for parte in ruta.split("/"):
        if not isinstance(d, dict):
            return None
        d = d.get(parte)
    return d


def num(v) -> int:
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 0


def limpio(v) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip()


def titulo(nombre: str) -> str:
    """``CUEVA CUSTODIO JOSE (1711...)`` -> ``Cueva Custodio Jose`` (sin cedula)."""
    nombre = re.sub(r"\(.*?\)", "", nombre or "").strip()
    return " ".join(p.capitalize() for p in nombre.split())


def cedula_de(submitter: str) -> str:
    m = re.search(r"\((\d{6,})\)", submitter or "")
    return m.group(1) if m else ""


def fecha_de(v) -> date | None:
    try:
        return date.fromisoformat(str(v)[:10])
    except ValueError:
        return None


def lunes(d: date) -> date:
    return d - timedelta(days=d.weekday())


def fotos_incompletas(reg: dict) -> int:
    s = reg.get("__system") or {}
    esperadas, presentes = s.get("attachmentsExpected"), s.get("attachmentsPresent")
    if esperadas is None or presentes is None:
        return 0
    return max(0, num(esperadas) - num(presentes))


def _clave(texto: str) -> str:
    sin_tildes = unicodedata.normalize("NFD", texto or "").encode("ascii", "ignore").decode()
    return sin_tildes.upper()


def clasificar(texto: str) -> tuple[str, str]:
    """Producto -> (negocio, categoria) segun CATALOGO; la primera regla gana."""
    t = _clave(texto)
    for patron, neg, cat in _REGLAS:
        if patron.search(t):
            return neg or SIN_NEGOCIO, cat
    return SIN_NEGOCIO, "SIN CATEGORÍA"


def categorias_de(texto: str) -> list[str]:
    """Todas las categorias que menciona un texto (un mueble puede tener varias)."""
    t = _clave(texto)
    return [cat for patron, _, cat in _REGLAS if patron.search(t)]


def ordenados(valores) -> list[str]:
    return sorted({v for v in valores if v})


def mas_comun(valores) -> str:
    c = Counter(v for v in valores if v)
    return c.most_common(1)[0][0] if c else ""


# --- personal -------------------------------------------------------------------


class Personal:
    """Cedula -> nombre y supervisor, sin exponer la cedula en la pagina."""

    def __init__(self):
        self.nombre: dict[str, Counter] = defaultdict(Counter)
        self.sup: dict[str, Counter] = defaultdict(Counter)
        # Respaldos cuando la cedula no esta en ninguna lista con supervisor.
        self.sup_local: dict[str, Counter] = defaultdict(Counter)
        self.sup_zona: dict[str, Counter] = defaultdict(Counter)

    def agregar(self, cedula: str, nombre: str = "", supervisor: str = "",
                local: str = "", zona: str = ""):
        supervisor = limpio(supervisor)
        if supervisor and supervisor != "Sin Alcance":
            if local:
                self.sup_local[local][supervisor] += 1
            if zona:
                self.sup_zona[limpio(zona)][supervisor] += 1
        if not cedula:
            return
        if nombre:
            self.nombre[cedula][titulo(nombre)] += 1
        if supervisor:
            self.sup[cedula][supervisor] += 1

    def de(self, cedula: str, local: str = "", zona: str = "") -> tuple[str, str]:
        n = self.nombre.get(cedula)
        s = self.sup.get(cedula) or self.sup_local.get(local) or self.sup_zona.get(limpio(zona))
        return (n.most_common(1)[0][0] if n else "Sin asignar",
                s.most_common(1)[0][0] if s else "Sin asignar")


# --- estados --------------------------------------------------------------------

# ok = registrado sin novedad, nov = registrado con novedad,
# pen = sin registro y el periodo sigue abierto, no = sin registro y el periodo cerro,
# prog = el periodo aun no empieza.


def estado(registrado: bool, con_novedad: bool, abierto: bool, futuro: bool = False) -> str:
    if registrado:
        return "nov" if con_novedad else "ok"
    if futuro:
        return "prog"
    return "pen" if abierto else "no"


# --- visibilidad ----------------------------------------------------------------


def corto_mueble(label: str) -> str:
    """``RISTRAS | RISTRAS | MARS | CONFITERIA | ... | Cantidad: 1`` -> ``RISTRAS · MARS``."""
    partes = [p.strip() for p in (label or "").split("|")]
    if len(partes) >= 3:
        return f"{partes[0]} · {partes[2]}"
    return partes[0] if partes else ""


def negocio_categoria_mueble(label: str) -> tuple[str, list[str]]:
    """``TIPO | MUEBLE | NEGOCIO | CATEGORIA | MARCAS | Cantidad`` -> (negocio, categorias)."""
    partes = [p.strip() for p in (label or "").split("|")]
    negocio = partes[2].upper() if len(partes) > 2 and partes[2] else SIN_NEGOCIO
    texto = " ".join(partes[3:5])
    cats = categorias_de(texto)
    if not cats:
        cats = [partes[3].upper()] if len(partes) > 3 and partes[3] else ["SIN CATEGORÍA"]
    return negocio, cats


def novedades_vis(reg: dict) -> list[str]:
    out = []
    for k in range(1, 10):
        label = reg.get(f"item_{k}_label")
        if not label:
            continue
        grp = reg.get(f"grp_item_{k}") or {}
        m = corto_mueble(label)
        if grp.get(f"item_{k}_presente") == "no":
            tipos = ", ".join(MOTIVOS_VIS.get(t, t) for t in (grp.get(f"item_{k}_motivo_tipo") or "").split())
            det = limpio(grp.get(f"item_{k}_motivo"))
            out.append(f"{m}: no está en el local" + (f" ({tipos})" if tipos else "") + (f" — {det}" if det else ""))
            continue
        if grp.get(f"item_{k}_lleno") == "no":
            out.append(f"{m}: mueble sin producto / no está lleno")
        if grp.get(f"item_{k}_danio") == "si":
            tipos = ", ".join(DANIOS.get(t, t) for t in (grp.get(f"item_{k}_danio_tipo") or "").split())
            det = limpio(grp.get(f"item_{k}_danio_detalle"))
            out.append(f"{m}: dañado" + (f" ({tipos})" if tipos else "") + (f" — {det}" if det else ""))
    faltan = fotos_incompletas(reg)
    if faltan:
        out.append(f"{faltan} foto(s) sin terminar de subir al servidor")
    return out


def armar_visibilidad(odk: Odk, personal: Personal, hoy: date) -> dict:
    universo = odk.lista(FORM_VIS, "muebles_por_local_lista.csv")
    inicio = hoy.replace(day=1)
    for _ in range(MESES_VIS - 1):
        inicio = (inicio - timedelta(days=1)).replace(day=1)
    registros = odk.registros(FORM_VIS, desde=inicio - timedelta(days=1))
    for r in registros:
        personal.agregar(cedula_de(g(r, "__system/submitterName")), g(r, "__system/submitterName"))

    meses = []
    m = inicio
    while m <= hoy:
        meses.append(m.strftime("%Y-%m"))
        m = (m.replace(day=28) + timedelta(days=4)).replace(day=1)

    por_local_mes: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for r in registros:
        f = fecha_de(r.get("fecha_hoy") or g(r, "__system/submissionDate"))
        if f:
            por_local_mes[(r.get("local_id"), f.strftime("%Y-%m"))].append(r)

    llaves = {u["llave"] for u in universo}
    filas = []
    # Quien levanta cada local, por si su cedula no esta en las otras listas.
    subio_local: dict[str, Counter] = defaultdict(Counter)
    for r in registros:
        subio_local[r.get("local_id")][titulo(g(r, "__system/submitterName"))] += 1

    for u in universo:
        merc, sup = personal.de(u.get("acc_1", ""), u["llave"], u.get("acc_zona", ""))
        if merc == "Sin asignar" and subio_local.get(u["llave"]):
            merc = subio_local[u["llave"]].most_common(1)[0][0]
        labels = [u.get(f"item_{k}_label") for k in range(1, 10) if u.get(f"item_{k}_label")]
        muebles = [corto_mueble(l) for l in labels]
        clases = [negocio_categoria_mueble(l) for l in labels]
        negocios = ordenados(n for n, _ in clases)
        categorias = ordenados(c for _, cs in clases for c in cs)
        for mes in meses:
            regs = sorted(por_local_mes.get((u["llave"], mes), []), key=lambda r: g(r, "__system/submissionDate") or "")
            ultimo = regs[-1] if regs else None
            novs = novedades_vis(ultimo) if ultimo else []
            abierto = mes == hoy.strftime("%Y-%m")
            filas.append({
                "p": mes,
                "id": u["llave"],
                "loc": limpio(u.get("local_label")),
                "cad": limpio(u.get("cadena_label")),
                "zon": limpio(u.get("acc_zona")),
                "sup": sup,
                "mer": merc,
                "neg": negocios,
                "cat": categorias,
                "obj": " | ".join(muebles),
                "e": estado(bool(ultimo), bool(novs), abierto),
                "f": (ultimo.get("fecha_hoy") or "")[:10] if ultimo else "",
                "por": titulo(g(ultimo, "__system/submitterName")) if ultimo else "",
                "n": len(regs),
                "nov": novs[:MAX_DETALLE],
            })

    extra = [r for r in registros if r.get("local_id") not in llaves]
    return {
        "titulo": "Visibilidad",
        "subtitulo": "Muebles alquilados · un registro por local cada mes",
        "unidad": "locales",
        "periodos": [{"id": m, "label": etiqueta_mes(m)} for m in meses][::-1],
        "periodo_tipo": "mes",
        "filas": filas,
        "fuera": fuera_de_lista(extra, "local_id", "local_label", "fecha_hoy"),
        "registros": len(registros),
    }


MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


def etiqueta_mes(m: str) -> str:
    a, mm = m.split("-")
    return f"{MESES[int(mm) - 1].capitalize()} {a}"


def fuera_de_lista(regs: list[dict], k_id: str, k_label: str, k_fecha: str) -> list[dict]:
    out = []
    for r in regs[-200:]:
        out.append({"id": g(r, k_id) or "", "loc": limpio(g(r, k_label)),
                    "f": str(g(r, k_fecha) or "")[:10],
                    "por": titulo(g(r, "__system/submitterName"))})
    return out


# --- acuerdos comerciales -------------------------------------------------------


def novedades_ac(reg: dict) -> list[str]:
    out = []
    for k in range(1, 6):
        nombre = reg.get(f"mueble_{k}_nombre")
        if not nombre:
            continue
        grp = reg.get(f"grp_mueble_{k}") or {}
        ej = grp.get(f"mueble_{k}_ejecutado")
        if ej == "no":
            motivos = ", ".join(MOTIVOS_AC.get(t, t) for t in (grp.get(f"mueble_{k}_motivo") or "").split())
            com = limpio(grp.get(f"mueble_{k}_comentario"))
            out.append(f"{nombre}: no ejecutado" + (f" ({motivos})" if motivos else "") + (f" — {com}" if com else ""))
            continue
        if ej != "si":
            continue
        for i in range(1, 5):
            item = reg.get(f"mueble_{k}_item_{i}_label")
            gi = reg.get(f"grp_mueble_{k}_item_{i}") or {}
            mot = gi.get(f"mueble_{k}_item_{i}_motivo")
            if item and mot:
                motivos = ", ".join(MOTIVOS_AC.get(t, t) for t in mot.split())
                com = limpio(gi.get(f"mueble_{k}_item_{i}_comentario"))
                out.append(f"{nombre} · {item}: falta ({motivos})" + (f" — {com}" if com else ""))
    if str(g(reg, "grp_local/vigente")) == "0":
        out.append("Registrado fuera de la vigencia del acuerdo")
    faltan = fotos_incompletas(reg)
    if faltan:
        out.append(f"{faltan} foto(s) sin terminar de subir al servidor")
    return out


def periodo_ac(d: date) -> date:
    """Inicio del periodo de acuerdos: semana ISO antes de INICIO_QUINCENA_AC, luego bloques de 14 dias."""
    if d >= INICIO_QUINCENA_AC:
        return INICIO_QUINCENA_AC + timedelta(days=14 * ((d - INICIO_QUINCENA_AC).days // 14))
    return lunes(d)


def dias_ac(s: date) -> int:
    return 14 if s >= INICIO_QUINCENA_AC else 7


def etiqueta_ac(s: date) -> str:
    """``Semanas 46-47 · 9 nov al 22 nov`` (o la etiqueta semanal antes del cambio)."""
    if dias_ac(s) == 7:
        return etiqueta_semana(s)
    fin = s + timedelta(days=13)
    return (f"Semanas {s.isocalendar().week}-{fin.isocalendar().week} · "
            f"{s.day} {MESES[s.month - 1][:3]} al {fin.day} {MESES[fin.month - 1][:3]}")


def armar_acuerdos(odk: Odk, personal: Personal, hoy: date) -> dict:
    universo = odk.lista(FORM_AC, "acuerdos_por_local_lista.csv")
    desde = min((fecha_de(u.get("vigencia_desde")) for u in universo if fecha_de(u.get("vigencia_desde"))), default=None)
    registros = odk.registros(FORM_AC, desde=desde - timedelta(days=1) if desde else None)
    for u in universo:
        personal.agregar(u.get("acc_1", ""), u.get("mercaderista"), u.get("supervisor"),
                         local=u.get("local_id", ""), zona=u.get("acc_zona", ""))

    # Registros por acuerdo y periodo de la visita (semana ISO o bloque de 2 semanas).
    por_llave_sem: dict[tuple[str, date], list[dict]] = defaultdict(list)
    for r in registros:
        f = fecha_de(r.get("fecha"))
        if f:
            por_llave_sem[(g(r, "grp_local/acuerdo_id"), periodo_ac(f))].append(r)

    semana_hoy = periodo_ac(hoy)
    semanas = [semana_hoy]
    while len(semanas) < PERIODOS_AC:
        semanas.append(periodo_ac(semanas[-1] - timedelta(days=1)))
    # Antes de la primera semana con registros el formulario no existia.
    con_datos = {s for (_, s) in por_llave_sem}
    semanas = [s for s in semanas if not con_datos or s >= min(con_datos)]

    filas = []
    for u in universo:
        desde, hasta = fecha_de(u.get("vigencia_desde")), fecha_de(u.get("vigencia_hasta"))
        if not (desde and hasta):
            continue
        muebles = [u.get(f"mueble_{k}_nombre") for k in range(1, 6) if u.get(f"mueble_{k}_nombre")]
        productos = [u.get(f"mueble_{k}_item_{i}_label") for k in range(1, 6) for i in range(1, 5)
                     if u.get(f"mueble_{k}_item_{i}_label")]
        clases = [clasificar(p) for p in productos]
        # Se espera un registro en cada periodo que toque la vigencia del acuerdo.
        for s in semanas:
            if not (desde <= s + timedelta(days=dias_ac(s) - 1) and hasta >= s):
                continue
            regs = sorted(por_llave_sem.get((u["llave"], s), []), key=lambda r: g(r, "__system/submissionDate") or "")
            vigentes = [r for r in regs if str(g(r, "grp_local/vigente")) != "0"]
            ultimo = (vigentes or regs or [None])[-1]
            novs = novedades_ac(ultimo) if ultimo else []
            filas.append({
                "p": s.isoformat(),
                "acu": u.get("acuerdo", ""),
                "id": u.get("local_id", ""),
                "loc": limpio(u.get("local_nombre")) + (f" · {u['ciudad']}" if u.get("ciudad") else ""),
                "cad": limpio(u.get("cadena_label")),
                "zon": limpio(u.get("acc_zona")),
                "sup": limpio(u.get("supervisor")) or "Sin asignar",
                "mer": titulo(u.get("mercaderista")) or "Sin asignar",
                "neg": ordenados(n for n, _ in clases) or [SIN_NEGOCIO],
                "cat": ordenados(c for _, c in clases) or ["SIN CATEGORÍA"],
                "obj": " | ".join(muebles) + (f" · {', '.join(dict.fromkeys(productos))}" if productos else ""),
                "vig": limpio(u.get("vigencia_label")),
                "e": estado(bool(ultimo), bool(novs), s == semana_hoy),
                "f": (ultimo.get("fecha") or "")[:10] if ultimo else "",
                "por": titulo(g(ultimo, "__system/submitterName")) if ultimo else "",
                "n": len(regs),
                "nov": novs[:MAX_DETALLE],
            })

    con_filas = {f["p"] for f in filas}
    llaves = {u["llave"] for u in universo}
    extra = [r for r in registros if g(r, "grp_local/acuerdo_id") not in llaves]
    return {
        "titulo": "Acuerdos comerciales",
        "subtitulo": "Un registro por local y acuerdo cada 2 semanas mientras esté vigente (hasta el 8 nov era cada semana)",
        "unidad": "acuerdos",
        "periodos": [{"id": s.isoformat(), "label": etiqueta_ac(s)} for s in semanas if s.isoformat() in con_filas],
        "periodo_tipo": "semana",
        "filas": filas,
        "fuera": fuera_de_lista(extra, "grp_local/local_id", "grp_local/local_nombre", "fecha"),
        "registros": len(registros),
    }


# --- promociones ----------------------------------------------------------------


def novedades_pr(reg: dict) -> tuple[list[str], Counter]:
    out, cuenta = [], Counter()
    pantallas = g(reg, "grp_promos/pantalla") or []
    for pant in pantallas if isinstance(pantallas, list) else []:
        for clave, p in pant.items():
            if not (isinstance(p, dict) and re.fullmatch(r"p\d+", clave)):
                continue
            sfx = clave[1:]
            est = p.get(f"estado_{sfx}")
            if not est:
                continue
            cuenta[est] += 1
            if est in ESTADOS_PR and est != "sin_alcance":
                prod = limpio(p.get(f"producto_{sfx}"))
                desc = limpio(p.get(f"descuento_{sfx}"))
                txt = f"{prod} ({desc}): {ESTADOS_PR[est]}"
                sentido = p.get(f"sentido_{sfx}")
                if est == "activo_distinto" and sentido:
                    txt += f", {'mayor' if sentido == 'mayor' else 'menor'} al pactado"
                out.append(txt)
    if not cuenta:  # sin detalle de pantallas: usar los totales del cierre
        cuenta["activo"] = num(g(reg, "grp_cierre/n_activas"))
        for k, est in [("n_inactivas", "inactivo"), ("n_agotadas", "agotado"),
                       ("n_distintas", "activo_distinto"), ("n_sin_alcance", "sin_alcance")]:
            n = num(g(reg, f"grp_cierre/{k}"))
            if n:
                cuenta[est] += n
                if est != "sin_alcance":
                    out.append(f"{n} {'promoción' if n == 1 else 'promociones'}: {ESTADOS_PR[est]}")
        menor = num(g(reg, "grp_cierre/n_menor"))
        if menor:
            out.append(f"{menor} con descuento menor al pactado")
    return out, cuenta


def armar_promociones(odk: Odk, personal: Personal, hoy: date) -> dict:
    universo = odk.lista(FORM_PR, "locales_promociones_lista.csv")
    # Promociones pactadas por cadena (grupo) con su vigencia: de ahi salen
    # el negocio y la categoria que tiene que revisar cada local cada semana.
    promos_grupo: dict[str, list[tuple[str, str, str, str]]] = defaultdict(list)
    for it in odk.lista(FORM_PR, "promociones_items.csv"):
        neg, cat = clasificar(f"{it.get('producto', '')} {it.get('categoria', '')}")
        promos_grupo[it.get("grupo", "")].append((it.get("inicio", ""), it.get("fin", ""), neg, cat))
    semana_hoy = lunes(hoy)
    semanas = [semana_hoy - timedelta(weeks=i) for i in range(SEMANAS_PROMO)]
    # Sin $expand: el detalle producto por producto pesa >100 MB; los totales
    # del cierre bastan para saber si hubo novedad y de que tipo.
    registros = odk.registros(FORM_PR, desde=semanas[-1] - timedelta(days=1))
    for u in universo:
        personal.agregar(u.get("acc_1", ""), u.get("mercaderistas"), u.get("supervisor"),
                         zona=u.get("zona") or u.get("acc_zona", ""))

    por_local_sem: dict[tuple[str, date], list[dict]] = defaultdict(list)
    for r in registros:
        f = fecha_de(r.get("fecha"))
        if f:
            por_local_sem[(g(r, "grp_local/local_id"), lunes(f))].append(r)

    # La primera semana con datos marca el arranque del formulario.
    con_datos = {s for (_, s) in por_local_sem}
    semanas = [s for s in semanas if not con_datos or s >= min(con_datos)]

    filas = []
    for u in universo:
        for s in semanas:
            regs = sorted(por_local_sem.get((u["llave"], s), []), key=lambda r: g(r, "__system/submissionDate") or "")
            ultimo = regs[-1] if regs else None
            novs, cuenta = novedades_pr(ultimo) if ultimo else ([], Counter())
            ini, fin = s.isoformat(), (s + timedelta(days=6)).isoformat()
            vigentes = [(n, c) for d, h, n, c in promos_grupo.get(u.get("grupo", ""), []) if d <= fin and h >= ini]
            filas.append({
                "p": s.isoformat(),
                "id": u["llave"],
                "loc": limpio(u.get("local_nombre")) + (f" · {u['ciudad']}" if u.get("ciudad") else ""),
                "cad": limpio(u.get("grupo_label") or u.get("cadena_label")),
                "zon": limpio(u.get("zona") or u.get("acc_zona")),
                "sup": limpio(u.get("supervisor")) or "Sin asignar",
                "mer": titulo(u.get("mercaderistas")) or "Sin asignar",
                "neg": ordenados(n for n, _ in vigentes) or ["Sin promociones vigentes"],
                "cat": ordenados(c for _, c in vigentes) or ["SIN PROMOCIONES VIGENTES"],
                "obj": limpio(u.get("formato")),
                "e": estado(bool(ultimo), bool(novs), s == semana_hoy),
                "f": (ultimo.get("fecha") or "")[:10] if ultimo else "",
                "por": titulo(g(ultimo, "__system/submitterName")) if ultimo else "",
                "n": len(regs),
                "nov": novs[:MAX_DETALLE] + ([f"y {len(novs) - MAX_DETALLE} más"] if len(novs) > MAX_DETALLE else []),
                "c": {k: cuenta.get(k, 0) for k in ("activo", "inactivo", "agotado", "activo_distinto", "sin_alcance")},
            })

    llaves = {u["llave"] for u in universo}
    extra = [r for r in registros if g(r, "grp_local/local_id") not in llaves]
    return {
        "titulo": "Promociones",
        "subtitulo": "Un registro por local cada semana (lunes a domingo)",
        "unidad": "locales",
        "periodos": [{"id": s.isoformat(), "label": etiqueta_semana(s)} for s in semanas],
        "periodo_tipo": "semana",
        "filas": filas,
        "fuera": fuera_de_lista(extra, "grp_local/local_id", "grp_local/local_nombre", "fecha"),
        "registros": len(registros),
    }


# --- encuestas mensuales (SOS, planes, check outs, KBDS) -------------------------
#
# Un registro por local cada mes calendario: el objetivo se reinicia el dia 1.
# El universo es la lista de locales que trae el propio formulario, asi que si
# cambia la lista en ODK la meta cambia sola.


def zona_txt(v) -> str:
    """``sierra_norte`` / ``SIERRA NORTE`` -> ``SIERRA NORTE``."""
    return limpio(str(v or "").replace("_", " ")).upper()


def meses_desde(hoy: date, n: int) -> tuple[date, list[str]]:
    inicio = hoy.replace(day=1)
    for _ in range(n - 1):
        inicio = (inicio - timedelta(days=1)).replace(day=1)
    meses, m = [], inicio
    while m <= hoy:
        meses.append(m.strftime("%Y-%m"))
        m = (m.replace(day=28) + timedelta(days=4)).replace(day=1)
    return inicio, meses


def armar_mensual(odk: Odk, personal: Personal, hoy: date, asignado: dict[str, tuple[str, str]], *,
                  form: str, titulo_: str, subtitulo: str, unidad: str, universo: list[dict],
                  claves_reg, novedades, campo_local: str, campo_nombre: str,
                  can_titulo: str = "") -> dict:
    """``universo``: filas {k, id, loc, cad, zon, can, obj}; ``claves_reg(r)``: las k que cubre un registro."""
    inicio, meses = meses_desde(hoy, MESES_MENS)
    registros = odk.registros(form, desde=inicio - timedelta(days=1))
    por: dict[tuple[str, str], list[dict]] = defaultdict(list)
    subio_local: dict[str, Counter] = defaultdict(Counter)
    for r in registros:
        f = fecha_de(r.get("fecha") or g(r, "__system/submissionDate"))
        if not f:
            continue
        for k in claves_reg(r):
            por[(k, f.strftime("%Y-%m"))].append(r)
        subio_local[g(r, campo_local) or ""][titulo(g(r, "__system/submitterName"))] += 1
    # Los meses antes del primer registro son de antes de que existiera la encuesta.
    con_datos = {m for (_, m) in por}
    meses = [m for m in meses if not con_datos or m >= min(con_datos)]

    mes_hoy = hoy.strftime("%Y-%m")
    filas = []
    for u in universo:
        mer, sup = asignado.get(u["id"], ("Sin asignar", "Sin asignar"))
        if sup == "Sin asignar":
            sup = personal.de("", u["id"], u["zon"])[1]
        if mer == "Sin asignar" and subio_local.get(u["id"]):
            mer = subio_local[u["id"]].most_common(1)[0][0]
        for mes in meses:
            regs = sorted(por.get((u["k"], mes), []), key=lambda r: g(r, "__system/submissionDate") or "")
            ultimo = regs[-1] if regs else None
            novs = novedades(ultimo) if ultimo else []
            filas.append({
                "p": mes, "id": u["id"], "loc": u["loc"], "cad": u["cad"], "zon": u["zon"],
                "can": u.get("can", ""), "sup": sup, "mer": mer, "neg": [], "cat": [],
                "obj": u.get("obj", ""),
                "e": estado(bool(ultimo), bool(novs), mes == mes_hoy),
                "f": str(ultimo.get("fecha") or g(ultimo, "__system/submissionDate") or "")[:10] if ultimo else "",
                "por": titulo(g(ultimo, "__system/submitterName")) if ultimo else "",
                "n": len(regs),
                "nov": novs[:MAX_DETALLE],
            })

    claves = {u["k"] for u in universo}
    extra = [r for r in registros if not set(claves_reg(r)) & claves]
    return {
        "titulo": titulo_,
        "subtitulo": subtitulo,
        "unidad": unidad,
        "can_titulo": can_titulo,
        "periodos": [{"id": m, "label": etiqueta_mes(m)} for m in meses][::-1],
        "periodo_tipo": "mes",
        "filas": filas,
        "fuera": fuera_de_lista(extra, campo_local, campo_nombre, "fecha"),
        "registros": len(registros),
    }


def con_ciudad(nombre, ciudad) -> str:
    nombre, ciudad = limpio(nombre), limpio(ciudad)
    return f"{nombre} · {ciudad.title()}" if ciudad and ciudad.upper() not in nombre.upper() else nombre


def armar_sos(odk: Odk, personal: Personal, hoy: date, asignado) -> dict:
    universo = [{"k": u["llave"], "id": u["llave"], "loc": con_ciudad(u.get("local_nombre"), u.get("ciudad")),
                 "cad": limpio(u.get("cadena_nombre")), "zon": zona_txt(u.get("zona")), "obj": limpio(u.get("formato"))}
                for u in odk.lista(FORM_SOS, "tiendas_moderno_1008.csv")]
    motivos = {k: v["label"] for k, v in opciones_xml(odk.xml(FORM_SOS), "motivos_sin_categoria").items()}

    def novs(r):
        out = []
        if g(r, "sos_posible/puede_sos") == "no":
            m = limpio(g(r, "no_se_pudo/motivo_no_sos"))
            out.append("No se pudo medir la percha" + (f": {m}" if m else ""))
        if (m := limpio(g(r, "faltantes/faltantes_motivo"))):
            n = num(r.get("n_faltan"))
            out.append(f"{n or 'Hay'} categoría(s) que ya no están en la tienda: " + ", ".join(motivos.get(x, x) for x in m.split()))
        if (faltan := fotos_incompletas(r)):
            out.append(f"{faltan} foto(s) sin terminar de subir al servidor")
        return out

    return armar_mensual(odk, personal, hoy, asignado, form=FORM_SOS, titulo_="SOS percha",
                         subtitulo="Participación en percha · moderno · un registro por local cada mes",
                         unidad="locales", universo=universo,
                         claves_reg=lambda r: [g(r, "buscar_tienda/local_id") or ""], novedades=novs,
                         campo_local="buscar_tienda/local_id", campo_nombre="local_datos/local_nombre")


PLANES = {"ninja": ("plan_n", "Plan Ninja"), "diamante": ("plan_d", "Plan Aliado Diamante")}


def armar_planes(odk: Odk, personal: Personal, hoy: date, asignado) -> dict:
    universo = []
    for u in odk.lista(FORM_PL, "locales_planes_1028.csv"):
        for plan, (col, nombre) in PLANES.items():
            if u.get(col) == "1":
                universo.append({"k": f"{u['llave']}|{plan}", "id": u["llave"],
                                 "loc": con_ciudad(u.get("pdv"), u.get("ciudad")), "cad": limpio(u.get("cliente")),
                                 "zon": zona_txt(u.get("zona_nom") or u.get("zona")), "can": nombre,
                                 "obj": limpio(u.get("formato"))})
    motivos = {k: v["label"] for k, v in opciones_xml(odk.xml(FORM_PL), "motivos_sin_girasol").items()}

    def novs(r):
        out = []
        if g(r, "disponibilidad/hay_girasol") == "no":
            m = g(r, "sin_girasol/sin_girasol_motivo")
            det = limpio(g(r, "sin_girasol/sin_girasol_detalle"))
            out.append("Sin aceite de girasol" + (f": {motivos.get(m, m)}" if m else "") + (f" — {det}" if det else ""))
        if (faltan := fotos_incompletas(r)):
            out.append(f"{faltan} foto(s) sin terminar de subir al servidor")
        return out

    def claves(r):
        loc = g(r, "datos_local/local_id") or ""
        return [f"{loc}|{p}" for p in (g(r, "plan_grp/plan") or "").split()]

    return armar_mensual(odk, personal, hoy, asignado, form=FORM_PL, titulo_="Planes comerciales",
                         subtitulo="Plan Ninja y Plan Aliado Diamante · un registro por local y plan cada mes",
                         unidad="locales × plan", universo=universo, claves_reg=claves, novedades=novs,
                         campo_local="datos_local/local_id", campo_nombre="datos_local/local_nombre",
                         can_titulo="Todos los planes")


def armar_checkouts(odk: Odk, personal: Personal, hoy: date, asignado) -> dict:
    xml = odk.xml(FORM_CO)
    zonas = {k: v["label"] for k, v in opciones_xml(xml, "zonas").items()}
    canales = {"farma": "Farmacias", "gas": "Gasolineras"}
    universo = [{"k": u["name"], "id": u["name"], "loc": limpio(u.get("label")), "cad": limpio(u.get("cliente")),
                 "zon": zona_txt(zonas.get(u.get("zona"), u.get("zona"))),
                 "can": canales.get(u.get("canal"), limpio(u.get("canal"))),
                 "obj": "Con percha" if u.get("percha") == "1" else ""}
                for u in odk.lista(FORM_CO, "locales_checkouts.csv")]

    def novs(r):
        out = []
        if g(r, "farma_mod_grp/modulos") == "ninguno":
            out.append("No tiene check out")
        if (faltan := fotos_incompletas(r)):
            out.append(f"{faltan} foto(s) sin terminar de subir al servidor")
        return out

    return armar_mensual(odk, personal, hoy, asignado, form=FORM_CO, titulo_="Check outs",
                         subtitulo="Farmacias y gasolineras · un registro por local cada mes",
                         unidad="locales", universo=universo,
                         claves_reg=lambda r: [g(r, "datos_local/local_id") or ""], novedades=novs,
                         campo_local="datos_local/local_id", campo_nombre="local_nombre",
                         can_titulo="Todos los canales")


def armar_kbds(odk: Odk, personal: Personal, hoy: date, asignado) -> dict:
    universo = [{"k": u["llave"], "id": u["llave"], "loc": con_ciudad(u.get("pdv"), u.get("ciudad")),
                 "cad": limpio(u.get("cadena")), "zon": zona_txt(u.get("zona_nom") or u.get("zona")),
                 "obj": " · ".join(x for x in (limpio(u.get("formato")), limpio(u.get("cluster"))) if x)}
                for u in odk.lista(FORM_KB, "locales_kbds.csv")]
    motivos = {k: v["label"] for k, v in opciones_xml(odk.xml(FORM_KB), "motivos_no").items()}

    def novs(r):
        out = []
        if g(r, "levantar_grp/puede_levantar") == "no":
            m = g(r, "levantar_grp/motivo_no")
            otro = limpio(g(r, "levantar_grp/motivo_otro"))
            out.append("No se pudo levantar" + (f": {motivos.get(m, m)}" if m else "") + (f" — {otro}" if otro else ""))
        if (faltan := fotos_incompletas(r)):
            out.append(f"{faltan} foto(s) sin terminar de subir al servidor")
        return out

    return armar_mensual(odk, personal, hoy, asignado, form=FORM_KB, titulo_="KBDS",
                         subtitulo="Moderno · un registro por local cada mes",
                         unidad="locales", universo=universo,
                         claves_reg=lambda r: [g(r, "datos_local/local_id") or ""], novedades=novs,
                         campo_local="datos_local/local_id", campo_nombre="datos_local/local_nombre")


def asignacion(*formularios: dict) -> dict[str, tuple[str, str]]:
    """Codigo de local -> (mercaderista, supervisor) segun las pestañas que ya los conocen."""
    out: dict[str, tuple[str, str]] = {}
    for f in formularios:
        actual = f["periodos"][0]["id"] if f.get("periodos") else None
        for r in f["filas"]:
            if r.get("p") != actual:
                continue
            mer, sup = out.get(r["id"], ("Sin asignar", "Sin asignar"))
            if mer == "Sin asignar" and r["mer"] and not r["mer"].startswith("Sin "):
                mer = r["mer"]
            if sup == "Sin asignar" and r["sup"] and not r["sup"].startswith("Sin "):
                sup = r["sup"]
            out[r["id"]] = (mer, sup)
    return out


# --- exhibiciones ganadas -------------------------------------------------------
#
# Cada exhibicion ganada es una entidad de ganadas_moderno: se crea al ganarla
# (fecha_alta) y en cada visita el mercaderista confirma si sigue o la da de
# baja (fecha_baja, motivo). Si en una semana no se sube encuesta, la
# exhibicion sigue activa: solo una baja la termina.


def opciones_xml(xml: str, instancia: str) -> dict[str, dict[str, str]]:
    """Items de una lista de opciones del formulario: name -> {label, ...}."""
    import xml.etree.ElementTree as ET

    raiz = ET.fromstring(xml)
    out: dict[str, dict[str, str]] = {}
    for inst in raiz.iter():
        if inst.tag.endswith("}instance") and inst.get("id") == instancia:
            for item in inst.iter():
                if item.tag.endswith("item") or item.tag == "item":
                    campos = {c.tag.split("}")[-1]: (c.text or "").strip() for c in item}
                    if campos.get("name"):
                        out[campos["name"]] = campos
    return out


def preparar_foto(original: bytes) -> tuple[bytes, bytes, tuple[int, int]]:
    """Foto de ODK -> (grande, miniatura, (ancho, alto)), ya derecha segun EXIF."""
    from PIL import Image, ImageOps

    im = Image.open(io.BytesIO(original))
    derecha = im.format == "JPEG" and im.getexif().get(0x0112, 1) == 1
    im = ImageOps.exif_transpose(im).convert("RGB")
    if derecha:
        grande = original  # el formulario ya la limita a 1024 px
    else:
        b = io.BytesIO(); im.save(b, "JPEG", quality=90); grande = b.getvalue()
    mini = im.copy()
    mini.thumbnail((MINIATURA, MINIATURA))
    b = io.BytesIO(); mini.save(b, "JPEG", quality=55, optimize=True)
    return grande, b.getvalue(), im.size


def _git(args: list[str], cwd: Path, entrada: bytes | None = None, env: dict | None = None) -> bytes:
    r = subprocess.run(["git", *args], cwd=cwd, input=entrada, capture_output=True, env=env)
    if r.returncode:
        raise RuntimeError(f"git {args[0]}: {r.stderr.decode(errors='ignore').strip()[:300]}")
    return r.stdout


def repo_github() -> str | None:
    if os.environ.get("GITHUB_REPOSITORY"):
        return os.environ["GITHUB_REPOSITORY"]
    try:
        url = _git(["remote", "get-url", "origin"], AQUI).decode().strip()
    except (RuntimeError, OSError):
        return None
    m = re.search(r"github\.com[:/]([^/]+/[^/]+?)(\.git)?$", url)
    return m.group(1) if m else None


def _aes(clave: str, sal: bytes, iteraciones: int):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    return AESGCM(PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=sal, iterations=iteraciones).derive(clave.encode()))


def sincronizar_fotos(odk: Odk, origen: dict[str, tuple[str, str]], clave: str | None) -> dict:
    """Sube a la rama de fotos las que faltan (grande + miniatura) y devuelve
    lo que la pagina necesita para mostrarlas: {base, cifrado, sal, it, ext, dims}.

    La rama solo recibe commits cuando hay fotos nuevas, y se baja sin el
    contenido de las fotos (--filter=blob:none): cada corrida descarga de ODK
    solo lo nuevo. Con SEGUIMIENTO_CLAVE cada archivo va cifrado (AES-GCM,
    12 bytes de IV + datos) con una clave derivada de la sal del manifest.
    """
    repo = repo_github()
    if not repo:
        print("Aviso: sin repositorio de GitHub, no se publican fotos")
        return {}
    from cryptography.exceptions import InvalidTag

    b64 = lambda b: base64.b64encode(b).decode()
    tmp = Path(tempfile.mkdtemp(prefix="fotos_"))
    git = lambda *a, **k: _git(list(a), tmp, **k)
    manifest: dict = {}
    nuevas: dict[str, tuple[bytes, bytes]] = {}
    try:
        git("init", "-q", "--bare")
        # FOTOS_REMOTO / FOTOS_BASE: para probar con un repositorio y un servidor locales.
        git("remote", "add", "origin", os.environ.get("FOTOS_REMOTO") or f"https://github.com/{repo}.git")
        for k, v in [("core.repositoryformatversion", "1"), ("extensions.partialClone", "origin"),
                     ("remote.origin.promisor", "true"), ("remote.origin.partialclonefilter", "blob:none")]:
            git("config", k, v)
        # En GitHub Actions el token de actions/checkout vive en la config del repo.
        try:
            for linea in _git(["config", "--get-regexp", r"^http\..*\.extraheader$"], AQUI).decode().splitlines():
                k, v = linea.split(" ", 1)
                git("config", k, v)
        except RuntimeError:
            pass

        padre = None
        if git("ls-remote", "--heads", "origin", FOTOS_RAMA).strip():
            git("fetch", "-q", "--depth", "1", "--filter=blob:none", "origin", FOTOS_RAMA)
            padre = git("rev-parse", "FETCH_HEAD").decode().strip()
            try:
                manifest = json.loads(git("cat-file", "-p", f"{padre}:manifest.json"))
            except (RuntimeError, ValueError):
                manifest = {}

        aes = None
        valido = bool(manifest) and manifest.get("cifrado") == bool(clave)
        if valido and clave:
            aes = _aes(clave, base64.b64decode(manifest["sal"]), manifest["it"])
            chk = base64.b64decode(manifest["check"])
            try:
                valido = aes.decrypt(chk[:12], chk[12:], None) == b"ok"
            except InvalidTag:
                valido = False
        if not valido:  # primera vez o cambio la clave: se rehace la rama
            manifest, padre = {"cifrado": bool(clave), "fotos": {}}, None
            if clave:
                sal = os.urandom(16)
                aes = _aes(clave, sal, ITERACIONES)
                iv = os.urandom(12)
                manifest.update(sal=b64(sal), it=ITERACIONES, check=b64(iv + aes.encrypt(iv, b"ok", None)))

        def sellar(datos: bytes) -> bytes:
            if not aes:
                return datos
            iv = os.urandom(12)
            return iv + aes.encrypt(iv, datos, None)

        dims: dict = manifest.setdefault("fotos", {})
        for ent, (instancia, nombre) in origen.items():
            if ent in dims:
                continue
            try:
                original = odk._get(f"{odk._form(FORM_EX)}/submissions/{instancia}/attachments/{nombre}").content
                grande, mini, tam = preparar_foto(original)
            except Exception as e:  # una foto mala no detiene el resto
                print(f"Aviso: foto {nombre} no disponible ({str(e)[:80]})")
                continue
            nuevas[ent] = (grande, mini)
            dims[ent] = list(tam)

        if nuevas or (not padre and dims):
            ext = ".bin" if aes else ".jpg"
            env = {**os.environ, "GIT_INDEX_FILE": str(tmp / "indice"),
                   "GIT_AUTHOR_NAME": "seguimiento", "GIT_AUTHOR_EMAIL": "seguimiento@users.noreply.github.com",
                   "GIT_COMMITTER_NAME": "seguimiento", "GIT_COMMITTER_EMAIL": "seguimiento@users.noreply.github.com"}
            git("read-tree", padre or "--empty", env=env)

            def agregar(ruta: str, datos: bytes):
                sha = git("hash-object", "-w", "--stdin", entrada=datos).decode().strip()
                git("update-index", "--add", "--cacheinfo", f"100644,{sha},{ruta}", env=env)

            for ent, (grande, mini) in nuevas.items():
                agregar(f"g/{ent}{ext}", sellar(grande))
                agregar(f"t/{ent}{ext}", sellar(mini))
            agregar("manifest.json", json.dumps(manifest, separators=(",", ":")).encode())
            arbol = git("write-tree", env=env).decode().strip()
            commit = git("commit-tree", arbol, "-m", f"{len(nuevas)} foto(s) nuevas",
                         *(["-p", padre] if padre else []), env=env).decode().strip()
            git("push", "-q", *([] if padre else ["-f"]), "origin", f"{commit}:refs/heads/{FOTOS_RAMA}")
            print(f"Fotos: {len(nuevas)} nuevas, {len(dims)} en total")
    except Exception as e:
        print(f"Aviso: no se pudieron actualizar las fotos ({str(e)[:200]})")
        # Solo se muestran las que ya estaban publicadas.
        for ent in nuevas:
            manifest.get("fotos", {}).pop(ent, None)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if not manifest.get("fotos"):
        return {}
    return {
        "base": os.environ.get("FOTOS_BASE") or f"https://raw.githubusercontent.com/{repo}/{FOTOS_RAMA}/",
        "cifrado": manifest.get("cifrado", False),
        "sal": manifest.get("sal"),
        "it": manifest.get("it"),
        "ext": ".bin" if manifest.get("cifrado") else ".jpg",
        "dims": manifest["fotos"],
    }


def armar_exhibiciones(odk: Odk, personal: Personal, hoy: date, clave: str | None = None,
                       con_fotos: bool = True) -> dict:
    xml = odk.xml(FORM_EX)
    categorias = {k: v.get("label", "").upper() for k, v in opciones_xml(xml, "categorias").items()}
    sku_por_ean = {v.get("ean", ""): v for v in opciones_xml(xml, "skus").values()}
    motivos = {k: v.get("label", "") for k, v in opciones_xml(xml, "motivos_baja").items()}

    entidades = odk.lista(FORM_EX, "ganadas_moderno.csv")
    registros = odk.registros(FORM_EX, expandir=True)
    # Quien la gano: el que subio el registro que creo la entidad.
    gano: dict[str, str] = {}
    # Respaldo si ese registro no esta: el ultimo que levanto el local.
    ultimo_local: dict[str, str] = {}
    # Fotos: una al ganarla y, si sigue, otra cada 7 dias o mas (hasta 2 al
    # mes). foto -> (registro, archivo); entidad -> [(fecha, foto, es_seguimiento)].
    # La de alta se llama como la entidad; las de seguimiento, entidad_AAAAMMDD.
    origen_foto: dict[str, tuple[str, str]] = {}
    fotos_de: dict[str, list[tuple[str, str, int]]] = defaultdict(list)
    for r in sorted(registros, key=lambda r: g(r, "__system/submissionDate") or ""):
        ultimo_local[g(r, "grp_local/local_id") or ""] = g(r, "__system/submitterName") or ""
        fecha = str(r.get("fecha") or "")[:10]
        for e in r.get("exhibiciones") or []:
            if e.get("modo") == "nueva" and e.get("nuevo_id"):
                gano[e["nuevo_id"]] = g(r, "__system/submitterName") or ""
                if g(e, "grp_foto/foto") and e["nuevo_id"] not in EXCLUIR_EX:
                    origen_foto[e["nuevo_id"]] = (r["__id"], g(e, "grp_foto/foto"))
                    fotos_de[e["nuevo_id"]].append((fecha, e["nuevo_id"], 0))
            seg = g(e, "grp_seguimiento/foto_seguimiento")
            if e.get("modo") == "seguimiento" and seg and e.get("ent_id") and e["ent_id"] not in EXCLUIR_EX:
                pid = f"{e['ent_id']}_{fecha.replace('-', '')}"
                if pid not in origen_foto:
                    fotos_de[e["ent_id"]].append((fecha, pid, 1))
                origen_foto[pid] = (r["__id"], seg)

    filas = []
    for u in entidades:
        alta = fecha_de(u.get("fecha_alta"))
        if not alta or u.get("name") in EXCLUIR_EX:
            continue
        baja = fecha_de(u.get("fecha_baja")) if u.get("estado") == "baja" else None
        eans = (u.get("eans") or "").split()
        nombres = [s.strip() for s in (u.get("skus") or "").split(" + ")]
        skus, cats = [], []
        for i, ean in enumerate(eans):
            info = sku_por_ean.get(ean, {})
            skus.append(limpio(info.get("label") or (nombres[i] if i < len(nombres) else ean)))
            cats.append(categorias.get(info.get("categoria", ""), "SIN CATEGORÍA"))
        submitter = gano.get(u.get("name", "")) or ultimo_local.get(u.get("local_id", ""), "")
        merc, sup = personal.de(cedula_de(submitter), u.get("local_id", ""), u.get("zona", ""))
        if merc == "Sin asignar" and submitter:
            merc = titulo(submitter)
        filas.append({
            "id": u.get("name", ""),
            "lid": u.get("local_id", ""),
            "loc": limpio(u.get("local_nombre")),
            "cad": limpio(u.get("cadena")),
            # Antes de agregar canales el formulario era solo de canal moderno.
            "can": limpio(u.get("canal")) or "Moderno",
            "zon": limpio(u.get("zona")),
            "sup": sup,
            "mer": merc,
            "tipo": limpio(u.get("tipo_label") or u.get("tipo")),
            "skus": list(dict.fromkeys(skus)) or [limpio(u.get("skus"))],
            "neg": ordenados(n.strip().upper() for n in (u.get("negocios") or "").split("+")) or [SIN_NEGOCIO],
            "cat": ordenados(cats) or ["SIN CATEGORÍA"],
            "alta": alta.isoformat(),
            "baja": baja.isoformat() if baja else "",
            "vista": (u.get("ultima_vista") or "")[:10],
            "mot": motivos.get(u.get("motivo_baja", ""), limpio(u.get("motivo_baja"))),
            "com": limpio(u.get("comentario_baja")),
        })

    fotos = sincronizar_fotos(odk, origen_foto, clave) if con_fotos else {}
    dims = fotos.pop("dims", {})
    for f in filas:
        # Solo las publicadas, de la mas vieja a la mas nueva: f = fecha, d = [ancho, alto], s = de seguimiento.
        f["fotos"] = [{"id": pid, "f": fecha, "d": dims[pid], "s": s}
                      for fecha, pid, s in sorted(fotos_de.get(f["id"], [])) if pid in dims]

    semana_hoy = lunes(hoy)
    primera = lunes(min((date.fromisoformat(f["alta"]) for f in filas), default=hoy))
    semanas = [semana_hoy - timedelta(weeks=i) for i in range(SEMANAS_EX)]
    semanas = [s for s in semanas if s >= primera]
    return {
        "titulo": "Exhibiciones ganadas",
        "subtitulo": "Moderno, tradicional y farmacias · una exhibición sigue activa hasta que se registra su baja",
        "unidad": "exhibiciones",
        "periodos": [{"id": s.isoformat(), "label": etiqueta_semana(s)} for s in semanas],
        "periodo_tipo": "semana",
        "filas": filas,
        "fotos": fotos,
        "registros": len(registros),
    }


def etiqueta_semana(s: date) -> str:
    """Semana ISO: ``Semana 44 · 26 oct al 1 nov``."""
    fin = s + timedelta(days=6)
    return f"Semana {s.isocalendar().week} · {s.day} {MESES[s.month - 1][:3]} al {fin.day} {MESES[fin.month - 1][:3]}"


# --- salida ---------------------------------------------------------------------


def cifrar(texto: str, clave: str) -> dict:
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    sal, iv = os.urandom(16), os.urandom(12)
    iteraciones = 250_000
    k = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=sal, iterations=iteraciones).derive(clave.encode())
    datos = AESGCM(k).encrypt(iv, texto.encode("utf-8"), None)
    b64 = lambda b: base64.b64encode(b).decode()
    return {"sal": b64(sal), "iv": b64(iv), "it": iteraciones, "datos": b64(datos)}


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--env", help="archivo .env con ODK_EMAIL y ODK_PASSWORD")
    ap.add_argument("--salida", default=str(SALIDA))
    ap.add_argument("--sin-fotos", action="store_true", help="no descargar ni publicar fotos")
    args = ap.parse_args()
    if args.env:
        leer_env(args.env)

    email, password = os.environ.get("ODK_EMAIL"), os.environ.get("ODK_PASSWORD")
    if not (email and password):
        raise SystemExit("Faltan ODK_EMAIL / ODK_PASSWORD (variables de entorno o --env).")
    odk = Odk(os.environ.get("ODK_BASE", "https://odk.campiq.space"),
              int(os.environ.get("ODK_PROJECT", "1")), email, password)

    ahora = datetime.now(ECUADOR)
    hoy = ahora.date()
    personal = Personal()
    # Acuerdos y promociones primero: sus listas traen nombre y supervisor de
    # cada cedula, que visibilidad no tiene.
    ac = armar_acuerdos(odk, personal, hoy)
    pr = armar_promociones(odk, personal, hoy)
    vi = armar_visibilidad(odk, personal, hoy)
    clave = os.environ.get("SEGUIMIENTO_CLAVE")
    ex = armar_exhibiciones(odk, personal, hoy, clave, con_fotos=not args.sin_fotos)
    asig = asignacion(pr, vi, ac)
    mensuales = {}
    for k, armar in (("sos", armar_sos), ("pl", armar_planes), ("co", armar_checkouts), ("kb", armar_kbds)):
        mensuales[k] = armar(odk, personal, hoy, asig)

    datos = {
        "actualizado": ahora.isoformat(timespec="minutes"),
        "hoy": hoy.isoformat(),
        "formularios": {"vis": vi, "ac": ac, "pr": pr, "ex": ex, **mensuales},
    }
    texto = json.dumps(datos, ensure_ascii=False, separators=(",", ":"))
    carga = {"cifrado": cifrar(texto, clave)} if clave else {"plano": datos}

    html = PLANTILLA.read_text(encoding="utf-8")
    bloque = json.dumps(carga, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = html.replace("/*__CARGA__*/null", bloque)
    salida = Path(args.salida)
    salida.parent.mkdir(parents=True, exist_ok=True)
    salida.write_text(html, encoding="utf-8")
    # La pagina publicada consulta este archivo para recargarse cuando hay datos nuevos.
    (salida.parent / "version.json").write_text(json.dumps({"actualizado": datos["actualizado"]}), encoding="utf-8")
    shutil.copyfile(PLANTILLA_PPT, salida.parent / PLANTILLA_PPT.name)

    for clave_f, f in datos["formularios"].items():
        if clave_f == "ex":
            activas = sum(1 for x in f["filas"] if not x["baja"])
            print(f"{f['titulo']:<22} {f['registros']:>5} registros | {len(f['filas'])} exhibiciones, {activas} activas")
            continue
        e = Counter(x["e"] for x in f["filas"] if x["p"] == f["periodos"][0]["id"])
        print(f"{f['titulo']:<22} {f['registros']:>5} registros | {f['periodos'][0]['label']}: {dict(e)}")
    print(f"Escrito {salida} ({salida.stat().st_size / 1e6:.1f} MB){' · cifrado' if clave else ''}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
