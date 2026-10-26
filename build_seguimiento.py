"""Genera docs/index.html: seguimiento de levantamientos ODK en un solo archivo.

Lee de ODK Central tres formularios y sus listas de entidades (el universo de
lo que cada mercaderista tiene que levantar) y produce una pagina estatica con
los datos embebidos. No necesita servidor: GitHub Actions la regenera y GitHub
Pages la sirve.

    Formulario                               Universo (lista de entidades)      Periodo esperado
    seguimiento_muebles_alquilados_uiux_v4   muebles_por_local_lista            un registro por local cada mes
    acuerdos_comerciales                     acuerdos_por_local_lista           un registro por local y acuerdo en su vigencia
    promociones_mercaderista                 locales_promociones_lista          un registro por local cada semana (lun-dom)

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
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests

AQUI = Path(__file__).resolve().parent
PLANTILLA = AQUI / "plantilla.html"
SALIDA = AQUI / "docs" / "index.html"

ECUADOR = timezone(timedelta(hours=-5))  # sin horario de verano
TIMEOUT = 180

FORM_VIS = "seguimiento_muebles_alquilados_uiux_v4"
FORM_AC = "acuerdos_comerciales"
FORM_PR = "promociones_mercaderista"

MESES_VIS = 2                      # meses que muestra visibilidad (el actual y el anterior)
SEMANAS_PROMO = 4                  # semanas hacia atras que muestra promociones
SEMANAS_AC = 8                     # semanas ISO hacia atras que muestra acuerdos
MAX_DETALLE = 6                    # novedades listadas por registro

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
            r = requests.get(url, headers=self.h, timeout=TIMEOUT)
            r.raise_for_status()
            cuerpo = r.json()
            filas.extend(cuerpo.get("value", []))
            url = cuerpo.get("@odata.nextLink")
            if not url:
                break
        return filas

    def lista(self, form: str, nombre: str) -> list[dict]:
        """Adjunto CSV del formulario (las listas de entidades se publican asi)."""
        r = requests.get(f"{self._form(form)}/attachments/{nombre}", headers=self.h, timeout=TIMEOUT)
        r.raise_for_status()
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


def armar_acuerdos(odk: Odk, personal: Personal, hoy: date) -> dict:
    universo = odk.lista(FORM_AC, "acuerdos_por_local_lista.csv")
    desde = min((fecha_de(u.get("vigencia_desde")) for u in universo if fecha_de(u.get("vigencia_desde"))), default=None)
    registros = odk.registros(FORM_AC, desde=desde - timedelta(days=1) if desde else None)
    for u in universo:
        personal.agregar(u.get("acc_1", ""), u.get("mercaderista"), u.get("supervisor"),
                         local=u.get("local_id", ""), zona=u.get("acc_zona", ""))

    # Registros por acuerdo y semana ISO de la visita (lunes a domingo).
    por_llave_sem: dict[tuple[str, date], list[dict]] = defaultdict(list)
    for r in registros:
        f = fecha_de(r.get("fecha"))
        if f:
            por_llave_sem[(g(r, "grp_local/acuerdo_id"), lunes(f))].append(r)

    semana_hoy = lunes(hoy)
    semanas = [semana_hoy - timedelta(weeks=i) for i in range(SEMANAS_AC)]
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
        # Se espera un registro en cada semana que toque la vigencia del acuerdo.
        for s in semanas:
            if not (desde <= s + timedelta(days=6) and hasta >= s):
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
        "subtitulo": "Un registro por local y acuerdo cada semana mientras esté vigente",
        "unidad": "acuerdos",
        "periodos": [{"id": s.isoformat(), "label": etiqueta_semana(s)} for s in semanas if s.isoformat() in con_filas],
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

    datos = {
        "actualizado": ahora.isoformat(timespec="minutes"),
        "hoy": hoy.isoformat(),
        "formularios": {"vis": vi, "ac": ac, "pr": pr},
    }
    texto = json.dumps(datos, ensure_ascii=False, separators=(",", ":"))
    clave = os.environ.get("SEGUIMIENTO_CLAVE")
    carga = {"cifrado": cifrar(texto, clave)} if clave else {"plano": datos}

    html = PLANTILLA.read_text(encoding="utf-8")
    bloque = json.dumps(carga, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = html.replace("/*__CARGA__*/null", bloque)
    salida = Path(args.salida)
    salida.parent.mkdir(parents=True, exist_ok=True)
    salida.write_text(html, encoding="utf-8")
    # La pagina publicada consulta este archivo para recargarse cuando hay datos nuevos.
    (salida.parent / "version.json").write_text(json.dumps({"actualizado": datos["actualizado"]}), encoding="utf-8")

    for clave_f, f in datos["formularios"].items():
        e = Counter(x["e"] for x in f["filas"] if x["p"] == f["periodos"][0]["id"])
        print(f"{f['titulo']:<22} {f['registros']:>5} registros | {f['periodos'][0]['label']}: {dict(e)}")
    print(f"Escrito {salida} ({salida.stat().st_size / 1e6:.1f} MB){' · cifrado' if clave else ''}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
