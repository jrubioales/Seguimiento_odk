"""Actualizar las encuestas de ODK con un Excel subido desde la página de seguimiento.

Cubre Promociones, Acuerdos comerciales y Visibilidad. Para cada una:

* ``exportar`` arma las hojas que la página ofrece para descargar (lo que hay
  hoy en ODK, sin cédulas).
* ``planear`` compara un Excel subido con lo que hay en ODK y devuelve qué
  entidades crear o cambiar y, en Promociones, los adjuntos nuevos del
  formulario. No escribe nada.
* ``aplicar`` ejecuta ese plan en ODK Central.

Nunca borra: un local, acuerdo o mueble que se quite del Excel queda como está
en ODK. Las promociones sí se reemplazan completas (las vigentes de hoy en
adelante salen solo del Excel).

Uso:
    python cargas_odk.py probar --env ../ODK\\ MUEBLES\\ VISIBILIDAD/.env --encuesta ac --json hojas.json
    python cargas_odk.py procesar --dir cargas      # lo usa GitHub Actions
"""

from __future__ import annotations

import argparse
import base64
import csv
import io
import json
import math
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests

ECUADOR = timezone(timedelta(hours=-5))
TIMEOUT = 180
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]

FORM_PR = "promociones_mercaderista"
FORM_AC = "acuerdos_comerciales"
FORM_VIS = "seguimiento_muebles_alquilados_uiux_v4"
L_PR = "locales_promociones_lista"
L_AC = "acuerdos_por_local_lista"
L_AC_LOC = "acuerdos_locales_lista"
L_AC_IT = "acuerdos_items_lista"
L_VIS = "muebles_por_local_lista"

# Promociones: orden de los pasillos en la encuesta y su nombre en pantalla.
PASILLOS = [
    ("ACEITE", "ACEITES"), ("ACHIOTE", "ACHIOTE"), ("MANTECA", "MANTECAS"),
    ("SAZONADOR", "SAZONADORES Y CALDOS"), ("MARGARINA", "MARGARINAS"), ("DETERGENTE", "DETERGENTES"),
    ("JABON DE LAVAR", "JABÓN DE LAVAR"), ("QUITAMANCHAS", "QUITAMANCHAS"), ("LAVAVAJILLAS", "LAVAVAJILLAS"),
    ("JABON DE TOCADOR", "JABÓN DE TOCADOR"),
]
MAX_PANTALLA = 15          # un pasillo con más promociones se parte en partes iguales
MAX_MUEBLES_AC = 5         # muebles por acuerdo (columnas mueble_1..5 de la lista)
MAX_MUEBLES_VIS = 9        # muebles por local en visibilidad (item_1..9)

# Columnas de cada hoja del Excel: (columna en el Excel, obligatoria).
HOJAS = {
    "pr": {
        "Locales": ["Código", "Local", "Grupo de promociones", "Cadena", "Formato", "Ciudad", "Provincia",
                    "Zona", "Mercaderista", "Supervisor", "Cédula mercaderista nuevo"],
        "Promociones": ["Grupo de promociones", "EAN", "Producto", "Categoría", "Descuento", "Campaña",
                        "Inicio", "Fin", "Precio promoción", "Exhibición", "ID promoción"],
    },
    "ac": {
        "Locales": ["Código", "Local", "Cadena", "Ciudad", "Provincia", "Zona", "Mercaderista", "Supervisor",
                    "Cédula mercaderista nuevo"],
        "Acuerdos": ["Código", "Local", "Acuerdo", "Desde", "Hasta", "Mueble", "Producto"],
    },
    "vis": {
        "Locales": ["Código", "Local", "Cadena", "Zona", "Mercaderista", "Supervisor", "Cédula mercaderista nuevo"],
        "Muebles": ["Código", "Local", "Tipo", "Subtipo", "Proveedor", "Categorías", "Marcas", "Cantidad", "Modalidad"],
    },
}
TITULOS = {"pr": "Promociones", "ac": "Acuerdos comerciales", "vis": "Visibilidad"}


# --- utilidades -----------------------------------------------------------------


def limpio(x) -> str:
    return re.sub(r"\s+", " ", str(x if x is not None else "")).strip()


def clave(x) -> str:
    x = unicodedata.normalize("NFKD", str(x or "")).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]+", " ", x).strip()


def slug(x) -> str:
    return clave(x).lower().replace(" ", "_")


def titulo(x) -> str:
    return " ".join(p.capitalize() for p in limpio(x).split())


def cedula(x) -> str:
    t = limpio(x).split(".")[0]
    return t.zfill(10) if t.isdigit() else t


def fecha(x) -> date | None:
    """Fecha de una celda: ISO, dd/mm/aaaa o número de serie de Excel."""
    t = limpio(x)
    if not t:
        return None
    if re.fullmatch(r"\d{4,5}(\.\d+)?", t):           # serie de Excel
        return date(1899, 12, 30) + timedelta(days=int(float(t)))
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.strptime(t[:10] if formato.startswith("%Y") else t, formato).date()
        except ValueError:
            pass
    return None


def rango(a: date, b: date) -> str:
    """``del 2 al 30 nov`` · ``del 1 oct al 30 nov``."""
    if (a.year, a.month) == (b.year, b.month):
        return f"del {a.day} al {b.day} {MESES[b.month - 1]}"
    return f"del {a.day} {MESES[a.month - 1]} al {b.day} {MESES[b.month - 1]}"


def numero(x) -> str:
    """``20.9`` · ``15`` (sin ceros de más)."""
    try:
        v = float(str(x).replace(",", "."))
    except ValueError:
        return ""
    return f"{v:.4f}".rstrip("0").rstrip(".")


# --- ODK ------------------------------------------------------------------------


class Odk:
    def __init__(self, base: str, proyecto: int | str, email: str, password: str):
        self.base = base.rstrip("/")
        self.proyecto = proyecto
        r = requests.post(f"{self.base}/v1/sessions", json={"email": email, "password": password}, timeout=60)
        if r.status_code >= 400:
            raise SystemExit(f"Login ODK fallido ({r.status_code})")
        self.s = requests.Session()
        self.s.headers["Authorization"] = f"Bearer {r.json()['token']}"
        self._cache: dict[str, list[dict]] = {}

    def url(self, ruta: str) -> str:
        return f"{self.base}/v1/projects/{self.proyecto}/{ruta}"

    def _ok(self, r: requests.Response) -> requests.Response:
        if r.status_code >= 400:
            raise RuntimeError(f"ODK {r.status_code}: {r.text[:300]}")
        return r

    def entidades(self, lista: str) -> list[dict]:
        if lista not in self._cache:
            r = self._ok(self.s.get(self.url(f"datasets/{lista}/entities.csv"), timeout=TIMEOUT))
            self._cache[lista] = list(csv.DictReader(io.StringIO(r.content.decode("utf-8-sig"))))
        return self._cache[lista]

    def propiedades(self, lista: str) -> list[str]:
        r = self._ok(self.s.get(self.url(f"datasets/{lista}"), timeout=60))
        return [p["name"] for p in r.json().get("properties", [])]

    def adjunto(self, form: str, nombre: str) -> str:
        r = self._ok(self.s.get(self.url(f"forms/{form}/attachments/{nombre}"), timeout=TIMEOUT))
        return r.content.decode("utf-8-sig")

    def crear(self, lista: str, label: str, datos: dict) -> None:
        self._ok(self.s.post(self.url(f"datasets/{lista}/entities"), json={"label": label, "data": datos}, timeout=60))

    def actualizar(self, lista: str, uuid: str, version: str, datos: dict, label: str | None = None) -> None:
        cuerpo = {"data": datos}
        if label is not None:
            cuerpo["label"] = label
        self._ok(self.s.patch(self.url(f"datasets/{lista}/entities/{uuid}"), params={"baseVersion": version},
                              json=cuerpo, timeout=60))

    def hay_borrador(self, form: str) -> bool:
        return self.s.get(self.url(f"forms/{form}/draft"), timeout=60).status_code == 200

    def publicar_adjuntos(self, form: str, archivos: dict[str, str], version: str) -> None:
        """Borrador desde la versión publicada, adjuntos nuevos y publicar."""
        self._ok(self.s.post(self.url(f"forms/{form}/draft"), timeout=TIMEOUT))
        for nombre, texto in archivos.items():
            self._ok(self.s.post(self.url(f"forms/{form}/draft/attachments/{nombre}"), data=texto.encode("utf-8"),
                                 headers={"Content-Type": "text/csv"}, timeout=TIMEOUT))
        self._ok(self.s.post(self.url(f"forms/{form}/draft/publish"), params={"version": version}, timeout=TIMEOUT))


# --- personal (nombre -> cédula) ------------------------------------------------------


class Personal:
    """Cédula de cada mercaderista según las listas de ODK; la página no las tiene."""

    def __init__(self, odk: Odk):
        self.por_nombre: dict[str, Counter] = defaultdict(Counter)
        for lista, col_nombre, col_ced in ((L_PR, "mercaderistas", "acc_1"), (L_AC, "mercaderista", "acc_1"),
                                           (L_VIS, "mercaderista", "acc_1"),
                                           ("ganadas_moderno", "mercaderista", "cod_mercaderista")):
            try:
                ents = odk.entidades(lista)
            except RuntimeError:
                continue
            for e in ents:
                n, c = clave(e.get(col_nombre)), cedula(e.get(col_ced))
                if n and c:
                    self.por_nombre[n][c] += 1
        self.zonas = {limpio(e.get("acc_zona")).upper() for l in (L_PR, L_AC, L_VIS) for e in odk.entidades(l)} - {""}

    def cedula_de(self, nombre: str, nueva: str = "") -> str | None:
        if cedula(nueva):
            return cedula(nueva)
        c = self.por_nombre.get(clave(nombre))
        return c.most_common(1)[0][0] if c else None


# --- hojas del Excel ------------------------------------------------------------


def leer_hoja(hojas: dict, enc: str, nombre: str, errores: list[str]) -> list[tuple[int, dict]] | None:
    """Filas de una hoja como (número de fila en Excel, {columna: texto})."""
    filas = hojas.get(nombre)
    if filas is None:
        return None
    if filas and isinstance(filas[0], dict):
        cols = {clave(k): k for k in filas[0].keys()}
        faltan = [c for c in HOJAS[enc][nombre] if clave(c) not in cols and "nuevo" not in c.lower()]
        if faltan:
            errores.append(f"Hoja «{nombre}»: faltan las columnas {', '.join(faltan)}")
            return []
    salida = []
    for i, f in enumerate(filas):
        g = {c: limpio(next((v for k, v in f.items() if clave(k) == clave(c)), "")) for c in HOJAS[enc][nombre]}
        if any(g.values()):
            salida.append((i + 2, g))
    return salida


def asignacion_locales(filas, personal: Personal, errores: list[str], actuales: dict[str, str],
                       hoja="Locales") -> dict[str, dict]:
    """Código de local -> datos de la hoja Locales. La cédula (_ced) solo se busca
    si el mercaderista cambia; actuales = código -> mercaderista de hoy en ODK."""
    salida = {}
    for n, f in filas:
        cod = f["Código"]
        if not cod:
            errores.append(f"{hoja}, fila {n}: falta el código del local")
            continue
        if cod in salida:
            errores.append(f"{hoja}, fila {n}: el código {cod} está repetido")
            continue
        ced = ""
        if f.get("Mercaderista") and clave(f["Mercaderista"]) != clave(actuales.get(cod)):
            ced = personal.cedula_de(f["Mercaderista"], f.get("Cédula mercaderista nuevo", ""))
            if not ced:
                errores.append(f"{hoja}, fila {n}: no conozco a «{f['Mercaderista']}»; si es nuevo, escribe su cédula en "
                               "«Cédula mercaderista nuevo»")
                continue
        zona = f.get("Zona", "").upper()
        if zona and personal.zonas and zona not in personal.zonas:
            errores.append(f"{hoja}, fila {n}: zona «{f['Zona']}» desconocida (usa: {', '.join(sorted(personal.zonas))})")
            continue
        salida[cod] = {**f, "Zona": zona, "_ced": ced, "_fila": n}
    return salida


class Plan:
    def __init__(self, enc: str):
        self.enc = enc
        self.errores: list[str] = []
        self.avisos: list[str] = []
        self.ops: list[dict] = []
        self.conteo: Counter = Counter()

    def crear(self, lista: str, label: str, datos: dict, que: str):
        self.ops.append({"tipo": "crear", "lista": lista, "label": label, "datos": {k: v for k, v in datos.items() if v != ""}})
        self.conteo[f"{que} nuevos"] += 1

    def cambiar(self, lista: str, ent: dict, nuevos: dict, label: str | None, que: str):
        """Solo lo que cambia de verdad (espacios y mayúsculas no cuentan)."""
        datos = {k: v for k, v in nuevos.items() if limpio(ent.get(k)) != limpio(v)}
        lab = label if label is not None and limpio(ent.get("label")) != limpio(label) else None
        if datos or lab:
            self.ops.append({"tipo": "cambiar", "lista": lista, "uuid": ent["__id"], "version": ent["__version"],
                             "datos": datos, "label": lab})
            self.conteo[f"{que} cambiados"] += 1

    def resumen(self) -> dict:
        return {"encuesta": self.enc, "errores": self.errores, "avisos": self.avisos,
                "conteo": dict(self.conteo), "operaciones": len(self.ops)}


def acc(ced: str, zona: str, libre: bool) -> dict:
    return {"acc_1": ced, "acc_zona": zona, "acc_libre": zona if libre else "", "acc_todos": "si"}


def usa_libre(ents: list[dict]) -> bool:
    con = [e for e in ents if e.get("acc_zona")]
    return bool(con) and sum(1 for e in con if e.get("acc_libre") == e.get("acc_zona")) > len(con) / 2


# --- PROMOCIONES ----------------------------------------------------------------


def _orden_pasillo(cat: str) -> tuple[int, str]:
    for i, (c, _) in enumerate(PASILLOS):
        if clave(c) == clave(cat):
            return i, ""
    return len(PASILLOS), clave(cat)


def _nombre_pasillo(cat: str) -> str:
    for c, n in PASILLOS:
        if clave(c) == clave(cat):
            return n
    return limpio(cat).upper()


def _mem(desc: str) -> str:
    return re.sub(r"\s+", "", desc).lower()


def armar_promociones(promos: list[dict], hoy: date) -> tuple[list[dict], list[dict]]:
    """Promociones (una por fila) -> filas de promociones_items.csv y promociones_tramos.csv."""
    por_grupo = defaultdict(list)
    for p in promos:
        if p["fin"] >= hoy:
            por_grupo[p["grupo"]].append(p)
    items, tramos = [], []
    for grupo in sorted(por_grupo):
        ps = por_grupo[grupo]
        cortes = sorted({hoy} | {p["inicio"] for p in ps if p["inicio"] > hoy} | {p["fin"] + timedelta(days=1) for p in ps})
        for a, b in zip(cortes, cortes[1:]):
            fin = b - timedelta(days=1)
            activas = [p for p in ps if p["inicio"] <= a and p["fin"] >= fin]
            if not activas:
                continue
            # Mismo producto y mismo descuento: una sola promoción.
            juntas: dict[tuple, dict] = {}
            for p in sorted(activas, key=lambda p: (p["inicio"], p["ids"])):
                k = (p["ean"], _mem(p["descuento"]))
                if k in juntas:
                    j = juntas[k]
                    j["ids"] = j["ids"] + [i for i in p["ids"] if i not in j["ids"]]
                    if p["campana"] and p["campana"] not in j["campanas"]:
                        j["campanas"].append(p["campana"])
                    j["inicio"], j["fin"] = min(j["inicio"], p["inicio"]), max(j["fin"], p["fin"])
                    j["precio"] = j["precio"] or p["precio"]
                    j["exhibicion"] = j["exhibicion"] or p["exhibicion"]
                else:
                    juntas[k] = {**p, "ids": list(p["ids"]), "campanas": [p["campana"]] if p["campana"] else []}
            lista = sorted(juntas.values(), key=lambda p: (_orden_pasillo(p["categoria"]), clave(p["producto"]), p["ean"], p["descuento"]))
            tramo = f"{grupo}-{a:%Y%m%d}"
            # Pantallas: una por pasillo; si pasa de MAX_PANTALLA se parte en partes iguales.
            pasillos = defaultdict(list)
            for p in lista:
                pasillos[_orden_pasillo(p["categoria"])].append(p)
            seq, pantalla = 0, 0
            filas_tramo = []
            for orden in sorted(pasillos):
                ps_p = pasillos[orden]
                nombre = _nombre_pasillo(ps_p[0]["categoria"])
                partes = math.ceil(len(ps_p) / MAX_PANTALLA)
                base, extra = divmod(len(ps_p), partes)
                tamanos = [base + (1 if i < extra else 0) for i in range(partes)]
                k = 0
                for parte, tam in enumerate(tamanos, 1):
                    pantalla += 1
                    cuantas = f"{tam} promociones" if tam > 1 else "1 promocion"
                    etiqueta = f"🧭 {nombre} · parte {parte} de {partes} · {cuantas}" if partes > 1 else f"🧭 {nombre} · {cuantas}"
                    for pos in range(1, tam + 1):
                        p = ps_p[k]; k += 1; seq += 1
                        otras = [o for o in lista if o["ean"] == p["ean"] and o is not p]
                        desc = p["descuento"]
                        campana = " + ".join(p["campanas"])
                        detalle = f"EAN {p['ean']} · {desc}" + (f" · Promo ${p['precio']}" if p["precio"] else "") + \
                                  f" · {rango(p['inicio'], p['fin'])} · {campana}"
                        filas_tramo.append({
                            "clave": f"{tramo}#{seq}", "tramo": tramo, "grupo": grupo, "seq": str(seq),
                            "promo_id": p["ids"][0] if p["ids"] else "", "promo_ids": " ".join(p["ids"]),
                            "mem_clave": f"{p['ean']}|{_mem(desc)}", "ean": p["ean"], "producto": p["producto"],
                            "descuento": desc, "campana": campana, "categoria": p["categoria"],
                            "pasillo": f"🧭 {nombre} · {k} de {len(ps_p)} en este pasillo",
                            "inicio": p["inicio"].isoformat(), "fin": p["fin"].isoformat(),
                            "ficha_titulo": f"🏷️ {p['producto']}", "ficha_detalle": detalle,
                            "otra_promo": ("⚠️ Hay otra promoción de este producto: " + "; ".join(
                                f"{o['descuento']} ({' + '.join(o['campanas'])})" for o in otras)
                                + ". Fíjate cuál está en la cenefa.") if otras else "",
                            "exhibicion": p["exhibicion"], "clave_p": f"{tramo}#{pantalla}#{pos}", "pantalla": str(pantalla),
                            "pos": str(pos), "pantalla_n": str(tam), "pantalla_label": etiqueta,
                            "desc_tipo": "pct" if desc.endswith("%") else "",
                            "desc_num": numero(desc[:-1]) if desc.endswith("%") else "",
                            "precio_original": "", "precio_promocion": p["precio"], "precio_meta": "", "meta_txt": "",
                        })
            items.extend(filas_tramo)
            tramos.append({"tramo": tramo, "grupo": grupo, "desde": a.isoformat(), "hasta": fin.isoformat(),
                           "desde_num": f"{a:%Y%m%d}", "hasta_num": f"{fin:%Y%m%d}", "n": str(len(filas_tramo)),
                           "label": rango(a, fin), "n_pantallas": str(pantalla)})
    return items, tramos


def promos_de_items(items: list[dict], hoy: date) -> list[dict]:
    """Promociones vigentes (una por producto, descuento y periodo) a partir del adjunto."""
    vistas = {}
    for i in items:
        fin = fecha(i["fin"])
        if not fin or fin < hoy:
            continue
        k = (i["grupo"], i["ean"], _mem(i["descuento"]), i["inicio"], i["fin"])
        if k not in vistas:
            vistas[k] = i
    return list(vistas.values())


def a_csv(filas: list[dict], columnas: list[str]) -> str:
    out = io.StringIO()
    w = csv.DictWriter(out, fieldnames=columnas, extrasaction="ignore", lineterminator="\n")
    w.writeheader()
    for f in filas:
        w.writerow({c: f.get(c, "") for c in columnas})
    return out.getvalue()


def vista_promos(items: list[dict], tramos: list[dict], desde: date) -> dict:
    """Qué ve la encuesta cada día (para saber si algo cambió de verdad)."""
    por_tramo = defaultdict(list)
    for i in items:
        por_tramo[i["tramo"]].append(tuple(i.get(c, "") for c in (
            "seq", "ean", "producto", "descuento", "campana", "categoria", "pasillo", "inicio", "fin", "ficha_detalle",
            "otra_promo", "pantalla", "pos", "pantalla_n", "pantalla_label", "desc_tipo", "desc_num", "promo_id",
            "promo_ids", "mem_clave")))
    salida = {}
    for t in tramos:
        a, b = fecha(t["desde"]), fecha(t["hasta"])
        d = max(a, desde)
        while d <= b:
            salida[(t["grupo"], d)] = sorted(por_tramo[t["tramo"]], key=lambda x: int(x[0]))
            d += timedelta(days=1)
    return salida


def exportar_pr(odk: Odk, hoy: date) -> dict:
    locales = odk.entidades(L_PR)
    grupos = {e["grupo"]: e["grupo_label"] for e in locales}
    items = list(csv.DictReader(io.StringIO(odk.adjunto(FORM_PR, "promociones_items.csv"))))
    filas_loc = [{
        "Código": e["llave"], "Local": e["local_nombre"], "Grupo de promociones": e["grupo_label"],
        "Cadena": e["cadena_label"], "Formato": e["formato"], "Ciudad": e["ciudad"], "Provincia": e["provincia"],
        "Zona": e["zona"] or e["acc_zona"], "Mercaderista": e["mercaderistas"], "Supervisor": e["supervisor"],
    } for e in sorted(locales, key=lambda e: (e["grupo_label"], e["local_nombre"]))]
    filas_pr = [{
        "Grupo de promociones": grupos.get(i["grupo"], i["grupo"]), "EAN": i["ean"], "Producto": i["producto"],
        "Categoría": i["categoria"], "Descuento": i["descuento"], "Campaña": i["campana"], "Inicio": i["inicio"],
        "Fin": i["fin"], "Precio promoción": i["precio_promocion"], "Exhibición": i["exhibicion"], "ID promoción": i["promo_ids"],
    } for i in sorted(promos_de_items(items, hoy), key=lambda i: (i["grupo"], _orden_pasillo(i["categoria"]), i["producto"].lower()))]
    return {"Locales": filas_loc, "Promociones": filas_pr}


def planear_pr(odk: Odk, hojas: dict, hoy: date, personal: Personal) -> Plan:
    plan = Plan("pr")
    ents = odk.entidades(L_PR)
    por_llave = {e["llave"]: e for e in ents}
    grupos = {clave(e["grupo_label"]): (e["grupo"], e["grupo_label"]) for e in ents}
    grupos.update({clave(e["grupo"]): (e["grupo"], e["grupo_label"]) for e in ents})
    cadenas = {clave(e["cadena_label"]): (e["cadena"], e["cadena_label"]) for e in ents}
    libre = usa_libre(ents)

    filas = leer_hoja(hojas, "pr", "Locales", plan.errores)
    if filas is not None:
        locales = asignacion_locales(filas, personal, plan.errores, {e["llave"]: e["mercaderistas"] for e in ents})
        for cod, f in locales.items():
            g = grupos.get(clave(f["Grupo de promociones"]))
            if not g:
                plan.errores.append(f"Locales, fila {f['_fila']}: grupo «{f['Grupo de promociones']}» no existe en la encuesta "
                                    f"(usa: {', '.join(sorted({v[1] for v in grupos.values()}))})")
                continue
            cad = cadenas.get(clave(f["Cadena"]), (slug(f["Cadena"]), f["Cadena"]))
            datos = {"llave": cod, "grupo": g[0], "grupo_label": g[1], "cadena": cad[0], "cadena_label": cad[1],
                     "local_nombre": f["Local"], "formato": f["Formato"], "ciudad": f["Ciudad"], "provincia": f["Provincia"],
                     "zona": f["Zona"], "supervisor": f["Supervisor"], "mercaderistas": f["Mercaderista"]}
            label = f"{f['Local']} · {f['Ciudad']} ({cod})" if f["Ciudad"] else f"{f['Local']} ({cod})"
            e = por_llave.get(cod)
            if e:
                if f["Mercaderista"] and clave(f["Mercaderista"]) != clave(e["mercaderistas"]):
                    datos["acc_1"] = f["_ced"]
                if f["Zona"] and f["Zona"] != e["acc_zona"]:
                    datos["acc_zona"] = f["Zona"]
                    if libre:
                        datos["acc_libre"] = f["Zona"]
                plan.cambiar(L_PR, e, datos, label, "locales")
            else:
                if not (f["Local"] and f["Zona"] and f["Mercaderista"]):
                    plan.errores.append(f"Locales, fila {f['_fila']}: un local nuevo necesita Local, Zona y Mercaderista")
                    continue
                plan.crear(L_PR, label, {**datos, **acc(f["_ced"], f["Zona"], libre)}, "locales")

    filas = leer_hoja(hojas, "pr", "Promociones", plan.errores)
    if filas is not None:
        promos = []
        for n, f in filas:
            g = grupos.get(clave(f["Grupo de promociones"]))
            ini, fin = fecha(f["Inicio"]), fecha(f["Fin"])
            ean = re.sub(r"\D", "", f["EAN"].split(".")[0])
            problemas = [m for m, mal in (
                (f"grupo «{f['Grupo de promociones']}» no existe", not g), ("falta el EAN", not ean),
                ("falta el producto", not f["Producto"]), ("falta la categoría", not f["Categoría"]),
                ("falta el descuento", not f["Descuento"]), ("fecha de inicio no válida", not ini),
                ("fecha de fin no válida", not fin), ("el fin es antes del inicio", ini and fin and fin < ini)) if mal]
            if problemas:
                plan.errores.append(f"Promociones, fila {n}: {'; '.join(problemas)}")
                continue
            desc = f["Descuento"]
            if re.fullmatch(r"0?[.,]\d+", desc):            # 0.15 de Excel -> 15%
                desc = numero(float(desc.replace(",", ".")) * 100) + "%"
            precio = numero(f["Precio promoción"].replace("$", "")) if f["Precio promoción"] else ""
            if precio:
                precio = f"{float(precio):.2f}"
            promos.append({"grupo": g[0], "ean": ean, "producto": f["Producto"], "categoria": f["Categoría"].upper(),
                           "descuento": desc, "campana": f["Campaña"].upper(), "inicio": ini, "fin": fin,
                           "precio": precio, "exhibicion": f["Exhibición"],
                           "ids": f["ID promoción"].split() or [f"X{n:05d}"]})
        if not promos and not plan.errores:
            plan.errores.append("La hoja Promociones está vacía: así la encuesta quedaría sin promociones. "
                                "Si solo cambias locales, borra la hoja Promociones del Excel antes de subirlo.")
        sin_cat = sorted({p["categoria"] for p in promos if _orden_pasillo(p["categoria"])[0] == len(PASILLOS)})
        if sin_cat:
            plan.avisos.append(f"Categorías sin pasillo conocido (van al final de la encuesta): {', '.join(sin_cat)}")
        if not plan.errores:
            items, tramos = armar_promociones(promos, hoy)
            viejos_i = list(csv.DictReader(io.StringIO(odk.adjunto(FORM_PR, "promociones_items.csv"))))
            viejos_t = list(csv.DictReader(io.StringIO(odk.adjunto(FORM_PR, "promociones_tramos.csv"))))
            if vista_promos(items, tramos, hoy) != vista_promos(viejos_i, viejos_t, hoy):
                cols_i = list(viejos_i[0].keys()) if viejos_i else list(items[0].keys())
                cols_t = list(viejos_t[0].keys()) if viejos_t else list(tramos[0].keys())
                plan.ops.append({"tipo": "adjuntos", "form": FORM_PR, "archivos": {
                    "promociones_items.csv": a_csv(items, cols_i), "promociones_tramos.csv": a_csv(tramos, cols_t)}})
                plan.conteo["promociones vigentes"] = len({(i["grupo"], i["ean"], i["descuento"]) for i in items})
                plan.conteo["tramos"] = len(tramos)
    return plan


# --- ACUERDOS -------------------------------------------------------------------


def _vig_label(a: date, b: date) -> str:
    return f"Del {a:%d/%m/%Y} al {b:%d/%m/%Y}"


def exportar_ac(odk: Odk, hoy: date) -> dict:
    pl = odk.entidades(L_AC)
    locs = {e["llave"]: e for e in odk.entidades(L_AC_LOC)}
    asig = {}
    for e in pl:
        asig.setdefault(e["local_id"], e)
    filas_loc = []
    for cod in sorted(set(locs) | set(asig), key=lambda c: (limpio((locs.get(c) or asig[c])["local_nombre"]))):
        l, a = locs.get(cod) or asig[cod], asig.get(cod) or {}
        filas_loc.append({"Código": cod, "Local": l["local_nombre"], "Cadena": l["cadena_label"], "Ciudad": l["ciudad"],
                          "Provincia": l["provincia"], "Zona": a.get("acc_zona") or l.get("acc_zona", ""),
                          "Mercaderista": a.get("mercaderista", ""), "Supervisor": a.get("supervisor", "")})
    filas_ac = []
    for e in sorted(pl, key=lambda e: (e["acuerdo"], e["local_nombre"])):
        for k in range(1, MAX_MUEBLES_AC + 1):
            if not e.get(f"mueble_{k}_nombre"):
                continue
            prods = [e[f"mueble_{k}_item_{j}_label"] for j in range(1, 20) if e.get(f"mueble_{k}_item_{j}_label")] or [""]
            for p in prods:
                filas_ac.append({"Código": e["local_id"], "Local": e["local_nombre"], "Acuerdo": e["acuerdo"],
                                 "Desde": e["vigencia_desde"], "Hasta": e["vigencia_hasta"],
                                 "Mueble": e[f"mueble_{k}_nombre"], "Producto": p})
    return {"Locales": filas_loc, "Acuerdos": filas_ac}


def planear_ac(odk: Odk, hojas: dict, hoy: date, personal: Personal) -> Plan:
    plan = Plan("ac")
    pl, locs_e, items_e = odk.entidades(L_AC), odk.entidades(L_AC_LOC), odk.entidades(L_AC_IT)
    max_items = max([int(m.group(1)) for p in odk.propiedades(L_AC) for m in [re.match(r"mueble_1_item_(\d+)_id$", p)] if m] or [3])
    pl_llave = {e["llave"]: e for e in pl}
    loc_llave = {e["llave"]: e for e in locs_e}
    it_llave = {e["llave"]: e for e in items_e}
    base_loc = {}
    for e in pl:
        base_loc.setdefault(e["local_id"], e)
    for e in locs_e:
        base_loc.setdefault(e["llave"], e)

    # Locales (asignación y datos del local).
    filas = leer_hoja(hojas, "ac", "Locales", plan.errores)
    locales = asignacion_locales(filas or [], personal, plan.errores,
                                 {c: e.get("mercaderista", "") for c, e in base_loc.items()})
    # Cómo queda cada local: lo de ODK con lo del Excel encima.
    def datos_local(cod: str) -> dict | None:
        b, f = base_loc.get(cod), locales.get(cod)
        if not b and not f:
            return None
        d = {"local_nombre": limpio((b or {}).get("local_nombre")), "cadena_label": limpio((b or {}).get("cadena_label")),
             "cadena": limpio((b or {}).get("cadena")), "ciudad": limpio((b or {}).get("ciudad")),
             "provincia": limpio((b or {}).get("provincia")), "acc_zona": limpio((b or {}).get("acc_zona")),
             "mercaderista": limpio((b or {}).get("mercaderista")), "supervisor": limpio((b or {}).get("supervisor")),
             "acc_1": limpio((b or {}).get("acc_1"))}
        if f:
            for col, k in (("Local", "local_nombre"), ("Cadena", "cadena_label"), ("Ciudad", "ciudad"),
                           ("Provincia", "provincia"), ("Zona", "acc_zona"), ("Mercaderista", "mercaderista"),
                           ("Supervisor", "supervisor")):
                if f[col]:
                    d[k] = f[col]
            if f["Cadena"] and clave(f["Cadena"]) != clave((b or {}).get("cadena_label")):
                d["cadena"] = slug(f["Cadena"])
            if f["_ced"]:
                d["acc_1"] = f["_ced"]
        return d

    # Acuerdos: una fila por local, acuerdo, mueble y producto.
    filas = leer_hoja(hojas, "ac", "Acuerdos", plan.errores)
    acuerdos: dict[str, dict] = {}
    for n, f in filas or []:
        cod, nom = f["Código"], f["Acuerdo"]
        ini, fin = fecha(f["Desde"]), fecha(f["Hasta"])
        problemas = [m for m, mal in (("falta el código", not cod), ("falta el acuerdo", not nom),
                                      ("fecha Desde no válida", not ini), ("fecha Hasta no válida", not fin),
                                      ("Hasta es antes de Desde", ini and fin and fin < ini), ("falta el mueble", not f["Mueble"]))
                     if mal]
        if not problemas and not datos_local(cod):
            problemas.append(f"el local {cod} no existe; agrégalo en la hoja Locales")
        if problemas:
            plan.errores.append(f"Acuerdos, fila {n}: {'; '.join(problemas)}")
            continue
        llave = f"{cod}-{slug(nom)}"
        a = acuerdos.setdefault(llave, {"local": cod, "acuerdo": nom, "ini": ini, "fin": fin, "muebles": {}, "fila": n})
        if (a["ini"], a["fin"]) != (ini, fin):
            plan.errores.append(f"Acuerdos, fila {n}: el acuerdo «{nom}» del local {cod} tiene otras fechas en la fila {a['fila']}")
            continue
        prods = a["muebles"].setdefault(f["Mueble"], [])
        if f["Producto"] and f["Producto"] not in prods:
            prods.append(f["Producto"])
    for llave, a in acuerdos.items():
        if len(a["muebles"]) > MAX_MUEBLES_AC:
            plan.errores.append(f"Acuerdos, fila {a['fila']}: «{a['acuerdo']}» en {a['local']} tiene {len(a['muebles'])} muebles (máximo {MAX_MUEBLES_AC})")
        for m, ps in a["muebles"].items():
            if len(ps) > max_items:
                plan.errores.append(f"Acuerdos, fila {a['fila']}: el mueble «{m}» de «{a['acuerdo']}» en {a['local']} tiene {len(ps)} productos (máximo {max_items})")
    if plan.errores:
        return plan

    libre = usa_libre(pl)
    tocados = set(locales) | {a["local"] for a in acuerdos.values()}
    items_hechos: set[str] = set()
    # Locales cuyo mercaderista o zona cambia en este Excel: sus productos se reasignan.
    cambio_asig = {c for c, f in locales.items()
                   if f["_ced"] or (f["Zona"] and f["Zona"] != limpio((base_loc.get(c) or {}).get("acc_zona")))}
    # acuerdos_por_local_lista
    for llave, a in acuerdos.items():
        d = datos_local(a["local"])
        vl = _vig_label(a["ini"], a["fin"])
        datos = {"llave": llave, "local_id": a["local"], "acuerdo": a["acuerdo"], "cadena": d["cadena"] or slug(d["cadena_label"]),
                 "cadena_label": d["cadena_label"], "local_nombre": d["local_nombre"], "ciudad": d["ciudad"],
                 "provincia": d["provincia"], "supervisor": d["supervisor"], "mercaderista": d["mercaderista"],
                 "vigencia_desde": a["ini"].isoformat(), "vigencia_hasta": a["fin"].isoformat(),
                 "vigencia_desde_num": f"{a['ini']:%Y%m%d}", "vigencia_hasta_num": f"{a['fin']:%Y%m%d}",
                 "vigencia_label": vl, "mueble_count": str(len(a["muebles"])),
                 "item_count": str(sum(len(p) for p in a["muebles"].values())),
                 "acc_1": d["acc_1"], "acc_zona": d["acc_zona"]}
        for k in range(1, MAX_MUEBLES_AC + 1):
            datos.update({f"mueble_{k}_{c}": "" for c in ("id", "nombre", "items_n", "items_label", "items_todos")})
            datos.update({f"mueble_{k}_item_{j}_{c}": "" for j in range(1, max_items + 1) for c in ("id", "label")})
        for k, (mueble, prods) in enumerate(a["muebles"].items(), 1):
            mid = f"{llave}-{slug(mueble)}"
            ids = [f"{mid}-{slug(p)}" for p in prods]
            datos.update({f"mueble_{k}_id": mid, f"mueble_{k}_nombre": mueble, f"mueble_{k}_items_n": str(len(prods)),
                          f"mueble_{k}_items_label": " · ".join(prods), f"mueble_{k}_items_todos": " ".join(ids)})
            for j, (iid, p) in enumerate(zip(ids, prods), 1):
                items_hechos.add(iid)
                datos.update({f"mueble_{k}_item_{j}_id": iid, f"mueble_{k}_item_{j}_label": p})
                it = {"llave": iid, "mueble_id": mid, "local_id": a["local"], "acuerdo": a["acuerdo"],
                      "mueble_nombre": mueble, "orden": str(j)}
                if a["local"] in cambio_asig:
                    it.update({"acc_1": d["acc_1"], "acc_zona": d["acc_zona"]})
                if iid in it_llave:
                    plan.cambiar(L_AC_IT, it_llave[iid], it, p, "productos de acuerdo")
                else:
                    plan.crear(L_AC_IT, p, {**it, **acc(d["acc_1"], d["acc_zona"], libre)}, "productos de acuerdo")
            # (los productos nuevos siempre llevan la asignación del local)
        label = f"{a['acuerdo']} · {vl}"
        if llave in pl_llave:
            e = pl_llave[llave]
            # Columnas vacías que ODK ya tiene vacías no se mandan.
            plan.cambiar(L_AC, e, {k: v for k, v in datos.items() if v or e.get(k)}, label, "acuerdos")
        else:
            plan.crear(L_AC, label, {**datos, **acc(d["acc_1"], d["acc_zona"], libre)}, "acuerdos")

    # Locales cambiados sin tocar sus acuerdos: asignación y datos en los acuerdos que ya tienen.
    for e in pl:
        cod = e["local_id"]
        if cod in locales and e["llave"] not in acuerdos:
            d = datos_local(cod)
            plan.cambiar(L_AC, e, {k: d[k] for k in ("local_nombre", "cadena_label", "ciudad", "provincia", "supervisor",
                                                       "mercaderista", "acc_1", "acc_zona")}, None, "acuerdos")
    for e in items_e:
        if e["local_id"] in cambio_asig and e["llave"] not in items_hechos:
            d = datos_local(e["local_id"])
            if limpio(e.get("acc_1")) != d["acc_1"] or limpio(e.get("acc_zona")) != d["acc_zona"]:
                plan.cambiar(L_AC_IT, e, {"acc_1": d["acc_1"], "acc_zona": d["acc_zona"]}, None, "productos de acuerdo")

    # acuerdos_locales_lista: cuántos acuerdos tiene cada local y hasta cuándo.
    todos = defaultdict(dict)
    for e in pl:
        todos[e["local_id"]][e["llave"]] = fecha(e["vigencia_hasta"]) or date.min
    for llave, a in acuerdos.items():
        todos[a["local"]][llave] = a["fin"]
    for cod in sorted(tocados):
        d = datos_local(cod)
        if not d or not todos.get(cod):
            continue
        acs = todos[cod]
        datos = {"llave": cod, "cadena": d["cadena"] or slug(d["cadena_label"]), "cadena_label": d["cadena_label"],
                 "local_nombre": d["local_nombre"], "ciudad": d["ciudad"], "provincia": d["provincia"],
                 "acuerdos_n": str(len(acs)), "acuerdo_unico": next(iter(acs)) if len(acs) == 1 else "",
                 "vigencia_hasta_num": f"{max(acs.values()):%Y%m%d}", "acc_1": d["acc_1"], "acc_zona": d["acc_zona"]}
        label = f"{d['local_nombre']} · {d['ciudad']} ({cod})" if d["ciudad"] else f"{d['local_nombre']} ({cod})"
        if cod in loc_llave:
            e = loc_llave[cod]
            plan.cambiar(L_AC_LOC, e, {k: v for k, v in datos.items() if v or e.get(k)}, label, "locales")
        else:
            plan.crear(L_AC_LOC, label, {**datos, **acc(d["acc_1"], d["acc_zona"], libre)}, "locales")
    return plan


# --- VISIBILIDAD ----------------------------------------------------------------


def exportar_vis(odk: Odk, hoy: date) -> dict:
    ents = sorted(odk.entidades(L_VIS), key=lambda e: (e["cadena_label"], e["local_label"]))
    filas_loc = [{"Código": e["llave"], "Local": e["local_label"], "Cadena": e["cadena_label"], "Zona": e["acc_zona"],
                  "Mercaderista": e.get("mercaderista", ""), "Supervisor": e.get("supervisor", "")} for e in ents]
    filas_m = []
    for e in ents:
        for k in range(1, MAX_MUEBLES_VIS + 1):
            lab = e.get(f"item_{k}_label")
            if not lab:
                continue
            partes = [p.strip() for p in lab.split(" | ")]
            partes = (partes[:5] + [""] * 5)[:5]
            filas_m.append({"Código": e["llave"], "Local": e["local_label"], "Tipo": partes[0], "Subtipo": partes[1],
                            "Proveedor": partes[2], "Categorías": partes[3], "Marcas": partes[4],
                            "Cantidad": e.get(f"item_{k}_cant", ""), "Modalidad": e.get(f"item_{k}_modalidad", "")})
    return {"Locales": filas_loc, "Muebles": filas_m}


def planear_vis(odk: Odk, hojas: dict, hoy: date, personal: Personal) -> Plan:
    plan = Plan("vis")
    ents = odk.entidades(L_VIS)
    por_llave = {e["llave"]: e for e in ents}
    libre = usa_libre(ents)
    locales = asignacion_locales(leer_hoja(hojas, "vis", "Locales", plan.errores) or [], personal, plan.errores,
                                 {e["llave"]: e.get("mercaderista", "") for e in ents})
    muebles: dict[str, list[dict]] = defaultdict(list)
    for n, f in leer_hoja(hojas, "vis", "Muebles", plan.errores) or []:
        cod = f["Código"]
        cant = re.sub(r"\.0+$", "", f["Cantidad"])
        problemas = [m for m, mal in (("falta el código", not cod), ("falta el tipo", not f["Tipo"]),
                                      ("cantidad no válida (usa un número entero)", cant and not cant.isdigit())) if mal]
        if not problemas and cod not in por_llave and cod not in locales:
            problemas.append(f"el local {cod} no existe; agrégalo en la hoja Locales")
        if problemas:
            plan.errores.append(f"Muebles, fila {n}: {'; '.join(problemas)}")
            continue
        muebles[cod].append({**f, "Cantidad": cant, "_fila": n})
    for cod, ms in muebles.items():
        if len(ms) > MAX_MUEBLES_VIS:
            plan.errores.append(f"Muebles, fila {ms[0]['_fila']}: el local {cod} tiene {len(ms)} muebles (máximo {MAX_MUEBLES_VIS})")
    if plan.errores:
        return plan
    for cod in sorted(set(locales) | set(muebles)):
        e, f = por_llave.get(cod), locales.get(cod)
        datos = {}
        if f:
            datos.update({"llave": cod, "local_label": f["Local"] or (e or {}).get("local_label", ""),
                          "cadena_label": f["Cadena"] or (e or {}).get("cadena_label", "")})
            if f["Cadena"] and (not e or clave(f["Cadena"]) != clave(e.get("cadena_label"))):
                datos["cadena_code"] = slug(f["Cadena"])
            for col, k in (("Mercaderista", "mercaderista"), ("Supervisor", "supervisor")):
                if f[col]:
                    datos[k] = f[col]
            if f["_ced"] and (not e or clave(f["Mercaderista"]) != clave(e.get("mercaderista"))):
                datos["acc_1"] = f["_ced"]
            if f["Zona"]:
                datos["acc_zona"] = f["Zona"]
                if libre:
                    datos["acc_libre"] = f["Zona"]
        if cod in muebles:
            ms = muebles[cod]
            datos["mueble_count"] = str(len(ms))
            for k in range(1, MAX_MUEBLES_VIS + 1):
                m = ms[k - 1] if k <= len(ms) else None
                datos[f"item_{k}_label"] = " | ".join([m["Tipo"], m["Subtipo"], m["Proveedor"], m["Categorías"], m["Marcas"],
                                                       f"Cantidad: {m['Cantidad']}"]) if m else ""
                datos[f"item_{k}_cant"] = m["Cantidad"] if m else ""
                datos[f"item_{k}_modalidad"] = m["Modalidad"].lower() if m else ""
        nombre = datos.get("local_label") or (e or {}).get("local_label", "")
        label = f"{nombre} · Código {cod}"
        if e:
            plan.cambiar(L_VIS, e, {k: v for k, v in datos.items() if v or e.get(k)}, label, "locales")
        else:
            if not (f and f["Local"] and f["Zona"] and f["Mercaderista"] and f["Cadena"]):
                plan.errores.append(f"Locales: el local nuevo {cod} necesita Local, Cadena, Zona y Mercaderista")
                continue
            plan.crear(L_VIS, label, {**datos, "cadena_code": slug(f["Cadena"]), **acc(f["_ced"], f["Zona"], libre)}, "locales")
    return plan


# --- entrada única --------------------------------------------------------------


EXPORTAR = {"pr": exportar_pr, "ac": exportar_ac, "vis": exportar_vis}
PLANEAR = {"pr": planear_pr, "ac": planear_ac, "vis": planear_vis}


def exportar(odk: Odk, hoy: date) -> dict:
    """Lo que la página ofrece para descargar: {encuesta: {titulo, columnas, hojas}}."""
    salida = {}
    for enc, fn in EXPORTAR.items():
        try:
            salida[enc] = {"titulo": TITULOS[enc], "columnas": HOJAS[enc], "hojas": fn(odk, hoy)}
        except Exception as e:  # una lista que falla no tumba la página
            print(f"Aviso: no se pudo exportar {enc}: {str(e)[:200]}")
    return salida


def planear(odk: Odk, enc: str, hojas: dict, hoy: date) -> Plan:
    if enc not in PLANEAR:
        p = Plan(enc); p.errores.append(f"Encuesta desconocida: {enc}"); return p
    return PLANEAR[enc](odk, hojas, hoy, Personal(odk))


def aplicar(odk: Odk, plan: Plan, ahora: datetime) -> list[str]:
    """Ejecuta el plan; devuelve los errores de ODK (los demás cambios quedan hechos)."""
    fallas = []
    for op in plan.ops:
        try:
            if op["tipo"] == "crear":
                odk.crear(op["lista"], op["label"], op["datos"])
            elif op["tipo"] == "cambiar":
                odk.actualizar(op["lista"], op["uuid"], op["version"], op["datos"], op.get("label"))
            elif op["tipo"] == "adjuntos":
                if odk.hay_borrador(op["form"]):
                    fallas.append("El formulario de Promociones tiene un borrador abierto en ODK: publícalo o descártalo y vuelve a subir.")
                    continue
                odk.publicar_adjuntos(op["form"], op["archivos"], ahora.strftime("%Y%m%d%H%M"))
        except Exception as e:
            fallas.append(f"{op['tipo']} {op.get('lista') or op.get('form')} {op.get('label') or ''}: {str(e)[:200]}")
    return fallas


# --- cifrado (mismo formato que la página) -------------------------------------------


def _aes(clave_txt: str, sal: bytes, it: int):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    return AESGCM(PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=sal, iterations=it).derive(clave_txt.encode()))


def descifrar(c: dict, clave_txt: str) -> dict:
    b = base64.b64decode
    return json.loads(_aes(clave_txt, b(c["sal"]), c["it"]).decrypt(b(c["iv"]), b(c["datos"]), None))


def cifrar(obj, clave_txt: str) -> dict:
    sal, iv, it = os.urandom(16), os.urandom(12), 250_000
    datos = _aes(clave_txt, sal, it).encrypt(iv, json.dumps(obj, ensure_ascii=False).encode(), None)
    b = lambda x: base64.b64encode(x).decode()
    return {"sal": b(sal), "iv": b(iv), "it": it, "datos": b(datos)}


# --- CLI ------------------------------------------------------------------------


def _odk_env() -> Odk:
    return Odk(os.environ.get("ODK_BASE", "https://odk.campiq.space"), os.environ.get("ODK_PROJECT", "1"),
               os.environ["ODK_EMAIL"], os.environ["ODK_PASSWORD"])


def procesar(carpeta: Path) -> None:
    """Aplica cada carga pendiente (cargas/<id>.json) y escribe <id>.resultado.json."""
    clave_txt = os.environ.get("SEGUIMIENTO_CLAVE", "")
    pendientes = sorted(p for p in carpeta.glob("*.json") if not p.name.endswith(".resultado.json")
                        and not (carpeta / p.name.replace(".json", ".resultado.json")).exists())
    if not pendientes:
        print("Sin cargas pendientes")
        return
    odk = _odk_env()
    for p in pendientes:
        ahora = datetime.now(ECUADOR)
        carga: dict = {}
        try:
            sobre = json.loads(p.read_text(encoding="utf-8"))
            carga = descifrar(sobre["cifrado"], clave_txt) if "cifrado" in sobre else sobre["plano"]
            plan = planear(odk, carga["encuesta"], carga["hojas"], ahora.date())
            res = plan.resumen()
            if not plan.errores:
                fallas = aplicar(odk, plan, ahora)
                res["fallas"] = fallas
                res["aplicado"] = not fallas
            else:
                res["aplicado"] = False
        except Exception as e:
            res = {"aplicado": False, "errores": [f"No se pudo procesar la carga: {str(e)[:300]}"], "conteo": {}}
        res["procesado"] = ahora.isoformat(timespec="seconds")
        res["archivo"] = carga.get("archivo", "")
        salida = {"cifrado": cifrar(res, clave_txt)} if clave_txt else {"plano": res}
        p.with_name(p.name.replace(".json", ".resultado.json")).write_text(json.dumps(salida), encoding="utf-8")
        print(f"{p.name}: {'aplicado' if res.get('aplicado') else 'NO aplicado'} · {res.get('conteo')} · "
              f"{len(res.get('errores', []))} errores · {len(res.get('fallas', []))} fallas")


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    pr = sub.add_parser("probar", help="planear una carga sin escribir en ODK")
    pr.add_argument("--env"); pr.add_argument("--encuesta", required=True); pr.add_argument("--json", required=True)
    pr.add_argument("--hoy")
    ex = sub.add_parser("exportar", help="escribir en JSON lo que la página ofrece descargar")
    ex.add_argument("--env"); ex.add_argument("--salida", required=True); ex.add_argument("--hoy")
    pc = sub.add_parser("procesar", help="aplicar las cargas pendientes de una carpeta")
    pc.add_argument("--dir", default="cargas"); pc.add_argument("--env")
    args = ap.parse_args()
    if getattr(args, "env", None):
        for linea in Path(args.env).read_text(encoding="utf-8-sig").splitlines():
            if "=" in linea and not linea.lstrip().startswith("#"):
                k, v = linea.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
    hoy = date.fromisoformat(args.hoy) if getattr(args, "hoy", None) else datetime.now(ECUADOR).date()
    if args.cmd == "procesar":
        procesar(Path(args.dir))
    elif args.cmd == "exportar":
        Path(args.salida).write_text(json.dumps(exportar(_odk_env(), hoy), ensure_ascii=False), encoding="utf-8")
    else:
        hojas = json.loads(Path(args.json).read_text(encoding="utf-8"))
        hojas = hojas.get(args.encuesta, {}).get("hojas", hojas)
        plan = planear(_odk_env(), args.encuesta, hojas, hoy)
        print(json.dumps(plan.resumen(), ensure_ascii=False, indent=1))
        # Solo los campos que cambian (sin valores: acc_1 es la cédula).
        for op in plan.ops[:25]:
            datos = op.get("datos") or {}
            print(op["tipo"], op.get("lista") or op.get("form"), "·", op.get("label") or "",
                  "·", sorted(datos) if op["tipo"] == "cambiar" else (len(datos) if datos else list(op.get("archivos", {}))))


if __name__ == "__main__":
    main()
