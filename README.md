# Seguimiento ODK

Página única (`docs/index.html`) para ver qué subió cada mercaderista a ODK, qué
le falta y qué subió con novedad, en tres formularios:

| Pestaña | Formulario ODK | Universo (lista de entidades del formulario) | Qué se espera |
|---|---|---|---|
| Visibilidad | `seguimiento_muebles_alquilados_uiux_v4` | `muebles_por_local_lista` (718 locales) | 1 registro por local **cada mes** |
| Acuerdos comerciales | `acuerdos_comerciales` | `acuerdos_por_local_lista` (829 local × acuerdo) | 1 registro por local y acuerdo **cada semana ISO** (lunes a domingo) mientras el acuerdo esté vigente |
| Promociones | `promociones_mercaderista` | `locales_promociones_lista` (920 locales) | 1 registro por local **cada semana** (lunes a domingo) |

El universo sale de las mismas listas que usan los formularios, así que cuando
se actualiza una lista en ODK la página se ajusta sola en la siguiente corrida.

## Estados

| Estado | Significa |
|---|---|
| ✓ Completo | Registrado sin novedades |
| ! Con novedad | Registrado, pero reporta un problema (ver abajo) |
| … Pendiente | Sin registro y el periodo sigue abierto |
| ✕ No realizado | Sin registro y el periodo ya cerró |
| · Programado | El acuerdo todavía no empieza (solo Acuerdos) |

**Avance** = (Completos + Con novedad) ÷ esperados.

Qué cuenta como novedad:

* **Visibilidad**: mueble que no está en el local (con su motivo), mueble sin
  producto o no lleno, mueble dañado (tipo y detalle).
* **Acuerdos**: mueble no ejecutado (sin stock, sin alcance, espacio limitado,
  no autoriza el administrador + comentario), producto del mueble que falta,
  registro fuera de vigencia.
* **Promociones**: promociones no activas, agotadas, con precio distinto o con
  descuento menor al pactado (totales del cierre de cada registro). «Sin
  alcance» no se marca como novedad.
* En los tres: fotos que no terminaron de subir al servidor.

## Cómo se publica (sin servidor)

ODK Central no permite que una página web lo consulte directamente desde el
navegador (no envía cabeceras CORS). Por eso los datos se descargan cada 10 minutos
en GitHub Actions, se guardan dentro del HTML y la página se sirve gratis en
GitHub Pages. No depende de ninguna computadora encendida.

### Puesta en marcha (una sola vez)

1. Crear un repositorio en GitHub, por ejemplo `seguimiento-odk`, y subir esta
   carpeta:

   ```powershell
   git remote add origin https://github.com/<usuario>/seguimiento-odk.git
   git push -u origin main
   ```

2. En *Settings → Secrets and variables → Actions → New repository secret*:

   | Secret | Valor |
   |---|---|
   | `ODK_EMAIL` | usuario web de ODK Central con lectura del proyecto 1 |
   | `ODK_PASSWORD` | su contraseña |
   | `SEGUIMIENTO_CLAVE` | clave que pedirá la página para abrirse (recomendado) |

3. *Actions → Actualizar seguimiento → Run workflow*. Espera a que termine
   (unos 3 minutos).

4. *Settings → Pages*: Source `Deploy from a branch`, rama **`gh-pages`**,
   carpeta `/ (root)`. El link queda en
   `https://<usuario>.github.io/seguimiento-odk/`.

Después se actualiza solo cada 10 minutos, de 06:00 a 21:59 (hora de Ecuador), y la página abierta se recarga sola cuando hay datos nuevos (revisa cada 2 minutos). Entre que alguien sube un registro y aparece pasan normalmente 10 a 15 minutos.

### Privacidad

* La página **no** contiene cédulas, RUC ni coordenadas.
* Sí contiene nombres de supervisores, mercaderistas, locales y los
  comentarios de novedad. Con `SEGUIMIENTO_CLAVE` todo eso va **cifrado**
  (AES-256) dentro del HTML: sin la clave no se puede leer, aunque el
  repositorio sea público. El navegador recuerda la clave después del primer
  ingreso.
* Para cambiar la clave, edita el secret y vuelve a correr el workflow.

## Generar la página en esta computadora

```powershell
python -m pip install -r requirements.txt
python build_seguimiento.py --env "..\ODK MUEBLES VISIBILIDAD\.env"
```

Escribe `docs/index.html` (unos 2,5 MB), que se abre con doble clic.

## Ajustes en el código

En [`build_seguimiento.py`](build_seguimiento.py):

| Constante | Qué cambia |
|---|---|
| `INICIO_VIS` | primer mes que se muestra en Visibilidad |
| `SEMANAS_PROMO` | cuántas semanas hacia atrás muestra Promociones |
| `MOTIVOS_AC`, `MOTIVOS_VIS`, `ESTADOS_PR` | cómo se redactan los motivos de novedad |

La interfaz está en [`plantilla.html`](plantilla.html). El generador reemplaza
`/*__CARGA__*/null` por los datos.
