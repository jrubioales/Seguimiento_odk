# Seguimiento ODK

Página única (`docs/index.html`) para ver qué subió cada mercaderista a ODK, qué
le falta y qué subió con novedad, en estos formularios:

| Pestaña | Formulario ODK | Universo (lista de entidades del formulario) | Qué se espera |
|---|---|---|---|
| Visibilidad | `seguimiento_muebles_alquilados_uiux_v4` | `muebles_por_local_lista` (718 locales) | 1 registro por local **cada mes** |
| Acuerdos comerciales | `acuerdos_comerciales` | `acuerdos_por_local_lista` (829 local × acuerdo) | 1 registro por local y acuerdo **cada 2 semanas** (bloques de lunes a domingo de la semana siguiente, desde el 2 nov 2026: semanas 45-46, 47-48…; antes era cada semana ISO) mientras el acuerdo esté vigente |
| Promociones | `promociones_mercaderista` | `locales_promociones_lista` (920 locales) | 1 registro por local **cada semana** (lunes a domingo) |
| Exhibiciones ganadas | `extravisibilidad_ganados_moderno` | `ganadas_moderno` (una entidad por exhibición) | No hay registros esperados: mide exhibiciones **activas, ganadas y perdidas** por semana ISO |
| SOS percha | `sos_participacion_percha_moderno` | `tiendas_moderno_1008.csv` (447 locales) | 1 registro por local **cada mes** |
| Planes comerciales | `planes_comerciales` | `locales_planes_1028.csv` (`plan_n` = Plan Ninja, `plan_d` = Plan Aliado Diamante) | 1 registro por local **y plan** cada mes |
| Check outs | `check_outs` | `locales_checkouts.csv` (farmacias y gasolineras) | 1 registro por local **cada mes** |
| KBDS | `kbds` | `locales_kbds.csv` (442 locales) | 1 registro por local **cada mes** |

### Encuestas mensuales (SOS, planes, check outs, KBDS)

La meta es la lista de locales del propio formulario y se reinicia el día 1
de cada mes: en diciembre todo vuelve a «falta subir». Se ve el avance por
zona, por cadena / cliente, por mercaderista o local por local, con filtros de
zona, cadena / cliente y, en Planes y Check outs, plan o canal. El
mercaderista y el supervisor salen de las listas de Promociones, Visibilidad
y Acuerdos (mismo código de local); si el local no está en ellas, se usa
quien lo subió. Planes comerciales usa otros códigos de local, por eso casi
todos sus locales quedan «Sin asignar». Los meses antes del primer registro
de una encuesta no se muestran.

Cuenta como novedad: SOS que no se pudo medir o categorías que ya no están,
local de Planes sin aceite de girasol, check out inexistente, KBDS que no se
pudo levantar, y fotos sin terminar de subir.

### Exhibiciones ganadas

Cada exhibición ganada en canal moderno se registra una vez (fecha de alta). En
las visitas siguientes el mercaderista confirma si sigue o la da de baja con
un motivo. **Si una semana no se sube encuesta, la exhibición sigue activa**:
solo una baja la termina.

Por cada semana se muestra:

* **Activas**: ganadas hasta el cierre de la semana (o hasta hoy) y sin baja.
* **Ganadas / Perdidas**: altas y bajas dentro de la semana.
* **Llevan activas**: días promedio desde el alta de las activas.
* **Duran (perdidas)**: días promedio entre alta y baja de las que se perdieron.

Se agrupa por negocio, categoría, SKU, zona, cadena, tipo de exhibición, local
o mercaderista. Una exhibición con varios productos cuenta en cada negocio,
categoría o SKU que tiene. La categoría sale del catálogo de SKU del propio
formulario. Esta pestaña tiene además filtros por canal (Moderno,
Tradicional, Farmacias / especiales) y cadena; las exhibiciones registradas
antes de que el formulario tuviera canal cuentan como Moderno.

#### Fotos y PPT

La vista **Fotos** muestra la foto que se tomó al ganar cada exhibición. Si
sigue activa, el formulario pide otra foto de seguimiento cuando pasaron 7 días
o más desde la última (hasta 2 al mes): esas exhibiciones se ven con dos fotos,
la de cuando se ganó y la de seguimiento más reciente; al ampliar se puede
pasar por todas. La galería carga miniaturas livianas (unos 20 KB); al ampliar
o descargar una foto se usa la original (hasta 1024 px, como la limita el
formulario).

**Armar PPT** genera en el navegador una presentación con el formato de
`formato.pptx`: la portada con el logo y luego, por cadena, diapositivas de 4
fotos con su estado (nueva, activa o perdida), local, tipo, categoría, canal,
SKU y fecha. Una exhibición con foto de seguimiento ocupa dos espacios: la foto
de cuando se ganó y la de seguimiento, cada una con su fecha. Usa la semana, los filtros y la lista elegidos (activas, ganadas en la
semana, perdidas o todas).

* `plantilla_ppt.pptx` es `formato.pptx` con solo la portada (las fotos de
  ejemplo pesaban 27 MB). Si cambias el formato, recórtalo igual y reemplázalo.
* Las fotos se guardan en la rama **`fotos`** del repositorio (`t/` miniaturas,
  `g/` originales). Cada corrida descarga de ODK solo las nuevas. Con
  `SEGUIMIENTO_CLAVE` van cifradas con la misma clave de la página; si la
  clave cambia, la rama se rehace sola en la siguiente corrida.
* Para probar sin GitHub: `FOTOS_REMOTO` (repositorio git donde subirlas) y
  `FOTOS_BASE` (URL desde donde las lee la página). `--sin-fotos` las omite.

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

* La página **no** contiene cédulas, RUC ni coordenadas. Las fotos de
  exhibiciones pueden mostrar la ubicación que imprime la cámara; van cifradas
  igual que los datos.
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
| `PERIODOS_AC`, `INICIO_QUINCENA_AC` | cuántos periodos muestra Acuerdos y desde qué lunes se mide cada 2 semanas |
| `SEMANAS_EX` | cuántas semanas hacia atrás muestra Exhibiciones ganadas |
| `MESES_MENS` | cuántos meses muestran SOS, planes, check outs y KBDS |
| `PLANES` | qué columna de la lista marca cada plan comercial |
| `MOTIVOS_AC`, `MOTIVOS_VIS`, `ESTADOS_PR` | cómo se redactan los motivos de novedad |

La interfaz está en [`plantilla.html`](plantilla.html). El generador reemplaza
`/*__CARGA__*/null` por los datos.
