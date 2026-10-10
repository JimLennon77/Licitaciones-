# Licitaciones Mercado Público → GitHub

Exporta automáticamente licitaciones de Mercado Público (ChileCompra) a Excel y guarda cada exportación en este repositorio, de modo que el historial de commits queda como registro de trazabilidad.

## Qué hace

| Cuándo | Qué pasa |
|---|---|
| Todos los días ~07:45 (Chile) | Consulta la API, filtra por palabras clave y genera `datos/diario/licitaciones_AAAA-MM-DD.xlsx` (+ `.csv`) |
| Cada vez que subes un Excel a `entrada/` | Lo mueve a `datos/entrada_procesada/` con fecha y hora, crea una copia `.csv` y lo registra |
| Botón **Run workflow** (pestaña Actions) | Ejecuta ambas cosas a demanda |

Siempre se actualizan:
- `datos/historial.csv`: una fila por licitación, con la primera y la última vez que apareció.
- `REGISTRO.md`: bitácora de cada exportación (fecha, origen, archivo, total, nuevas).
- Un commit con fecha y hora. En GitHub puedes ver qué cambió entre días comparando los `.csv`.

## Puesta en marcha (una vez)

1. **Ticket de la API**: solicítalo en https://www.chilecompra.cl/api/ (con tu RUT o ClaveÚnica).
2. **Guardar el ticket como secreto**: en el repositorio, ve a *Settings → Secrets and variables → Actions → New repository secret*, con nombre `MP_TICKET` y el ticket como valor.
3. **Permisos de escritura**: en *Settings → Actions → General → Workflow permissions*, marca **Read and write permissions**.
4. **Primera prueba**: en *Actions → Licitaciones Mercado Público → Run workflow*.

## Ajustar filtros

Edita `scripts/config.py`:
- `SERVICIOS`: las líneas de servicio de DATAELECT (eléctrica y potencia, telecomunicaciones, control e instrumentación, estructural, construcción y montaje, dibujo técnico, traducción y capacitación), cada una con sus términos de búsqueda. El Excel indica en qué servicio calzó cada licitación y trae una hoja de resumen por servicio.
- `PATRONES_EXCLUIDOS`: para descartar ruido (por ejemplo, "tablero sensorial" o equipos médicos).
- `REGIONES`: si la dejas vacía, incluye todo Chile.
- `ESTADO`: por defecto `activas`.

## Subir un Excel a mano

Arrastra el archivo a la carpeta `entrada/` desde la web de GitHub (*Add file → Upload files*) y haz commit. El proceso corre solo en un par de minutos.

## Ejecutar en tu PC (opcional)

```bash
pip install -r requirements.txt
set MP_TICKET=tu_ticket          # Windows (PowerShell: $env:MP_TICKET="tu_ticket")
python scripts/licitaciones.py api
```

## Tablero en GitHub Projects

Cada ejecución diaria agrega las licitaciones **nuevas** como tarjetas en el proyecto
[JimLennon77 / proyecto 1](https://github.com/users/JimLennon77/projects/1):

- Título `[Código] Nombre`, con organismo, fechas, monto, descripción y enlace en el cuerpo.
- Campos: Código, Organismo, Región, Servicio DATAELECT, Cierre, **Visita a terreno**, **Monto estimado**, URL (se crean solos si no existen).
- La fecha de visita sale del detalle oficial de la licitación (`FechaVisitaTerreno`); si viene vacía, se busca en la descripción. En el Excel, la columna `FuenteVisita` indica de dónde salió.
- Cierre, visita y monto se actualizan también en las tarjetas que ya estaban, sin tocar su Status.
- **Cada día se eliminan del tablero las tarjetas en estado Done** (o Terminado/Listo). Antes de eliminarlas se guardan en `datos/tablero_done.csv` y no se vuelven a agregar. Para desactivarlo, pon `ELIMINAR_DONE: "0"` en el workflow.
- Entran en la primera columna del Status (o en "Nueva"/"Por revisar" si existe). Desde ahí las mueves tú.
- No se duplican: si la tarjeta con ese código ya está, se salta.

Requiere un token personal guardado como secreto `PROJECT_TOKEN`:
1. GitHub → foto de perfil → *Settings → Developer settings → Personal access tokens → Tokens (classic) → Generate new token (classic)*.
2. Nombre: `licitaciones-project`, expiración a elección, y marca solo el permiso **project**.
3. Copia el token y guárdalo en el repo: *Settings → Secrets and variables → Actions → New repository secret*, nombre `PROJECT_TOKEN`.

Sin ese secreto, el paso se omite y el resto sigue funcionando igual.
