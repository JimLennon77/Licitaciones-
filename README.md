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
