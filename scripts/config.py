"""Configuración de filtros. Edita estas listas para ajustar qué licitaciones se guardan."""

# Se guarda una licitación si su nombre o descripción contiene AL MENOS UNA de estas palabras
# (sin distinguir mayúsculas ni tildes).
PALABRAS_CLAVE = [
    "electric",          # eléctrico, eléctrica, electricidad, electricista
    "subestacion",
    "media tension",
    "baja tension",
    "alta tension",
    "transformador",
    "tablero",
    "proteccion",        # protecciones eléctricas
    "scada",
    "automatizacion",
    "instrumentacion",
    "telecomunicacion",
    "fibra optica",
    "radiocomunicacion",
    "cctv",
    "control de acceso",
    "fotovoltaic",
    "alumbrado",
    "grupo electrogeno",
    "ups",
    "malla a tierra",
    "puesta a tierra",
]

# Licitaciones que contengan estas palabras se descartan aunque calcen arriba.
PALABRAS_EXCLUIDAS = [
    "insumos de oficina",
    "articulos de aseo",
]

# Regiones a conservar (texto contenido en RegionUnidad). Lista vacía = todo Chile.
REGIONES = []

# Estado a consultar en la API: "activas" (publicadas y abiertas) es lo habitual.
ESTADO = "activas"

# Pausa entre consultas de detalle (segundos) para respetar los límites de la API.
PAUSA_SEGUNDOS = 0.6
