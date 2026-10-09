"""Filtros basados en los servicios de DATAELECT (www.dataelect.cl).

Cada servicio tiene una lista de patrones. Una licitación se guarda si su nombre o
descripción calza con al menos un patrón; en el Excel aparece en qué servicio(s) calzó.

Los patrones se comparan sin mayúsculas ni tildes y son expresiones regulares:
  - "electric"            calza con eléctrico, eléctrica, electricidad, electricista...
  - "\\bups\\b"           solo la palabra UPS completa
  - "curso.{0,60}protec"  "curso" seguido (a menos de 60 caracteres) de "protec..."
Cada patrón debe calzar al comienzo de una palabra.
"""

SERVICIOS = {
    "Ingeniería eléctrica y potencia": [
        r"electric", r"subestacion", r"transformador", r"switchgear",
        r"\bccm\b", r"centro de control de motores", r"sala electrica",
        r"tableros? (electric|de fuerza|de distribucion|general)",
        r"media tension", r"alta tension", r"baja tension",
        r"cortocircuito", r"flujo de carga", r"coordinacion de protecciones",
        r"protecciones electricas", r"rele[s]? de proteccion",
        r"puesta a tierra", r"malla a tierra", r"sistema de tierra",
        r"iluminacion", r"alumbrado", r"luminaria",
        r"grupos? electrogeno", r"generador(es)? electric", r"respaldo de energia",
        r"\bups\b", r"banco de baterias", r"\bbess\b", r"almacenamiento de energia",
        r"fotovoltaic", r"paneles? solar", r"empalme electric", r"\bte1\b",
        r"unilineal", r"linea(s)? de (transmision|distribucion)", r"eficiencia energetica",
    ],
    "Telecomunicaciones": [
        r"telecomunicacion", r"fibra optica", r"cableado estructurado",
        r"radiocomunicacion", r"radio enlace", r"radioenlace", r"\bvhf\b", r"\buhf\b",
        r"telefonia ip", r"\bvoip\b", r"\bwi-?fi\b", r"red(es)? de datos",
        r"\bcctv\b", r"camaras de (vigilancia|seguridad|televigilancia)", r"televigilancia",
        r"control de acceso", r"sistemas? de seguridad electronic", r"sistemas? de alarma",
        r"deteccion de incendio", r"redes? (lan|wan)", r"sala de (servidores|comunicaciones)",
        r"\brack\b", r"enlace de datos", r"canalizacion(es)? (electric|de datos|de fibra)",
    ],
    "Control e instrumentación": [
        r"\bplc\b", r"\bdcs\b", r"\bscada\b", r"\bhmi\b", r"automatizacion",
        r"instrumentacion", r"control de procesos", r"telemetria", r"telecontrol",
        r"iec ?61850", r"modbus", r"\bdnp3\b", r"iec ?60870",
        r"redes? industrial", r"ciberseguridad (ot|industrial)", r"variadores? de frecuencia",
    ],
    "Ingeniería estructural": [
        r"estructuras? metalica", r"fundaciones? (para|de) (equipos|transformador)",
        r"soportes? de bandeja", r"bandeja portaconductor", r"sala electrica modular",
        r"memoria de calculo estructural",
    ],
    "Construcción y montaje": [
        r"instalacion(es)? electric", r"normalizacion electric", r"regularizacion electric",
        r"certificacion (te1|sec)", r"montaje electric", r"montaje de (tableros|equipos|luminarias)",
        r"inspeccion tecnica de obra", r"\bito\b", r"puesta en servicio",
        r"mantencion (electric|de tableros|de subestacion|de grupo)",
    ],
    "Dibujo técnico 2D/3D": [
        r"dibujo tecnico", r"digitalizacion de planos", r"planos? (electric|as.?built)",
        r"\bas.?built\b", r"\bcad\b", r"modelamiento 3d", r"levantamiento de planos",
    ],
    "Traducción técnica": [
        r"traduccion(es)? tecnic", r"traduccion.{0,60}(normas?|manuales?)",
    ],
    "Capacitación técnica": [
        r"(capacitacion|curso|diplomado|taller).{0,60}(electric|protecciones|transformador|"
        r"potencia|automatizacion|instrumentacion|scada|plc|fibra optica|telecomunicacion|"
        r"lectura de planos|alta tension|media tension)",
    ],
}

# Si alguno de estos patrones aparece, la licitación se descarta (ruido frecuente).
PATRONES_EXCLUIDOS = [
    r"tablero sensorial", r"material(es)? didactic", r"juguete", r"\btermos?\b",
    r"electrodomestic", r"electrocardiograf", r"electroencefalo", r"electrobisturi",
    r"electromedicina", r"vehiculos? electric", r"bicicleta", r"cigarrillo",
]

# Regiones a conservar (texto contenido en RegionUnidad). Lista vacía = todo Chile.
REGIONES = []

# Estado a consultar en la API: "activas" (publicadas y abiertas) es lo habitual.
ESTADO = "activas"

# Pausa entre consultas de detalle (segundos) para respetar los límites de la API.
PAUSA_SEGUNDOS = 0.6
