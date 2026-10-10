"""
Exporta licitaciones de Mercado Público (ChileCompra) a Excel y deja trazabilidad en el repo.

Modos:
  python scripts/licitaciones.py api       -> consulta la API, filtra y genera el Excel del día
  python scripts/licitaciones.py entrada   -> procesa los Excel subidos a la carpeta entrada/

Requiere la variable de entorno MP_TICKET (ticket de la API de Mercado Público) para el modo api.
"""
import csv
import datetime as dt
import gzip
import json
import os
import re
import shutil
import sys
import time
import unicodedata
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).parent))
import config  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "datos"
DIARIO = DATOS / "diario"
ENTRADA = RAIZ / "entrada"
ENTRADA_PROC = DATOS / "entrada_procesada"
HISTORIAL = DATOS / "historial.csv"
REGISTRO = RAIZ / "REGISTRO.md"

API = "https://api.mercadopublico.cl/servicios/v1/publico/licitaciones.json"
TZ = dt.timezone(dt.timedelta(hours=-3))  # Chile continental (aprox.; solo para nombrar archivos)

COLUMNAS = [
    ("Codigo", 18), ("Nombre", 60), ("Estado", 12), ("Tipo", 8), ("Organismo", 40),
    ("Unidad", 30), ("Region", 22), ("Comuna", 18), ("FechaPublicacion", 18),
    ("FechaCierre", 18), ("FechaVisita", 18), ("VisitaObligatoria", 10), ("FuenteVisita", 14),
    ("MontoEstimado", 18), ("Moneda", 8), ("MontoVisible", 10), ("Descripcion", 80),
    ("ServicioDATAELECT", 30), ("PalabrasCalzadas", 30), ("URL", 50),
]
CAMPOS_HIST = ["Codigo", "Nombre", "Organismo", "Region", "FechaCierre", "Estado",
               "PrimeraVez", "UltimaVez"]


def normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return texto.lower()


def ahora():
    return dt.datetime.now(TZ)


# ---------------------------------------------------------------- API
def get_json(params, ticket, intentos=4):
    params = {**params, "ticket": ticket}
    for i in range(intentos):
        r = requests.get(API, params=params, timeout=60)
        if r.status_code == 200:
            datos = r.json()
            if "Codigo" in datos and "Listado" not in datos:  # error de la API (p. ej. 10500)
                print(f"  aviso API: {datos.get('Mensaje')}", file=sys.stderr)
                time.sleep(5 * (i + 1))
                continue
            return datos
        time.sleep(5 * (i + 1))
    raise RuntimeError(f"La API no respondió correctamente tras {intentos} intentos: {params.get('codigo') or params}")


_SERVICIOS = {srv: [re.compile(r"\b" + p) for p in pats] for srv, pats in config.SERVICIOS.items()}
_EXCLUIDOS = [re.compile(r"\b" + p) for p in config.PATRONES_EXCLUIDOS]


def calza(texto):
    """Devuelve {servicio: [términos encontrados]} según los servicios de config.py."""
    t = " ".join(normalizar(texto).split())
    if any(rx.search(t) for rx in _EXCLUIDOS):
        return {}
    out = {}
    for srv, regs in _SERVICIOS.items():
        hallados = []
        for rx in regs:
            m = rx.search(t)
            if m and m.group(0).strip() not in hallados:
                hallados.append(m.group(0).strip())
        if hallados:
            out[srv] = hallados
    return out


_MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
          "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}


def _fecha_api(valor):
    """'2026-10-13T10:00:00' -> '2026-10-13 10:00'. Ignora fechas vacías o de relleno (año 1900/0001)."""
    v = (valor or "").strip()
    if not v or v[:4] in ("0001", "1900"):
        return ""
    return v[:16].replace("T", " ")


def visita_terreno(lic):
    """Devuelve (fecha 'AAAA-MM-DD HH:MM' o '', obligatoria 'Sí'/'No'/'', fuente)."""
    fechas = lic.get("Fechas") or {}
    texto = " ".join(str(lic.get(k) or "") for k in ("Descripcion", "Nombre"))
    t = normalizar(texto)
    flag = lic.get("VisitaTerreno")
    obligatoria = ""
    if "visita" in t and "obligatori" in t:
        obligatoria = "Sí"
    elif str(flag) in ("1", "True", "true"):
        obligatoria = "Sí"
    # 1) campo oficial del detalle de la licitación
    for clave in ("FechaVisitaTerreno", "FechaVisita"):
        f = _fecha_api(fechas.get(clave) or lic.get(clave))
        if f:
            return f, obligatoria or "Sí", "Mercado Público"
    # 2) respaldo: fecha escrita en la descripción cerca de la palabra "visita"
    for m in re.finditer(r"visita", t):
        trozo = t[m.start(): m.start() + 220]
        n = re.search(r"(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{2,4})", trozo)
        if n:
            d, mth, y = int(n.group(1)), int(n.group(2)), int(n.group(3))
        else:
            n = re.search(r"(\d{1,2})\s+de\s+(" + "|".join(_MESES) + r")(?:\s+(?:de|del)\s+(\d{4}))?", trozo)
            if not n:
                continue
            d, mth = int(n.group(1)), _MESES[n.group(2)]
            y = int(n.group(3)) if n.group(3) else ahora().year
        if y < 100:
            y += 2000
        try:
            fecha = dt.date(y, mth, d)
        except ValueError:
            continue
        h = re.search(r"(\d{1,2})[:.](\d{2})\s*(?:hrs|horas|h\b)", trozo)
        hora = f" {int(h.group(1)):02d}:{h.group(2)}" if h and int(h.group(1)) < 24 else ""
        return fecha.isoformat() + hora, obligatoria, "Descripción"
    if str(flag) in ("0", "False", "false") and not obligatoria:
        return "", "No", "Mercado Público"
    return "", obligatoria, ""


def fila_desde_detalle(lic, palabras):
    comprador = lic.get("Comprador") or {}
    fechas = lic.get("Fechas") or {}
    codigo = lic.get("CodigoExterno", "")
    visita = visita_terreno(lic)
    try:
        monto = float(lic.get("MontoEstimado") or 0) or ""
    except (TypeError, ValueError):
        monto = ""
    return {
        "Codigo": codigo,
        "Nombre": lic.get("Nombre", ""),
        "Estado": lic.get("Estado", ""),
        "Tipo": lic.get("Tipo", ""),
        "Organismo": comprador.get("NombreOrganismo", ""),
        "Unidad": comprador.get("NombreUnidad", ""),
        "Region": comprador.get("RegionUnidad", ""),
        "Comuna": comprador.get("ComunaUnidad", ""),
        "FechaPublicacion": (fechas.get("FechaPublicacion") or "")[:16].replace("T", " "),
        "FechaCierre": (fechas.get("FechaCierre") or lic.get("FechaCierre") or "")[:16].replace("T", " "),
        "FechaVisita": visita[0],
        "VisitaObligatoria": visita[1],
        "FuenteVisita": visita[2],
        "MontoEstimado": monto,
        "Moneda": lic.get("Moneda", ""),
        "MontoVisible": {1: "Sí", 0: "No"}.get(lic.get("VisibilidadMonto"), ""),
        "Descripcion": (lic.get("Descripcion") or "").strip(),
        "ServicioDATAELECT": " | ".join(palabras),
        "PalabrasCalzadas": ", ".join(sorted({w for v in palabras.values() for w in v}))[:200],
        "URL": f"https://www.mercadopublico.cl/Procurement/Modules/RFB/DetailsAcquisition.aspx?idlicitacion={codigo}",
    }


def modo_api():
    ticket = os.environ.get("MP_TICKET", "").strip()
    if not ticket:
        sys.exit("Falta MP_TICKET (ticket de la API de Mercado Público).")

    print(f"Consultando licitaciones en estado '{config.ESTADO}'...")
    listado = get_json({"estado": config.ESTADO}, ticket).get("Listado", [])
    print(f"  {len(listado)} licitaciones en el listado.")

    # Primer filtro barato por nombre; luego se confirma con la descripción del detalle.
    candidatas = [l for l in listado if calza(l.get("Nombre"))]
    # Las que no calzan por nombre pero podrían calzar por descripción se revisan solo
    # si el listado es manejable (evita miles de consultas).
    print(f"  {len(candidatas)} calzan por nombre; descargando detalle...")

    filas, crudos = [], []
    for i, l in enumerate(candidatas, 1):
        det = get_json({"codigo": l["CodigoExterno"]}, ticket).get("Listado") or [l]
        lic = det[0]
        crudos.append(lic)
        palabras = calza(f"{lic.get('Nombre')} {lic.get('Descripcion')}")
        if not palabras:
            continue
        fila = fila_desde_detalle(lic, palabras)
        if config.REGIONES and not any(normalizar(r) in normalizar(fila["Region"]) for r in config.REGIONES):
            continue
        filas.append(fila)
        if i % 25 == 0:
            print(f"    {i}/{len(candidatas)}")
        time.sleep(config.PAUSA_SEGUNDOS)

    filas.sort(key=lambda f: f["FechaCierre"] or "9999")
    hoy = ahora().strftime("%Y-%m-%d")
    DIARIO.mkdir(parents=True, exist_ok=True)
    xlsx = DIARIO / f"licitaciones_{hoy}.xlsx"
    escribir_excel(filas, xlsx)
    # Detalle completo tal como lo entrega Mercado Público (respaldo y trazabilidad)
    (DATOS / "detalle").mkdir(parents=True, exist_ok=True)
    with gzip.open(DATOS / "detalle" / f"detalle_{hoy}.json.gz", "wt", encoding="utf-8") as fh:
        json.dump(crudos, fh, ensure_ascii=False)
    con_visita = sum(1 for f in filas if f["FechaVisita"])
    print(f"  Con fecha de visita: {con_visita} · con monto: {sum(1 for f in filas if f['MontoEstimado'])}")
    escribir_csv(filas, DIARIO / f"licitaciones_{hoy}.csv")
    nuevas = actualizar_historial(filas)
    registrar(f"API Mercado Público ({config.ESTADO})", xlsx, len(filas), nuevas)
    print(f"Listo: {len(filas)} licitaciones ({len(nuevas)} nuevas) -> {xlsx.relative_to(RAIZ)}")


# ---------------------------------------------------------------- Excel subidos a mano
def modo_entrada():
    archivos = sorted(p for p in ENTRADA.glob("*.xls*") if not p.name.startswith("~$"))
    if not archivos:
        print("No hay Excel nuevos en entrada/.")
        return
    ENTRADA_PROC.mkdir(parents=True, exist_ok=True)
    sello = ahora().strftime("%Y-%m-%d_%H%M")
    for arch in archivos:
        destino = ENTRADA_PROC / f"{sello}_{arch.name}"
        filas = leer_excel_generico(arch)
        shutil.move(str(arch), destino)
        # Copia en CSV para que GitHub muestre las diferencias fila por fila entre versiones.
        escribir_csv(filas, destino.with_suffix(".csv"), columnas=list(filas[0].keys()) if filas else [])
        nuevas = actualizar_historial(filas) if filas and "Codigo" in filas[0] else []
        registrar(f"Excel subido: {arch.name}", destino, len(filas), nuevas)
        print(f"Procesado {arch.name}: {len(filas)} filas -> {destino.relative_to(RAIZ)}")


def leer_excel_generico(ruta):
    """Lee la primera hoja; detecta la fila de encabezados y mapea el código de licitación."""
    wb = load_workbook(ruta, read_only=True, data_only=True)
    hoja = wb.worksheets[0]
    filas_crudas = [list(r) for r in hoja.iter_rows(values_only=True)]
    wb.close()
    # Encabezado = primera fila con al menos 3 celdas de texto
    idx = next((i for i, r in enumerate(filas_crudas)
                if sum(isinstance(c, str) and c.strip() != "" for c in r) >= 3), 0)
    enc = [str(c).strip() if c is not None else f"col{j}" for j, c in enumerate(filas_crudas[idx])]
    # Renombra la columna de código (Mercado Público la llama "Código", "ID", "CodigoExterno", etc.)
    for j, h in enumerate(enc):
        n = normalizar(h)
        if n in {"codigo", "codigo externo", "codigoexterno", "id licitacion", "id", "nro licitacion"}:
            enc[j] = "Codigo"
            break
    filas = []
    for r in filas_crudas[idx + 1:]:
        if all(c in (None, "") for c in r):
            continue
        filas.append({enc[j]: ("" if c is None else c) for j, c in enumerate(r[:len(enc)])})
    return filas


# ---------------------------------------------------------------- Salidas y trazabilidad
def escribir_excel(filas, ruta):
    wb = Workbook()
    ws = wb.active
    ws.title = "Licitaciones"
    enc = [c for c, _ in COLUMNAS]
    ws.append(enc)
    for f in filas:
        ws.append([f.get(c, "") for c in enc])
    azul = PatternFill("solid", fgColor="1F4E78")
    for j, (_, ancho) in enumerate(COLUMNAS, 1):
        celda = ws.cell(row=1, column=j)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = azul
        celda.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = ancho
    url_col = enc.index("URL") + 1
    for i in range(2, ws.max_row + 1):
        c = ws.cell(row=i, column=url_col)
        if c.value:
            c.hyperlink = c.value
            c.font = Font(color="0563C1", underline="single")
        for j in (enc.index("Nombre") + 1, enc.index("Descripcion") + 1):
            ws.cell(row=i, column=j).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=i, column=enc.index("MontoEstimado") + 1).number_format = '#,##0'
        v = ws.cell(row=i, column=enc.index("FechaVisita") + 1)
        if v.value:
            v.fill = PatternFill("solid", fgColor="FFF2CC")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    res = wb.create_sheet("Por servicio")
    res.append(["Servicio DATAELECT", "Licitaciones"])
    for srv in config.SERVICIOS:
        res.append([srv, sum(srv in str(f.get("ServicioDATAELECT", "")) for f in filas)])
    res.column_dimensions["A"].width = 36
    for c in res[1]:
        c.font = Font(bold=True)
    info = wb.create_sheet("Info")
    info.append(["Generado", ahora().strftime("%Y-%m-%d %H:%M")])
    info.append(["Fuente", "API Mercado Público (ChileCompra)"])
    info.append(["Estado consultado", config.ESTADO])
    info.append(["Servicios filtrados", ", ".join(config.SERVICIOS)])
    info.append(["Total", len(filas)])
    wb.save(ruta)


def escribir_csv(filas, ruta, columnas=None):
    columnas = columnas or [c for c, _ in COLUMNAS]
    with open(ruta, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=columnas, extrasaction="ignore")
        w.writeheader()
        w.writerows(filas)


def actualizar_historial(filas):
    """Historial maestro: una fila por licitación con primera y última vez vista."""
    hist = {}
    if HISTORIAL.exists():
        with open(HISTORIAL, encoding="utf-8-sig") as fh:
            hist = {r["Codigo"]: r for r in csv.DictReader(fh)}
    hoy = ahora().strftime("%Y-%m-%d")
    nuevas = []
    for f in filas:
        cod = str(f.get("Codigo", "")).strip()
        if not cod:
            continue
        previo = hist.get(cod)
        if previo is None:
            nuevas.append(cod)
            previo = {"PrimeraVez": hoy}
        hist[cod] = {
            "Codigo": cod,
            "Nombre": f.get("Nombre", previo.get("Nombre", "")),
            "Organismo": f.get("Organismo", previo.get("Organismo", "")),
            "Region": f.get("Region", previo.get("Region", "")),
            "FechaCierre": f.get("FechaCierre", previo.get("FechaCierre", "")),
            "Estado": f.get("Estado", previo.get("Estado", "")),
            "PrimeraVez": previo["PrimeraVez"],
            "UltimaVez": hoy,
        }
    DATOS.mkdir(parents=True, exist_ok=True)
    with open(HISTORIAL, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=CAMPOS_HIST)
        w.writeheader()
        for r in sorted(hist.values(), key=lambda r: (r["PrimeraVez"], r["Codigo"]), reverse=True):
            w.writerow(r)
    return nuevas


def registrar(origen, archivo, total, nuevas):
    if not REGISTRO.exists():
        REGISTRO.write_text(
            "# Registro de exportaciones\n\n"
            "| Fecha | Origen | Archivo | Licitaciones | Nuevas |\n"
            "|---|---|---|---|---|\n", encoding="utf-8")
    rel = archivo.relative_to(RAIZ).as_posix()
    linea = f"| {ahora():%Y-%m-%d %H:%M} | {origen} | [{archivo.name}]({rel}) | {total} | {len(nuevas)} |\n"
    with open(REGISTRO, "a", encoding="utf-8") as fh:
        fh.write(linea)
    resumen = os.environ.get("GITHUB_STEP_SUMMARY")
    if resumen:
        with open(resumen, "a", encoding="utf-8") as fh:
            fh.write(f"### {origen}\n- Archivo: `{rel}`\n- Licitaciones: {total}\n- Nuevas: {len(nuevas)}\n")
            if nuevas:
                fh.write("- Códigos nuevos: " + ", ".join(nuevas[:50]) + "\n")


if __name__ == "__main__":
    modo = sys.argv[1] if len(sys.argv) > 1 else "api"
    {"api": modo_api, "entrada": modo_entrada}.get(modo, lambda: sys.exit(f"Modo desconocido: {modo}"))()
