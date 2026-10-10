"""
Publica las licitaciones del día como tarjetas en el GitHub Project del usuario.

Uso:  python scripts/proyecto.py datos/diario/licitaciones_AAAA-MM-DD.csv

Variables de entorno:
  PROJECT_TOKEN   token personal (classic) con permiso "project"
  PROJECT_OWNER   usuario dueño del proyecto (por defecto JimLennon77)
  PROJECT_NUMBER  número del proyecto (por defecto 1)

Qué hace:
  - Crea (si faltan) los campos: Código, Organismo, Región, Servicio DATAELECT, Cierre,
    Visita a terreno, Monto estimado y URL. Si el proyecto ya tiene un campo con ese nombre, lo reutiliza.
  - Agrega cada licitación nueva como borrador "[Código] Nombre", con el detalle en el cuerpo.
  - No duplica: si ya existe una tarjeta con ese código, la salta.
  - Deja el Status en la primera opción (p. ej. "Todo" / "Nueva"), salvo que exista una
    opción cuyo nombre esté en STATUS_INICIAL.
  - Actualiza en las tarjetas existentes los campos que pueden cambiar (Cierre, Visita a
    terreno, Monto estimado) sin tocar el Status ni los campos propios del usuario.
  - Elimina del tablero las tarjetas en estado "Done" (o Terminado/Listo/Finalizado), dejando
    antes un registro en datos/tablero_done.csv; esas licitaciones no se vuelven a agregar.
"""
import csv
import os
import unicodedata
import re
import sys
import time

import requests

API = "https://api.github.com/graphql"
TOKEN = os.environ.get("PROJECT_TOKEN", "").strip()
OWNER = os.environ.get("PROJECT_OWNER", "JimLennon77")
NUMBER = int(os.environ.get("PROJECT_NUMBER", "1"))
MAX_NUEVAS = int(os.environ.get("PROJECT_MAX_NUEVAS", "300"))
STATUS_INICIAL = ["Nueva", "Nuevas", "Por revisar", "Backlog", "Todo", "To do"]
STATUS_DONE = {"done", "terminado", "terminada", "listo", "lista", "finalizado", "finalizada", "hecho"}
ELIMINAR_DONE = os.environ.get("ELIMINAR_DONE", "1") == "1"
RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
REGISTRO_DONE = os.path.join(RAIZ, "datos", "tablero_done.csv")

CAMPOS = {  # nombre en el proyecto -> (tipo, columna del CSV)
    "Código": ("TEXT", "Codigo"),
    "Organismo": ("TEXT", "Organismo"),
    "Región": ("TEXT", "Region"),
    "Servicio DATAELECT": ("SINGLE_SELECT", "ServicioDATAELECT"),
    "Cierre": ("DATE", "FechaCierre"),
    "Visita a terreno": ("DATE", "FechaVisita"),
    "Monto estimado": ("TEXT", "MontoTexto"),
    "URL": ("TEXT", "URL"),
}
# Campos que se actualizan también en tarjetas ya existentes (pueden cambiar o venir vacíos antes).
CAMPOS_ACTUALIZABLES = ["Cierre", "Visita a terreno", "Monto estimado"]
SERVICIOS = ["Ingeniería eléctrica y potencia", "Telecomunicaciones", "Control e instrumentación",
             "Ingeniería estructural", "Construcción y montaje", "Dibujo técnico 2D/3D",
             "Traducción técnica", "Capacitación técnica"]
COLORES = ["BLUE", "GREEN", "PURPLE", "GRAY", "ORANGE", "PINK", "YELLOW", "RED"]


def gql(query, **variables):
    for intento in range(6):
        try:
            r = requests.post(API, json={"query": query, "variables": variables},
                              headers={"Authorization": f"Bearer {TOKEN}"}, timeout=60)
        except requests.RequestException as e:
            print(f"  red: {e}; reintentando", flush=True)
            time.sleep(15 * (intento + 1))
            continue
        texto = r.text.lower()
        limite = ("rate limit" in texto or "too quickly" in texto or "abuse" in texto
                  or r.status_code == 429)
        if r.status_code in (500, 502, 503, 504) or limite:
            espera = int(r.headers.get("Retry-After", 0) or 0) or (60 if limite else 10) * (intento + 1)
            print(f"  GitHub pide esperar ({r.status_code}); pausa de {espera}s", flush=True)
            time.sleep(espera)
            continue
        try:
            datos = r.json()
        except ValueError:
            raise RuntimeError(f"Respuesta no válida de GitHub (HTTP {r.status_code}): {r.text[:300]}")
        if r.status_code == 401:
            raise RuntimeError("Token rechazado (401): revisa el secreto PROJECT_TOKEN.")
        if datos.get("errors"):
            raise RuntimeError(f"HTTP {r.status_code}: {datos['errors']}")
        return datos["data"]
    raise RuntimeError("GitHub no respondió")


def cargar_proyecto():
    q = """query($login:String!,$n:Int!){ user(login:$login){ projectV2(number:$n){ id title
      fields(first:50){ nodes{ ... on ProjectV2FieldCommon{ id name dataType }
        ... on ProjectV2SingleSelectField{ id name dataType options{ id name } } } } } } }"""
    p = gql(q, login=OWNER, n=NUMBER)["user"]["projectV2"]
    if not p:
        sys.exit(f"No encontré el proyecto {NUMBER} de {OWNER} (¿el token tiene permiso 'project'?).")
    return p


def asegurar_campos(p):
    existentes = {f["name"].lower(): f for f in p["fields"]["nodes"] if f}
    for nombre, (tipo, _) in CAMPOS.items():
        previo = existentes.get(nombre.lower())
        if previo and nombre == "Monto estimado" and previo.get("dataType") != tipo:
            # Versión anterior era numérica; GitHub rechaza montos grandes, se reemplaza por texto.
            gql("""mutation($f:ID!){ deleteProjectV2Field(input:{fieldId:$f}){ clientMutationId } }""",
                f=previo["id"])
            print(f"  campo reemplazado: {nombre} ({previo.get('dataType')} -> {tipo})")
            previo = None
        if previo:
            continue
        variables = {"pid": p["id"], "name": nombre, "type": tipo}
        if tipo == "SINGLE_SELECT":
            variables["opts"] = [{"name": s, "color": COLORES[i % len(COLORES)], "description": ""}
                                 for i, s in enumerate(SERVICIOS)]
            q = """mutation($pid:ID!,$name:String!,$type:ProjectV2CustomFieldType!,
                   $opts:[ProjectV2SingleSelectFieldOptionInput!]){
                   createProjectV2Field(input:{projectId:$pid,name:$name,dataType:$type,
                   singleSelectOptions:$opts}){ clientMutationId } }"""
        else:
            q = """mutation($pid:ID!,$name:String!,$type:ProjectV2CustomFieldType!){
                   createProjectV2Field(input:{projectId:$pid,name:$name,dataType:$type}){ clientMutationId } }"""
        gql(q, **variables)
        print(f"  campo creado: {nombre}")
    return cargar_proyecto()


def tarjetas_existentes(pid):
    """Devuelve {código: {"id", "titulo", "valores": {campo: valor}}} de todas las tarjetas."""
    q = """query($pid:ID!,$after:String){ node(id:$pid){ ... on ProjectV2{
      items(first:100, after:$after){ pageInfo{ hasNextPage endCursor }
        nodes{ id content{ ... on DraftIssue{ title } ... on Issue{ title } }
          fieldValues(first:30){ nodes{
            ... on ProjectV2ItemFieldTextValue{ text field{ ... on ProjectV2FieldCommon{ name } } }
            ... on ProjectV2ItemFieldDateValue{ date field{ ... on ProjectV2FieldCommon{ name } } }
            ... on ProjectV2ItemFieldNumberValue{ number field{ ... on ProjectV2FieldCommon{ name } } }
            ... on ProjectV2ItemFieldSingleSelectValue{ name field{ ... on ProjectV2FieldCommon{ name } } }
          } } } } } } }"""
    tarjetas, after = {}, None
    while True:
        items = gql(q, pid=pid, after=after)["node"]["items"]
        for n in items["nodes"]:
            titulo = (n.get("content") or {}).get("title") or ""
            m = re.match(r"\s*\[([^\]]+)\]", titulo)
            if not m:
                continue
            valores = {}
            for fv in (n.get("fieldValues") or {}).get("nodes") or []:
                nombre = ((fv or {}).get("field") or {}).get("name")
                if nombre:
                    valores[nombre] = next((fv[k] for k in ("text", "date", "number", "name") if k in fv), None)
            tarjetas[m.group(1).strip()] = {"id": n["id"], "titulo": titulo, "valores": valores}
        if not items["pageInfo"]["hasNextPage"]:
            return tarjetas
        after = items["pageInfo"]["endCursor"]


def _norm(t):
    return unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower().strip()


def codigos_done_registrados():
    if not os.path.exists(REGISTRO_DONE):
        return set()
    with open(REGISTRO_DONE, encoding="utf-8-sig") as fh:
        return {r["Codigo"] for r in csv.DictReader(fh) if r.get("Codigo")}


def eliminar_done(p, tarjetas, status):
    """Registra y elimina del tablero las tarjetas con Status Done. Devuelve los códigos eliminados."""
    if not (ELIMINAR_DONE and status):
        return []
    hechas = [(c, t) for c, t in tarjetas.items() if _norm(t["valores"].get(status["name"])) in STATUS_DONE]
    if not hechas:
        return []
    cols = ["FechaEliminacion", "Codigo", "Titulo", "Status", "Organismo", "Region", "Cierre",
            "Visita a terreno", "Monto estimado", "Servicio DATAELECT", "URL"]
    nuevo = not os.path.exists(REGISTRO_DONE)
    os.makedirs(os.path.dirname(REGISTRO_DONE), exist_ok=True)
    borrar = """mutation($pid:ID!,$item:ID!){ deleteProjectV2Item(input:{projectId:$pid,itemId:$item}){ deletedItemId } }"""
    eliminadas = []
    with open(REGISTRO_DONE, "a", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        if nuevo:
            w.writeheader()
        for codigo, t in hechas:
            v = t["valores"]
            # Primero se registra (trazabilidad) y luego se elimina.
            w.writerow({"FechaEliminacion": time.strftime("%Y-%m-%d"), "Codigo": codigo, "Titulo": t["titulo"],
                        "Status": v.get(status["name"], ""), "Organismo": v.get("Organismo", ""),
                        "Region": v.get("Región", ""), "Cierre": v.get("Cierre", ""),
                        "Visita a terreno": v.get("Visita a terreno", ""), "Monto estimado": v.get("Monto estimado", ""),
                        "Servicio DATAELECT": v.get("Servicio DATAELECT", ""), "URL": v.get("URL", "")})
            fh.flush()
            try:
                gql(borrar, pid=p["id"], item=t["id"])
                eliminadas.append(codigo)
            except RuntimeError as e:
                print(f"  no se pudo eliminar {codigo}: {e}")
            time.sleep(0.5)
    for c in eliminadas:
        tarjetas.pop(c, None)
    print(f"Eliminadas del tablero (Done): {len(eliminadas)} — registradas en datos/tablero_done.csv")
    return eliminadas


def monto_texto(fila):
    """Monto legible: '$ 4.486.154.225' (o con la moneda si no es CLP)."""
    try:
        n = float(fila.get("MontoEstimado") or 0)
    except ValueError:
        n = 0
    if n > 0:
        moneda = (fila.get("Moneda") or "CLP").strip()
        cifra = f"{n:,.0f}".replace(",", ".") if moneda in ("CLP", "") else f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"$ {cifra}" if moneda in ("CLP", "") else f"{cifra} {moneda}"
    if (fila.get("MontoVisible") or "") == "No":
        return "No publicado"
    return ""


def valor_campo(campo, fila):
    tipo = campo["dataType"]
    col = next((c for n, (_, c) in CAMPOS.items() if n.lower() == campo["name"].lower()), None)
    bruto = monto_texto(fila) if col == "MontoTexto" else str(fila.get(col, "") or "").strip()
    if not bruto:
        return None
    if tipo == "DATE":
        return {"date": bruto[:10]} if re.match(r"\d{4}-\d{2}-\d{2}", bruto) else None
    if tipo == "NUMBER":
        try:
            n = float(bruto)
        except ValueError:
            return None
        # GitHub no acepta números sobre ~2.147 millones de millones; el monto completo queda en el cuerpo.
        return {"number": n} if 0 <= n < 2_147_483_647 else None
    if tipo == "SINGLE_SELECT":
        primero = bruto.split("|")[0].strip()
        op = next((o for o in campo.get("options", []) if o["name"].lower() == primero.lower()), None)
        return {"singleSelectOptionId": op["id"]} if op else None
    if tipo == "TEXT":
        return {"text": bruto[:1000]}
    return None


def escribir_log(titulo, total, previas, ok, errores, actualizadas=0, eliminadas=0):
    """Deja constancia en el repo (datos/proyecto_log.md) de cada publicación en el tablero."""
    ruta = os.path.join(os.path.dirname(__file__), "..", "datos", "proyecto_log.md")
    nuevo = not os.path.exists(ruta)
    with open(ruta, "a", encoding="utf-8") as fh:
        if nuevo:
            fh.write("# Publicaciones en GitHub Project\n\n")
        fh.write(f"## {time.strftime('%Y-%m-%d %H:%M UTC', time.gmtime())} · {titulo}\n"
                 f"- En archivo: {total} · ya en tablero: {previas} · agregadas: {ok} · actualizadas: {actualizadas}"
                 f" · eliminadas (Done): {eliminadas} · errores: {len(errores)}\n")
        for e in errores[:30]:
            fh.write(f"  - {e[:400]}\n")
        fh.write("\n")


def cuerpo(f):
    monto = monto_texto(f) or "no informado"
    visita = f.get("FechaVisita") or "no informada"
    if f.get("VisitaObligatoria") == "Sí":
        visita += " (obligatoria)"
    return (
        f"**Organismo:** {f.get('Organismo','')}  \n"
        f"**Unidad:** {f.get('Unidad','')} — {f.get('Comuna','')}, {f.get('Region','')}  \n"
        f"**Tipo:** {f.get('Tipo','')} · **Estado:** {f.get('Estado','')}  \n"
        f"**Publicación:** {f.get('FechaPublicacion','')} · **Cierre:** {f.get('FechaCierre','')}  \n"
        f"**Visita a terreno:** {visita}  \n"
        f"**Monto estimado:** {monto}  \n"
        f"**Servicio DATAELECT:** {f.get('ServicioDATAELECT','')}  \n"
        f"**Términos encontrados:** {f.get('PalabrasCalzadas','')}\n\n"
        f"### Descripción\n{(f.get('Descripcion') or '').strip()[:4000]}\n\n"
        f"[Ver en Mercado Público]({f.get('URL','')})"
    )


def main(ruta_csv):
    if not TOKEN:
        print("PROJECT_TOKEN no configurado: se omite la publicación en GitHub Projects.")
        return
    with open(ruta_csv, encoding="utf-8-sig") as fh:
        filas = list(csv.DictReader(fh))
    p = asegurar_campos(cargar_proyecto())
    print(f"Proyecto: {p['title']}")
    campos = [f for f in p["fields"]["nodes"] if f and f["name"].lower() in {n.lower() for n in CAMPOS}]
    status = next((f for f in p["fields"]["nodes"] if f and f["name"] == "Status"), None)
    op_status = None
    if status and status.get("options"):
        op_status = next((o for o in status["options"] if o["name"].lower() in
                          [s.lower() for s in STATUS_INICIAL]), status["options"][0])

    tarjetas = tarjetas_existentes(p["id"])
    eliminadas = eliminar_done(p, tarjetas, status)
    excluidas = codigos_done_registrados()
    ya = set(tarjetas)
    nuevas = [f for f in filas if f.get("Codigo") and f["Codigo"] not in ya
              and f["Codigo"] not in excluidas][:MAX_NUEVAS]
    print(f"{len(filas)} licitaciones en el archivo, {len(ya)} ya en el proyecto, "
          f"{len(excluidas)} ya cerradas (Done), {len(nuevas)} por agregar.")

    add = """mutation($pid:ID!,$t:String!,$b:String){
      addProjectV2DraftIssue(input:{projectId:$pid,title:$t,body:$b}){ projectItem{ id } } }"""
    upd = """mutation($pid:ID!,$item:ID!,$field:ID!,$v:ProjectV2FieldValue!){
      updateProjectV2ItemFieldValue(input:{projectId:$pid,itemId:$item,fieldId:$field,value:$v}){ clientMutationId } }"""
    errores, ok = [], 0
    for i, f in enumerate(nuevas, 1):
        titulo = f"[{f['Codigo']}] {f.get('Nombre','').strip()}"[:250]
        try:
            item = gql(add, pid=p["id"], t=titulo, b=cuerpo(f))["addProjectV2DraftIssue"]["projectItem"]["id"]
            for c in campos:
                v = valor_campo(c, f)
                if v:
                    try:
                        gql(upd, pid=p["id"], item=item, field=c["id"], v=v)
                    except RuntimeError as e:
                        errores.append(f"{f['Codigo']} campo {c['name']}: {e}")
            if op_status:
                gql(upd, pid=p["id"], item=item, field=status["id"], v={"singleSelectOptionId": op_status["id"]})
            ok += 1
        except RuntimeError as e:
            errores.append(f"{f['Codigo']}: {e}")
            if len(errores) >= 15 and ok == 0:
                break  # error sistemático: no insistir
        if i % 20 == 0:
            print(f"  {i}/{len(nuevas)}", flush=True)
        time.sleep(1.0)  # respeta el límite de GitHub (~80 creaciones por minuto)
    print(f"Listo: {ok} tarjetas agregadas al proyecto; {len(errores)} errores.")

    # Actualiza Cierre / Visita / Monto en las tarjetas que ya estaban (no toca Status ni otros campos).
    por_codigo = {f["Codigo"]: f for f in filas if f.get("Codigo")}
    actualizables = [c for c in campos if c["name"] in CAMPOS_ACTUALIZABLES]
    actualizadas = 0
    for codigo, t in tarjetas.items():
        f = por_codigo.get(codigo)
        if not f:
            continue
        cambio = False
        for c in actualizables:
            v = valor_campo(c, f)
            if not v:
                continue
            nuevo_valor = next(iter(v.values()))
            if str(t["valores"].get(c["name"]) or "") == str(nuevo_valor):
                continue
            try:
                gql(upd, pid=p["id"], item=t["id"], field=c["id"], v=v)
                cambio = True
            except RuntimeError as e:
                errores.append(f"{codigo} actualizar {c['name']}: {e}")
            time.sleep(0.7)
        actualizadas += cambio
    print(f"Tarjetas existentes actualizadas: {actualizadas}")
    escribir_log(p["title"], len(filas), len(ya), ok, errores, actualizadas, len(eliminadas))
    resumen = os.environ.get("GITHUB_STEP_SUMMARY")
    if resumen:
        with open(resumen, "a", encoding="utf-8") as fh:
            fh.write(f"### GitHub Project «{p['title']}»\n- Tarjetas nuevas: {ok}\n"
                     f"- Actualizadas: {actualizadas}\n- Eliminadas (Done): {len(eliminadas)}\n- Errores: {len(errores)}\n")
            for e in errores[:20]:
                fh.write(f"  - {e[:300]}\n")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("Uso: python scripts/proyecto.py <archivo.csv>")
    try:
        main(sys.argv[1])
    except Exception as e:  # deja el motivo en el repo para poder revisarlo
        escribir_log("(sin conectar)", 0, 0, 0, [f"Fallo general: {e}"])
        raise
