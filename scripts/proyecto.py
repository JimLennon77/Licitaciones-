"""
Publica las licitaciones del día como tarjetas en el GitHub Project del usuario.

Uso:  python scripts/proyecto.py datos/diario/licitaciones_AAAA-MM-DD.csv

Variables de entorno:
  PROJECT_TOKEN   token personal (classic) con permiso "project"
  PROJECT_OWNER   usuario dueño del proyecto (por defecto JimLennon77)
  PROJECT_NUMBER  número del proyecto (por defecto 1)

Qué hace:
  - Crea (si faltan) los campos: Código, Organismo, Región, Servicio DATAELECT, Cierre,
    Monto estimado y URL. Si el proyecto ya tiene un campo con ese nombre, lo reutiliza.
  - Agrega cada licitación nueva como borrador "[Código] Nombre", con el detalle en el cuerpo.
  - No duplica: si ya existe una tarjeta con ese código, la salta.
  - Deja el Status en la primera opción (p. ej. "Todo" / "Nueva"), salvo que exista una
    opción cuyo nombre esté en STATUS_INICIAL.
"""
import csv
import os
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

CAMPOS = {  # nombre en el proyecto -> (tipo, columna del CSV)
    "Código": ("TEXT", "Codigo"),
    "Organismo": ("TEXT", "Organismo"),
    "Región": ("TEXT", "Region"),
    "Servicio DATAELECT": ("SINGLE_SELECT", "ServicioDATAELECT"),
    "Cierre": ("DATE", "FechaCierre"),
    "Monto estimado": ("NUMBER", "MontoEstimado"),
    "URL": ("TEXT", "URL"),
}
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
        if nombre.lower() in existentes:
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


def codigos_existentes(pid):
    q = """query($pid:ID!,$after:String){ node(id:$pid){ ... on ProjectV2{
      items(first:100, after:$after){ pageInfo{ hasNextPage endCursor }
        nodes{ content{ ... on DraftIssue{ title } ... on Issue{ title } } } } } } }"""
    codigos, after = set(), None
    while True:
        items = gql(q, pid=pid, after=after)["node"]["items"]
        for n in items["nodes"]:
            m = re.match(r"\s*\[([^\]]+)\]", (n.get("content") or {}).get("title") or "")
            if m:
                codigos.add(m.group(1).strip())
        if not items["pageInfo"]["hasNextPage"]:
            return codigos
        after = items["pageInfo"]["endCursor"]


def valor_campo(campo, fila):
    tipo = campo["dataType"]
    col = next((c for n, (_, c) in CAMPOS.items() if n.lower() == campo["name"].lower()), None)
    bruto = str(fila.get(col, "") or "").strip()
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


def escribir_log(titulo, total, previas, ok, errores):
    """Deja constancia en el repo (datos/proyecto_log.md) de cada publicación en el tablero."""
    ruta = os.path.join(os.path.dirname(__file__), "..", "datos", "proyecto_log.md")
    nuevo = not os.path.exists(ruta)
    with open(ruta, "a", encoding="utf-8") as fh:
        if nuevo:
            fh.write("# Publicaciones en GitHub Project\n\n")
        fh.write(f"## {time.strftime('%Y-%m-%d %H:%M UTC', time.gmtime())} · {titulo}\n"
                 f"- En archivo: {total} · ya en tablero: {previas} · agregadas: {ok} · errores: {len(errores)}\n")
        for e in errores[:30]:
            fh.write(f"  - {e[:400]}\n")
        fh.write("\n")


def cuerpo(f):
    monto = f.get("MontoEstimado") or "no informado"
    return (
        f"**Organismo:** {f.get('Organismo','')}  \n"
        f"**Unidad:** {f.get('Unidad','')} — {f.get('Comuna','')}, {f.get('Region','')}  \n"
        f"**Tipo:** {f.get('Tipo','')} · **Estado:** {f.get('Estado','')}  \n"
        f"**Publicación:** {f.get('FechaPublicacion','')} · **Cierre:** {f.get('FechaCierre','')}  \n"
        f"**Monto estimado:** {monto} {f.get('Moneda','')}  \n"
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

    ya = codigos_existentes(p["id"])
    nuevas = [f for f in filas if f.get("Codigo") and f["Codigo"] not in ya][:MAX_NUEVAS]
    print(f"{len(filas)} licitaciones en el archivo, {len(ya)} ya en el proyecto, {len(nuevas)} por agregar.")

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
    nuevas = nuevas[:ok] if ok else []
    print(f"Listo: {ok} tarjetas agregadas al proyecto; {len(errores)} errores.")
    escribir_log(p["title"], len(filas), len(ya), ok, errores)
    resumen = os.environ.get("GITHUB_STEP_SUMMARY")
    if resumen:
        with open(resumen, "a", encoding="utf-8") as fh:
            fh.write(f"### GitHub Project «{p['title']}»\n- Tarjetas nuevas: {ok}\n- Errores: {len(errores)}\n")
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
