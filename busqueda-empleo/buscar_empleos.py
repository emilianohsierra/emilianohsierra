#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 buscar_empleos.py  —  Buscador de trabajo remoto (Asistente Virtual / Soporte)
================================================================================

Autor original: agente "Jolteon seeker" para Emiliano
Objetivo: encontrar vacantes REMOTAS de asistente virtual y servicio al cliente
          (ideal bilingüe español/inglés) abiertas para México / LATAM / mundo,
          desde varias bolsas de trabajo, y guardarlas en un CSV.

¿Por qué usa APIs y no "scraping" de LinkedIn/OCC/Indeed?
--------------------------------------------------------
Raspar el HTML de esos sitios: (1) viola sus Términos de Servicio, (2) te puede
bloquear la IP, y (3) se rompe cada vez que cambian el diseño. En cambio, estas
bolsas ofrecen APIs PÚBLICAS y gratuitas pensadas para consumirse por programa.
Es la forma correcta, estable y legal de hacerlo. Fuentes usadas:

    - Remotive         https://remotive.com/api/remote-jobs
    - RemoteOK         https://remoteok.com/api
    - Jobicy           https://jobicy.com/api/v2/remote-jobs
    - Arbeitnow        https://www.arbeitnow.com/api/job-board-api
    - We Work Remotely https://weworkremotely.com/categories/*.rss
    - Torre.ai (LATAM) https://search.torre.co/opportunities/_search

Cómo se usa
-----------
    1) Instala la única dependencia:   pip install requests
    2) Corre:                          python buscar_empleos.py
    3) Abre el archivo generado:       empleos_remotos.csv  (Excel lo abre)

Puedes ajustar la búsqueda en la sección CONFIGURACIÓN de abajo.
================================================================================
"""

import csv
import html
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

try:
    import requests
except ImportError:
    print("Falta la librería 'requests'. Instálala con:  pip install requests")
    sys.exit(1)


# =============================================================================
#  CONFIGURACIÓN  — cambia estas listas a tu gusto
# =============================================================================

# PERFILES de búsqueda, según el perfil real de Emiliano: asistente virtual y
# servicio al cliente bilingüe (español nativo + inglés intermedio). Cada
# vacante se etiqueta con el perfil al que coincide. Con que aparezca UNA
# palabra del perfil (en el título o resumen), la vacante pasa el filtro.
PERFILES = {
    "Asistente virtual": [
        "virtual assistant", "asistente virtual", "executive assistant",
        "administrative assistant", "admin assistant", "personal assistant",
        "office assistant", "back office", "data entry", "captura de datos",
        "scheduler", "appointment setter", "receptionist", "recepcionista",
        "dispatcher", "operations assistant", "remote assistant",
    ],
    "Servicio al cliente": [
        "customer service", "customer support", "customer success",
        "customer care", "customer experience", "client support",
        "client success", "support specialist", "support agent",
        "support representative", "help desk", "helpdesk", "service desk",
        "technical support", "soporte", "atención al cliente",
        "atencion al cliente", "call center", "contact center",
        "concierge", "onboarding specialist", "chat support",
        "content moderator", "content moderation", "community support",
        "asesor", "asesor bilingüe", "asesor telefónico", "ejecutivo telefónico",
        "representante", "agente bilingüe",
    ],
}

# Señales de que la vacante pide ESPAÑOL (además de inglés). Deben mencionar
# español de forma explícita — no basta "bilingual" solo (podría ser francés).
BILINGUE = [
    "spanish", "español", "espanol", "castellano", "hispanohablante",
    "bilingüe", "bilingue",            # palabras en español => contexto español
    "spanish-english", "english-spanish", "spanish/english", "english/spanish",
    "spanish and english", "english and spanish",
]

# Si es True, SOLO se muestran vacantes que pidan español + inglés (tu perfil).
# Cámbialo a False si algún día quieres ver también roles de soporte en inglés.
SOLO_BILINGUE = True

# Descarta vacantes más viejas que estos días (algunas bolsas dejan publicadas
# ofertas de hace años que ya están cerradas). Las de fecha desconocida SÍ pasan.
DIAS_MAX_ANTIGUEDAD = 365

# Regiones aceptadas. La vacante pasa si su ubicación menciona alguna de estas
# o si es abierta a todo el mundo. Así evitamos vacantes "solo EE.UU./Europa".
REGIONES_OK = [
    "mexico", "méxico", "latam", "latin america", "latinoamerica",
    "americas", "north america", "worldwide", "anywhere", "global",
    "remote", "remoto",
]

# Si el texto de la vacante menciona alguno de estos, la aceptamos AUNQUE la
# ubicación sea vaga: son señales de que contratan bilingües / desde LATAM.
MARCADORES_LATAM = [
    "spanish", "español", "bilingual", "bilingüe", "bilingue",
    "latam", "latin america", "latinoamerica", "mexico", "méxico",
]

# La vacante está ABIERTA (no restringida a un país) si dice algo de esto:
ABIERTO_A_LATAM = [
    "mexico", "méxico", "latam", "latin america", "latinoamerica",
    "worldwide", "anywhere", "global", "americas",
]

# Ubicaciones/países que RESTRINGEN a un lugar que no es México/LATAM.
# Si la ubicación menciona uno de estos (y NO dice también LATAM/México), se descarta.
RESTRICCION_PAIS_UBICACION = [
    "canada", "united states", "usa", "u.s.a", "u.s.", "united kingdom",
    "england", "ireland", "australia", "india", "philippines", "germany",
    "france", "spain", "portugal", "europe", "emea", "apac",
]

# Frases en la descripción que dejan claro que solo aplican residentes de otro país.
RESTRICCION_PAIS_TEXTO = [
    "canadian residents", "residents of canada", "canada residents",
    "u.s. residents", "us residents only", "united states residents",
    "residents of the united states", "uk residents", "residents of the uk",
    "must reside in the united states", "must reside in canada",
    "must be located in the united states", "must be located in canada",
    "must be based in the united states", "must be based in canada",
    "authorized to work in the united states", "authorized to work in the us",
    "must be authorized to work in the u", "us work authorization",
    "only to canadian residents", "only to us residents", "only to u.s. residents",
    "open only to canadian", "open only to us", "must live in the united states",
]

# Términos que DESCARTAN una vacante aunque tenga palabras clave (opcional).
# Incluye idiomas que Emiliano NO maneja (alemán, francés, etc.) para no ver ruido.
EXCLUIR = [
    "senior", "sr.", "lead", "staff", "principal", "manager", "head of",
    "10+ years", "10 años", "crypto", "trader", "trading",
    "deutsch", "german speaker", "french c1", "français", "c1/c2",
    "portuguese speaker", "portugués", "native german", "native french",
]

ARCHIVO_SALIDA = "empleos_remotos.csv"
ARCHIVO_VISTOS = "vistos.json"   # URLs ya vistas, para detectar vacantes NUEVAS
TIEMPO_ESPERA = 20               # segundos máximo por petición HTTP


# =============================================================================
#  UTILIDADES
# =============================================================================

def limpiar(texto):
    """Quita etiquetas HTML y espacios de sobra de un texto."""
    if not texto:
        return ""
    texto = re.sub(r"<[^>]+>", " ", str(texto))   # borra <p>, <br>, etc.
    texto = html.unescape(texto)                   # &amp; -> &
    return re.sub(r"\s+", " ", texto).strip()


def perfil_de(texto):
    """Devuelve el nombre del perfil que coincide, o None si ninguno.
    Si coincide con varios, gana el primero declarado en PERFILES."""
    t = texto.lower()
    for nombre, palabras in PERFILES.items():
        if any(k in t for k in palabras):
            return nombre
    return None


def region_ok(texto):
    t = (texto or "").lower()
    if not t:
        return True  # si no dicen región, la dejamos pasar (suelen ser globales)
    return any(r in t for r in REGIONES_OK)


def restringido_a_otro_pais(ubicacion, texto):
    """True si la vacante solo aplica para residentes de un país que no es
    México/LATAM (ej. 'open only to Canadian residents'). Estas se descartan
    aunque sean bilingües."""
    u = (ubicacion or "").lower()
    t = (texto or "").lower()
    abierto_ubic = any(a in u for a in ABIERTO_A_LATAM)
    # Ubicación amarrada a un país no-LATAM (y sin decir que también abre a LATAM)
    if not abierto_ubic and any(p in u for p in RESTRICCION_PAIS_UBICACION):
        return True
    # Frase explícita de "solo residentes de X" en la descripción
    if any(f in t for f in RESTRICCION_PAIS_TEXTO):
        if not any(a in t for a in ["latam", "latin america", "latinoamerica",
                                    "mexico", "méxico"]):
            return True
    return False


def excluida(texto):
    t = texto.lower()
    return any(x in t for x in EXCLUIR)


def cargar_vistos():
    """Lee el set de URLs vistas en corridas anteriores (vacío si no existe)."""
    try:
        with open(ARCHIVO_VISTOS, encoding="utf-8") as f:
            return set(json.load(f))
    except (FileNotFoundError, json.JSONDecodeError):
        return set()


def guardar_vistos(urls):
    """Guarda el set de URLs vistas para la próxima corrida."""
    try:
        with open(ARCHIVO_VISTOS, "w", encoding="utf-8") as f:
            json.dump(sorted(urls), f)
    except OSError as e:
        print(f"   ⚠  No se pudo guardar {ARCHIVO_VISTOS}: {e}")


def get_json(url, headers=None):
    """Descarga y devuelve JSON, o None si falla (sin tirar el programa)."""
    try:
        r = requests.get(url, headers=headers or {}, timeout=TIEMPO_ESPERA)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"   ⚠  No se pudo consultar {url.split('/')[2]}: {e}")
        return None


# =============================================================================
#  UN "FETCH" POR CADA FUENTE  — cada uno devuelve una lista de dicts normalizados
#  Formato normalizado: fuente, titulo, empresa, ubicacion, fecha, url, resumen
# =============================================================================

def fetch_remotive():
    print("→ Remotive...")
    # Combinamos la categoría software-dev con varias búsquedas para traer más
    # volumen relevante; después quitamos duplicados por id de la vacante.
    urls = [
        "https://remotive.com/api/remote-jobs?category=customer-support&limit=100",
        "https://remotive.com/api/remote-jobs?search=bilingual%20spanish",
        "https://remotive.com/api/remote-jobs?search=spanish%20speaking",
        "https://remotive.com/api/remote-jobs?search=spanish%20support",
    ]
    out, vistos = [], set()
    for url in urls:
        data = get_json(url)
        if not data:
            continue
        for j in data.get("jobs", []):
            if j.get("id") in vistos:
                continue
            vistos.add(j.get("id"))
            out.append({
                "fuente": "Remotive",
                "titulo": limpiar(j.get("title")),
                "empresa": limpiar(j.get("company_name")),
                "ubicacion": limpiar(j.get("candidate_required_location")),
                "fecha": (j.get("publication_date") or "")[:10],
                "url": j.get("url", ""),
                "resumen": limpiar(j.get("description"))[:300],
            })
    return out


def fetch_remoteok():
    print("→ RemoteOK...")
    # RemoteOK pide un User-Agent normal o rechaza la petición.
    data = get_json("https://remoteok.com/api",
                    headers={"User-Agent": "Mozilla/5.0 (buscador-empleos)"})
    if not data:
        return []
    out = []
    for j in data:
        if not isinstance(j, dict) or "position" not in j:
            continue  # el primer elemento es aviso legal, se salta
        out.append({
            "fuente": "RemoteOK",
            "titulo": limpiar(j.get("position")),
            "empresa": limpiar(j.get("company")),
            "ubicacion": limpiar(j.get("location")) or "Worldwide",
            "fecha": (j.get("date") or "")[:10],
            "url": j.get("url", ""),
            "resumen": limpiar(j.get("description"))[:300],
        })
    return out


def fetch_jobicy():
    print("→ Jobicy...")
    # El filtro por tag es muy estricto; pedimos varias etiquetas y también sin
    # filtro (industry dev) para ampliar, y deduplicamos por id.
    urls = [
        # Slugs válidos en Jobicy para tu perfil
        "https://jobicy.com/api/v2/remote-jobs?count=50&industry=supporting",
        "https://jobicy.com/api/v2/remote-jobs?count=50&industry=admin",
        "https://jobicy.com/api/v2/remote-jobs?count=50&tag=virtual-assistant",
    ]
    out, vistos = [], set()
    for url in urls:
        data = get_json(url)
        if not data:
            continue
        for j in data.get("jobs", []):
            jid = j.get("id") or j.get("url")
            if jid in vistos:
                continue
            vistos.add(jid)
            out.append({
                "fuente": "Jobicy",
                "titulo": limpiar(j.get("jobTitle")),
                "empresa": limpiar(j.get("companyName")),
                "ubicacion": limpiar(j.get("jobGeo")),
                "fecha": (j.get("pubDate") or "")[:10],
                "url": j.get("url", ""),
                "resumen": limpiar(j.get("jobExcerpt"))[:300],
            })
    return out


def fetch_arbeitnow():
    print("→ Arbeitnow...")
    data = get_json("https://www.arbeitnow.com/api/job-board-api")
    if not data:
        return []
    out = []
    for j in data.get("data", []):
        ts = j.get("created_at")
        fecha = ""
        if ts:
            try:
                fecha = datetime.fromtimestamp(int(ts), tz=timezone.utc).strftime("%Y-%m-%d")
            except Exception:
                fecha = ""
        ubic = "Remote" if j.get("remote") else limpiar(j.get("location"))
        out.append({
            "fuente": "Arbeitnow",
            "titulo": limpiar(j.get("title")),
            "empresa": limpiar(j.get("company_name")),
            "ubicacion": ubic,
            "fecha": fecha,
            "url": j.get("url", ""),
            "resumen": limpiar(j.get("description"))[:300],
        })
    return out


def _texto_rss(url):
    """Descarga un feed RSS como texto (o None si falla)."""
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0 (buscador-empleos)"},
                          timeout=TIEMPO_ESPERA)
        r.raise_for_status()
        return r.text
    except Exception as e:
        print(f"   ⚠  No se pudo consultar {url.split('/')[2]}: {e}")
        return None


def fetch_weworkremotely():
    print("→ We Work Remotely...")
    # Feeds RSS por categoría, enfocados a tu perfil: soporte al cliente y
    # "todos los demás" (que incluye asistente virtual, admin y captura de datos).
    feeds = [
        "https://weworkremotely.com/categories/remote-customer-support-jobs.rss",
        "https://weworkremotely.com/categories/all-other-remote-jobs.rss",
    ]
    out, vistos = [], set()
    for url in feeds:
        xml = _texto_rss(url)
        if not xml:
            continue
        try:
            raiz = ET.fromstring(xml)
        except ET.ParseError:
            continue
        for item in raiz.iter("item"):
            link = (item.findtext("link") or "").strip()
            if link in vistos:
                continue
            vistos.add(link)
            # El título viene como "Empresa: Puesto"; lo separamos.
            titulo_full = limpiar(item.findtext("title"))
            if ": " in titulo_full:
                empresa, titulo = titulo_full.split(": ", 1)
            else:
                empresa, titulo = "", titulo_full
            region = limpiar(item.findtext("region")) or limpiar(item.findtext("country"))
            fecha = ""
            pub = item.findtext("pubDate")
            if pub:
                # Formato tipo "Wed, 02 Jul 2026 10:00:00 +0000" -> tomamos la fecha
                m = re.search(r"(\d{1,2})\s+(\w{3})\s+(\d{4})", pub)
                meses = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05","Jun":"06",
                         "Jul":"07","Aug":"08","Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
                if m:
                    fecha = f"{m.group(3)}-{meses.get(m.group(2),'01')}-{int(m.group(1)):02d}"
            out.append({
                "fuente": "WeWorkRemotely",
                "titulo": titulo,
                "empresa": empresa,
                "ubicacion": region or "Remote",
                "fecha": fecha,
                "url": link,
                "resumen": limpiar(item.findtext("description"))[:300],
            })
    return out


def fetch_torre():
    print("→ Torre.ai (LATAM)...")
    # Torre es una plataforma LATAM con muchos roles bilingües/remotos. Su API
    # de búsqueda acepta POST con el rol buscado; consultamos varios y dedupe.
    roles = ["bilingual spanish customer service", "spanish english customer service",
             "bilingual customer service", "asesor bilingüe", "atención al cliente",
             "bilingual virtual assistant", "spanish english support",
             "customer service", "virtual assistant", "call center"]
    out, vistos = [], set()
    for rol in roles:
        cuerpo = {
            "and": [
                {"skill/role": {"text": rol, "experience": "potential-to-develop"}},
                {"remote": {"term": True}},
            ],
            "offset": 0, "size": 20,
        }
        try:
            r = requests.post("https://search.torre.co/opportunities/_search",
                              json=cuerpo, timeout=TIEMPO_ESPERA,
                              headers={"User-Agent": "Mozilla/5.0 (buscador-empleos)"})
            r.raise_for_status()
            resultados = r.json().get("results", [])
        except Exception as e:
            print(f"   ⚠  No se pudo consultar torre.co: {e}")
            continue
        for j in resultados:
            jid = j.get("id")
            if not jid or jid in vistos:
                continue
            vistos.add(jid)
            orgs = j.get("organizations") or [{}]
            # Salario, si Torre lo hace visible
            salario = ""
            comp = (j.get("compensation") or {})
            if comp.get("visible") and comp.get("data"):
                d = comp["data"]
                lo = d.get("minAmount")
                hi = d.get("maxAmount") or lo
                if lo:
                    salario = (f" | {d.get('currency','USD')} "
                               f"{int(lo)}-{int(hi)}/{d.get('periodicity','month')}")
            out.append({
                "fuente": "Torre",
                "titulo": limpiar(j.get("objective")),
                "empresa": limpiar(orgs[0].get("name")),
                "ubicacion": "Remote (LATAM-friendly)",
                "fecha": (j.get("created") or "")[:10],
                "url": f"https://torre.ai/post/{jid}",
                "resumen": (limpiar(j.get("tagline")) + salario)[:300],
            })
    return out


FUENTES = [fetch_remotive, fetch_remoteok, fetch_jobicy,
           fetch_arbeitnow, fetch_weworkremotely, fetch_torre]


# =============================================================================
#  PROGRAMA PRINCIPAL
# =============================================================================

def main():
    print("=" * 70)
    print(" Vacantes REMOTAS para Emiliano — Soporte/CS (hoy) + Web/Dev (futuro)")
    print("=" * 70)

    # 1) Recolectar de todas las fuentes
    crudas = []
    for fetch in FUENTES:
        try:
            crudas.extend(fetch())
        except Exception as e:
            print(f"   ⚠  Error en fuente: {e}")

    print(f"\nTotal descargadas (sin filtrar): {len(crudas)}")

    # 1.5) Cargar URLs vistas en corridas anteriores (para detectar NUEVAS)
    vistas_previas = cargar_vistos()

    # Fecha de corte por antigüedad (formato AAAA-MM-DD para comparar como texto)
    from datetime import timedelta
    fecha_corte = (datetime.now(timezone.utc) - timedelta(days=DIAS_MAX_ANTIGUEDAD)).strftime("%Y-%m-%d")

    # 2) Filtrar por perfil + región - exclusiones, etiquetar y quitar duplicados
    claves_vistas = set()
    resultados = []
    for v in crudas:
        texto = f"{v['titulo']} {v['resumen']}"
        perfil = perfil_de(texto)
        if perfil is None:
            continue
        if excluida(texto):
            continue
        # RECHAZO DURO: restringida a otro país (ej. "solo residentes de Canadá"),
        # aunque sea bilingüe. Esto evita mostrarte vacantes donde no puedes aplicar.
        if restringido_a_otro_pais(v["ubicacion"], texto):
            continue
        # Pasa si la ubicación es elegible, O si el texto marca bilingüe/LATAM
        marca_latam = any(m in texto.lower() for m in MARCADORES_LATAM)
        if not region_ok(v["ubicacion"]) and not marca_latam:
            continue
        es_bilingue = any(b in texto.lower() for b in BILINGUE)
        # Filtro OBLIGATORIO: solo vacantes que pidan español + inglés.
        if SOLO_BILINGUE and not es_bilingue:
            continue
        # Descarta vacantes demasiado viejas (probablemente ya cerradas).
        if v.get("fecha") and v["fecha"] < fecha_corte:
            continue
        # Dedupe por URL y también por título+empresa (algunas bolsas repiten la
        # misma vacante con distinta URL/ID).
        clave = v["url"] or f"{v['titulo']}|{v['empresa']}"
        clave_texto = f"{v['titulo'].lower().strip()}|{v['empresa'].lower().strip()}"
        if clave in claves_vistas or clave_texto in claves_vistas:
            continue
        claves_vistas.add(clave)
        claves_vistas.add(clave_texto)
        v["perfil"] = perfil
        v["nueva"] = "SI" if clave not in vistas_previas else ""
        v["bilingue"] = "SI" if es_bilingue else ""
        resultados.append(v)

    # 3) Ordenar (sort estable, se aplica de menos a más importante):
    #    fecha desc  ->  bilingües primero (tu fortaleza)  ->  NUEVAS primero.
    resultados.sort(key=lambda v: v["fecha"], reverse=True)
    resultados.sort(key=lambda v: v["bilingue"] == "SI", reverse=True)
    resultados.sort(key=lambda v: v["nueva"] == "SI", reverse=True)

    # 4) Mostrar en pantalla, agrupado por perfil.
    #    ★ = vacante nueva desde la última corrida   🗣 = pide español/bilingüe
    for perfil in PERFILES:
        grupo = [v for v in resultados if v["perfil"] == perfil]
        nuevas_g = sum(1 for v in grupo if v["nueva"] == "SI")
        print(f"\n───── {perfil}  ({len(grupo)} vacantes, {nuevas_g} nuevas) " + "─" * 12)
        for v in grupo[:15]:
            marcas = ("★ NUEVA " if v["nueva"] == "SI" else "") + ("🗣 " if v["bilingue"] == "SI" else "")
            print(f"[{v['fecha'] or '  ?  '}] {marcas}{v['titulo']}")
            print(f"        {v['empresa']}  ·  {v['ubicacion']}  ·  ({v['fuente']})")
            print(f"        {v['url']}")

    # 5) Guardar TODO en CSV (columnas para filtrar en Excel)
    campos = ["nueva", "bilingue", "perfil", "fecha", "titulo", "empresa",
              "ubicacion", "fuente", "url", "resumen"]
    with open(ARCHIVO_SALIDA, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        for v in resultados:
            w.writerow({k: v.get(k, "") for k in campos})

    # 6) Actualizar el registro de URLs vistas y escribir un resumen para el cron
    todas_urls = vistas_previas | {v["url"] for v in resultados if v["url"]}
    guardar_vistos(todas_urls)

    nuevas = sum(1 for v in resultados if v["nueva"] == "SI")
    bilingues = sum(1 for v in resultados if v["bilingue"] == "SI")
    por_perfil = {p: sum(1 for v in resultados if v["perfil"] == p) for p in PERFILES}

    with open("resumen.json", "w", encoding="utf-8") as f:
        json.dump({"total": len(resultados), "nuevas": nuevas,
                   "bilingues": bilingues, "por_perfil": por_perfil},
                  f, ensure_ascii=False)

    detalle = "   ".join(f"· {p}: {n}" for p, n in por_perfil.items())
    filtro = "SOLO bilingües español-inglés" if SOLO_BILINGUE else "todas (bilingües marcadas 🗣)"
    print("\n" + "=" * 70)
    print(f" ✔  Guardadas {len(resultados)} vacantes en:  {ARCHIVO_SALIDA}   [{filtro}]")
    print(f"     {detalle}")
    print(f"     · NUEVAS desde la última corrida: {nuevas}")
    print(f"     Ábrelo con Excel/Google Sheets (columnas 'perfil', 'bilingue', 'nueva').")
    print("=" * 70)


if __name__ == "__main__":
    main()
