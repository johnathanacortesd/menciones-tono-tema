# -*- coding: utf-8 -*-
"""
Analizador de menciones: Tono + Tema + Sub-tema
===============================================

Sube un XLSX de monitoreo de medios (GlobalNews u otro export), indica la marca/entidad,
sus voceros y sus alias, y descarga el mismo XLSX con tres columnas nuevas:
Tono (Positivo/Neutro/Negativo), Tema (cubo que agrupa sub-temas parecidos) y
Sub-tema (3 a 7 palabras).

POR QUE ESTA APP ACIERTA MEJOR QUE UN PROMPT SUELTO
No es el prompt: son cinco piezas deterministas alrededor del LLM.

1. AGRUPACION. Las notas iguales o casi iguales (el mismo comunicado en 20 medios) se
   agrupan ANTES de etiquetar. La etiqueta se decide por grupo, no por fila, así que dos
   notas iguales nunca pueden salir con tono distinto. Esto es lo que rompe la mayoria de
   los intentos con API: se etiqueta fila por fila y el modelo dice "Neutro" unas veces y
   "Positivo" otras para el mismo texto.
2. EL SUB-TEMA MANDA. Primero se escribe el sub-tema (una frase corta que resume el hecho)
   y solo despues el tono, leyendo ese resumen junto con el titular. Igual que un analista:
   resumir y luego juzgar. Ademas, los sub-temas ya asignados se pasan en cada lote como
   CANDIDATOS para que los casos repetidos reutilicen EXACTAMENTE el mismo texto.
3. VALIDADOR CON REGLAS DURAS + REPARACION. Cada etiqueta se valida (3-7 palabras, sin
   verbo conjugado al inicio, sin terminar en preposicion, sin `:` `;` `|` ni comillas,
   sin rotulos genericos como "noticias generales", tono dentro del vocabulario) y los
   fallos se devuelven al modelo para que los corrija. Sin este paso quedan etiquetas
   de 9 palabras y rangos de titulo: es el error mas comun al usar un LLM suelto.
4. EL TEMA NO LO INVENTA EL LLM. Los Temas son una lista cerrada de cubos por cliente
   (gobierno territorial, gremio/sector) y se asignan por reglas lexicas, mirando primero
   el sub-tema y solo despues el titulo (y solo si el titulo es corto de verdad). El LLM
   solo interviene cuando ninguna regla coincide, y entonces elige DENTRO de la lista (o
   propone un cubo nuevo especifico si tu lo autorizas). Nunca puede escribir "Otros":
   la app bloquea la descarga mientras haya un grupo sin cubo.
5. DOS PASADAS CONTRA EL RUIDO. Si la columna de titulo en realidad trae el cuerpo
   completo, palabras incidentales ("contratos", "contratiempos") disparan cubos que no
   corresponden. Por eso el sub-tema va primero y el titulo solo entra si mide <=160
   caracteres.

LIMITES HONESTOS
- El tono es juicio: la app acierta el criterio en la mayoria de los casos, pero casos
  limite (una critica dicha con ironia, una obra anunciada y nunca ejecutada) pueden
  diferir de una lectura humana. La app marca esos casos y te deja reasignarlos antes
  de descargar.
- Revisa siempre la hoja "Resumen" y la tabla de auditoria del final.
- Un LLM con temperatura 0 es casi determinista, pero no identico entre corridas: guarda
  el XLSX que entregas como version final del periodo.

Requisitos: streamlit, openpyxl, pandas, requests, rapidfuzz, numpy.
"""

import collections
from collections.abc import Mapping
import hashlib
import html
import hmac
import io
import json
import re
import time
import unicodedata

import numpy as np
import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

st.set_page_config(page_title='Tono, Tema y Sub-tema de menciones', page_icon='📰', layout='wide')


# ============================================================================
# 1. UTILIDADES DE TEXTO
# ============================================================================
def ctrl(s):
    """Quita caracteres de control: rompen el XLSX de salida."""
    return ''.join(ch for ch in str(s or '') if (ch >= ' ' or ch in '\n\t') and ch not in '\ufffe\uffff')


def nz(s):
    s = unicodedata.normalize('NFKD', ctrl(s))
    s = ''.join(c for c in s if not unicodedata.combining(c)).lower()
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9ñ ]+', ' ', s)).strip()


def words(s):
    s = unicodedata.normalize('NFKD', ctrl(s))
    s = ''.join(c for c in s if not unicodedata.combining(c)).lower()
    return re.findall(r'[a-z0-9ñ]+', s)


def grams(w, k):
    return set(tuple(w[i:i + k]) for i in range(max(0, len(w) - k + 1)))


def raiz(t):
    t = nz(t)
    if len(t) > 4 and t.endswith('es'):
        return t[:-2]
    if len(t) > 3 and t.endswith('s'):
        return t[:-1]
    return t


def sq(s):
    return re.sub(r'\s+', ' ', ctrl(s)).strip()


# ============================================================================
# 2. AGRUPACION DE NOTAS IGUALES O SIMILARES  (pieza 1)
# ============================================================================
GENERIC = set("""a al algo algunos ante antes aqui asi aun aunque bien cada como con contra cual cuando
de del desde donde dos el ella ellas ellos en entre era eran es esa esas ese eso esos esta estaba
estan este esto estos fue fueron ha hace hacia hasta hay la las le le los lo los mas me mi mis
mucho muy nos o os otra otras otro otros para pero poco por porque que quien se sea segun ser si
sin sobre son su sus tal tambien tan tanto te tiene tienen todo todos tu tus un una uno unos y ya""".split())
K_BODY = 5
UMBRAL_TITULO_POR_DEFECTO = 92      # % de similitud de titulos
UMBRAL_CUERPO_POR_DEFECTO = 85      # % de 5-gramas en comun del texto mas corto
MIN_GRAMAS = 30
MIN_PALABRAS_TITULO = 3


def agrupar(filas, umbral_titulo=UMBRAL_TITULO_POR_DEFECTO, umbral_cuerpo=UMBRAL_CUERPO_POR_DEFECTO):
    """Union-find sobre similitud de titulo (palabras de contenido) y de cuerpo (5-gramas).

    umbral_titulo bajo (p. ej. 85) fusiona mas agresivamente campanas repetidas; 92 es el
    valor con el que se validaron los casos reales.
    """
    from rapidfuzz import fuzz, process
    for f in filas:
        f['ctit'] = set(w for w in words(f['titulo']) if w not in GENERIC)
        f['g5'] = grams(words(f['texto']), K_BODY)
    par = list(range(len(filas)))

    def find(x):
        while par[x] != x:
            par[x] = par[par[x]]
            x = par[x]
        return x

    def uni(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            par[max(ra, rb)] = min(ra, rb)

    tit = [' '.join(sorted(f['ctit'])) for f in filas]
    if len(filas) > 1:
        t1 = process.cdist(tit, tit, scorer=fuzz.ratio, workers=-1) / 100.0
        t2 = process.cdist(tit, tit, scorer=fuzz.token_sort_ratio, workers=-1) / 100.0
        T = np.maximum(t1, t2)
        for i in range(len(filas)):
            for j in np.where(T[i] >= umbral_titulo / 100.0)[0]:
                if j > i and len(filas[i]['ctit'] & filas[j]['ctit']) >= MIN_PALABRAS_TITULO:
                    uni(i, j)
    inv = collections.defaultdict(set)
    for j, f in enumerate(filas):
        for g in f['g5']:
            inv[g].add(j)
    for i, f in enumerate(filas):
        if len(f['g5']) < MIN_GRAMAS:
            continue
        hits = collections.Counter()
        for g in f['g5']:
            for j in inv.get(g, ()):
                if j != i:
                    hits[j] += 1
        for j, inter in hits.items():
            if j > i and min(len(f['g5']), len(filas[j]['g5'])) >= MIN_GRAMAS and \
                    inter / min(len(f['g5']), len(filas[j]['g5'])) >= umbral_cuerpo / 100.0:
                uni(i, j)
    gr = collections.defaultdict(list)
    for i in range(len(filas)):
        gr[find(i)].append(i)
    return sorted(gr.values(), key=lambda v: -len(v))


def construir_grupos(filas, umbral_titulo, umbral_cuerpo):
    idxs = agrupar(filas, umbral_titulo, umbral_cuerpo)
    grupos, mapa = [], {}
    for k, ix in enumerate(idxs, 1):
        cnt = collections.Counter(filas[i]['titulo'] for i in ix)
        rep = cnt.most_common(1)[0][0] or sq(filas[ix[0]]['texto'])[:120]
        cuerpo = max((filas[i]['texto'] for i in ix), key=len)
        alts = [t for t in sorted(set(filas[i]['titulo'] for i in ix)) if t and t != rep][:3]
        grupos.append({'grupo': k, 'n': len(ix), 'titulo': rep, 'titulos_alt': alts,
                       'filas': [filas[i]['_fila_n'] for i in ix], 'ids': [filas[i]['id'] for i in ix],
                       'texto': sq(cuerpo)[:700]})
        for i in ix:
            mapa[filas[i]['id']] = k
    return grupos, mapa


# ============================================================================
# 3. VALIDADOR DE ETIQUETAS  (pieza 3)
# ============================================================================
PREP_FIN = {'de', 'del', 'la', 'el', 'los', 'las', 'un', 'una', 'unos', 'unas', 'para', 'en', 'con',
            'por', 'y', 'o', 'a', 'al', 'que', 'su', 'sus', 'sin', 'sobre', 'entre', 'tras', 'ni'}
CONECT = PREP_FIN | {'se', 'lo', 'le', 'les', 'es', 'son', 'como', 'mas', 'más', 'muy'}
VERBOS1 = set("""entregan anuncian avanza avanzan instalan reconocen firman inauguran denuncian alertan
piden exigen rechazan critican senalan señalan aseguran confirman advierten solicitan denunciaron
anunciaron instalaron avanzaron reconocieron inauguraron entregaron logran obtienen reciben presentan
lideran realizan adelantan ejecutan mejoran aumentan disminuyen reducen destinan aprueban celebran
conmemoran rinden posesionan nombran designan ratifican reafirman garantizan benefician participan
visitan recorren supervisan verifican socializan capacitan forman graduan certifican arranca culmina
inicia termina sigue siguen mantiene mantienen trabaja trabajan llevan deja dejan abre abren obtiene
recibe recibio recibieron pidieron exigieron superviso entrego anuncio advirtio aseguro confirmo
denuncio critico lidero presento aprobo celebro conmemoro designo nombro""".split())
RE_VERBO = re.compile(r'(aron|ieron|ió|eó|arán|erán|irán|aba|aban|amos|imos)$')
ROTULO_GEN = {
    'gestion gubernamental', 'gestion institucional', 'gestion departamental', 'actividad institucional',
    'actividad gubernamental', 'noticias de la entidad', 'noticias generales', 'temas generales',
    'cobertura informativa', 'panorama regional', 'informacion general', 'actualidad departamental',
    'eventos institucionales', 'otros', 'varios', 'general', 'miscelaneo', 'miscelanea',
}
MARCO = {'noticias', 'informacion', 'información', 'cobertura', 'actualidad', 'eventos', 'actividades',
         'destacados', 'panorama', 'menciones', 'informe'}
FILLER = MARCO | {'importantes', 'relevantes', 'generales', 'varias', 'diversos', 'diversas',
                  'departamental', 'institucional', 'gubernamental', 'regional', 'recientes', 'varios'}
MIN_PAL, MAX_PAL = 3, 7
TONOS = ('Positivo', 'Neutro', 'Negativo')


def validar(sub_tema, tono, fuentes, min_pal=MIN_PAL, max_pal=MAX_PAL):
    """Devuelve la lista de problemas. Vacio = etiqueta valida."""
    p = []
    sub_tema = str(sub_tema or '').strip()
    if not sub_tema:
        return ['vacio']
    if re.search(r'[:;|"\'«»—–]', sub_tema):
        p.append('caracter_marcador')
    w = sub_tema.split()
    if len(w) < min_pal:
        p.append('corto(%d)' % len(w))
    if len(w) > max_pal:
        p.append('largo(%d)' % len(w))
    if nz(w[-1]) in PREP_FIN:
        p.append('termina_preposicion')
    prim = nz(w[0])
    nexo2 = len(w) > 1 and nz(w[1]) in PREP_FIN
    if not nexo2 and (prim in VERBOS1 or (len(prim) > 4 and RE_VERBO.search(prim))):
        p.append('verbo_inicial(%s)' % prim)
    if nz(sub_tema) in ROTULO_GEN:
        p.append('rotulo_generico')
    toks = [nz(t) for t in w]
    if toks and toks[0] in MARCO and all(t in FILLER for t in toks[1:]):
        p.append('marco_vacio')
    if tono not in TONOS:
        p.append('tono_invalido(%s)' % tono)
    fuente = ' '.join(nz(f) for f in (fuentes or []))
    ft = set(re.findall(r'[a-z0-9ñ]+', fuente))
    fr = {raiz(t) for t in ft}
    faltan = [t for t in toks if len(t) >= 4 and t not in CONECT and t not in ft and raiz(t) not in fr]
    if len(faltan) > 1:
        p.append('revisar_anclaje(%s)' % ','.join(faltan[:4]))
    return p


# ============================================================================
# 4. TAXONOMIA DE TEMAS  (pieza 4)
# ============================================================================
TAX_GOBIERNO = {
 "nota": "Cubos para gobernaciones, alcaldias y entidades publicas territoriales. El orden es la prioridad: lo especifico antes que lo generico.",
 "temas": [
  "Educación superior y universidad",
  "Deporte y Juegos Nacionales",
  "Electoral y político",
  "Política nacional e internacional",
  "Control, justicia y contratación",
  "Reconocimientos y liderazgo",
  "Gestión del riesgo, lluvias e inundaciones",
  "Seguridad y convivencia",
  "Salud y red hospitalaria",
  "Educación y primera infancia",
  "Vías, obra pública e infraestructura",
  "Agua, saneamiento y servicios públicos",
  "Energía, gas y transición energética",
  "Ayudas sociales y atención a comunidades",
  "Agro y desarrollo económico",
  "Turismo, cultura y patrimonio",
  "Juventud, género y participación",
  "Vivienda y hábitat",
  "Cooperación, paz y derechos humanos",
  "Gestión institucional y comunitaria",
  "Otros"
 ],
 "reglas": [
  {
   "tema": "Control, justicia y contratación",
   "claves": [
    "contraloria*",
    "procuraduria*",
    "fiscalia*",
    "disciplinari*",
    "sancion*",
    "irregular*",
    "hallazgo*",
    "corrupcion",
    "investigacion*",
    "captura",
    "carcel",
    "judicial",
    "accion popular",
    "veeduria",
    "responsabilidad fiscal",
    "contrat*",
    "licitacion*",
    "interventoria",
    "demanda",
    "proceso fiscal",
    "juzgado",
    "fiscal",
    "clientelismo",
    "fotomulta*"
   ]
  },
  {
   "tema": "Educación superior y universidad",
   "claves": [
    "universidad*",
    "unisucre",
    "consejo superior",
    "docente*",
    "rector*",
    "nomina*",
    "pregrado",
    "alma mater",
    "carrera profesional",
    "estudiantes universitarios",
    "profesor*",
    "docencia"
   ]
  },
  {
   "tema": "Deporte y Juegos Nacionales",
   "claves": [
    "juegos nacionales",
    "juegos deportivos",
    "villa olimpic*",
    "deportiv*",
    "liga*",
    "federacion deportiva",
    "softbol",
    "futbol",
    "balon",
    "deporte",
    "torneo*",
    "cancha*",
    "partido de",
    "escuadron femenino",
    "atleta*",
    "ciclismo",
    "patinaje",
    "olimpic*"
   ]
  },
  {
   "tema": "Política nacional e internacional",
   "claves": [
    "venezuela",
    "maduro",
    "cancilleria",
    "la guajira",
    "frontera",
    "cepeda",
    "barreras",
    "vicepresidencial*",
    "petro",
    "gobierno nacional",
    "vargas lleras",
    "presidencial*",
    "candidato presidencial",
    "cumbre de gobernadores",
    "tribunal de garantias",
    "uribe",
    "gustavo petro",
    "nacional e internacional"
   ]
  },
  {
   "tema": "Electoral y político",
   "claves": [
    "eleccion*",
    "electoral",
    "candidat*",
    "junta de accion",
    "jac",
    "registraduria",
    "voto",
    "votacion",
    "congreso",
    "senado",
    "senador*",
    "campana",
    "coalicion*",
    "partido",
    "aspirante*",
    "encuesta*",
    "parlamentaria",
    "bancada",
    "favorabilidad",
    "estragatega",
    "campanas politicas",
    "paloma valencia",
    "aranas",
    "aran a",
    "arana"
   ]
  },
  {
   "tema": "Reconocimientos y liderazgo",
   "claves": [
    "reconocimiento*",
    "forbes",
    "mujeres poderosas",
    "premio*",
    "galardon*",
    "homenaje*",
    "distincion*",
    "liderazgo",
    "aprobacion del gobernador",
    "mejor gobernador",
    "ranking",
    "perfil de liderazgo",
    "destacad*",
    "exaltad*",
    "proyeccion"
   ]
  },
  {
   "tema": "Gestión del riesgo, lluvias e inundaciones",
   "claves": [
    "inundacion*",
    "crecient*",
    "alerta roja",
    "alerta amarilla",
    "frente frio",
    "lluvia*",
    "ola invernal",
    "emergencia*",
    "damnificad*",
    "desbordamiento",
    "rio cauca",
    "san jorge",
    "mojana",
    "calamidad",
    "aguacero",
    "vendaval",
    "oleaje",
    "afectacion*",
    "caregato",
    "cara de gato",
    "boquete",
    "riada",
    "temporada de lluvias",
    "afectad*",
    "desastre",
    "ungrd",
    "evacuacion*",
    "condiciones adversas",
    "invierno",
    "clima*"
   ]
  },
  {
   "tema": "Seguridad y convivencia",
   "claves": [
    "seguridad(?! alimentaria)",
    "homicidio*",
    "asesinato*",
    "sicari*",
    "abigeato",
    "contrabando",
    "microtrafico",
    "policia*",
    "escuadron*",
    "patrullaje*",
    "convivencia",
    "armado",
    "criminal*",
    "delincuencia",
    "hurto*",
    "extorsion",
    "reten",
    "fuerza publica",
    "multicrimen",
    "clan del golfo",
    "dispositivo de seguridad",
    "incautacion*",
    "incautan",
    "operativo*",
    "violencia",
    "armada",
    "infanteria",
    "militar*",
    "ejercito",
    "soldado*",
    "antidron*",
    "armas",
    "anticonrabando",
    "anticontrabando*",
    "banda*",
    "generacion delincuencial"
   ]
  },
  {
   "tema": "Salud y red hospitalaria",
   "claves": [
    "hospital*",
    "ese(?= (de|centro|hospital|caimito|san|municipal))",
    "minsalud",
    "ministerio de salud",
    "fusi*",
    "salud",
    "acoso*",
    "enfermer*",
    "clinica*",
    "eps",
    "pai",
    "vacuna*",
    "dengue",
    "cancer*",
    "leishmaniasis",
    "red hospitalaria",
    "deuda hospitalaria",
    "ambulancia*",
    "hemocentro",
    "consultorio*",
    "cirugia",
    "nutricional",
    "salud mental",
    "droga*",
    "alimentos",
    "bebidas",
    "higiene",
    "hus",
    "hospitalario"
   ]
  },
  {
   "tema": "Educación y primera infancia",
   "claves": [
    "educacion*",
    "educativ*",
    "escolar*",
    "pae",
    "colegio*",
    "institucion educativa",
    "estudiante*",
    "nino*",
    "nina*",
    "primera infancia",
    "aula*",
    "matricula*",
    "escuela*",
    "anio escolar",
    "clases",
    "becas",
    "icfes",
    "rectoria",
    "jardin infantil",
    "infancia",
    "pedagogic*",
    "oferta educativa",
    "ciencia*",
    "tecnologia",
    "entornos digitales"
   ]
  },
  {
   "tema": "Vías, obra pública e infraestructura",
   "claves": [
    "via(?! al)",
    "vial*",
    "pavimentacion*",
    "obra*",
    "puente*",
    "calle*",
    "malecón",
    "malecon*",
    "anden*",
    "dique",
    "variante",
    "terraplen*",
    "acueducto*",
    "arroyo",
    "megaproyecto",
    "canalizacion*",
    "infraestructura",
    "corredor*",
    "tramo*",
    "bacheo",
    "intervencion*",
    "terminacion*",
    "estadio",
    "parque*",
    "peaje",
    "viaducto",
    "placa huella",
    "alumbrado",
    "hidroelectrica",
    "aeropuerto",
    "terminal de transporte",
    "hospital en construccion"
   ]
  },
  {
   "tema": "Agua, saneamiento y servicios públicos",
   "claves": [
    "agua potable",
    "aguas de sucre",
    "saneamiento",
    "alcantarillado*",
    "servicios publicos",
    "afinia",
    "energia electrica",
    "residuos",
    "aseo",
    "recurso hidrico",
    "cobertura total",
    "aseguramiento"
   ]
  },
  {
   "tema": "Energía, gas y transición energética",
   "claves": [
    "gas natural",
    "gas domiciliario",
    "energia solar",
    "fotovoltaic*",
    "transicion energetica",
    "biocombustible*",
    "bioenergetic*",
    "panel solar",
    "electrificacion*",
    "potencia energetica",
    "mineria",
    "hidrocarburo*"
   ]
  },
  {
   "tema": "Vivienda y hábitat",
   "claves": [
    "vivienda*",
    "habitacional*",
    "predio*",
    "titulacion*",
    "legalizacion*",
    "mejoramiento de vivienda",
    "techo*",
    "habitantes del corregimiento",
    "urbanismo"
   ]
  },
  {
   "tema": "Agro y desarrollo económico",
   "claves": [
    "agro*",
    "agricol*",
    "campesin*",
    "ganaderia",
    "ganadero*",
    "cultivo*",
    "piscicultura",
    "gallina*",
    "emprendimiento*",
    "emprendedor*",
    "empresario*",
    "economia familiar",
    "seguridad alimentaria",
    "empleo",
    "desempleo",
    "comercio",
    "productiv*",
    "pescador*",
    "alevinos",
    "finca*",
    "cafe",
    "agropecuari*",
    "empresarial*",
    "empresas verdes",
    "gremio*",
    "abigeato no"
   ]
  },
  {
   "tema": "Turismo, cultura y patrimonio",
   "claves": [
    "turismo",
    "turístic*",
    "turistic*",
    "cultur*",
    "patrimonio",
    "festival*",
    "porro",
    "cuadros vivos",
    "regata*",
    "mural*",
    "arte",
    "musica*",
    "carnaval",
    "hotel*",
    "visitante*",
    "playa*",
    "libro",
    "escritor*",
    "biblioteca",
    "gastronomia",
    "museo",
    "feria*",
    "danza",
    "vallenato",
    "bandera",
    "golfo de morrosquillo",
    "velero*",
    "cabalgata*",
    "fiestas",
    "aniversario",
    "conmemoracion*",
    "celebracion*",
    "identidad",
    "tradicion*"
   ]
  },
  {
   "tema": "Ayudas sociales y atención a comunidades",
   "claves": [
    "ayuda*",
    "canasta*",
    "donacion*",
    "sillas de ruedas",
    "subsidio*",
    "entrega de alimentos",
    "albergue*",
    "bonos",
    "jornada de atencion",
    "atencion a la comunidad",
    "apoyo social",
    "gestora social",
    "vejez",
    "discapacidad",
    "adulto mayor",
    "hambre",
    "solidaridad",
    "familias beneficiadas",
    "jornada buen futuro",
    "economias para la vida",
    "prosperidad social",
    "poblacion vulnerable",
    "beneficiari*"
   ]
  },
  {
   "tema": "Juventud, género y participación",
   "claves": [
    "joven*",
    "juventud*",
    "genero",
    "mujer*",
    "violencia politica",
    "participacion ciudadana",
    "consejeros juveniles",
    "mesa de juventudes",
    "igualdad",
    "victimas",
    "migrante*",
    "consejo municipal de juventud",
    "protagonistas del cambio",
    "consejo departamental de juventud",
    "juntas de accion",
    "poblacion migrante"
   ]
  },
  {
   "tema": "Cooperación, paz y derechos humanos",
   "claves": [
    "cooperacion",
    "internacional*",
    "embajador*",
    "embajada",
    "australia",
    "italia",
    "union europea",
    "laboratorio de paz",
    "paz",
    "derechos humanos",
    "defensoria",
    "desplazamiento*",
    "onu",
    "acnur",
    "unesco",
    "relaciones exteriores",
    "despojo",
    "alerta temprana",
    "proteccion de la poblacion",
    "comunidad indigena",
    "cabildo indigena",
    "etnia",
    "zenu",
    "poblacion civil",
    "conflicto armado",
    "acompanamiento a la comunidad"
   ]
  },
  {
   "tema": "Gestión institucional y comunitaria",
   "claves": [
    "gobierno",
    "consejo de gobierno",
    "alcalde",
    "alcaldia",
    "asamblea",
    "diputad*",
    "rendicion de cuentas",
    "comite*",
    "foro*",
    "visita*",
    "audiencia*",
    "sesion*",
    "posesion*",
    "nombramiento*",
    "encargo*",
    "delegado*",
    "gestion",
    "institucional",
    "convenio*",
    "alianza*",
    "comunidad*",
    "escuadron",
    "junta de accion",
    "dialogo*",
    "prioridades",
    "despacho",
    "fondo mixto",
    "impuesto*",
    "tributaria*",
    "hacienda",
    "tasa de seguridad",
    "recaudo",
    "presupuesto*",
    "pasaporte*",
    "movilidad",
    "programa",
    "columna*",
    "opinion",
    "personajes del dia",
    "ronda pais",
    "tema general",
    "especial region",
    "enfoque",
    "balance de gestion",
    "acto administrativo`"
   ]
  }
 ]
}
TAX_GREMIO = {
 "nota": "Cubos para gremios, federaciones y camaras (sector productivo). El orden es la prioridad: lo especifico antes que lo generico.",
 "temas": [
  "Ayudas y solidaridad por emergencias",
  "Precios e inflación del pollo y el huevo",
  "Cadena productiva del maíz",
  "Cría y producción avícola",
  "Política sectorial y gremios",
  "Tecnología, innovación y retos hacia 2030",
  "Congreso, foros y eventos del gremio",
  "Consumo y promoción del pollo",
  "Consumo y promoción del huevo",
  "Exportaciones y mercados internacionales",
  "Competitividad y clima de inversión",
  "Empleo y aporte del sector a la economía",
  "Seguridad y delitos contra el sector",
  "Sostenibilidad y temas ambientales",
  "Economía y política nacional",
  "Otros"
 ],
 "reglas": [
  {
   "tema": "Ayudas y solidaridad por emergencias",
   "claves": [
    "terremoto",
    "sismo",
    "damnificad*",
    "fenaviton",
    "solidarid*",
    "incendios"
   ]
  },
  {
   "tema": "Precios e inflación del pollo y el huevo",
   "claves": [
    "precio*",
    "inflacion",
    "alza",
    "caro",
    "costo* de produccion",
    "salario minimo",
    "indice del pollo",
    "subida"
   ]
  },
  {
   "tema": "Cadena productiva del maíz",
   "claves": [
    "maiz",
    "alimento animal",
    "soya",
    "siembras",
    "cadena productiva"
   ]
  },
  {
   "tema": "Cría y producción avícola",
   "claves": [
    "cria de pollos",
    "economia familiar",
    "galpon*",
    "granjas",
    "produccion de pollo"
   ]
  },
  {
   "tema": "Política sectorial y gremios",
   "claves": [
    "plan para exportar pollo",
    "medidas financieras",
    "medidas para el sector",
    "linea de credito",
    "agenda",
    "beneficios tributarios",
    "tributari*",
    "minagricultura",
    "ministro de agricultura",
    "de la espriella",
    "credito*",
    "industriales",
    "sac",
    "agro",
    "gremio*",
    "amcham"
   ]
  },
  {
   "tema": "Tecnología, innovación y retos hacia 2030",
   "claves": [
    "inteligencia artificial",
    "ia(?= )",
    "automatizacion",
    "tecnolog*",
    "digital",
    "innovacion",
    "2030",
    "transformacion",
    "continuidad del negocio",
    "hoja de ruta",
    "datos y"
   ]
  },
  {
   "tema": "Congreso, foros y eventos del gremio",
   "claves": [
    "congreso",
    "foro",
    "panel",
    "conferencista*",
    "asamblea",
    "seminario",
    "respaldo"
   ]
  },
  {
   "tema": "Consumo y promoción del pollo",
   "claves": [
    "pollo week",
    "semana del pollo",
    "pollo asado",
    "pollo frito",
    "platos de pollo",
    "restaurante*",
    "gastronom*",
    "tulio",
    "festival",
    "comer pollo",
    "consumo de pollo"
   ]
  },
  {
   "tema": "Consumo y promoción del huevo",
   "claves": [
    "huevo*",
    "nevera",
    "ovoproducto*",
    "conservar los huevos"
   ]
  },
  {
   "tema": "Exportaciones y mercados internacionales",
   "claves": [
    "exportacion*",
    "exportar",
    "exportadora",
    "estados unidos",
    "japon",
    "china",
    "emiratos",
    "corea",
    "caribe",
    "mercados internacionales",
    "comercio internacional",
    "abrir mercados"
   ]
  },
  {
   "tema": "Competitividad y clima de inversión",
   "claves": [
    "competitividad",
    "inversion*",
    "reglas claras",
    "estabilidad y seguridad",
    "empresarios",
    "consejo gremial"
   ]
  },
  {
   "tema": "Empleo y aporte del sector a la economía",
   "claves": [
    "empleo*",
    "aporte*",
    "pib",
    "economia",
    "crecimiento",
    "industria avicola",
    "papel de la avicultura",
    "genera"
   ]
  },
  {
   "tema": "Seguridad y delitos contra el sector",
   "claves": [
    "robo*",
    "hurto*",
    "delincuente*",
    "armados",
    "furgones",
    "granja avicola",
    "seguridad"
   ]
  },
  {
   "tema": "Sostenibilidad y temas ambientales",
   "claves": [
    "olores",
    "gallinero*",
    "vertimiento*",
    "ambiental",
    "residuales",
    "sostenibilidad",
    "contaminacion",
    "demanda*"
   ]
  },
  {
   "tema": "Economía y política nacional",
   "claves": [
    "dolar",
    "apagon",
    "protesta*",
    "consejo de estado",
    "el nino",
    "energia",
    "transmilenio",
    "cooperacion"
   ]
  }
 ]
}
CUBO_PROHIBIDO = {'otros', 'otro', 'varios', 'varias', 'general', 'generales', 'miscelaneo',
                  'miscelanea', 'sin clasificar', 'no clasificado', 'informacion general',
                  'sin categoria', 'pendiente'}
META_CUBO = {'entidad', 'cliente', 'empresa', 'compania', 'organizacion', 'institucion', 'institucional',
             'gubernamental', 'departamental', 'regional', 'municipal', 'general', 'generales', 'varios',
             'varias', 'otros', 'otras', 'miscelaneo', 'miscelanea', 'temas', 'asuntos'}


def patron(kw):
    if '(' in kw or '?' in kw:
        return r'(?<![a-z])' + kw
    if kw.endswith('*'):
        return r'(?<![a-z])' + re.escape(nz(kw[:-1])) + r'[a-z]*'
    return r'(?<![a-z])' + re.escape(nz(kw)) + r'(?![a-z])'


def tema_de(sub_tema, titulo, tax):
    """Dos pasadas (pieza 5): manda el sub-tema; el titulo solo si parece titular corto."""
    for txt in (nz(sub_tema), nz(titulo) if len(str(titulo or '')) <= 160 else ''):
        if not txt:
            continue
        for r in tax['reglas']:
            for k in r['claves']:
                if re.search(patron(k), txt):
                    return r['tema'], k
    return None, None


def cubo_valido(nombre, tax, permitir_nuevos=True):
    """Acepta un cubo de la lista del cliente, o uno NUEVO solo si es específico de verdad.

    Rechaza: 'Otros' y sus primos, rótulos vacíos ('noticias generales'), y cualquier nombre
    construido con palabras marco ('actividad institucional', 'boletín interno de la entidad').
    """
    nombre = sq(nombre)
    if not nombre:
        return None
    n = nz(nombre)
    if n in CUBO_PROHIBIDO or n in ROTULO_GEN:
        return None
    en_lista = next((t for t in tax['temas'] if nz(t) == n), None)
    if en_lista:
        return en_lista
    if not permitir_nuevos:
        return None
    toks = [t for t in n.split() if t]
    if not (2 <= len(nombre.split()) <= 5):
        return None
    if any(t in META_CUBO for t in toks):
        return None
    contenido = [t for t in toks if t not in CONECT and t not in FILLER and t not in MARCO and len(t) > 3]
    if not contenido:
        return None
    return nombre


# ============================================================================
# 5. CAPA LLM
# ============================================================================
PROVEEDORES = {
    'Groq': ('https://api.groq.com/openai/v1', 'llama-3.3-70b-versatile'),
    'OpenAI': ('https://api.openai.com/v1', 'gpt-4.1-mini'),
    'OpenAI (nano)': ('https://api.openai.com/v1', 'gpt-4.1-nano-2025-04-14'),
    'DeepSeek': ('https://api.deepseek.com/v1', 'deepseek-chat'),
    'Otro (compatible OpenAI)': ('', ''),
}
CLAVES_SECRETS = ('llm_api_key', 'api_key', 'openai_api_key', 'OPENAI_API_KEY', 'groq_api_key',
                  'GROQ_API_KEY')
CLAVES_PASSWORD = ('app_password', 'APP_PASSWORD', 'password', 'PASSWORD')


def _secciones_secrets(secrets_obj):
    """Devuelve [raíz] + cada sección [nombre] de los secrets, para aceptar las claves sueltas.

    Ojo: las secciones de st.secrets NO son `dict` sino un Mapping de Streamlit
    (AttrDict); con isinstance(v, dict) se ignoraban y ni la contraseña ni la api key
    del bloque [general] se leían.
    """
    try:
        s = dict(secrets_obj)
    except Exception:
        return []
    return [s] + [v for v in s.values() if isinstance(v, Mapping)]


def password_esperada(secrets_obj):
    """Contraseña configurada en los secrets, o None si no hay.

    Acepta `app_password` (como las otras apps del usuario) o `APP_PASSWORD`, en la raíz o dentro
    de una sección [general]. Admite texto plano o un hash con el prefijo `sha256:`.
    """
    for sec in _secciones_secrets(secrets_obj):
        for k in CLAVES_PASSWORD:
            v = sec.get(k)
            if v:
                return str(v)
    return None


def coincide_password(ingresada, esperada):
    """True si la contraseña ingresada corresponde. Soporta `sha256:...` en la configurada."""
    if not esperada:
        return False
    ing = str(ingresada or '')
    if esperada.startswith('sha256:'):
        return hmac.compare_digest(hashlib.sha256(ing.encode('utf-8')).hexdigest(),
                                   esperada.split(':', 1)[1].strip().lower())
    return hmac.compare_digest(ing, esperada)


def leer_secrets():
    """Lee proveedor/modelo/base_url/api_key de st.secrets (Streamlit Cloud o secrets.toml local).

    Acepta las claves sueltas o dentro de una sección [general]. Si no hay nada, devuelve {} y la
    app funciona igual escribiendo la key en la barra lateral.
    """
    cfg = {}
    try:
        s = dict(st.secrets)
    except Exception:
        return cfg
    for sec in _secciones_secrets(s):
        for k in ('proveedor', 'base_url', 'modelo', 'criterio'):
            if k in sec and not cfg.get(k):
                cfg[k] = str(sec[k])
        if not cfg.get('api_key'):
            for k in CLAVES_SECRETS:
                if k in sec and sec[k]:
                    cfg['api_key'] = str(sec[k])
                    break
    return cfg

EJEMPLOS = [
 {
  "titulo": "Sucre lo hace de nuevo: 40 mil niños y niñas inician sus clases con alimentación escolar desde el primer día",
  "sub_tema": "Inicio de clases con alimentación escolar",
  "tono": "Positivo"
 },
 {
  "titulo": "Gobernación de Sucre impulsa economía familiar y seguridad alimentaria en Toluviejo con 2 mil gallinas ponedoras",
  "sub_tema": "Gallinas ponedoras para economía familiar",
  "tono": "Positivo"
 },
 {
  "titulo": "Ciudad Natural del Golfo de Morrosquillo: la estrategia de Sucre para dinamizar el turismo",
  "sub_tema": "Ciudad Natural del Golfo de Morrosquillo",
  "tono": "Positivo"
 },
 {
  "titulo": "En Sucre destruyen más de 250 mil productos de contrabando valorados en más de 670 millones de pesos",
  "sub_tema": "Destrucción de productos de contrabando",
  "tono": "Positivo"
 },
 {
  "titulo": "Sucre abre nuevas rutas de cooperación internacional tras visita de la embajadora de Australia, Anna Chrisp",
  "sub_tema": "Cooperación internacional con Australia",
  "tono": "Positivo"
 },
 {
  "titulo": "Gobernación de Sucre aprobó más de $39 mil millones para la construcción de la Variante Sampués - Segovia- Sincelejo",
  "sub_tema": "Aprobación de recursos para la Variante Sampués",
  "tono": "Positivo"
 },
 {
  "titulo": "Por demoras en el PAE, Procuraduría suspende a exsecretario de Educación de Sucre",
  "sub_tema": "Sanción a exsecretario de Educación",
  "tono": "Negativo"
 },
 {
  "titulo": "Sancionan a exsecretario de Educación de Sucre, por demora en el PAE",
  "sub_tema": "Sanción por retraso en el PAE",
  "tono": "Negativo"
 },
 {
  "titulo": "Vía al Llano, una obra que está estancada",
  "sub_tema": "Estancamiento de la vía al Llano",
  "tono": "Negativo"
 },
 {
  "titulo": "Gobernadores del Caribe y la ANI evalúan proyecto del Canal del Dique",
  "sub_tema": "Avance del proyecto Canal del Dique",
  "tono": "Neutro"
 },
 {
  "titulo": "Alcalde de Cartagena y gobernadores de Bolívar, Sucre y Atlántico revisaron la operación del Canal del Dique",
  "sub_tema": "Operación del Canal del Dique",
  "tono": "Neutro"
 },
 {
  "titulo": "Yamil Arana, el mejor gobernador de la Región Caribe, según encuesta de Datanálisis",
  "sub_tema": "Encuesta de aprobación al gobernador Arana",
  "tono": "Neutro"
 },
 {
  "titulo": "Defensoría advierte riesgo inminente por violencia y desplazamientos en El Roble",
  "sub_tema": "Alerta temprana por desplazamientos en El Roble",
  "tono": "Neutro"
 },
 {
  "titulo": "Consolidación de la región Caribe como potencia bioenergética es inaplazable: gobernador Verano",
  "sub_tema": "Región Caribe como potencia bioenergética",
  "tono": "Neutro"
 },
 {
  "titulo": "Telecaribe/ El Reportero del Campo",
  "sub_tema": "Programa El Reportero del Campo",
  "tono": "Neutro"
 }
]

EJEMPLOS_TEMA = [
    {'titulo': 'La entidad presentó el informe Panorama de la Juventud 2026: desempleo juvenil y salud '
               'mental en alerta',
     'sub_tema': 'Informe sobre juventud y desempleo', 'tono': 'Neutro'},
    {'titulo': 'El gremio advierte que la informalidad laboral sigue creciendo en el país',
     'sub_tema': 'Informe sobre informalidad laboral', 'tono': 'Neutro'},
    {'titulo': 'El estudio de la entidad revela brechas de salud mental en los jóvenes',
     'sub_tema': 'Estudio sobre salud mental juvenil', 'tono': 'Neutro'},
    {'titulo': 'Más de la mitad de los intentos de suicidio en Colombia corresponden a jóvenes entre 15 y 29 años',
     'sub_tema': 'Intentos de suicidio en jóvenes', 'tono': 'Neutro'},
    {'titulo': 'Obras de manejo ambiental no dan espera en la ciénaga del Totumo',
     'sub_tema': 'Obras pendientes en la ciénaga', 'tono': 'Neutro'},
    {'titulo': 'Muere una atleta y los especialistas en cuidados intensivos reaccionan en redes',
     'sub_tema': 'Reacciones por muerte de atleta', 'tono': 'Neutro'},
    {'titulo': 'Vecinos denuncian que la universidad no ha terminado la obra del bloque nuevo',
     'sub_tema': 'Denuncia por obra sin terminar', 'tono': 'Negativo'},
    {'titulo': 'Roban 180.000 huevos en una granja del Atlántico y la Policía recupera los camiones',
     'sub_tema': 'Robo a granja del Atlántico', 'tono': 'Neutro'},
    {'titulo': 'La Contraloría cuestiona los sobrecostos en la obra que ejecuta la entidad',
     'sub_tema': 'Cuestionamientos por sobrecostos', 'tono': 'Negativo'},
]

EJEMPLOS_SECTOR = [
    {'titulo': 'FENAVI realizará su congreso de 2028 en Barranquilla',
     'sub_tema': 'Congreso avícola 2028 en Barranquilla', 'tono': 'Positivo'},
    {'titulo': 'Gobierno impulsa plan para convertir a Colombia en potencia exportadora de pollo y huevo',
     'sub_tema': 'Plan para exportar pollo y huevo', 'tono': 'Positivo'},
    {'titulo': 'Fenavitón: el gremio avícola busca ayudar a familias damnificadas por el terremoto',
     'sub_tema': 'Ayudas de Fenavi por el terremoto', 'tono': 'Positivo'},
    {'titulo': 'Expertos analizan el impacto de la inteligencia artificial en la industria durante el congreso',
     'sub_tema': 'Paneles de IA en el congreso avícola', 'tono': 'Positivo'},
    {'titulo': '180.000 huevos y 15 hombres armados: el millonario robo a una granja avícola del Atlántico',
     'sub_tema': 'Robo a granja avícola en Sabanalarga', 'tono': 'Neutro'},
    {'titulo': 'El precio del pollo asado sigue subiendo y estas son las causas',
     'sub_tema': 'Aumento del precio del pollo asado', 'tono': 'Neutro'},
    {'titulo': 'Campesinos denuncian que una empresa vierte aguas residuales en una quebrada',
     'sub_tema': 'Denuncia por vertimientos de Mac Pollo', 'tono': 'Negativo'},
    {'titulo': 'Alcalde de Cartagena y gobernadores evalúan el proyecto del Canal del Dique',
     'sub_tema': 'Revisión del Canal del Dique', 'tono': 'Neutro'},
]
CRITERIOS_TONO = {
    'Aspectual estricto (recomendado)': (
        "El tono mide SOLO lo que se dice de la entidad, de su vocero o de sus funcionarios.\n"
        "- Positivo: la entidad o su vocero es sujeto de un hecho favorable (obra entregada, avance,\n"
        "  beneficio para la comunidad, programa, reconocimiento, cifra buena, declaracion que los deja bien).\n"
        "- Negativo: existe critica, reclamo, sancion, denuncia o evaluacion negativa DIRIGIDA a la entidad,\n"
        "  a su administracion o a sus funcionarios.\n"
        "- LA PREGUNTA CLAVE antes de escribir Negativo: ¿de quien habla la nota? Si la entidad o su vocero\n"
        "  NO aparece como responsable, señalado o protagonista de la critica, el tono es Neutro. Los temas\n"
        "  tristes o graves NO son Negativo para la entidad: muertes, suicidio, delincuencia, desempleo,\n"
        "  pobreza, inundaciones, obras inconclusas de terceros, quejas contra otros. Una nota puede\n"
        "  mencionar a la entidad y seguir siendo Neutro (la entidad estudia el problema, participa en un\n"
        "  foro, firma una alianza o es una voz mas entre varias).\n"
        "- EL TEMA NO DECIDE EL TONO. Si la entidad publica un informe, estudio, encuesta o campaña sobre un\n"
        "  problema (desempleo, salud mental, pobreza, violencia, inseguridad, medio ambiente), el tono es\n"
        "  Neutro, y Positivo si la entidad aparece como autora de un aporte (diagnostico, propuesta,\n"
        "  solucion, alianza). Que el tema sea grave o triste NO hace Negativo a quien lo investiga.\n"
        "  Negativo exige siempre un ataque, critica o señalamiento CONTRA la entidad o su vocero.\n"
        "- Neutro: todo lo demas. Incluye hechos malos sin responsable institucional (inundaciones,\n"
        "  homicidios, accidentes, robos, alzas de precios), la cobertura de OTRA entidad del mismo\n"
        "  territorio, y los casos en que el vocero denuncia a un tercero.\n"
        "Ante duda entre Positivo y Neutro, o entre Negativo y Neutro, elige Neutro."
    ),
    'Favorabilidad del sector (para gremios)': (
        "COMO DECIDIR EL TONO (en este orden; el primero que se cumpla gana)\n"
        "P1. La nota deja bien al sector o a la entidad: congreso o evento del gremio, campaña de\n"
        "    consumo de sus productos, exportaciones o mercados nuevos, crecimiento o cifras buenas,\n"
        "    reconocimiento, modernizacion, tecnologia, innovacion, competitividad, agenda o plan a\n"
        "    futuro, o un plan oficial que beneficia al sector aunque lo anuncie un ministerio -> Positivo.\n"
        "    Si la nota trata del sector en tono positivo o informativo y NO es una critica, es Positivo.\n"
        "P2. La nota tiene una critica, denuncia, sancion o señalamiento contra el gremio, su vocero o una\n"
        "    empresa del sector, o un hecho que se le atribuye y daña su imagen (contaminacion, malas\n"
        "    practicas, incumplimiento) -> Negativo.\n"
        "P3. Neutro SOLO si la nota no trata del sector ni lo afecta: politica nacional, otro gremio, otro\n"
        "    sector, economia del pais, resultados de otra entidad. Tambien son Neutro: las notas de\n"
        "    servicio o consejos al consumidor, las alertas economicas o de seguridad general, los datos\n"
        "    de precios y la agenda de una entidad distinta.\n"
        "NO son Neutro las notas del sector sobre tecnologia, congresos o planes: esas son Positivo.\n"
        "EL TEMA NO DECIDE EL TONO: si el gremio o la entidad publica un informe o estudio sobre un\n"
        "problema (desempleo, salud mental, pobreza, precios, inseguridad), no es Negativo; es Neutro o\n"
        "Positivo segun su encuadre. Negativo exige critica o hecho atribuible al sector.\n"
        "Los robos, hurtos y delitos contra granjas o empresas del sector son Neutro: son la victima,\n"
        "no la falta. Ante duda, elige Neutro."
    ),
}

REGLAS_SUBTEMA = (
    "El Sub-tema es el HECHO concreto de la nota, en 3 a 5 palabras (nunca mas de 7).\n"
    "- Frase nominal, sin verbo conjugado al inicio (bien: 'Entrega del parque'; mal: 'Entregaron el\n"
    "  parque'). Si empiezas con un sustantivo de accion esta bien: 'Anuncio de inversiones'.\n"
    "- Sin terminar en preposicion o nexo, y sin dos puntos, comas, comillas ni barras.\n"
    "- No copies el titular ni recortes una frase del texto: sintetiza el hecho.\n"
    "- Sin repetir el nombre del gremio, de la entidad ni del medio.\n"
    "- Sin etiquetas de categoria. MAL: 'Exportaciones y mercados internacionales', 'Transformacion\n"
    "  digital en el sector'. BIEN: 'Exportacion de pollo a Estados Unidos', 'Paneles de IA en el\n"
    "  congreso avicola'.\n"
    "- Incluye el actor o el lugar cuando son lo que distingue el hecho: no 'Visita internacional' sino\n"
    "  'Visita de la embajadora de Australia'.\n"
    "- Prohibido rotulos vacios: 'noticias generales', 'gestion institucional', 'varios'.\n"
    "- Si el hecho ya esta en CANDIDATOS, copia ese texto EXACTO (mismas palabras y mayusculas).\n"
    "  Nunca crees una variante nueva de un hecho que ya tiene sub-tema."
)


def prompt_sistema(cfg):
    entidad = cfg['entidad'] or 'la entidad analizada'
    voceros = ', '.join(cfg['voceros']) or '(no definido)'
    alias = ', '.join(cfg['alias']) or '(sin alias adicionales)'
    lineas = [
        'Eres analista senior de monitoreo de medios en Colombia. Trabajas sobre notas en espanol y tu',
        'trabajo es etiquetar CADA GRUPO de notas con Sub-tema y Tono.',
        '',
        'Un GRUPO es una nota publicada por varios medios (o notas casi iguales): se etiqueta UNA sola vez',
        'y esa etiqueta aplica a todas sus menciones. Nunca cambies la etiqueta entre menciones del mismo grupo.',
        '',
        'ENTIDAD OBJETIVO DEL ANALISIS: %s' % entidad,
        'VOCERO(S): %s' % voceros,
        'ALIAS Y FORMAS DE NOMBRARLA EN LOS MEDIOS: %s' % alias,
        '',
        'REGLA DE TONO',
        CRITERIOS_TONO[cfg['criterio']],
        'Si la entidad no aparece mencionada ni participa, el tono es Neutro sin mas analisis.',
        '',
        'REGLA DE SUB-TEMA',
        REGLAS_SUBTEMA,
        '',
        'EJEMPLOS YA ETIQUETADOS',
    ]
    ejemplos = list(EJEMPLOS) + EJEMPLOS_TEMA
    if str(cfg.get('criterio', '')).startswith('Favorabilidad'):
        ejemplos += EJEMPLOS_SECTOR
    for e in ejemplos:
        lineas.append('  TITULAR: %s' % e['titulo'])
        lineas.append('  ->  sub_tema: "%s"  |  tono: %s' % (e['sub_tema'], e['tono']))
    lineas += [
        '',
        'SALIDA',
        'Responde UNICAMENTE con JSON valido, sin markdown y sin explicaciones, con esta forma:',
        '{"resultados":[{"id":<numero de grupo>,"sub_tema":"<3 a 7 palabras>","tono":"Positivo|Neutro|Negativo"}]}',
        'Debes devolver un objeto por cada grupo recibido, con su id exacto.',
    ]
    return '\n'.join(lineas)


def prompt_lote(grupos_lote, candidatos):
    bloques = []
    for g in grupos_lote:
        b = ['GRUPO id=%d (%d menciones)' % (g['grupo'], g['n']),
             'TITULAR: %s' % sq(g['titulo'])[:220]]
        if g.get('titulos_alt'):
            b.append('OTROS TITULARES DEL MISMO GRUPO: %s' % ' // '.join(sq(t)[:120] for t in g['titulos_alt']))
        b.append('TEXTO: %s' % sq(g['texto'])[:700])
        bloques.append('\n'.join(b))
    msg = '\n\n'.join(bloques)
    msg += ('\n\nRecuerda: el sub_tema de cada grupo debe tener entre 3 y 5 palabras, y solo JSON.')
    if candidatos:
        msg += ('\n\nCANDIDATOS (sub-temas ya usados; reutiliza el mismo texto si el hecho es el mismo):\n'
                + '\n'.join('- %s' % c for c in candidatos[-120:]))
    return msg


def llamar_llm(cfg, mensajes, json_mode=True, max_tokens=4000, temperatura=0.0, intentos=3):
    url = cfg['base_url'].rstrip('/') + '/chat/completions'
    payload = {'model': cfg['modelo'], 'messages': mensajes, 'temperature': temperatura,
               'max_tokens': max_tokens}
    if json_mode:
        payload['response_format'] = {'type': 'json_object'}
    cab = {'Authorization': 'Bearer %s' % cfg['api_key'], 'Content-Type': 'application/json'}
    ultimo = ''
    for k in range(intentos):
        try:
            r = requests.post(url, headers=cab, json=payload, timeout=cfg.get('timeout', 120))
            if r.status_code in (429, 500, 502, 503):
                ultimo = 'HTTP %s' % r.status_code
                time.sleep(2 + 3 * k)
                continue
            if r.status_code != 200:
                raise RuntimeError('HTTP %s: %s' % (r.status_code, r.text[:300]))
            return r.json()['choices'][0]['message']['content']
        except requests.RequestException as e:
            ultimo = str(e)[:200]
            time.sleep(2 + 3 * k)
    raise RuntimeError('Fallo la llamada al modelo: %s' % ultimo)


def _json_loose(txt):
    txt = re.sub(r'^```(?:json)?|```$', '', str(txt or '').strip(), flags=re.M).strip()
    try:
        return json.loads(txt)
    except Exception:
        pass
    m = re.search(r'\{.*\}', txt, flags=re.S)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            return None
    return None


def etiquetar_lote(cfg, grupos_lote, candidatos):
    msgs = [{'role': 'system', 'content': prompt_sistema(cfg)},
            {'role': 'user', 'content': prompt_lote(grupos_lote, candidatos)}]
    txt = llamar_llm(cfg, msgs)
    data = _json_loose(txt)
    salida = {}
    for r in (data or {}).get('resultados', []) or []:
        try:
            gid = int(r.get('id'))
        except Exception:
            continue
        if any(g['grupo'] == gid for g in grupos_lote):
            salida[gid] = {'sub_tema': sq(r.get('sub_tema')), 'tono': sq(r.get('tono')).capitalize()}
    return salida


def reparar_lote(cfg, fallos):
    """Devuelve las etiquetas invalidas al modelo con el problema exacto, para que las corrija."""
    detalle = []
    for f in fallos:
        detalle.append('GRUPO id=%d\nTITULAR: %s\nTEXTO: %s\nSUB-TEMA ACTUAL: "%s"\nPROBLEMAS: %s'
                       % (f['grupo'], sq(f['titulo'])[:180], sq(f['texto'])[:450],
                          f['sub_tema'], '; '.join(f['problemas'])))
    msgs = [{'role': 'system', 'content': prompt_sistema(cfg)},
            {'role': 'user', 'content':
             'Corrige SOLO estos sub-temas. Devuelve el mismo tono salvo que el tono no este permitido.\n'
             'Un sub-tema valido tiene 3 a 5 palabras (maximo 7) en frase nominal, no empieza con verbo\n'
             'conjugado, no termina en preposicion y no lleva marcadores ni rotulos vacios. Si el problema\n'
             'dice largo(N), recorta a 5 palabras sin perder el hecho.\n\n'
             + '\n\n'.join(detalle)
             + '\n\nResponde UNICAMENTE con {"resultados":[{"id":<grupo>,"sub_tema":"...","tono":"..."}]}'}]
    txt = llamar_llm(cfg, msgs)
    data = _json_loose(txt) or {}
    out = {}
    for r in data.get('resultados', []) or []:
        try:
            out[int(r.get('id'))] = {'sub_tema': sq(r.get('sub_tema')), 'tono': sq(r.get('tono')).capitalize()}
        except Exception:
            continue
    return out


def _voto_mayoria(pasadas, ids_lote):
    """Combina varias pasadas: gana el tono mas votado (empate -> Neutro) y el sub-tema mas repetido."""
    salida = {}
    for gid in ids_lote:
        tonos = [v[gid]['tono'] for v in pasadas if gid in v and v[gid].get('tono')]
        subs = [v[gid]['sub_tema'] for v in pasadas if gid in v and v[gid].get('sub_tema')]
        if not tonos and not subs:
            continue
        top = collections.Counter(tonos).most_common()
        if not top:
            tono = 'Neutro'
        elif len(top) > 1 and top[0][1] == top[1][1]:
            tono = 'Neutro' if 'Neutro' in (top[0][0], top[1][0]) else top[0][0]
        else:
            tono = top[0][0]
        cs = collections.Counter(nz(x) for x in subs)
        if not cs:
            sub = ''
        else:
            rep = max(cs.values())
            sub = min([x for x in subs if cs[nz(x)] == rep], key=len)
        salida[gid] = {'sub_tema': sub, 'tono': tono}
    return salida


def etiquetar_todo(cfg, grupos, progreso=None, tam_lote=15, max_reparaciones=2, votos=2):
    """Etiqueta todos los grupos: lotes -> votacion -> validacion -> reparacion."""
    etiquetas, bitacora = {}, []
    candidatos = []
    total = len(grupos)
    hechos = 0
    votos = max(1, int(votos or 1))
    for i in range(0, total, tam_lote):
        lote = grupos[i:i + tam_lote]
        pasadas = []
        for _v in range(votos):
            try:
                pasadas.append(etiquetar_lote(cfg, lote, candidatos))
            except Exception:
                pasadas.append({})
        labels = _voto_mayoria(pasadas, [g['grupo'] for g in lote]) if votos > 1 \
            else (pasadas[0] if pasadas else {})
        for g in lote:
            etiquetas[g['grupo']] = labels.get(g['grupo'], {'sub_tema': '', 'tono': ''})
        # validacion + reparacion
        for ronda in range(max_reparaciones + 1):
            fallos = []
            for g in lote:
                e = etiquetas[g['grupo']]
                pr = validar(e['sub_tema'], e['tono'], [g['titulo']] + g.get('titulos_alt', []) + [g['texto']])
                duros = [x for x in pr if not x.startswith('revisar_anclaje')]
                if duros or not e['sub_tema']:
                    fallos.append({'grupo': g['grupo'], 'titulo': g['titulo'], 'texto': g['texto'],
                                   'sub_tema': e['sub_tema'], 'problemas': duros or ['vacio']})
            if not fallos or ronda == max_reparaciones:
                for f in fallos:
                    bitacora.append({'grupo': f['grupo'], 'titulo': f['titulo'][:70],
                                     'sub_tema': etiquetas[f['grupo']]['sub_tema'],
                                     'problemas': '; '.join(f['problemas'])})
                break
            corr = reparar_lote(cfg, fallos)
            for gid, v in corr.items():
                if v.get('sub_tema'):
                    etiquetas[gid] = {'sub_tema': v['sub_tema'],
                                      'tono': v['tono'] if v['tono'] in TONOS else etiquetas[gid]['tono']}
            bitacora.append({'grupo': 0, 'titulo': 'lote %d..%d' % (lote[0]['grupo'], lote[-1]['grupo']),
                             'sub_tema': '', 'problemas': 'reparacion ronda %d (%d fallos)' % (ronda + 1, len(fallos))})
        for g in lote:
            st_ = etiquetas[g['grupo']]['sub_tema']
            if st_:
                candidatos.append(st_)
        hechos += len(lote)
        if progreso:
            progreso(hechos, total, 'etiquetados %d de %d grupos' % (hechos, total))
    return etiquetas, bitacora


def elegir_cubo(cfg, pendientes, tax, permitir_nuevos=True):
    """El LLM solo elige DENTRO de la lista de cubos (o propone un cubo nuevo especifico)."""
    lista = '\n'.join('- %s' % t for t in tax['temas'] if nz(t) not in CUBO_PROHIBIDO)
    bloques = []
    for p in pendientes:
        bloques.append('GRUPO id=%d\nSUB-TEMA: %s\nTITULAR: %s\nTEXTO: %s'
                       % (p['grupo'], p['sub_tema'], sq(p['titulo'])[:180],
                          sq(p.get('texto', ''))[:260]))
    extra = ('Si ningun cubo sirve, propón uno NUEVO en 2 a 5 palabras que describa el asunto concreto\n'
             '(por ejemplo "Tramite de pasaportes"). No se acepta un cubo generico.\n' if permitir_nuevos
             else 'No propongas cubos nuevos: elige siempre uno de la lista.\n')
    msgs = [{'role': 'system', 'content': 'Clasificas notas de prensa en cubos tematicos cerrados.\n'
             'Cubos disponibles:\n' + lista + '\n\n' + extra +
             'Responde UNICAMENTE con JSON: {"resultados":[{"id":<grupo>,"cubo":"<nombre exacto del cubo o cubo nuevo>"}]}'},
            {'role': 'user', 'content': '\n\n'.join(bloques)}]
    out = {}
    try:
        data = _json_loose(llamar_llm(cfg, msgs)) or {}
        for r in data.get('resultados', []) or []:
            try:
                gid = int(r.get('id'))
            except Exception:
                continue
            out[gid] = cubo_valido(r.get('cubo'), tax, permitir_nuevos)
    except Exception as e:
        st.warning('No se pudo clasificar el tema de %d grupo(s): %s' % (len(pendientes), str(e)[:120]))
    return out


def cubo_de_respaldo(sub_tema, titulo):
    """Último recurso determinista: un cubo específico derivado del sub-tema (nunca "Otros")."""
    base = sq(sub_tema) or sq(titulo) or 'Asuntos del periodo'
    palabras = [w for w in re.split(r'\s+', base) if len(w) > 2][:4]
    nombre = ' '.join(palabras).strip() or 'Asuntos del periodo'
    return ('Asuntos especificos del periodo' if nz(nombre) in CUBO_PROHIBIDO else nombre)[:60]


def _mismo_cubo(a, b):
    """Dos cubos son el mismo si son casi iguales o si uno es el otro con un añadido."""
    from rapidfuzz import fuzz
    if fuzz.token_sort_ratio(nz(a), nz(b)) >= 90:
        return True
    ta, tb = set(nz(a).split()), set(nz(b).split())
    chico, grande = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    return bool(chico) and chico <= grande and len(chico) >= 2 and 0 < len(grande) - len(chico) <= 3


def canonizar_nuevos(temas, tax):
    """Unifica cubos NUEVOS que son variantes del mismo nombre (los de la lista no se tocan)."""
    en_lista = {nz(t) for t in tax.get('temas', [])}
    cambios = 0
    for a in list(temas):
        for b in list(temas):
            if a >= b:
                continue
            if nz(temas[a]) in en_lista or nz(temas[b]) in en_lista:
                continue
            if not _mismo_cubo(temas[a], temas[b]):
                continue
            # gana el nombre más corto, o el que ya tenga más grupos
            na = sum(1 for x in temas if temas[x] == temas[a])
            nb = sum(1 for x in temas if temas[x] == temas[b])
            ganador = temas[a] if (na > nb or (na == nb and len(temas[a]) <= len(temas[b]))) else temas[b]
            perdedor = temas[b] if ganador == temas[a] else temas[a]
            for x in list(temas):
                if temas[x] == perdedor:
                    temas[x] = ganador
            cambios += 1
    return cambios


def derivar_reglas(cubos):
    """Convierte los cubos en reglas léxicas para el primer pase determinista."""
    reglas = []
    for nombre in cubos:
        toks = [t for t in nz(nombre).split() if t]
        claves = [nz(nombre)]
        for t in toks:
            if len(t) >= 4 and t not in CONECT and t not in FILLER and t not in MARCO and t not in claves:
                claves.append(t)
        if len(claves) > 1:
            reglas.append({'tema': nombre, 'claves': claves})
    reglas.sort(key=lambda r: -len(nz(r['tema']).split()))
    return reglas


def _muestreo_grupos(grupos, etiquetas, por_bloque=35, max_bloques=12):
    lineas = []
    for g in grupos:
        st_ = (etiquetas.get(g['grupo']) or {}).get('sub_tema') or ''
        lineas.append('- %s | %s' % (st_[:60], sq(g['titulo'])[:110]))
    if len(lineas) > por_bloque * max_bloques:
        paso = max(1, len(lineas) // (por_bloque * max_bloques))
        lineas = lineas[::paso]
    return [lineas[i:i + por_bloque] for i in range(0, len(lineas), por_bloque)]


def proponer_taxonomia(cfg, grupos, etiquetas, objetivo=16, progreso=None):
    """Construye la lista de Temas A PARTIR DEL CONTENIDO del archivo, sin lista fija.

    Los clientes son muy distintos (universidades, sector público, privado, marcas), así que los
    cubos se proponen leyendo los hechos del propio archivo: por bloques y con una consolidación
    que elimina duplicados y solapamientos.
    """
    if progreso:
        progreso('Generando la lista de Temas a partir del archivo...')
    propuestas = []
    for bloque in _muestreo_grupos(grupos, etiquetas):
        msgs = [{'role': 'system', 'content':
                 'Eres analista de medios en Colombia. Agrupas hechos en cubos temáticos en JSON.'},
                {'role': 'user', 'content':
                 'Estos son hechos de un dossier de prensa:\n\n' + '\n'.join(bloque) +
                 '\n\nPropón entre 10 y 14 CUBOS TEMATICOS que los agrupen, pensando en un cliente '
                 'colombiano (puede ser universidad, entidad pública, empresa privada o marca).\n'
                 'Reglas: nombres de 2 a 5 palabras; específicos de ESTOS hechos, no genéricos; sin '
                 'solaparse entre sí; nada de "Otros", "Varios", "General" ni "Información".\n'
                 'Responde solo JSON: {"cubos":["Cubo uno","Cubo dos"]}'}]
        try:
            data = _json_loose(llamar_llm(cfg, msgs)) or {}
        except Exception:
            data = {}
        for c in data.get('cubos', []) or []:
            v = cubo_valido(c, {'temas': propuestas}, permitir_nuevos=True)
            if v and not any(_mismo_cubo(v, x) for x in propuestas):
                propuestas.append(v)
    if not propuestas:
        return {'nota': 'lista de respaldo', 'temas': list(TAX_GOBIERNO['temas']),
                'reglas': list(TAX_GOBIERNO['reglas'])}
    msgs = [{'role': 'system', 'content':
             'Eres analista de medios en Colombia. Consolidas listas de cubos temáticos en JSON.'},
            {'role': 'user', 'content':
             'Estos cubos fueron propuestos para el mismo dossier:\n\n'
             + '\n'.join('- %s' % x for x in propuestas) +
             '\n\nDevuelve la LISTA FINAL de %d cubos (menos si no hay materia): sin duplicados, '
             'sin solaparse, de 2 a 5 palabras, específicos, sin "Otros" ni genéricos.\n'
             'Responde solo JSON: {"cubos":["..."]}' % objetivo}]
    try:
        data = _json_loose(llamar_llm(cfg, msgs)) or {}
    except Exception:
        data = {}
    finales = []
    for c in data.get('cubos', []) or propuestas:
        if not isinstance(c, str):
            continue
        v = cubo_valido(c, {'temas': finales}, permitir_nuevos=True)
        if v and not any(_mismo_cubo(v, x) for x in finales):
            finales.append(v)
    if len(finales) < 3:
        finales = propuestas[:max(3, objetivo)]
    return {'nota': 'Cubos generados automáticamente a partir del contenido de este archivo.',
            'temas': finales, 'reglas': derivar_reglas(finales)}


def asignar_temas(cfg, grupos, etiquetas, tax, permitir_nuevos=True, overrides=None):
    overrides = overrides or {}
    temas, pendientes, origen = {}, [], {}
    for g in grupos:
        if g['grupo'] in overrides:
            temas[g['grupo']] = overrides[g['grupo']]
            origen[g['grupo']] = 'manual'
            continue
        e = etiquetas.get(g['grupo'], {})
        t, k = tema_de(e.get('sub_tema', ''), g['titulo'], tax)
        if t:
            temas[g['grupo']] = t
            origen[g['grupo']] = 'regla:%s' % k
        else:
            pendientes.append({'grupo': g['grupo'], 'sub_tema': e.get('sub_tema', ''),
                               'titulo': g['titulo'], 'texto': g.get('texto', '')})
    if pendientes:
        elegidos = elegir_cubo(cfg, pendientes, tax, permitir_nuevos)
        for p in pendientes:
            t = elegidos.get(p['grupo'])
            if t:
                temas[p['grupo']] = t
                origen[p['grupo']] = 'llm'
    return temas, [p for p in pendientes if p['grupo'] not in temas], origen


# ============================================================================
# 6. XLSX DE SALIDA
# ============================================================================
FILL = {'Positivo': PatternFill('solid', fgColor='C6EFCE'),
        'Neutro': PatternFill('solid', fgColor='F2F2F2'),
        'Negativo': PatternFill('solid', fgColor='FFC7CE')}


def construir_xlsx(cfg, header, datos, grupos, mapa, etiquetas, temas, bitacora):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Menciones'
    EXTRA = ['Fila original', 'Grupo de similitud', 'Tono', 'Tema', 'Sub-tema', 'Criterio del tono']
    for i, c in enumerate(list(header) + EXTRA, 1):
        cell = ws.cell(row=1, column=i, value=c)
        color = {'Tono': 'C00000', 'Tema': '1F6F43', 'Sub-tema': '4E7B2F'}.get(c, '1F4E79')
        if c.startswith('Grupo') or c.startswith('Fila'):
            color = '7030A0'
        cell.fill = PatternFill('solid', fgColor=color)
        cell.font = Font(bold=True, color='FFFFFF')
    crit = {'Positivo': 'Favorable a %s (obra, logro o declaracion)' % cfg['entidad'],
            'Negativo': 'Critica o senalamiento dirigido a %s' % cfg['entidad'],
            'Neutro': 'Sin logro ni critica dirigida a %s' % cfg['entidad']}
    for r in datos:
        gid = mapa.get(str(r.get('_id')))
        e = etiquetas.get(gid, {})
        ws.append([ctrl(v) for v in r['_fila']] + [r['_fila_n'], gid, e.get('tono', ''),
                                                   temas.get(gid, ''), e.get('sub_tema', ''),
                                                   crit.get(e.get('tono', ''), '')])
        ws.cell(row=ws.max_row, column=len(header) + 3).fill = FILL.get(e.get('tono'), FILL['Neutro'])
    for k, v in {'A': 6, 'B': 10}.items():
        ws.column_dimensions[k].width = v
    for col in range(3, len(header) + 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 26
    for col in range(len(header) + 1, len(header) + 7):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 30

    ws2 = wb.create_sheet('Temas (agrupa Sub-temas)')
    for i, c in enumerate(['Tema', 'Grupos', 'Menciones', 'Sub-temas agrupados'], 1):
        cell = ws2.cell(row=1, column=i, value=c)
        cell.fill = PatternFill('solid', fgColor='1F4E79')
        cell.font = Font(bold=True, color='FFFFFF')
    por = collections.defaultdict(list)
    for g in grupos:
        por[temas.get(g['grupo'], 'SIN ASIGNAR')].append(g)
    for t, gs in sorted(por.items(), key=lambda x: -sum(g['n'] for g in x[1])):
        ws2.append([t, len(gs), sum(g['n'] for g in gs),
                    ' | '.join(sorted(set(etiquetas.get(g['grupo'], {}).get('sub_tema', '') for g in gs)))])
    for k, v in {'A': 42, 'B': 9, 'C': 11, 'D': 130}.items():
        ws2.column_dimensions[k].width = v

    ws3 = wb.create_sheet('Grupos')
    for i, c in enumerate(['Grupo', 'Menciones', 'Medios', 'Título representativo', 'Tono', 'Tema', 'Sub-tema'], 1):
        cell = ws3.cell(row=1, column=i, value=c)
        cell.fill = PatternFill('solid', fgColor='1F4E79')
        cell.font = Font(bold=True, color='FFFFFF')
    for g in grupos:
        e = etiquetas.get(g['grupo'], {})
        ws3.append([g['grupo'], g['n'], '', ctrl(g['titulo'])[:200], e.get('tono', ''),
                    temas.get(g['grupo'], ''), e.get('sub_tema', '')])
        ws3.cell(row=ws3.max_row, column=5).fill = FILL.get(e.get('tono'), FILL['Neutro'])
    for k, v in {'A': 7, 'B': 10, 'C': 22, 'D': 54, 'E': 10, 'F': 34, 'G': 42}.items():
        ws3.column_dimensions[k].width = v

    ws4 = wb.create_sheet('Resumen')
    ct = collections.Counter(e.get('tono') for e in etiquetas.values())
    cm = collections.Counter(temas.values())
    def put(a, b=None):
        ws4.append([ctrl(a)] + ([ctrl(b)] if b is not None else []))
    put('Sentimiento (Tono), Tema y Sub-tema')
    put('Entidad / vocero', '%s · %s' % (cfg.get('entidad', ''), ', '.join(cfg.get('voceros', []) or [])))
    put('Alias considerados', ', '.join(cfg.get('alias', []) or []))
    put('Criterio del tono', cfg.get('criterio', ''))
    put('Modelo', '%s (%s)' % (cfg.get('modelo', ''), cfg.get('proveedor', '')))
    put('Texto base', '%s + %s' % (cfg.get('col_titulo', ''), cfg.get('col_texto', '')))
    put('')
    put('TONO')
    for k in TONOS:
        put('   ' + k, '%d grupos | %d menciones' % (ct.get(k, 0),
            sum(g['n'] for g in grupos if etiquetas.get(g['grupo'], {}).get('tono') == k)))
    put('')
    put('TEMA (cubo que agrupa sub-temas; ningun grupo queda sin cubo)')
    for k, v in cm.most_common():
        put('   ' + k, '%d grupos' % v)
    put('')
    put('SUB-TEMA')
    put('   distintos', str(len(set(e.get('sub_tema') for e in etiquetas.values()))))
    put('   largo medio', '%.1f palabras' % (sum(len(str(e.get('sub_tema', '')).split())
                                                for e in etiquetas.values()) / max(1, len(etiquetas))))
    tam = collections.Counter(g['n'] for g in grupos)
    put('')
    put('UNIFORMIDAD')
    put('   regla', 'menciones identicas o similares comparten Tono, Tema y Sub-tema')
    put('   grupos', '%d (un miembro: %d · con 2+ menciones: %d)'
        % (len(grupos), tam.get(1, 0), sum(v for k, v in tam.items() if k > 1)))
    if bitacora:
        put('')
        put('CONTROL DE CALIDAD')
        put('   reparaciones y avisos', str(len(bitacora)))
        for b in bitacora[:40]:
            put('   - G%s %s' % (b['grupo'] or '', b['titulo'][:60]), b['problemas'][:90])
    ws4.column_dimensions['A'].width = 30
    ws4.column_dimensions['B'].width = 110
    for row in ws4.iter_rows():
        for c in row:
            c.alignment = Alignment(vertical='top', wrap_text=True)
    b = io.BytesIO()
    wb.save(b)
    b.seek(0)
    return b


# ============================================================================
# 7. LECTURA DEL XLSX DE ENTRADA
# ============================================================================
CAND_TITULO = ('título', 'titulo', 'title', 'titular', 'headline', 'encabezado')
CAND_TEXTO = ('cuerpoes', 'cuerpo', 'contenido', 'texto', 'body', 'text', 'nota', 'resumen - aclaracion')
CAND_ID = ('noticiaid', 'id', 'ref', 'registro', 'nro', 'codigo')


def detectar(header, candidatos, obligatorio=False):
    for j, h in enumerate(header):
        if nz(h) in [nz(c) for c in candidatos]:
            return j
    if obligatorio:
        raise ValueError('No encontre la columna; disponibles: %s' % ', '.join(map(str, header)))
    return None


def leer_filas(hoja, col_tit, col_txt, col_id, extras):
    wb = load_workbook(io.BytesIO(hoja), read_only=True, data_only=True)
    return wb


def extraer(archivo, nombre_hoja, col_tit, col_txt, col_id, cols_extra):
    wb = load_workbook(io.BytesIO(archivo), read_only=True, data_only=True)
    ws = wb[nombre_hoja]
    it = ws.iter_rows(values_only=True)
    header = [ctrl(h) for h in next(it)]
    filas, saltadas = [], 0
    for i, r in enumerate(it):
        txt = ctrl(r[col_txt]) if r[col_txt] is not None else ''
        if not txt.strip():
            saltadas += 1
            continue
        tit = ctrl(r[col_tit]) if r[col_tit] is not None else ''
        filas.append({'_fila': [ctrl(v) for v in r], '_fila_n': i + 2,
                      '_id': str(r[col_id]) if col_id is not None else str(i + 2),
                      'id': str(r[col_id]) if col_id is not None else str(i + 2),
                      'titulo': tit, 'texto': txt,
                      'extra': {header[j]: ctrl(r[j]) for j in (cols_extra or [])}})
    return header, filas, saltadas


# ============================================================================
# 8. INTERFAZ
# ============================================================================
def exigir_password():
    """Puerta de entrada. Devuelve True si se puede ver la app.

    Si no hay contraseña en los secrets, deja pasar y avisa (modo local). Si hay, no se dibuja
    nada más de la app hasta que el usuario acierte.
    """
    esperada = password_esperada(dict(st.secrets)) if _hay_secrets() else None
    if not esperada:
        return True
    if st.session_state.get('_auth_ok'):
        return True
    st.markdown("""
    <div class="auth-wrap">
        <div class="auth-icon">◈</div>
        <div class="auth-title">Tono, Tema y Sub-tema de Menciones</div>
        <div class="auth-sub">Ingresa tus credenciales para continuar</div>
    </div>""", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 2, 1])
    with col:
        with st.form('_login'):
            pw = st.text_input('Contraseña', type='password', key='_pw_input')
            entrar = st.form_submit_button('Entrar', use_container_width=True, type='primary')
    if entrar:
        if coincide_password(pw, esperada):
            st.session_state['_auth_ok'] = True
            st.session_state['_intentos'] = 0
            st.rerun()
        else:
            st.session_state['_intentos'] = st.session_state.get('_intentos', 0) + 1
            st.error('Contraseña incorrecta.')
            time.sleep(min(1.0 + 0.5 * st.session_state['_intentos'], 3.0))
    st.stop()


def _hay_secrets():
    try:
        _ = dict(st.secrets)
        return True
    except Exception:
        return False


def aviso_sin_password():
    """Aviso visible cuando la app está publicada y no tiene contraseña configurada."""
    if not _hay_secrets():
        return
    if not password_esperada(dict(st.secrets)):
        st.warning('⚠️ Esta app no tiene contraseña configurada. Agrega `app_password` en los '
                   'Secrets (o en .streamlit/secrets.toml) para restringir el acceso.')


THEME_LIGHT_VARS = """
:root,[data-testid="stApp"]{
    --bg:#f8f9fa;--s1:#ffffff;--s2:#f1f3f4;--s3:#e8eaed;
    --border:#dadce0;--border2:#bdc1c6;--border-focus:#f97316;
    --text:#202124;--text2:#3c4043;--text3:#5f6368;--text4:#9aa0a6;
    --accent:#f97316;--accent2:#ea580c;--accent3:#c2410c;
    --accent-bg:#fff7ed;--accent-bg2:#ffedd5;--accent-bdr:#fed7aa;
    --green:#059669;--green2:#047857;--green-bg:#ecfdf5;--green-bdr:#a7f3d0;
    --red:#dc2626;--amber:#d97706;--blue:#1a73e8;
    --success-bg:linear-gradient(135deg,#ecfdf5,#d1fae5);
    --success-title:#047857;
    --icon-dossier-bg:#fff7ed;
    --r:8px;--r2:12px;--r3:16px;--r4:20px;
    --shadow-sm:0 1px 2px rgba(60,64,67,0.1),0 1px 3px rgba(60,64,67,0.08);
    --shadow-md:0 1px 3px rgba(60,64,67,0.12),0 4px 8px rgba(60,64,67,0.08);
    --shadow-lg:0 2px 6px rgba(60,64,67,0.1),0 8px 24px rgba(60,64,67,0.1);
    --transition:all 0.2s cubic-bezier(0.4,0,0.2,1);
}
"""

THEME_DARK_VARS = """
:root,[data-testid="stApp"]{
    --bg:#121418;--s1:#1c1f26;--s2:#252830;--s3:#2e333c;
    --border:#3d4450;--border2:#5c6370;--border-focus:#f97316;
    --text:#e8eaed;--text2:#c5c8ce;--text3:#9aa0a6;--text4:#6e7480;
    --accent:#f97316;--accent2:#fb923c;--accent3:#fdba74;
    --accent-bg:#2a1c10;--accent-bg2:#3d2814;--accent-bdr:#9a5b28;
    --green:#34d399;--green2:#6ee7b7;--green-bg:#0f291e;--green-bdr:#065f46;
    --red:#f87171;--amber:#fbbf24;--blue:#60a5fa;
    --success-bg:linear-gradient(135deg,#0f291e,#134e3a);
    --success-title:#6ee7b7;
    --icon-dossier-bg:#2a1c10;
    --r:8px;--r2:12px;--r3:16px;--r4:20px;
    --shadow-sm:0 1px 2px rgba(0,0,0,0.4),0 1px 3px rgba(0,0,0,0.25);
    --shadow-md:0 1px 3px rgba(0,0,0,0.45),0 4px 8px rgba(0,0,0,0.3);
    --shadow-lg:0 2px 6px rgba(0,0,0,0.4),0 8px 24px rgba(0,0,0,0.35);
    --transition:all 0.2s cubic-bezier(0.4,0,0.2,1);
}
"""

def _default_theme() -> str:
    try:
        theme_obj = getattr(getattr(st, "context", None), "theme", None)
        theme_type = getattr(theme_obj, "type", None)
        if theme_type in ("dark", "light"):
            return theme_type
    except Exception:
        pass
    return "light"

def current_ui_theme() -> str:
    theme = st.session_state.get("ui_theme")
    if theme in ("dark", "light"):
        return theme
    return _default_theme()

def load_custom_css():
    theme_vars = THEME_DARK_VARS if current_ui_theme() == "dark" else THEME_LIGHT_VARS
    st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Google+Sans+Text:wght@400;500;700&family=Roboto+Mono:wght@400;500&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
""" + theme_vars + """
html,body,[data-testid="stApp"]{
    background:var(--bg)!important;color:var(--text)!important;
    font-family:'Google Sans Text','Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
    font-size:14px;-webkit-font-smoothing:antialiased;letter-spacing:0.01em;
}
#MainMenu,footer,header{visibility:hidden}.stDeployButton{display:none}
.block-container{padding-top:1rem!important;padding-bottom:0!important}
[data-testid="stAppViewBlockContainer"]{padding-top:1rem!important}
.app-header{background:var(--s1);border:1px solid var(--border);border-radius:var(--r3);padding:1rem 1.5rem;margin-bottom:1rem;display:flex;align-items:center;gap:1rem;box-shadow:var(--shadow-sm);position:relative;overflow:hidden;}
.app-header::after{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,#f97316,#fb923c,#fdba74);}
.app-header-icon{width:40px;height:40px;background:linear-gradient(135deg,#f97316,#ea580c);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:1.2rem;color:white;flex-shrink:0;box-shadow:0 2px 8px rgba(249,115,22,0.3);}
.app-header-text{flex:1}
.app-header-title{font-family:'Google Sans',sans-serif;font-size:1.25rem;font-weight:700;color:var(--text);letter-spacing:-0.01em;line-height:1.3}
.app-header-version{font-family:'Roboto Mono',monospace;font-size:0.65rem;color:var(--text3);letter-spacing:0.03em;margin-top:0.15rem}
.app-header-badge{background:var(--accent-bg);border:1px solid var(--accent-bdr);color:var(--accent2);font-family:'Roboto Mono',monospace;font-size:0.6rem;font-weight:500;padding:0.25rem 0.75rem;border-radius:100px;letter-spacing:0.04em;text-transform:uppercase;white-space:nowrap;}
.metrics-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:0.6rem;margin:0.8rem 0}
.metric-card{background:var(--s1);border:1px solid var(--border);border-radius:var(--r2);padding:0.8rem 0.6rem;text-align:center;transition:var(--transition);box-shadow:var(--shadow-sm);position:relative;overflow:hidden;}
.metric-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:var(--r2) var(--r2) 0 0}
.metric-card.m-total::before{background:linear-gradient(90deg,#5f6368,#9aa0a6)}
.metric-card.m-unique::before{background:linear-gradient(90deg,#059669,#34d399)}
.metric-card.m-dup::before{background:linear-gradient(90deg,#f59e0b,#fbbf24)}
.metric-card.m-time::before{background:linear-gradient(90deg,#1a73e8,#4285f4)}
.metric-card:hover{transform:translateY(-2px);box-shadow:var(--shadow-lg)}
.metric-val{font-family:'Google Sans',sans-serif;font-size:1.5rem;font-weight:700;line-height:1;margin-bottom:0.3rem;letter-spacing:-0.01em}
.metric-lbl{font-family:'Roboto Mono',monospace;font-size:0.62rem;color:var(--text3);text-transform:uppercase;letter-spacing:0.08em;font-weight:500}
[data-testid="stForm"]{background:var(--s1)!important;border:1px solid var(--border)!important;border-radius:var(--r3)!important;padding:1.2rem 1.5rem!important;box-shadow:var(--shadow-md)!important;}
.sec-label{font-family:'Google Sans',sans-serif;font-size:0.72rem;font-weight:700;color:var(--text2);letter-spacing:0.08em;text-transform:uppercase;padding-bottom:0.3rem;border-bottom:2px solid var(--s3);margin:0.8rem 0 0.5rem;display:flex;align-items:center;gap:0.5rem;}
.sec-label::before{content:'';display:inline-block;width:3px;height:12px;background:linear-gradient(180deg,#f97316,#ea580c);border-radius:2px}
.upload-zone{display:grid;grid-template-columns:1fr;gap:0.6rem;margin:0.3rem 0}
.upload-zone-card{background:var(--s1);border:1.5px dashed var(--border);border-radius:var(--r2);padding:0.6rem 0.8rem;display:flex;align-items:center;gap:0.6rem;transition:var(--transition);}
.upload-zone-card:hover{border-color:var(--accent);border-style:solid;transform:translateY(-1px);box-shadow:var(--shadow-md)}
.upload-zone-icon{width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:1rem;flex-shrink:0;}
.upload-zone-icon.uz-dossier{background:var(--icon-dossier-bg);color:#f97316}
.upload-zone-icon.uz-pkl{background:var(--accent-bg);color:var(--accent2)}
.upload-zone-text{flex:1;min-width:0}
.upload-zone-title{font-family:'Google Sans',sans-serif;font-size:0.82rem;font-weight:700;color:var(--text);line-height:1.2}
.upload-zone-desc{font-size:0.7rem;color:var(--text3);line-height:1.3}
[data-testid="stFileUploader"]{background:var(--s1)!important;border:1.5px dashed var(--border)!important;border-radius:var(--r)!important;padding:0.4rem 0.6rem!important;transition:var(--transition)!important;min-height:auto!important;}
[data-testid="stFileUploader"]:hover{border-color:var(--accent)!important;border-style:solid!important;background:var(--accent-bg)!important;}
[data-testid="stFileUploader"] section{padding:0.2rem!important}
[data-testid="stFileUploader"] section>div{font-size:0.78rem!important;color:var(--text2)!important}
[data-testid="stFileUploader"] section small{font-size:0.7rem!important;color:var(--text3)!important}
[data-testid="stFileUploader"] button{background:var(--accent-bg)!important;border:1px solid var(--accent-bdr)!important;color:var(--accent2)!important;font-weight:500!important;font-size:0.75rem!important;border-radius:100px!important;padding:0.25rem 0.8rem!important;font-family:'Google Sans',sans-serif!important;transition:var(--transition)!important;}
[data-testid="stFileUploader"] button:hover{background:var(--accent)!important;color:white!important;border-color:var(--accent)!important}
[data-testid="stTextInput"] input{background:var(--s1)!important;border:1.5px solid var(--border)!important;color:var(--text)!important;border-radius:var(--r)!important;font-family:'Google Sans Text',sans-serif!important;font-size:0.9rem!important;padding:0.5rem 0.75rem!important;transition:var(--transition)!important;}
[data-testid="stTextInput"] input:focus{border-color:var(--accent)!important;box-shadow:0 0 0 3px rgba(249,115,22,0.12)!important;}
label[data-testid="stWidgetLabel"] p{font-family:'Google Sans',sans-serif!important;color:var(--text2)!important;font-size:0.82rem!important;font-weight:500!important;margin-bottom:0.15rem!important;}
.stButton>button,[data-testid="stDownloadButton"]>button{background:var(--s1)!important;border:1.5px solid var(--border)!important;color:var(--text)!important;border-radius:100px!important;font-family:'Google Sans',sans-serif!important;font-weight:500!important;font-size:0.88rem!important;transition:var(--transition)!important;padding:0.5rem 1.2rem!important;box-shadow:none!important;}
.stButton>button:hover,[data-testid="stDownloadButton"]>button:hover{border-color:var(--accent)!important;color:var(--accent2)!important;background:var(--accent-bg)!important;box-shadow:var(--shadow-sm)!important;transform:translateY(-1px)!important;}
.stButton>button[kind="primary"],[data-testid="stDownloadButton"]>button[kind="primary"]{background:var(--accent)!important;border:none!important;color:#fff!important;font-weight:500!important;font-size:0.92rem!important;padding:0.6rem 1.5rem!important;box-shadow:0 1px 3px rgba(249,115,22,0.3),0 4px 12px rgba(249,115,22,0.15)!important;letter-spacing:0.01em!important;}
.stButton>button[kind="primary"]:hover,[data-testid="stDownloadButton"]>button[kind="primary"]:hover{background:var(--accent2)!important;box-shadow:0 2px 6px rgba(234,88,12,0.35),0 8px 24px rgba(234,88,12,0.18)!important;transform:translateY(-1px)!important;color:#fff!important;}
.success-banner{background:var(--success-bg);border:1px solid var(--green-bdr);border-left:4px solid var(--green);border-radius:var(--r2);padding:0.8rem 1.2rem;margin:0.5rem 0 0.8rem;display:flex;align-items:center;gap:0.8rem;}
.success-icon{width:34px;height:34px;background:linear-gradient(135deg,#059669,#047857);border-radius:50%;display:flex;align-items:center;justify-content:center;color:white;font-size:1rem;flex-shrink:0;}
.success-title{font-family:'Google Sans',sans-serif;font-size:1rem;font-weight:700;color:var(--success-title);margin-bottom:0.1rem}
.success-sub{font-size:0.8rem;color:var(--text2)}
.auth-wrap{max-width:380px;margin:8vh auto 0;text-align:center}
.auth-icon{width:60px;height:60px;background:linear-gradient(135deg,#f97316,#ea580c);border-radius:16px;display:inline-flex;align-items:center;justify-content:center;font-size:1.6rem;color:white;margin-bottom:1rem;box-shadow:0 4px 166px rgba(249,115,22,0.3);}
.auth-title{font-family:'Google Sans',sans-serif;font-size:1.5rem;font-weight:700;color:var(--text);margin-bottom:0.3rem}
.auth-sub{font-size:0.85rem;color:var(--text3);margin-bottom:2rem}
[data-testid="stProgressBar"]>div>div{background:linear-gradient(90deg,#f97316,#fb923c,#fdba74)!important;border-radius:100px!important;height:5px!important;}
[data-testid="stDataFrame"]{border:1px solid var(--border)!important;border-radius:var(--r2)!important;box-shadow:var(--shadow-sm)!important;overflow:hidden!important;}
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:var(--s2);border-radius:3px}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px}
::-webkit-scrollbar-thumb:hover{background:var(--accent)}
.footer{font-family:'Roboto Mono',monospace;font-size:0.6rem;color:var(--text4);text-align:center;padding:0.8rem 0 0.5rem;letter-spacing:0.04em;border-top:1px solid var(--s3);margin-top:1rem;}
.stElementContainer{margin-bottom:0!important}
[data-testid="stVerticalBlock"]>div{gap:0.3rem!important}
[data-testid="stHorizontalBlock"]>div{gap:0.4rem!important}
hr{border-color:var(--s3)!important;margin:0.5rem 0!important}
.config-badge{display:inline-flex;align-items:center;gap:0.4rem;background:var(--s2);border:1px solid var(--border);border-radius:100px;padding:0.2rem 0.7rem;font-family:'Roboto Mono',monospace;font-size:0.62rem;color:var(--text3);margin-bottom:0.6rem;}
.live-panel{background:var(--s1);border:1px solid var(--border);border-radius:var(--r3);padding:1rem 1.2rem;margin:0.4rem 0 0.8rem;box-shadow:var(--shadow-md);position:relative;overflow:hidden;}
.live-panel::after{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,#f97316,#fb923c,#fdba74);}
.live-head{display:flex;align-items:center;gap:0.75rem;margin-bottom:0.75rem;}
.live-pulse{width:12px;height:12px;border-radius:50%;background:var(--accent);box-shadow:0 0 0 0 rgba(249,115,22,0.6);animation:livePulse 1.4s ease-out infinite;flex-shrink:0;}
@keyframes livePulse{0%{box-shadow:0 0 0 0 rgba(249,115,22,0.55)}70%{box-shadow:0 0 0 12px rgba(249,115,22,0)}100%{box-shadow:0 0 0 0 rgba(249,115,22,0)}}
.live-title{font-family:'Google Sans',sans-serif;font-size:1.02rem;font-weight:700;color:var(--text);line-height:1.2}
.live-sub{font-size:0.78rem;color:var(--text3);margin-top:0.15rem}
.live-metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:0.5rem;margin:0.4rem 0 0.7rem}
.live-metric{background:var(--s2);border:1px solid var(--border);border-radius:var(--r);padding:0.55rem 0.5rem;text-align:center}
.live-metric-val{font-family:'Google Sans',sans-serif;font-size:1.15rem;font-weight:700;color:var(--accent2);line-height:1.1}
.live-metric-lbl{font-family:'Roboto Mono',monospace;font-size:0.58rem;color:var(--text3);text-transform:uppercase;letter-spacing:0.06em;margin-top:0.2rem}
.step-list{display:flex;flex-direction:column;gap:0.28rem;margin:0.2rem 0 0.6rem}
.step-item{display:flex;align-items:center;gap:0.5rem;font-size:0.8rem;color:var(--text3);padding:0.22rem 0.15rem}
.step-item .dot{width:18px;height:18px;border-radius:50%;border:1.5px solid var(--border2);display:flex;align-items:center;justify-content:center;font-size:0.65rem;flex-shrink:0;background:var(--s1)}
.step-item.is-done{color:var(--green2);font-weight:500}
.step-item.is-done .dot{background:var(--green);border-color:var(--green);color:#fff}
.step-item.is-active{color:var(--accent2);font-weight:700}
.step-item.is-active .dot{border-color:var(--accent);background:var(--accent-bg);color:var(--accent2);animation:livePulse 1.4s ease-out infinite}
.live-hint{background:var(--accent-bg);border:1px solid var(--accent-bdr);color:var(--accent3);border-radius:var(--r);padding:0.55rem 0.75rem;font-size:0.78rem;line-height:1.35}
.live-detail{font-size:0.8rem;color:var(--text2);margin-top:0.45rem;font-family:'Google Sans Text',sans-serif}
.theme-bar{display:flex;justify-content:flex-end;align-items:center;margin:0 0 0.6rem;gap:0.4rem}
.theme-bar .stButton>button{padding:0.35rem 0.85rem!important;font-size:0.78rem!important}
.pkl-hint{font-size:0.78rem;color:var(--text3);margin:0.15rem 0 0.55rem;line-height:1.35}
div[data-testid="stAlert"]{border-radius:var(--r2)!important}
[data-testid="stCheckbox"] p,[data-testid="stToggle"] p{color:var(--text2)!important}
[data-baseweb="select"]>div,[data-baseweb="input"]{background:var(--s1)!important;color:var(--text)!important}
.stMarkdown,.stCaption{color:var(--text2)}
@media(max-width:768px){
    .metrics-grid{grid-template-columns:repeat(2,1fr)}
    .live-metrics{grid-template-columns:1fr 1fr 1fr}
    .app-header{flex-direction:column;text-align:center;gap:0.5rem;padding:1rem}
}
</style>
""", unsafe_allow_html=True)

def _on_theme_toggle():
    st.session_state["ui_theme"] = "dark" if st.session_state.get("theme_toggle") else "light"

def render_theme_toggle():
    if "ui_theme" not in st.session_state:
        st.session_state["ui_theme"] = _default_theme()
    if "theme_toggle" not in st.session_state:
        st.session_state["theme_toggle"] = st.session_state["ui_theme"] == "dark"
    _, col_theme = st.columns([6, 1])
    with col_theme:
        st.toggle(
            "Modo oscuro",
            key="theme_toggle",
            on_change=_on_theme_toggle,
            help="Cambia entre tema claro y oscuro. El naranja de marca se conserva.",
        )

# ======================================
# Autenticación Básica
# ======================================

PIPELINE_STEPS = [
    ("read", "Leer el Excel"),
    ("group", "Agrupar notas iguales"),
    ("ai", "Análisis IA (Tono y Sub-tema)"),
    ("themes", "Temas a partir del archivo"),
    ("export", "Generar archivo de resultado"),
]

def _fmt_elapsed(seconds: float) -> str:
    seconds = max(0, int(seconds))
    if seconds < 60:
        return f"{seconds} s"
    return f"{seconds // 60} min {seconds % 60:02d} s"

def _fmt_size(n_bytes: int) -> str:
    if not n_bytes: return ""
    mb = n_bytes / (1024 * 1024)
    if mb < 0.1: return f"{n_bytes / 1024:.0f} KB"
    return f"{mb:.1f} MB"

def _active_step(pct: int, msg: str) -> str:
    m = (msg or "").lower()
    if pct >= 100 or "listo" in m or "completad" in m:
        return "done"
    if pct >= 95 or "archivo de resultado" in m or "generando" in m:
        return "export"
    if pct >= 82 or "tema" in m:
        return "themes"
    if pct >= 22 or "analizando" in m or "ia" in m:
        return "ai"
    if pct >= 12 or "agrupando" in m:
        return "group"
    return "read"

def _render_live_html(pct, msg, elapsed, file_label, active_key):
    steps_html = []
    reached_active = False
    for key, label in PIPELINE_STEPS:
        if active_key == "done":
            cls, mark = "is-done", "✓"
        elif key == active_key:
            cls, mark = "is-active", "●"
            reached_active = True
        elif not reached_active:
            cls, mark = "is-done", "✓"
        else:
            cls, mark = "", ""
        steps_html.append(f'<div class="step-item {cls}"><span class="dot">{mark}</span>{label}</div>')
        
    file_line = f" · {html.escape(file_label)}" if file_label else ""
    title = "Análisis completado" if active_key == "done" else "Analizando las menciones"
    safe_msg = html.escape(str(msg or ""))
    
    return f"""
    <div class="live-panel">
      <div class="live-head">
        <div class="live-pulse"></div>
        <div>
          <div class="live-title">{title}</div>
          <div class="live-sub">El análisis sigue activo{file_line}. No cierres esta pestaña.</div>
        </div>
      </div>
      <div class="live-metrics">
        <div class="live-metric"><div class="live-metric-val">{int(pct)}%</div><div class="live-metric-lbl">Avance</div></div>
        <div class="live-metric"><div class="live-metric-val">{elapsed}</div><div class="live-metric-lbl">Tiempo</div></div>
        <div class="live-metric"><div class="live-metric-val">en curso</div><div class="live-metric-lbl">Estado</div></div>
      </div>
      <div class="step-list">{''.join(steps_html)}</div>
      <div class="live-hint">Las notas iguales o casi iguales se etiquetan una sola vez, así que el análisis es más rápido y consistente.</div>
      <div class="live-detail">{safe_msg}</div>
    </div>
    """

# ============================================================================
# Guarda determinista del tono: "el tema negativo no es tono negativo"
# ============================================================================
CRITICA_PAT = re.compile(
    r'(denunci|cuestion|sancion|critic|rechaz|exig|acusa|se[nñ]al|demand|investiga|irregular|'
    r'sobrecosto|corrup|incumpl|multa|reclam|responsabiliz|se le atribuye)', re.I)
VICTIMA_PAT = re.compile(
    r'(\brobo\b|roban|rob[oa]ron|hurto|atrac|asalt|accidente|\bmuert|fallec|herid|inundaci|'
    r'deslizamiento|incendio|sequ[ií]a|apag[oó]n|el ni[nñ]o|desempleo|suicid|\bprecio|alza|'
    r'aumento|protesta|delincuencia|homicid|violencia|aguas residuales en)', re.I)
BLANCO_EMPRESA = re.compile(r'(una empresa|una compa[nñ][ií]a|una firma|una industria|un frigor[ií]fico|'
                            r'una planta|un matadero|una av[ií]cola|la empresa|la compa[nñ][ií]a)', re.I)
NOMBRE_PROPIO = re.compile(r'(?<![.!?]\s)(?<![.!?])\b[A-ZÁÉÍÓÚÑ][a-záéíóúñ]{2,}')


def _tema_negativo(texto):
    return bool(VICTIMA_PAT.search(ctrl(texto)))


def _critica_dirigida(texto, brand, aliases):
    t = ctrl(texto)
    if not t:
        return False
    nt = nz(t)
    marcas = [m for m in [brand] + list(aliases or []) if m and len(str(m)) > 3]
    for m in CRITICA_PAT.finditer(t):
        cerca = nt[max(0, m.start() - 60): m.end() + 90]
        if any(nz(x) and nz(x) in cerca for x in marcas):
            return True
        ventana = t[m.end(): m.end() + 35]
        if BLANCO_EMPRESA.search(ventana) or NOMBRE_PROPIO.search(ventana):
            return True
    return False


def aplicar_guarda_tono(grupos, etiquetas, brand, aliases):
    """Baja a Neutro los Negativos que solo describen un hecho trágico, sin señalamiento dirigido."""
    corregidos = []
    for g in grupos:
        e = etiquetas.get(g['grupo'])
        if not e or e.get('tono') != 'Negativo':
            continue
        texto = '%s %s' % (g['titulo'], g.get('texto', ''))
        if _tema_negativo(texto) and not _critica_dirigida(texto, brand, aliases):
            e['tono'] = 'Neutro'
            corregidos.append(g['grupo'])
    return corregidos


def idx_of(header, nombre):
    """Índice de una columna por su nombre (el header puede traer duplicados)."""
    return header.index(nombre) if nombre in header else None


def run_analysis():
    """Analiza el archivo que quedó pendiente, con el panel de progreso de la interfaz de Grill."""
    pend = st.session_state.get('pending_analisis') or {}
    cfg, datos = pend['cfg'], pend['datos']
    panel = st.empty()
    t0 = time.time()

    def pintar(pct, msg):
        panel.markdown(_render_live_html(int(pct), msg, _fmt_elapsed(time.time() - t0),
                                         pend.get('nombre', ''), _active_step(int(pct), msg)),
                       unsafe_allow_html=True)

    def prog(h, t_, msg):
        pintar(12 + 68 * (h / max(1, t_)), msg)

    try:
        pintar(3, 'Leyendo el Excel…')
        hdr, filas, saltadas = extraer(datos, pend['hoja'], pend['col_titulo'], pend['col_texto'],
                                       pend['col_id'], pend['extras'])
        if not filas:
            st.session_state['pending_analisis'] = None
            st.error('No hay filas con texto en la columna elegida.')
            return
        pintar(8, 'Agrupando notas iguales o similares…')
        grupos, mapa = construir_grupos(filas, cfg['umbral_titulo'], cfg['umbral_cuerpo'])
        pintar(12, 'Analizando con IA…')
        etiquetas, bitacora = etiquetar_todo(cfg, grupos, prog, tam_lote=cfg['tam_lote'],
                                            max_reparaciones=cfg['max_rep'], votos=cfg['votos'])
        corregidos = aplicar_guarda_tono(grupos, etiquetas, cfg['entidad'], cfg['alias'])
        pintar(82, 'Generando la lista de Temas a partir del archivo…')
        tax = proponer_taxonomia(cfg, grupos, etiquetas, objetivo=cfg['cubos_objetivo'],
                                 progreso=lambda m: pintar(84, m))
        pintar(88, 'Asignando Temas…')
        temas, pendientes, origen = asignar_temas(cfg, grupos, etiquetas, tax, True)
        canonizar_nuevos(temas, tax)
        for _r in range(2):                  # lo que no case se resuelve por API según el texto
            faltan = [q for q in pendientes if q['grupo'] not in temas]
            if not faltan:
                break
            elegidos = elegir_cubo(cfg, faltan, tax, True)
            for q in faltan:
                if elegidos.get(q['grupo']):
                    temas[q['grupo']] = elegidos[q['grupo']]
                    origen[q['grupo']] = 'llm'
        for g in grupos:                     # último recurso determinista: nunca queda vacío
            if not temas.get(g['grupo']):
                temas[g['grupo']] = cubo_de_respaldo(
                    etiquetas.get(g['grupo'], {}).get('sub_tema', ''), g['titulo'])
                origen[g['grupo']] = 'respaldo'
        pintar(95, 'Generando el archivo de resultado…')
        out = construir_xlsx(cfg, hdr, filas, grupos, mapa, etiquetas, temas, bitacora)
        pintar(100, 'Listo')
    except Exception as e:
        st.session_state['pending_analisis'] = None
        st.error('Se interrumpió el análisis: %s' % str(e)[:400])
        st.info('Revisa la API key, el modelo o baja «Grupos por llamada» en los ajustes finos.')
        return

    st.session_state['res'] = {'cfg': cfg, 'header': hdr, 'filas': filas, 'grupos': grupos,
                               'mapa': mapa, 'etiquetas': etiquetas, 'temas': temas,
                               'bitacora': bitacora, 'tax': tax, 'origen': origen}
    st.session_state['salida'] = out.getvalue()
    st.session_state['salida_nombre'] = 'Menciones_%s_Tono_Tema_Subtemas.xlsx' % (
        re.sub(r'[^A-Za-z0-9]+', '_', cfg['entidad'])[:40].strip('_') or 'cliente')
    st.session_state['metricas'] = {'total': len(filas), 'grupos': len(grupos),
                                    'temas': len(set(temas.values())),
                                    'tiempo': _fmt_elapsed(time.time() - t0)}
    st.session_state['avisos'] = {
        'guarda': len(corregidos), 'saltadas': saltadas,
        'sin_tono': len([g for g in grupos if not etiquetas.get(g['grupo'], {}).get('tono')]),
        'respaldo': len([1 for g in grupos if origen.get(g['grupo']) == 'respaldo'])}
    st.session_state['analisis_completo'] = True
    st.session_state['pending_analisis'] = None


def main():
    st.set_page_config(page_title='Tono, Tema y Sub-tema de Menciones', page_icon='◈',
                       layout='wide', initial_sidebar_state='collapsed')
    load_custom_css()
    render_theme_toggle()
    exigir_password()
    aviso_sin_password()

    st.markdown("""
    <div class="app-header">
        <div class="app-header-icon">◈</div>
        <div class="app-header-text">
            <div class="app-header-title">Tono, Tema y Sub-tema de Menciones</div>
            <div class="app-header-version">v1.0 · Sub-tema primero, tono después · Tema por reglas + IA · Realizado por Johnathan Cortés</div>
        </div>
        <div class="app-header-badge">Clasificador + IA</div>
    </div>""", unsafe_allow_html=True)

    if st.session_state.get('pending_analisis'):
        run_analysis()
        st.rerun()

    if not st.session_state.get('analisis_completo'):
        with st.form('main_form'):
            st.markdown('<div class="sec-label">1. Sube el archivo de entrada</div>',
                        unsafe_allow_html=True)
            st.markdown("""
            <div class="upload-zone">
                <div class="upload-zone-card">
                    <div class="upload-zone-icon uz-dossier">📋</div>
                    <div class="upload-zone-text">
                        <div class="upload-zone-title">Export de monitoreo de medios</div>
                        <div class="upload-zone-desc">Sube el .xlsx del período. Luego eliges las columnas de título y de texto.</div>
                    </div>
                </div>
            </div>""", unsafe_allow_html=True)
            archivo = st.file_uploader('XLSX', type=['xlsx', 'xlsm'], label_visibility='collapsed')

            if not archivo:
                st.info('Sube el archivo del período para habilitar el análisis.')
            else:
                datos = archivo.getvalue()
                wb = load_workbook(io.BytesIO(datos), read_only=True)
                hojas = wb.sheetnames
                wb2 = load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
                hoja = st.selectbox('Hoja', hojas)
                header = [ctrl(h) for h in next(wb2[hoja].iter_rows(values_only=True))]
                c1, c2, c3 = st.columns(3)
                with c1:
                    col_titulo = st.selectbox('Columna de título', header,
                                              index=detectar(header, CAND_TITULO) or 0)
                with c2:
                    col_texto = st.selectbox('Columna de texto completo', header,
                                             index=detectar(header, CAND_TEXTO) or 0)
                with c3:
                    i_id = detectar(header, CAND_ID)
                    col_id = st.selectbox('Columna de id', ['(usar número de fila)'] + list(header),
                                          index=(i_id + 1) if i_id is not None else 0)
                extras = st.multiselect('Columnas extra a conservar en la hoja de grupos', header,
                                        [h for h in ('Medio', 'Fecha', 'Tipo de Medio') if h in header])

            st.markdown('<div class="sec-label">2. Cliente y criterio del análisis</div>',
                        unsafe_allow_html=True)
            _sec = leer_secrets()
            c_brand, c_alias = st.columns(2)
            with c_brand:
                entidad = st.text_input('Entidad, marca o persona*',
                                        placeholder='Ej: Universidad Simón Bolívar, Fenavi, Alcaldía de Sincelejo',
                                        help='El tono se mide solo sobre esta entidad, sus voceros y sus alias.')
            with c_alias:
                alias = st.text_input('Alias o formas de nombrarla (coma o punto y coma)',
                                      placeholder='Ej: Unisimón; la universidad; la alma mater',
                                      help='Variantes del nombre que deben atribuirse al cliente.')
            c_crit, c_voc = st.columns([3, 2])
            with c_crit:
                criterio = st.radio('Criterio del tono', list(CRITERIOS_TONO.keys()), index=0,
                                    help='Aspectual estricto: solo la crítica dirigida a la entidad es '
                                         'Negativo (gobiernos, alcaldías, universidades, entidades '
                                         'públicas). Favorabilidad del sector: cuenta cómo queda el '
                                         'sector aunque la marca no sea el actor (gremios, cámaras).')
            with c_voc:
                voceros = st.text_input('Vocero(s) de la entidad (opcional)',
                                        placeholder='Ej: Gonzalo Moreno; el rector',
                                        help='Personas cuyo nombre se atribuye a la entidad para el tono.')

            st.markdown('<div class="sec-label">3. Modelo</div>', unsafe_allow_html=True)
            _prov = _sec.get('proveedor') if _sec.get('proveedor') in PROVEEDORES else list(PROVEEDORES)[0]
            cp, cm = st.columns(2)
            with cp:
                proveedor = st.selectbox('Proveedor', list(PROVEEDORES.keys()),
                                         index=list(PROVEEDORES).index(_prov))
            base_def, modelo_def = PROVEEDORES[proveedor]
            with cm:
                modelo = st.text_input('Modelo', _sec.get('modelo') or modelo_def)
            base_url = st.text_input('base_url', _sec.get('base_url') or base_def)
            clave_secrets = _sec.get('api_key', '')
            if clave_secrets:
                api_key = clave_secrets
                st.caption('🔒 API key cargada desde los Secrets. No se muestra ni se envía al navegador.')
                with st.expander('Usar otra API key solo en esta sesión'):
                    tmp = st.text_input('API key temporal', value='', type='password',
                                        help='No se guarda: vive solo en esta pestaña del navegador.')
                    if tmp.strip():
                        api_key = tmp.strip()
                        st.caption('Usando la key temporal de esta sesión.')
            else:
                api_key = st.text_input('API key', value='', type='password',
                                        help='No está en los Secrets: se escribe aquí y vive solo en '
                                             'esta sesión. Recomendado: guárdala en los Secrets.')

            with st.expander('⚙ Ajustes finos del análisis (opcional)'):
                ca, cb, cc, cd = st.columns(4)
                with ca:
                    tam_lote = st.slider('Grupos por llamada', 5, 30, 10, 1,
                                         help='Con gpt-4.1-nano 10 funciona mejor.')
                with cb:
                    votos = st.slider('Verificaciones del tono por grupo', 1, 3, 2, 1,
                                      help='Cada grupo se etiqueta N veces y gana la mayoría; un empate '
                                           'cae a Neutro.')
                with cc:
                    umbral_titulo = st.slider('Similitud de titulares (%)', 75, 100,
                                              UMBRAL_TITULO_POR_DEFECTO, 1)
                with cd:
                    umbral_cuerpo = st.slider('Similitud de cuerpos (%)', 70, 100,
                                              UMBRAL_CUERPO_POR_DEFECTO, 1)
                ce, cf = st.columns(2)
                with ce:
                    cubos_objetivo = st.slider('Cubos objetivo de la lista de Temas', 8, 25, 16, 1,
                                               help='La lista se genera leyendo los hechos de este '
                                                    'archivo; este es el tamaño buscado.')
                with cf:
                    max_rep = st.slider('Máximo de reparaciones por lote', 0, 3, 2, 1)

            if st.form_submit_button('▶ Iniciar análisis', use_container_width=True, type='primary'):
                if not archivo:
                    st.error('Sube el archivo Excel del período.')
                elif not entidad.strip():
                    st.error('Indica la entidad, marca o persona: el tono se mide solo sobre ella.')
                elif not api_key.strip():
                    st.error('Falta la API key (en los Secrets o en «Usar otra API key solo en esta sesión»).')
                else:
                    st.session_state['pending_analisis'] = {
                        'datos': datos, 'nombre': archivo.name, 'hoja': hoja,
                        'col_titulo': idx_of(header, col_titulo), 'col_texto': idx_of(header, col_texto),
                        'col_id': None if str(col_id).startswith('(') else idx_of(header, col_id),
                        'extras': [idx_of(header, e) for e in extras],
                        'cfg': {'entidad': entidad.strip(),
                                'voceros': [v.strip() for v in re.split(r'[,;]', voceros) if v.strip()],
                                'alias': [a.strip() for a in re.split(r'[,\n;]', alias) if a.strip()],
                                'criterio': criterio, 'proveedor': proveedor,
                                'base_url': base_url.strip(), 'modelo': modelo.strip(),
                                'api_key': api_key.strip(), 'timeout': 120,
                                'tam_lote': int(tam_lote), 'votos': int(votos),
                                'umbral_titulo': int(umbral_titulo), 'umbral_cuerpo': int(umbral_cuerpo),
                                'cubos_objetivo': int(cubos_objetivo), 'max_rep': int(max_rep)}}
                    st.rerun()
    else:
        met = st.session_state.get('metricas') or {}
        av = st.session_state.get('avisos') or {}
        st.markdown("""
        <div class="success-banner"><div class="success-icon">✓</div>
        <div><div class="success-title">Análisis completado</div>
        <div class="success-sub">El archivo con Tono, Tema y Sub-tema por mención está listo para descargar</div></div></div>""",
                    unsafe_allow_html=True)
        avisos = []
        if av.get('guarda'):
            avisos.append('guarda del tono: %d Negativos sin señalamiento pasaron a Neutro'
                          % av['guarda'])
        if av.get('respaldo'):
            avisos.append('%d Temas asignados por respaldo determinista' % av['respaldo'])
        if av.get('sin_tono'):
            avisos.append('⚠️ %d grupos sin tono (revisa la API key o el modelo)' % av['sin_tono'])
        if avisos:
            st.info(' · '.join(avisos))
        st.markdown(f"""
        <div class="metrics-grid">
          <div class="metric-card m-total"><div class="metric-val" style="color:var(--text)">{met.get('total', 0)}</div><div class="metric-lbl">Total Registros</div></div>
          <div class="metric-card m-unique"><div class="metric-val" style="color:var(--green)">{met.get('grupos', 0)}</div><div class="metric-lbl">Hechos Únicos</div></div>
          <div class="metric-card m-dup"><div class="metric-val" style="color:var(--amber)">{met.get('temas', 0)}</div><div class="metric-lbl">Temas</div></div>
          <div class="metric-card m-time"><div class="metric-val" style="color:var(--blue)">{met.get('tiempo', '')}</div><div class="metric-lbl">Tiempo de Ejecución</div></div>
        </div>""", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        c1.download_button('⬇ Descargar Xlsx con Tono, Tema y Sub-tema',
                           data=st.session_state['salida'],
                           file_name=st.session_state['salida_nombre'],
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           use_container_width=True, type='primary')
        if c2.button('Nuevo análisis', use_container_width=True):
            ok, tema = st.session_state.get('_auth_ok'), st.session_state.get('ui_theme')
            st.session_state.clear()
            st.session_state['_auth_ok'] = ok
            if tema in ('dark', 'light'):
                st.session_state['ui_theme'] = tema
            st.rerun()

    st.markdown('<div class="footer">Tono, Tema y Sub-tema · Johnathan Cortés ©</div>',
                unsafe_allow_html=True)


if __name__ == '__main__':
    main()
