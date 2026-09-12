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
import datetime
from collections.abc import Mapping
import hashlib
import hmac
import io
import json
import re
import time
import unicodedata
import xlsxwriter

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
                       'texto': sq(cuerpo),
                       'autores': sorted({sq(str(filas[i].get('autor') or '')) for i in ix} - {''})})
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
        'El TONO se juzga SOLO con los pasajes que hablan de la entidad (o de sus voceros).',
        'El resto de la nota es contexto para el sub_tema, no para el tono. Si la nota no la',
        'menciona, el tono es Neutro.',
    ]
    return '\n'.join(lineas)


def prompt_lote(grupos_lote, candidatos, brand='', aliases=(), voceros=()):
    bloques = []
    for g in grupos_lote:
        b = ['GRUPO id=%d (%d menciones)' % (g['grupo'], g['n']),
             'TITULAR: %s' % sq(g['titulo'])[:220]]
        if g.get('titulos_alt'):
            b.append('OTROS TITULARES DEL MISMO GRUPO: %s' % ' // '.join(sq(t)[:120] for t in g['titulos_alt']))
        pasajes = _pasajes_entidad(g.get('texto', ''), g.get('titulo', ''), brand, aliases, voceros)
        if pasajes:
            b.append('LO QUE SE DICE DE LA ENTIDAD (decide el tono; nada mas cuenta): %s' % pasajes)
        else:
            b.append('LO QUE SE DICE DE LA ENTIDAD: (la nota no la menciona) -> el tono es Neutro')
        b.append('CONTEXTO DEL HECHO (sirve para el sub_tema; NO decide el tono): %s'
                 % sq(g.get('texto', ''))[:900])
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
            {'role': 'user', 'content': prompt_lote(grupos_lote, candidatos, cfg.get('entidad', ''),
                                                           cfg.get('alias') or [],
                                                           cfg.get('voceros') or [])}]
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


CAND_TITULO = ('título', 'titulo', 'title', 'titular', 'headline', 'encabezado')
CAND_TEXTO = ('cuerpoes', 'cuerpo', 'contenido', 'texto', 'body', 'text', 'nota', 'resumen - aclaracion')
CAND_ID = ('noticiaid', 'id', 'ref', 'registro', 'nro', 'codigo')


# --- nombres de columna con formato propio, igual que Grill-API ---
THOUSANDS_COLS = {'Nro. Pagina', 'Dimensión', 'Duración - Nro. Caracteres', 'Tier', 'Audiencia'}
CURRENCY_COLS = {'CPE', 'revalorización'}
ID_COLS = {'ID Noticia', 'ID duplicada', 'NoticiaId'}
HYPERLINK_LIKE = {'Link Nota', 'Link (Streaming - Imagen)'}
COLS_ANCHO_55 = {'Título', 'Resumen - Aclaracion', 'resumen corto', 'Contexto analizado'}


def _numero(valor):
    """Convierte a número el texto de una columna numérica, con la MISMA lógica que Grill-API.

    Respeta el formato colombiano: '1.234' son mil doscientos treinta y cuatro; '12,5' es doce y
    medio. Devuelve None cuando el texto no es numérico (y entonces se escribe tal cual).
    """
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return int(valor) if isinstance(valor, float) and valor.is_integer() else valor
    t = str(valor).strip()
    if not t:
        return None
    t = re.sub(r"[^\d.,\-eE]", "", t)
    if not t:
        return None
    puntos, comas = t.count('.'), t.count(',')
    if puntos > 1 and comas == 0:
        t = t.replace('.', '')
    elif comas > 1 and puntos == 0:
        t = t.replace(',', '')
    elif puntos > 0 and comas > 0:
        if t.rfind('.') > t.rfind(','):
            t = t.replace(',', '')
        else:
            t = t.replace('.', '').replace(',', '.')
    elif puntos == 1:
        partes = t.split('.')
        if len(partes[1]) == 3:
            t = t.replace('.', '')
    elif comas == 1:
        partes = t.split(',')
        if len(partes[1]) == 3:
            t = t.replace(',', '')
        else:
            t = t.replace(',', '.')
    try:
        f = float(t)
    except ValueError:
        return None
    return int(f) if f.is_integer() else f


def _contexto_analizado(titulo, texto, entidad, alias):
    """Fragmento del texto que menciona la entidad o alguno de sus alias (igual criterio que Grill)."""
    def limpio(t):
        return re.sub(r'https?://\S+', ' ', str(t or '')).strip()
    objetivos = [nz(x) for x in [entidad] + list(alias or []) if x and len(str(x)) > 2]
    oraciones = [x.strip() for x in re.split(r'(?<=[.!?\n])\s+', limpio(texto)) if x.strip()]
    bloques = []
    for i, o in enumerate(oraciones):
        if any(obj and obj in nz(o) for obj in objetivos):
            bloque = o
            if len(o.split()) < 10 and i + 1 < len(oraciones):
                bloque = '%s %s' % (o, oraciones[i + 1].strip())
            bloques.append(bloque)
        if sum(len(b) for b in bloques) > 700:
            break
    t = limpio(titulo)
    if not bloques and t and any(obj and obj in nz(t) for obj in objetivos):
        bloques = [t]
    return (' '.join(bloques)[:700] or '-') if bloques else '-'


def construir_xlsx(cfg, header, datos, grupos, mapa, etiquetas, temas, bitacora):
    """UNA sola hoja 'Resultado', con el mismo motor y formato que exporta Grill-API.

    xlsxwriter en modo streaming, encabezado en negrita, hipervínculos reales (los de Link Nota y
    Link (Streaming - Imagen) en negro sin subrayar, como Grill), fechas DD/MM/YYYY, miles y moneda
    con su formato e IDs como enteros puros. Sin hojas auxiliares.
    """
    cols = list(header) + ['Fila original', 'Grupo de similitud', 'Contexto analizado',
                           'Tono_IA', 'Tema_IA', 'Subtema_IA', 'Criterio del tono']
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {'constant_memory': True, 'strings_to_urls': False,
                                   'nan_inf_to_errors': False})
    ws = wb.add_worksheet('Resultado')
    fmt_header = wb.add_format({'bold': True})
    fmt_link = wb.add_format({'font_color': '#0563C1', 'underline': 1, 'align': 'left'})
    fmt_plain = wb.add_format({'font_color': '#000000', 'underline': False, 'align': 'left'})
    fmt_date = wb.add_format({'num_format': 'DD/MM/YYYY'})
    fmt_currency = wb.add_format({'num_format': '$#,##0'})
    fmt_thousands = wb.add_format({'num_format': '#,##0'})
    fmt_id = wb.add_format({'num_format': '0'})
    for i, c in enumerate(cols):
        if c in COLS_ANCHO_55:
            ws.set_column(i, i, 55)
        elif c in HYPERLINK_LIKE or str(c).lower().startswith('link'):
            ws.set_column(i, i, 15)
        elif c in ('Tema_IA', 'Subtema_IA'):
            ws.set_column(i, i, 28)
        elif c == 'Criterio del tono':
            ws.set_column(i, i, 30)
        elif c in ('Tono_IA', 'Grupo de similitud', 'Fila original'):
            ws.set_column(i, i, 14)
        else:
            ws.set_column(i, i, 20)
        ws.write(0, i, c, fmt_header)

    crit = {'Positivo': 'Favorable a %s (obra, logro o declaración)' % cfg['entidad'],
            'Negativo': 'Crítica o señalamiento dirigido a %s' % cfg['entidad'],
            'Neutro': 'Sin logro ni crítica dirigida a %s' % cfg['entidad']}
    for fila, r in enumerate(datos):
        gid = mapa.get(str(r.get('_id')))
        e = etiquetas.get(gid, {}) or {}
        tono = e.get('tono', '')
        ctx = _contexto_analizado(r.get('titulo', ''), r.get('texto', ''),
                                  cfg.get('entidad', ''), cfg.get('alias') or [])
        valores = [ctrl(v) for v in r['_fila']] + [r['_fila_n'], gid, ctx, tono, temas.get(gid, ''),
                                                  e.get('sub_tema', ''), crit.get(tono, '')]
        for cidx, h in enumerate(cols):
            val = valores[cidx] if cidx < len(valores) else None
            url = None
            if isinstance(val, str) and val.strip().lower().startswith('http'):
                url, cv = val.strip(), 'Link'
            elif h == 'Fecha' and val not in (None, ''):
                f = pd.to_datetime(str(val), dayfirst=True, errors='coerce')
                cv = str(val) if pd.isna(f) else f.to_pydatetime()
            elif h in ID_COLS:
                num = _numero(str(val).replace('.', '')) if val not in (None, '') else None
                cv = int(num) if isinstance(num, (int, float)) else (val if val not in (None, '') else None)
            elif h in THOUSANDS_COLS or h in CURRENCY_COLS:
                num = _numero(val)
                cv = num if num is not None else val
            else:
                cv = val

            r_excel = fila + 1
            if url:
                fmt = fmt_plain if h in HYPERLINK_LIKE else fmt_link
                ws.write_url(r_excel, cidx, url, fmt, string=str(cv or 'Link'))
            elif cv is None or cv == '':
                ws.write_blank(r_excel, cidx, None)
            elif h == 'Fecha' and isinstance(cv, datetime.datetime):
                ws.write_datetime(r_excel, cidx, cv, fmt_date)
            elif h in ID_COLS and isinstance(cv, int):
                ws.write_number(r_excel, cidx, cv, fmt_id)
            elif h in CURRENCY_COLS and isinstance(cv, (int, float)):
                ws.write_number(r_excel, cidx, cv, fmt_currency)
            elif h in THOUSANDS_COLS and isinstance(cv, (int, float)):
                ws.write_number(r_excel, cidx, cv, fmt_thousands)
            elif isinstance(cv, (int, float)) and not isinstance(cv, bool):
                ws.write(r_excel, cidx, cv)
            else:
                ws.write(r_excel, cidx, str(cv))
    wb.close()
    return buf


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


CAND_AUTOR = ('autor - conductor', 'autor', 'conductor', 'author', 'periodista')


def extraer(archivo, nombre_hoja, col_tit, col_txt, col_id, cols_extra):
    wb = load_workbook(io.BytesIO(archivo), read_only=True, data_only=True)
    ws = wb[nombre_hoja]
    it = ws.iter_rows(values_only=True)
    header = [ctrl(h) for h in next(it)]
    filas, saltadas = [], 0
    col_autor = detectar(header, CAND_AUTOR)
    for i, r in enumerate(it):
        txt = ctrl(r[col_txt]) if r[col_txt] is not None else ''
        if not txt.strip():
            saltadas += 1
            continue
        tit = ctrl(r[col_tit]) if r[col_tit] is not None else ''
        filas.append({'autor': ctrl(r[col_autor]) if col_autor is not None else '',
                      '_fila': [ctrl(v) for v in r], '_fila_n': i + 2,
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
    st.title('📰 Tono, Tema y Sub-tema de menciones')
    st.caption('Acceso restringido. Ingresa la contraseña para usar la aplicación.')
    with st.form('_login'):
        pw = st.text_input('Contraseña', type='password', key='_pw_input')
        entrar = st.form_submit_button('Entrar', type='primary')
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


def main():
    exigir_password()
    aviso_sin_password()
    st.title('📰 Tono, Tema y Sub-tema de menciones')
    st.caption('Sube el export de monitoreo, indica la marca y sus voceros, y descarga el XLSX '
               'con Tono, Tema y Sub-tema por mención. Los grupos de notas iguales comparten etiqueta.')

    with st.sidebar:
        _sec = leer_secrets()
        st.header('1. Cliente')
        entidad = st.text_input('Entidad / marca / persona', '')
        voceros = st.text_input('Vocero(s), separados por coma', '')
        alias = st.text_area('Alias y formas de nombrarla en los medios (coma o línea por alias)',
                             '', height=90,
                             help='Ej: Fenavi, Federación Nacional de Avicultores, el gremio avicultor, '
                                  'la avicultura colombiana')
        criterio = st.radio('Criterio del tono', list(CRITERIOS_TONO.keys()),
                            index=list(CRITERIOS_TONO.keys()).index(_sec['criterio'])
                            if _sec.get('criterio') in CRITERIOS_TONO else 0)

        st.header('2. Modelo')
        _prov = _sec.get('proveedor') if _sec.get('proveedor') in PROVEEDORES else list(PROVEEDORES)[0]
        proveedor = st.selectbox('Proveedor', list(PROVEEDORES.keys()), index=list(PROVEEDORES).index(_prov))
        base_def, modelo_def = PROVEEDORES[proveedor]
        base_url = st.text_input('base_url', _sec.get('base_url') or base_def)
        modelo = st.text_input('Modelo', _sec.get('modelo') or modelo_def)
        clave_secrets = _sec.get('api_key', '')
        if clave_secrets:
            # La key vive SOLO en el servidor: nunca se pasa como valor de un widget, porque
            # entonces viaja al navegador y se puede leer con las herramientas del desarrollador.
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
                                    help='No está en los Secrets: se escribe aquí y vive solo en esta '
                                         'sesión. Recomendado: guárdala en los Secrets de Streamlit.')
            if not api_key.strip():
                st.caption('Sin API key: se usará la que escribas aquí (solo esta sesión).')

        st.header('3. Agrupación y lotes')
        umbral_titulo = st.slider('Umbral de similitud de titulares (%)', 75, 100,
                                  UMBRAL_TITULO_POR_DEFECTO, 1,
                                  help='Bájalo (p. ej. 85) para fusionar una misma campaña publicada por '
                                       'muchos medios. Súbelo para separar notas parecidas pero distintas.')
        umbral_cuerpo = st.slider('Umbral de similitud de cuerpos (%)', 70, 100,
                                  UMBRAL_CUERPO_POR_DEFECTO, 1)
        votos = st.slider('Verificaciones del tono por grupo', 1, 3, 2, 1,
                          help='Cada grupo se etiqueta N veces y gana la mayoría; un empate cae a '
                               'Neutro. Con 2 se reducen los vaivenes del modelo.')
        tam_lote = st.slider('Grupos por llamada al modelo', 5, 30, 10, 1,
                             help='Con modelos pequeños (gpt-4.1-nano) 10 funciona mejor; con mini se '
                                  'puede subir a 15.')
        max_rep = st.slider('Máximo de reparaciones por lote', 0, 3, 2, 1)

        if st.session_state.get('_auth_ok'):
            st.divider()
            if st.button('Cerrar sesión'):
                st.session_state['_auth_ok'] = False
                st.session_state.pop('res', None)
                st.rerun()

    archivo = st.file_uploader('XLSX de menciones', type=['xlsx', 'xlsm'])
    if not archivo:
        st.info('Sube el archivo del período. Luego eliges las columnas y corres el análisis.')
        with st.expander('Cómo funciona por dentro (y por qué acierta)'):
            st.markdown(
                '1. **Agrupación**: las notas iguales o casi iguales se etiquetan una sola vez.\n'
                '2. **Sub-tema primero, tono después**: el resumen del hecho guía el juicio del tono.\n'
                '3. **Validador + reparación**: 3-7 palabras, sin verbo al inicio, sin rótulos vacíos; '
                'el modelo corrige lo que no pasa.\n'
                '4. **Temas**: la lista de cubos se genera desde este mismo archivo y el LLM elige '
                 'dentro de ella; lo que no case se resuelve por API.\n'
                '5. **Sin "Otros"**: la descarga se bloquea si algún grupo queda sin cubo.')
        return

    datos = archivo.getvalue()
    wb = load_workbook(io.BytesIO(datos), read_only=True)
    hojas = wb.sheetnames
    col1, col2 = st.columns([1, 3])
    with col1:
        hoja = st.selectbox('Hoja', hojas)
    wb2 = load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
    header = [ctrl(h) for h in next(wb2[hoja].iter_rows(values_only=True))]
    with st.expander('Columnas del archivo', expanded=True):
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
        extras = st.multiselect('Columnas extra a conservar en la hoja de grupos',
                                header, [h for h in ('Medio', 'Fecha', 'Tipo de Medio') if h in header])

    listo = entidad.strip() and modelo.strip() and api_key.strip()
    if not listo:
        st.warning('Completa entidad, modelo y API key en la barra lateral para correr el análisis.')
    correr = st.button('▶ Correr análisis', type='primary', disabled=not listo)

    if correr:
        cfg = {'entidad': entidad.strip(),
               # separa por coma, punto y coma o salto de linea: el usuario escribe de las tres formas
               'voceros': [v.strip() for v in re.split(r'[,;\n]', voceros) if v.strip()],
               'alias': [a.strip() for a in re.split(r'[,;\n]', alias) if a.strip()],
               'criterio': criterio, 'proveedor': proveedor, 'base_url': base_url.strip(),
               'modelo': modelo.strip(), 'api_key': api_key.strip(), 'timeout': 120,
               'col_titulo': col_titulo, 'col_texto': col_texto}
        idx = {h: j for j, h in enumerate(header)}
        barra = st.progress(0.0, 'Leyendo el archivo...')
        try:
            hdr, filas, saltadas = extraer(datos, hoja, idx[col_titulo], idx[col_texto],
                                           None if col_id.startswith('(') else idx[col_id],
                                           [idx[e] for e in extras])
            if not filas:
                st.error('No hay filas con texto en la columna elegida.')
                return
            barra.progress(0.1, 'Agrupando notas iguales o similares...')
            grupos, mapa = construir_grupos(filas, umbral_titulo, umbral_cuerpo)
            st.session_state['resumen_grupos'] = (len(filas), len(grupos), saltadas,
                                                 sum(1 for g in grupos if g['n'] > 1))
            st.info('%d filas con texto (%d sin texto, ignoradas) → **%d grupos** '
                    '(%d grupos con 2 o más menciones iguales).' %
                    (len(filas), saltadas, len(grupos), sum(1 for g in grupos if g['n'] > 1)))

            def prog(h, t, msg):
                barra.progress(0.1 + 0.7 * h / max(1, t), msg)
            etiquetas, bitacora = etiquetar_todo(cfg, grupos, prog, tam_lote=tam_lote,
                                                 max_reparaciones=max_rep, votos=votos)
            # --- guardas del tono, en orden (mismas reglas que el motor de Grill) ---
            sin_mencion = aplicar_guarda_mencion(grupos, etiquetas, cfg['entidad'], cfg['alias'],
                                                 cfg.get('voceros') or [])
            corregidos = aplicar_guarda_tono(grupos, etiquetas, cfg['entidad'], cfg['alias'])
            bajados, subidos = aplicar_guarda_actor(grupos, etiquetas, cfg['entidad'], cfg['alias'],
                                                    cfg.get('voceros') or [], cfg.get('criterio', ''))
            por_autor = aplicar_regla_autor(grupos, etiquetas, cfg.get('voceros') or [])
            st.session_state['guardas'] = {'sin_mencion': sin_mencion, 'negativos': corregidos,
                                         'sin_evidencia_actor': bajados, 'con_evidencia_actor': subidos,
                                         'por_autor': por_autor}
            st.session_state['corregidos_guarda'] = corregidos
            _av = []
            if sin_mencion:
                _av.append('%d sin mención a la entidad' % len(sin_mencion))
            if corregidos:
                _av.append('%d Negativos sin señalamiento' % len(corregidos))
            if bajados:
                _av.append('%d sin evidencia de actor' % len(bajados))
            if subidos:
                _av.append('%d con evidencia de actor' % len(subidos))
            if por_autor:
                _av.append('%d por autor vocero' % len(por_autor))
            if _av:
                st.caption('Guardas del tono: ' + ' · '.join(_av) + '.')
            barra.progress(0.82, 'Generando la lista de Temas a partir del archivo...')
            tax = proponer_taxonomia(cfg, grupos, etiquetas, objetivo=16,
                                     progreso=lambda m: barra.progress(0.82, m))
            barra.progress(0.88, 'Asignando Temas...')
            temas, pendientes, origen = asignar_temas(cfg, grupos, etiquetas, tax, True)
            canonizar_nuevos(temas, tax)
            # Grupos sin cubo: se resuelven por API según el texto, sin pedirle nada al usuario.
            for _ronda in range(2):
                faltan = [q for q in pendientes if q['grupo'] not in temas]
                if not faltan:
                    break
                elegidos = elegir_cubo(cfg, faltan, tax, True)
                for q in faltan:
                    if elegidos.get(q['grupo']):
                        temas[q['grupo']] = elegidos[q['grupo']]
                        origen[q['grupo']] = 'llm'
            for g in grupos:                      # último recurso determinista: nunca queda vacío
                if not temas.get(g['grupo']):
                    temas[g['grupo']] = cubo_de_respaldo(
                        etiquetas.get(g['grupo'], {}).get('sub_tema', ''), g['titulo'])
                    origen[g['grupo']] = 'respaldo'
            pendientes = []
        except Exception as e:
            msg = str(e)
            st.error('Se interrumpió el análisis: %s' % msg[:400])
            if '401' in msg or 'invalid_api_key' in msg.lower():
                st.info('La API key fue rechazada. Verifícala en la barra lateral (y que corresponda '
                        'al proveedor elegido).')
            elif '404' in msg or 'model' in msg.lower():
                st.info('Revisa el nombre del modelo y el base_url del proveedor.')
            elif 'JSON' in msg or 'json' in msg:
                st.info('El modelo devolvió algo que no es JSON: baja "Grupos por llamada al modelo" '
                        'o usa un modelo más grande.')
            else:
                st.info('Si el error se repite, baja el tamaño de lote y vuelve a intentar.')
            return
        barra.progress(1.0, 'Listo')
        st.session_state['res'] = {'cfg': cfg, 'header': header, 'filas': filas, 'grupos': grupos,
                                   'mapa': mapa, 'etiquetas': etiquetas, 'temas': temas,
                                   'pendientes': pendientes, 'bitacora': bitacora, 'origen': origen,
                                   'tax': tax}

    res = st.session_state.get('res')
    if not res:
        return
    etiquetas, temas, grupos = res['etiquetas'], res['temas'], res['grupos']
    cfg, tax = res['cfg'], res['tax']

    faltan_tono = [g for g in grupos if not etiquetas.get(g['grupo'], {}).get('tono')]
    if faltan_tono:
        st.warning('%d grupo(s) quedaron sin tono (revisa la API key o el modelo). Puedes descargar '
                   'igual.' % len(faltan_tono))
    out = construir_xlsx(cfg, res['header'], res['filas'], grupos, res['mapa'], etiquetas, temas,
                         res['bitacora'])
    nombre = 'Menciones_%s_Tono_Tema_Subtemas.xlsx' % re.sub(r'[^A-Za-z0-9]+', '_',
                                                             entidad.strip())[:40].strip('_')
    st.success('Listo. %d grupos, %d temas, ningún grupo sin cubo.' % (len(grupos), len(set(temas.values()))))
    st.download_button('⬇️ Descargar XLSX', out.getvalue(), file_name=nombre,
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


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



# ============================================================================
# Pasajes de la entidad y guardas del tono (mismas reglas que el motor de Grill)
# ============================================================================
def _mascara_otras_universidades(t):
    """Neutraliza nombres de OTRAS universidades para que un alias genérico no dispare por error."""
    otras = ('universidad de los andes', 'universidad del norte', 'universidad de atalaya',
             'universidad de la sabana', 'universidad del rosario', 'universidad de antioquia',
             'universidad de cartagena', 'universidad popular del cesar', 'universidad del cesar',
             'universidad san buenaventura', 'universidad de san buenaventura',
             'universidad autonoma del caribe', 'universidad de la costa', 'universidad ces',
             'universidad de concepcion', 'universidad jorge tadeo lozano', 'universidad del algarve',
             'universidad del magdalena', 'universidad de cordoba', 'universidad de sucre',
             'universidad tecnologico', 'universidad autonoma', 'universidad distrital')
    for o in otras:
        t = t.replace(o, ' otrauniversidad ')
    return re.sub(r'\buniversidad\s+(?:de|del)\s+\w+', ' otrauniversidad ', t)


def _objetivos_entidad(brand, aliases, voceros=()):
    return [nz(x) for x in [brand] + list(aliases or []) + list(voceros or []) if x and len(str(x)) > 3]


def _menciona_entidad(texto, brand, aliases, voceros=()):
    t = nz(texto)
    return bool(t) and any(o and o in t for o in _objetivos_entidad(brand, aliases, voceros))


def _pasajes_entidad(texto, titulo, brand, aliases, voceros=(), max_chars=1800):
    """Solo lo que se dice de la entidad: las oraciones que la mencionan (o a un vocero).

    El tono NO se juzga sobre la nota completa: se juzga sobre estos pasajes.
    """
    t = sq(texto)
    if not t:
        return sq(titulo)[:220]
    objetivos = _objetivos_entidad(brand, aliases, voceros)
    if not objetivos:
        return sq(t)[:max_chars]
    oraciones = [o.strip() for o in re.split(r'(?<=[.!?])\s+', t) if o.strip()]
    seleccion = []
    for i, o in enumerate(oraciones):
        if any(obj in nz(o) for obj in objetivos):
            bloque = [o]
            if len(o.split()) < 12 and i + 1 < len(oraciones):
                bloque.append(oraciones[i + 1])
            seleccion.append(' '.join(bloque))
    return ' '.join(seleccion)[:max_chars] if seleccion else ''


def _evidencia_actor(texto, brand, aliases, voceros=()):
    """True si algún pasaje muestra a la entidad organizando, como sede escogida, colaborando o hablando."""
    t = _mascara_otras_universidades(nz(texto))
    if not t:
        return False
    entes = [o for o in _objetivos_entidad(brand, aliases, voceros) if o and o != 'la universidad']
    if not entes:
        entes = [nz(brand)]
    ent = '(?:%s)' % '|'.join(re.escape(e) for e in entes if e)
    verbos = (r're[uú]ne|reuni[oó]|realiza|realiz[oó]|organiza|organiz[oó]|convoca|convoc[oó]|'
              r'recibe|recibi[oó]|ser[aá] sede|es sede|presenta|present[oó]|destaca|destac[oó]|'
              r'subraya|subray[oó]|anuncia|anunci[oó]|inaugura|inaugur[oó]|gradu[oó]|lidera|lanz[oó]|'
              r'firma|firm[oó]|aporta|aport[oó]|adelanta|ejecuta|escogida|elegida|seleccionada')
    patrones = [
        r'%s[^.]{0,30}\b(?:%s)' % (ent, verbos),
        r'(?:con el (?:acompañamiento|apoyo|respaldo)|en (?:alianza|articulaci[oó]n)|'
        r'colaboraci[oó]n|de la mano)[^.]{0,180}%s' % ent,
        r'(?:elaborad|publicad|desarrollad|realizad)\w*\s+por[^.]{0,140}%s' % ent,
        r'(?:escogida|elegida|seleccionada) como sede[^.]{0,60}%s' % ent,
        r'%s[^.]{0,60}(?:escogida|elegida|seleccionada) como sede' % ent,
        r'(?:decano|rector|director|directora|docente|investigador|investigadora|profesor|profesora|'
        r'presidente|vicepresidente)\b[^.]{0,60}\b(?:de la|de el|del)\s+%s' % ent,
    ]
    return any(re.search(p, t) for p in patrones)


def aplicar_guarda_mencion(grupos, etiquetas, brand, aliases, voceros=()):
    """Sin mención a la entidad no puede haber Positivo ni Negativo."""
    corregidos = []
    for g in grupos:
        e = etiquetas.get(g['grupo'])
        if not e or e.get('tono') not in ('Positivo', 'Negativo'):
            continue
        if not _menciona_entidad('%s %s' % (g.get('titulo', ''), g.get('texto', '')), brand, aliases, voceros):
            e['tono'] = 'Neutro'
            corregidos.append(g['grupo'])
    return corregidos


def aplicar_regla_autor(grupos, etiquetas, voceros=()):
    """Si la nota la firma un vocero de la entidad, es la entidad hablando en medios -> Positivo."""
    objetivos = [nz(v) for v in (voceros or []) if v and len(str(v)) > 5]
    if not objetivos:
        return []
    cambiados = []
    for g in grupos:
        e = etiquetas.get(g['grupo'])
        if not e or e.get('tono') == 'Positivo':
            continue
        for autor in g.get('autores') or []:
            a = nz(autor)
            if a and any(o and (o in a or a in o) for o in objetivos):
                e['tono'] = 'Positivo'
                cambiados.append(g['grupo'])
                break
    return cambiados


def aplicar_guarda_actor(grupos, etiquetas, brand, aliases, voceros=(), criterio=''):
    """Positivo exige evidencia de actor; si hay evidencia y salió Neutro, sube a Positivo.

    Solo en el criterio aspectual: en el de sector el tono depende de cómo queda el sector, no de que
    el gremio actúe.
    """
    if not str(criterio or '').lower().startswith('aspectual'):
        return [], []
    bajados, subidos = [], []
    for g in grupos:
        e = etiquetas.get(g['grupo'])
        if not e or e.get('tono') in (None, '', 'Duplicada'):
            continue
        # se revisa cada campo por separado: la normalizacion borra los puntos y la ventana
        # del patron cruzaria del titular al texto (falso Positivo por actor)
        ev = (_evidencia_actor(g.get('texto', ''), brand, aliases, voceros)
              or _evidencia_actor(g.get('titulo', ''), brand, aliases, voceros))
        if e['tono'] == 'Positivo' and not ev:
            e['tono'] = 'Neutro'
            bajados.append(g['grupo'])
        elif e['tono'] == 'Neutro' and ev:
            e['tono'] = 'Positivo'
            subidos.append(g['grupo'])
    return bajados, subidos


def aplicar_guarda_tono(grupos, etiquetas, brand, aliases):
    """Baja a Neutro los Negativos que solo describen un hecho tragico, sin señalamiento dirigido."""
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

if __name__ == '__main__':
    main()
