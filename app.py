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

CRITERIOS_TONO = {
    'Aspectual estricto (recomendado)': (
        "El tono mide SOLO lo que se dice de la entidad, de su vocero o de sus funcionarios.\n"
        "- Positivo: la entidad o su vocero es sujeto de un hecho favorable (obra entregada, avance,\n"
        "  beneficio para la comunidad, programa, reconocimiento, cifra buena, declaracion que los deja bien).\n"
        "- Negativo: existe critica, reclamo, sancion, denuncia o evaluacion negativa DIRIGIDA a la entidad,\n"
        "  a su administracion o a sus funcionarios.\n"
        "- Neutro: todo lo demas. Incluye hechos malos sin responsable institucional (inundaciones,\n"
        "  homicidios, accidentes, robos, alzas de precios), la cobertura de OTRA entidad del mismo\n"
        "  territorio, y los casos en que el vocero denuncia a un tercero.\n"
        "Ante duda entre Positivo y Neutro, o entre Negativo y Neutro, elige Neutro."
    ),
    'Favorabilidad del sector (para gremios)': (
        "El tono mide como queda parado el sector o la marca en la nota, aunque la entidad no sea el actor.\n"
        "- Positivo: la nota favorece al sector o a sus productos: promocion de consumo, crecimiento,\n"
        "  mercados nuevos o exportaciones, congresos y eventos del gremio, campañas, reconocimientos, y\n"
        "  tambien los anuncios de planes o medidas oficiales que benefician al sector (aunque el actor\n"
        "  sea un ministerio o el Gobierno).\n"
        "- Negativo: la nota daña al sector por un hecho ATRIBUIBLE a el o por una critica dirigida: una\n"
        "  denuncia o sancion contra el gremio, su vocero o una empresa del sector, contaminacion, malas\n"
        "  practicas, incumplimientos o escandalos que afectan su imagen.\n"
        "- Neutro: hechos adversos SIN responsable del sector ni institucional (robos, hurtos, delitos,\n"
        "  accidentes, incendios, inundaciones, clima), datos economicos, precios, y politica nacional sin\n"
        "  relacion con el sector. Un robo a una granja NO es Negativo: es una victima, no una falta.\n"
        "Ante duda, elige Neutro."
    ),
}

REGLAS_SUBTEMA = (
    "El Sub-tema resume EL HECHO de esa nota, no la categoria a la que pertenece, en una frase nominal\n"
    "de 3 a 5 palabras (maximo 7), coherente y especifica.\n"
    "- No empieces con verbo conjugado (nada de 'Entregan', 'Anuncian', 'Avanza', 'Denuncian').\n"
    "  Si empiezas con un sustantivo de accion esta bien: 'Entrega del parque', 'Anuncio de inversiones'.\n"
    "- No termines en preposicion o nexo.\n"
    "- No uses dos puntos, punto y coma, barra vertical, comillas ni guiones largos.\n"
    "- No copies el titular ni recortes una frase del texto: sintetiza el hecho.\n"
    "- Prohibido rotulos vacios: 'noticias generales', 'gestion institucional', 'varios', 'informacion'.\n"
    "- No repitas el nombre de la entidad, del gremio ni del medio, ni uses etiquetas de categoria.\n"
    "  MAL: 'Exportaciones y mercados internacionales de Fenavi' (es un tema, no un hecho),\n"
    "       'Transformacion digital en el sector avicola' (generico).\n"
    "  BIEN: 'Exportacion de pollo a Estados Unidos', 'Planes de IA en el congreso avicola'.\n"
    "- Incluye el actor o el lugar cuando son lo que distingue el hecho: no 'Visita internacional'\n"
    "  sino 'Visita de la embajadora de Australia'.\n"
    "- Si el hecho ya aparece en la lista CANDIDATOS, reutiliza EXACTAMENTE ese texto (misma mayuscula y\n"
    "  mismas palabras). Nunca crees una variante nueva de un hecho que ya tiene sub-tema."
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
        'EJEMPLOS REALES YA ETIQUETADOS (imitalos en criterio, brevedad y mayusculas)',
    ]
    for e in EJEMPLOS:
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
             'Un sub-tema valido tiene 3 a 7 palabras en frase nominal, no empieza con verbo conjugado,\n'
             'no termina en preposicion y no lleva marcadores ni rotulos vacios.\n\n'
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


def etiquetar_todo(cfg, grupos, progreso=None, tam_lote=15, max_reparaciones=2):
    """Etiqueta todos los grupos: lotes -> validacion -> reparacion. Devuelve etiquetas y bitacora."""
    etiquetas, bitacora = {}, []
    candidatos = []
    total = len(grupos)
    hechos = 0
    for i in range(0, total, tam_lote):
        lote = grupos[i:i + tam_lote]
        labels = etiquetar_lote(cfg, lote, candidatos)
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
        bloques.append('GRUPO id=%d\nSUB-TEMA: %s\nTITULAR: %s' % (p['grupo'], p['sub_tema'], sq(p['titulo'])[:180]))
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
            pendientes.append({'grupo': g['grupo'], 'sub_tema': e.get('sub_tema', ''), 'titulo': g['titulo']})
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
        api_key = st.text_input('API key', value=_sec.get('api_key', ''), type='password',
                                help='Si la guardas en los Secrets de Streamlit (llm_api_key), '
                                     'aparece aquí ya cargada.')
        if _sec.get('api_key'):
            st.caption('Usando la API key de los Secrets.')

        st.header('3. Agrupación y lotes')
        umbral_titulo = st.slider('Umbral de similitud de titulares (%)', 75, 100,
                                  UMBRAL_TITULO_POR_DEFECTO, 1,
                                  help='Bájalo (p. ej. 85) para fusionar una misma campaña publicada por '
                                       'muchos medios. Súbelo para separar notas parecidas pero distintas.')
        umbral_cuerpo = st.slider('Umbral de similitud de cuerpos (%)', 70, 100,
                                  UMBRAL_CUERPO_POR_DEFECTO, 1)
        tam_lote = st.slider('Grupos por llamada al modelo', 5, 30, 15, 1)
        max_rep = st.slider('Máximo de reparaciones por lote', 0, 3, 2, 1)

        st.header('4. Temas')
        tax_nombre = st.selectbox('Lista de Temas', ['Gobierno territorial (21 cubos)',
                                                     'Gremio o sector (16 cubos)'], index=0)
        permitir_nuevos = st.checkbox('Permitir que la IA cree cubos nuevos específicos', True,
                                      help='Si no hay cubo que sirva, la IA puede proponer uno nuevo '
                                           'de 2 a 5 palabras. Nunca puede escribir "Otros".')
        tax_json = st.text_area('Editar la lista de Temas (JSON)',
                                json.dumps(TAX_GOBIERNO if tax_nombre.startswith('Gobierno') else TAX_GREMIO,
                                           ensure_ascii=False, indent=1), height=150)
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
                '4. **Temas por reglas**: el LLM no inventa el Tema; elige dentro de una lista cerrada.\n'
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
        try:
            tax = json.loads(tax_json)
            assert 'temas' in tax and 'reglas' in tax
        except Exception as e:
            st.error('La lista de Temas (JSON) no es válida: %s' % e)
            return
        cfg = {'entidad': entidad.strip(),
               'voceros': [v.strip() for v in voceros.split(',') if v.strip()],
               'alias': [a.strip() for a in re.split(r'[,\n]', alias) if a.strip()],
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
                                                 max_reparaciones=max_rep)
            barra.progress(0.85, 'Asignando Temas...')
            temas, pendientes, origen = asignar_temas(cfg, grupos, etiquetas, tax, permitir_nuevos)
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

    st.subheader('Auditoría')
    ct = collections.Counter(e.get('tono') or 'SIN TONO' for e in etiquetas.values())
    cm = collections.Counter(temas.get(g['grupo'], 'SIN CUBO') for g in grupos)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric('Grupos', len(grupos))
    k2.metric('Positivo / Neutro / Negativo',
              '%d / %d / %d' % (ct.get('Positivo', 0), ct.get('Neutro', 0), ct.get('Negativo', 0)))
    k3.metric('Sub-temas distintos', len(set(e.get('sub_tema') for e in etiquetas.values())))
    k4.metric('Cubos usados', len(cm))

    tab = pd.DataFrame([{'Grupo': g['grupo'], 'Menciones': g['n'], 'Medio(s)': '',
                         'Título': g['titulo'][:110], 'Tono': etiquetas.get(g['grupo'], {}).get('tono', ''),
                         'Tema': temas.get(g['grupo'], 'SIN ASIGNAR'),
                         'Sub-tema': etiquetas.get(g['grupo'], {}).get('sub_tema', ''),
                         'Origen del tema': res['origen'].get(g['grupo'], '')} for g in grupos])
    st.dataframe(tab, use_container_width=True, height=430)
    with st.expander('Temas (cuántos sub-temas agrupa cada cubo)'):
        st.dataframe(pd.DataFrame([{'Tema': t, 'Grupos': len([1 for g in grupos if temas.get(g['grupo']) == t]),
                                    'Menciones': sum(g['n'] for g in grupos if temas.get(g['grupo']) == t),
                                    'Sub-temas': ' | '.join(sorted(set(etiquetas.get(g['grupo'], {}).get('sub_tema', '') for g in grupos if temas.get(g['grupo']) == t)))}
                                   for t in sorted(set(temas.values()))]), use_container_width=True)
    faltan_tono = [g for g in grupos if not etiquetas.get(g['grupo'], {}).get('tono')]
    if faltan_tono:
        st.error('Hay %d grupos sin tono (revisa la API key, el modelo o el tamaño de lote). '
                 'No descargues todavía.' % len(faltan_tono))
    if res['bitacora']:
        with st.expander('Control de calidad (%d avisos de reparación)' % len(res['bitacora'])):
            st.dataframe(pd.DataFrame(res['bitacora'][:200]), use_container_width=True)

    # ---- resolución obligatoria de grupos sin cubo: NUNCA "Otros" ----
    pend = [g for g in grupos if not temas.get(g['grupo'])]
    if pend:
        st.subheader('⚠️ Grupos sin cubo de Tema')
        st.caption('Ningún grupo puede quedar en "Otros". Asígnale un cubo existente o crea uno nuevo '
                   'específico (2 a 5 palabras).')
        opciones = [t for t in tax['temas'] if nz(t) not in CUBO_PROHIBIDO] + ['➕ Cubo nuevo...']
        with st.form('resolver'):
            nuevas = {}
            for g in pend:
                c1, c2, c3 = st.columns([1, 2, 2])
                c1.write('G%d (%d)' % (g['grupo'], g['n']))
                c2.write(str(g['titulo'])[:90] or etiquetas.get(g['grupo'], {}).get('sub_tema', ''))
                sel = c3.selectbox('Tema', opciones, key='sel_%d' % g['grupo'], label_visibility='collapsed')
                if sel == '➕ Cubo nuevo...':
                    nuevas[g['grupo']] = c3.text_input('Nuevo cubo', key='new_%d' % g['grupo'],
                                                       label_visibility='collapsed',
                                                       placeholder='Nombre del cubo nuevo')
                else:
                    nuevas[g['grupo']] = sel
            if st.form_submit_button('Aplicar y poder descargar'):
                rechazados = []
                for gid, t in nuevas.items():
                    v = cubo_valido(t, tax, permitir_nuevos=True)
                    if v:
                        temas[gid] = v
                    else:
                        rechazados.append(str(t)[:40])
                if rechazados:
                    st.warning('No se aceptaron (un Tema debe ser específico, sin "Otros" ni rótulos '
                               'vacíos): %s' % ', '.join(rechazados))
                st.rerun()
        return

    out = construir_xlsx(cfg, res['header'], res['filas'], grupos, res['mapa'], etiquetas, temas,
                         res['bitacora'])
    nombre = 'Menciones_%s_Tono_Tema_Subtemas.xlsx' % re.sub(r'[^A-Za-z0-9]+', '_',
                                                             entidad.strip())[:40].strip('_')
    st.success('Listo. %d grupos, %d temas, ningún grupo sin cubo.' % (len(grupos), len(cm)))
    st.download_button('⬇️ Descargar XLSX', out.getvalue(), file_name=nombre,
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    main()
