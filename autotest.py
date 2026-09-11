# -*- coding: utf-8 -*-
"""Autotest del pipeline: corre SIN API key (usa un modelo simulado).

    python autotest.py

Verifica las piezas que hacen la diferencia frente a un prompt suelto:
  1. agrupacion de notas iguales o similares
  2. validador + ciclo de reparacion (con etiquetas malas a proposito)
  3. temas: regla determinista + eleccion dentro de la lista cerrada
  4. prohibicion de "Otros" y de rotulos vacios como cubo nuevo
  5. XLSX de salida: hojas, uniformidad por grupo y columnas nuevas
"""
import io
import json
import re
import sys
import collections

from openpyxl import Workbook, load_workbook

import app

FALLOS = []


def check(nombre, cond, detalle=''):
    print(('  OK   ' if cond else '  FALLA') + '  ' + nombre + (('  -> ' + detalle) if detalle and not cond else ''))
    if not cond:
        FALLOS.append(nombre)


def dataset():
    """3 medios publican la misma nota + notas distintas + un caso que no cabe en tax_mini().

    Los cuerpos son largos a propósito: la agrupación por cuerpo exige 30 cinco-gramas como
    mínimo para evitar fusiones falsas en notas cortas.
    """
    comun = ('El presidente de la Federación Nacional de Avicultores de Colombia, Gonzalo Moreno, '
             'anunció durante el cierre del congreso del gremio que Barranquilla será la sede en 2028 '
             'del Congreso Internacional de Avicultores, un evento que reunirá a más de cuatro mil '
             'asistentes, cincuenta conferencistas y doscientos treinta patrocinadores de toda la '
             'cadena avícola del país, con una agenda centrada en tecnología, sanidad y apertura de '
             'mercados para el pollo y el huevo colombianos.')
    filas = [
        ('FENAVI realizará su congreso de 2028 en Barranquilla', comun, 'Medio A'),
        ('Barranquilla será sede del Congreso Internacional de Avicultores en 2028',
         comun + ' El gremio espera además la participación de delegaciones de Ecuador y Venezuela.', 'Medio B'),
        ('FENAVI realizará su congreso de 2028 en Barranquilla',
         comun + ' La organización confirmó que la sede se definió tras evaluar tres ciudades.', 'Medio C'),
        ('180.000 huevos y 15 hombres armados: el millonario robo a una granja avícola del Atlántico',
         'Un grupo de quince hombres armados ingresó a la granja El Brujal, en Sabanalarga, Atlántico, '
         'amordazó a los trabajadores de la bodega y huyó en cuatro furgones cargados con ciento '
         'ochenta mil huevos, según el reporte de las autoridades. La reacción de los propietarios y '
         'la denuncia oportuna de la comunidad permitieron que la Policía recuperara los vehículos '
         'horas después en el vecino departamento del Magdalena, donde fueron abandonados por los '
         'delincuentes que aún son buscados por las autoridades.', 'Medio D'),
        ('Sancionan a exsecretario de Educación por demora en el PAE',
         'La Procuraduría General de la Nación sancionó al exsecretario de Educación del departamento '
         'por las demoras acumuladas en la ejecución del Programa de Alimentación Escolar durante la '
         'vigencia anterior, luego de un proceso disciplinario que se extendió por varios meses y que '
         'incluyó testimonios de rectores y funcionarios de la secretaría, quienes advirtieron sobre '
         'los retrasos en la entrega de los complementos alimentarios a los colegios oficiales.', 'Medio E'),
        ('Crecen las inundaciones en La Mojana',
         'Las lluvias de las últimas semanas dejaron más de tres mil familias afectadas y pérdidas '
         'de cultivos en varios municipios de La Mojana, según el último reporte de gestión del '
         'riesgo. Los consejos municipales pidieron apoyo al Gobierno nacional para atender a los '
         'damnificados, mientras los organismos de socorro advirtieron que el nivel de los ríos '
         'seguirá subiendo durante los próximos días por el aumento de las precipitaciones.', 'Medio F'),
        ('Sube el precio del huevo en los mercados del país',
         'El precio del huevo registró un nuevo aumento en los mercados mayoristas del país durante '
         'la última semana, de acuerdo con el reporte de las centrales de abastos, que atribuyen el '
         'alza al costo del maíz importado y al incremento en los fletes. Los comerciantes señalaron '
         'que la tendencia podría mantenerse durante las próximas semanas si continúa la presión '
         'sobre los insumos de la cadena productiva.', 'Medio G'),
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = 'Table1'
    ws.append(['NoticiaId', 'Título', 'CuerpoEs', 'Medio', 'Fecha'])
    for i, (t, c, m) in enumerate(filas, 1):
        ws.append([1000 + i, t, c, m, '01/09/2026'])
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def stub_llm(cfg, mensajes, **kw):
    """Modelo simulado: meto errores a proposito para probar el validador."""
    prompt = ' '.join(m['content'] for m in mensajes)
    ids = [int(x) for x in re.findall(r'GRUPO id=(\d+)', prompt)]
    if 'Clasificas notas en cubos' in prompt or 'Clasificas notas de prensa en cubos' in prompt:
        out = []
        for k, i in enumerate(ids):
            out.append({'id': i, 'cubo': 'Otros' if k == 0 else 'Otros temas generales'})
        return json.dumps({'resultados': out}, ensure_ascii=False)
    if 'Corrige SOLO estos sub-temas' in prompt:
        return json.dumps({'resultados': [{'id': i, 'sub_tema': 'Cubo sin regla %d' % i, 'tono': 'Neutro'}
                                          for i in ids]}, ensure_ascii=False)
    out = []
    for k, i in enumerate(ids):
        if k == 0:
            et = ('Anuncio del congreso avícola para', 'Positivo')        # 6 palabras pero termina en prep.
        elif k == 1:
            et = ('Robo a granja avícola en Sabanalarga', 'Neutro')
        elif k == 2:
            et = ('Sancionan a exsecretario de Educación', 'Negativo')    # verbo conjugado al inicio
        elif k == 3:
            et = ('Inundaciones y pérdidas de cultivos', 'Neutro')
        else:
            et = ('Noticias generales', 'Positivo')                       # rotulo vacio
        out.append({'id': i, 'sub_tema': et[0], 'tono': et[1]})
    return json.dumps({'resultados': out}, ensure_ascii=False)


def tax_mini():
    return {
        'temas': ['Congreso y eventos', 'Seguridad y delitos', 'Control y sanciones', 'Gestión del riesgo', 'Otros'],
        'reglas': [
            {'tema': 'Seguridad y delitos', 'claves': ['robo*', 'hurto*', 'granja avicola']},
            {'tema': 'Control y sanciones', 'claves': ['sancion*', 'procuraduria*']},
            {'tema': 'Gestión del riesgo', 'claves': ['inundacion*', 'lluvias', 'afectad*']},
            {'tema': 'Congreso y eventos', 'claves': ['congreso', 'evento*']},
        ],
    }


def main():
    print('autotest — pipeline sin API key\n')
    app.llamar_llm = stub_llm
    datos = dataset()
    cfg = {'entidad': 'FENAVI', 'voceros': ['Gonzalo Moreno'], 'alias': ['Fenavi', 'el gremio avicultor'],
           'criterio': 'Aspectual estricto (recomendado)', 'proveedor': 'simulado', 'modelo': 'stub',
           'base_url': 'x', 'api_key': 'x', 'col_titulo': 'Título', 'col_texto': 'CuerpoEs'}

    print('1. lectura y agrupación')
    header = [app.ctrl(h) for h in next(load_workbook(io.BytesIO(datos), read_only=True, data_only=True)['Table1'].iter_rows(values_only=True))]
    ix = {h: j for j, h in enumerate(header)}
    hdr, filas, saltadas = app.extraer(datos, 'Table1', ix['Título'], ix['CuerpoEs'], ix['NoticiaId'],
                                       [ix['Medio'], ix['Fecha']])
    check('lee las 7 filas', len(filas) == 7, str(len(filas)))
    check('lee bien título y cuerpo (no los cruza)', filas[0]['titulo'].startswith('FENAVI realizará')
          and filas[0]['texto'].startswith('El presidente de la Federación'), filas[0]['titulo'][:40])
    grupos, mapa = app.construir_grupos(filas, 92, 85)
    g3 = [g for g in grupos if g['n'] == 3]
    check('agrupa las 3 versiones de la misma nota', len(g3) == 1, str([(g['grupo'], g['n']) for g in grupos]))
    check('total de grupos', len(grupos) == 5, str(len(grupos)))

    print('\n2. etiquetado con validador y reparación')
    etiquetas, bitacora = app.etiquetar_todo(cfg, grupos, None, tam_lote=10, max_reparaciones=2)
    check('etiquetó todos los grupos', len(etiquetas) == len(grupos), '%d/%d' % (len(etiquetas), len(grupos)))
    reparo = any('reparacion' in (b['problemas'] or '') for b in bitacora)
    check('el validador detectó etiquetas malas y pidió reparación', reparo, str(bitacora))
    duras = []
    for g in grupos:
        e = etiquetas[g['grupo']]
        pr = app.validar(e['sub_tema'], e['tono'], [g['titulo']] + g.get('titulos_alt', []) + [g['texto']])
        duras += [x for x in pr if not x.startswith('revisar_anclaje')]
    check('cero errores duros después de reparar', not duras, str(duras))
    check('tonos dentro del vocabulario', all(e['tono'] in app.TONOS for e in etiquetas.values()))

    print('\n3. temas: reglas y lista cerrada')
    temas, pend, _ = app.asignar_temas(cfg, grupos, etiquetas, tax_mini(), True)
    check('ningún grupo se llama "Otros"', not any(app.nz(t) == 'otros' for t in temas.values()), str(temas))
    check('"Otros" propuesto por el modelo fue rechazado',
          all(app.nz(t) not in app.CUBO_PROHIBIDO for t in temas.values()), str(temas))
    check('los grupos sin regla quedan pendientes y bloquean la descarga',
          len(pend) >= 1 and all(p['grupo'] not in temas for p in pend), str([p['grupo'] for p in pend]))

    print('\n4. guarda de cubos nuevos')
    malos = ['Otros', 'otros temas', 'Noticias generales', 'Actividad institucional',
             'Boletín interno de la entidad', 'Información general']
    check('rechaza rótulos vacíos y "Otros"', all(app.cubo_valido(m, tax_mini(), True) is None for m in malos),
          str([m for m in malos if app.cubo_valido(m, tax_mini(), True)]))
    check('acepta un cubo específico nuevo',
          app.cubo_valido('Trámite de pasaportes', tax_mini(), True) == 'Trámite de pasaportes')
    check('con permitir_nuevos=False solo acepta la lista',
          app.cubo_valido('Trámite de pasaportes', tax_mini(), False) is None)
    for g in grupos:
        temas.setdefault(g['grupo'], 'Congreso y eventos')

    print('\n5. XLSX de salida')
    out = app.construir_xlsx(cfg, hdr, filas, grupos, mapa, etiquetas, temas, bitacora)
    wb = load_workbook(io.BytesIO(out.getvalue()), read_only=True, data_only=True)
    check('4 hojas esperadas', wb.sheetnames == ['Menciones', 'Temas (agrupa Sub-temas)', 'Grupos', 'Resumen'],
          str(wb.sheetnames))
    rows = list(wb['Menciones'].iter_rows(values_only=True))
    H, D = list(rows[0]), rows[1:]
    i_g, i_t, i_tem, i_sub = len(H) - 5, len(H) - 4, len(H) - 3, len(H) - 2
    check('una fila por mención', len(D) == 7, str(len(D)))
    check('columnas nuevas al final', H[-6:] == ['Fila original', 'Grupo de similitud', 'Tono', 'Tema',
                                                 'Sub-tema', 'Criterio del tono'], str(H[-6:]))
    byg = collections.defaultdict(set)
    for r in D:
        byg[r[i_g]].add((r[i_t], r[i_tem], r[i_sub]))
    check('uniformidad: cada grupo con un solo (Tono, Tema, Sub-tema)',
          all(len(v) == 1 for v in byg.values()), str({k: v for k, v in byg.items() if len(v) != 1}))
    check('sin "Otros" en la columna Tema', not any(app.nz(r[i_tem]) == 'otros' for r in D))
    check('sub-temas de 3 a 7 palabras', all(3 <= len(str(r[i_sub]).split()) <= 7 for r in D),
          str([r[i_sub] for r in D if not 3 <= len(str(r[i_sub]).split()) <= 7]))

    print('\n' + ('TODO OK' if not FALLOS else 'FALLARON: %s' % FALLOS))
    return 1 if FALLOS else 0


if __name__ == '__main__':
    sys.exit(main())
