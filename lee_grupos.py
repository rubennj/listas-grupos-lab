# -*- coding: utf-8 -*-
"""
Created on Wed Jun  9 11:11:30 2021

@author: Ruben
"""

import logging
import string
from itertools import product
import pandas as pd
from pathlib import Path
from collections import deque

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(level='INFO', handlers=[
    logging.FileHandler("debug.log", mode='w'),
    logging.StreamHandler()]
)

SEMILLA_RND = 123
PATH_LISTAS = Path('listas_apolo')

asignaturas = pd.DataFrame(index=['instrumentacion', 'potencia', 'robotica','infind','automatizacion'],
                           data={'plazas_sesion': [8, 8, 20, 25, 30],
                                 'num_sesiones': [3, 2, 3, 6, 5],
                                 'horario_sesiones': [['MI09', 'MI11', 'JU09', 'JU11', 'VI11'], ['MI11', 'JU11', 'VI09'], ['MA11', 'MI09', 'MI11'],['LU09', 'LU11','MA09','MA11','MI09'],['MA09','MA11','JU09','JU11']],
                                 'num_subgrupos': [4, 7, 3, 2, 2],
                                 'semana_inicial': [3, 1, 3, 3, 5],
                                 })

grupos_grado = pd.DataFrame(index=['A302', 'A309', 'EE309'],
                            data={'limitaciones_sesion': [None, None, {'instrumentacion': 'MI11', 'potencia': 'MI11', 'robotica': 'MA11','infind':'MA09','automatizacion':'MA09'}],
                                  'prioridad_reparto': [3, 2, 1],
                                  })


def semanas_subgrupo(nombre_asignatura, subgrupo):
    semana_inicial = asignaturas.loc[nombre_asignatura, 'semana_inicial']
    offset_subgrupo = ord(subgrupo[-1]) - 65
    num_sesiones = asignaturas.loc[nombre_asignatura, 'num_sesiones']
    num_subgrupos = asignaturas.loc[nombre_asignatura, 'num_subgrupos']

    offset = semana_inicial+offset_subgrupo

    semanas_sesiones = [sem for sem in list(
        range(offset, num_subgrupos*num_sesiones+offset, num_subgrupos))]

    return semanas_sesiones


def lee_estudiantes_asignatura(asignatura):
    lista_todos_grupos_grado = []
    for _, grupo_grado in grupos_grado.iterrows():
        lista_grado = pd.read_excel(
            PATH_LISTAS / f'{asignatura.name} {grupo_grado.name}.xlsx', index_col='Email')[:-1]

        lista_grado['limitaciones_sesion_grupo_grado'] = None
        lista_grado['limitaciones_sesion_grupo_grado'] = lista_grado['limitaciones_sesion_grupo_grado'].apply(
            lambda x: grupo_grado['limitaciones_sesion'])

        lista_grado['prioridad_reparto_grupo_grado'] = grupo_grado['prioridad_reparto']

        lista_todos_grupos_grado.append(lista_grado)

    estudiantes_asignatura = pd.concat(lista_todos_grupos_grado)

    return estudiantes_asignatura


def lee_subgrupos_asignados_estudiante(subgrupos_tamaños, asignatura, lista_estudiantes_subgrupos, subgrupo_a_asignar, idx_estudiante):
    # cuenta plazas de cada subgrupo. La primera vez está vacío, por lo que se comprueba
    if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] < asignatura['plazas_sesion']:
        # si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
        if idx_estudiante in lista_estudiantes_subgrupos.index:
            subgrupos_ya_asignados = {
                subgrupo: asignatura_subgrupo for asignatura_subgrupo, subgrupo in lista_estudiantes_subgrupos.loc[idx_estudiante].items() if subgrupo != '-' and pd.notna(subgrupo)}
            logging.info('Asignatura %s - sesiones previamente asignadas %s',
                         asignatura.name, subgrupos_ya_asignados)
            sesiones_subgrupos_ya_asignados = [sesion.split(
                '-')[0] for sesion in subgrupos_ya_asignados]
        else:
            subgrupos_ya_asignados = []
            sesiones_subgrupos_ya_asignados = []

        return subgrupos_ya_asignados, sesiones_subgrupos_ya_asignados


def asigna_subgrupo_estudiante_semanas(subgrupos_ya_asignados, lista_estudiantes_asignatura, subgrupo_a_asignar, idx_estudiante):
    for subgrupo_ya_asignado in subgrupos_ya_asignados:
        asignatura_subgrupo_asignado = subgrupos_ya_asignados[subgrupo_ya_asignado].split('_')[1]
        semanas_subgrupos_ya_asignados = semanas_subgrupo(
            asignatura_subgrupo_asignado, subgrupo_ya_asignado)
        semanas_subgrupos_a_asignar = semanas_subgrupo(
            asignatura.name, subgrupo_a_asignar)
        if any(item in semanas_subgrupos_ya_asignados for item in semanas_subgrupos_a_asignar):
            logging.error(
                '%s no consigue asignar el subgrupo %s con semanas %s. Coinciden alguna semana con las del subgrupo %s de %s: %s', asignatura.name, subgrupo_a_asignar, semanas_subgrupos_a_asignar, subgrupo_ya_asignado, asignatura_subgrupo_asignado, semanas_subgrupos_ya_asignados)
        else:
            logging.info('%s asignado a %s (semanas %s)',
                         asignatura.name, subgrupo_a_asignar, semanas_subgrupos_a_asignar)
            lista_estudiantes_asignatura.at[idx_estudiante,
                                            f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
            return True


l = []
lista_total = []
lista_estudiantes_subgrupos = pd.DataFrame()
for _, asignatura in asignaturas.iterrows():

    estudiantes_asignatura = lee_estudiantes_asignatura(asignatura)

    # genera e inicializa a 0 en la lista de estudiantes de la asignatura columna con el subgrupo
    estudiantes_asignatura[f'subgrupo_{asignatura.name}'] = '-'

    num_subgrupos_sesion = asignatura['num_subgrupos']

    lista_subgrupos_asignatura = list(map(lambda x: str(x[0]) + '-' + str(x[1]), product(
        asignatura['horario_sesiones'], string.ascii_uppercase[:num_subgrupos_sesion])))

    ciclo_subgrupos_asignatura = deque(lista_subgrupos_asignatura)

    # baraja estudiantes
    estudiantes_asignatura = estudiantes_asignatura.sample(
        frac=1, random_state=SEMILLA_RND)

    # lista ordenando primero los estudiantes con mayor 'prioridad_reparto' según su grupo de grado
    lista_estudiantes_asignatura = estudiantes_asignatura.sort_values(
        by=['prioridad_reparto_grupo_grado'])

    for idx_estudiante, estudiante in lista_estudiantes_asignatura.iterrows():
        logging.info('\n\nEstudiante: %s , grupo: %s', idx_estudiante,
                      estudiante['Grupo matrícula'][-6:-1])
        # if idx_estudiante == 'miguel.aspiroz.delacalle@alumnos.upm.es':
        #     print('asdf')

        for subgrupo_a_asignar in list(ciclo_subgrupos_asignatura):
            ciclo_subgrupos_asignatura.rotate(1)
            logging.info('\n\nAsignatura %s Subgrupo: %s', asignatura.name, subgrupo_a_asignar)
            subgrupos_tamaños = lista_estudiantes_asignatura.groupby(
                f'subgrupo_{asignatura.name}').size()
            # cuenta plazas de cada subgrupo. La primera vez está vacío, por lo que se comprueba
            if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] < asignatura['plazas_sesion']:
                # si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
                subgrupos_ya_asignados, sesiones_subgrupos_ya_asignados = lee_subgrupos_asignados_estudiante(
                    subgrupos_tamaños, asignatura, lista_estudiantes_subgrupos, subgrupo_a_asignar, idx_estudiante)

                # si el subgrupo coincide (sesión+subgrupo) hay que comprobar que no coincidan las semanas
                # lista_subgrupos_ya_asignados = [subgrupo for subgrupo in subgrupos_ya_asignados]
                sesion_subgrupo_a_asignar = subgrupo_a_asignar.split('-')[0]

                if sesion_subgrupo_a_asignar in sesiones_subgrupos_ya_asignados:
                    if asigna_subgrupo_estudiante_semanas(
                        subgrupos_ya_asignados, lista_estudiantes_asignatura, subgrupo_a_asignar, idx_estudiante):
                        break
                else:
                    # estudiante sin restricciones: limitaciones_sesion está vacía
                    # estudiante con restricciones: se mira si no está vacía la lista de limitaciones_sesion y cual es su grupo horario. Si es compatible lo coge
                    if not estudiante['limitaciones_sesion_grupo_grado'] or (sesion_subgrupo_a_asignar in estudiante['limitaciones_sesion_grupo_grado'][f'{asignatura.name}']):
                        logging.info('%s asignado a %s',
                                     asignatura.name, subgrupo_a_asignar)
                        lista_estudiantes_asignatura.at[idx_estudiante,
                                                        f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
                        break
                    else:
                        logging.error(
                            '%s no consigue asignar el subgrupo %s. Hay restriccion %s', asignatura.name, subgrupo_a_asignar, estudiante['limitaciones_sesion_grupo_grado'][f'{asignatura.name}'])
                        continue
                    # si el subgrupo no es compatible, sale y busca en otro grupo
            else:
                logging.info('No hay plazas en el grupo %s',
                             subgrupo_a_asignar)

            # comprueba si no consigue asignar subgrupo
            if subgrupo_a_asignar == ciclo_subgrupos_asignatura[-1]:
                logging.error(
                    'Asignatura %s No hay subgrupo disponible para %s', asignatura.name, idx_estudiante)

    logging.error('\n\nLista estudiantes sin grupo de %s', asignatura.name)
    logging.error(
        lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-'])

    logging.info(lista_estudiantes_asignatura.groupby(
        by=f'subgrupo_{asignatura.name}').size())

    # lista_total.append(lista_estudiantes_asignatura)
    l.append(lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'])
    lista_estudiantes_subgrupos = pd.concat(l, axis=1)

# %%
# lista_total[0].join(lista_estudiantes_subgrupos, lsuffix='_').join(lista_total[1], lsuffix='-').to_excel('asdf.xlsx')
lista_estudiantes_subgrupos.to_excel('lista_subgrupos.xlsx')
