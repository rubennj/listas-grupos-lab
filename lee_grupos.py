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
import random

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(level='INFO', handlers=[
    logging.FileHandler("debug.log", mode='w'),
    logging.StreamHandler()]
)

SEMILLA_RND = 123
PATH_LISTAS = Path('listas_apolo')

asignaturas = pd.DataFrame(index=['instrumentacion', 'potencia', 'robotica'],
                           data={'plazas_sesion': [8, 8, 20],
                                 'num_sesiones': [3, 2, 3],
                                 'horario_sesiones': [['MI09', 'MI11', 'JU09', 'JU11', 'VI11'], ['MI11', 'JU11', 'VI09'], ['MA11', 'MI09', 'MI11']],
                                 'num_subgrupos': [4, 7, 3],
                                 'semana_inicial': [3, 3, 3]
                                 })

grupos_grado = pd.DataFrame(index=['A302', 'A309', 'EE309'],
                            data={'limitaciones': [None, None, {'instrumentacion': 'MI11', 'potencia': 'MI11', 'robotica': 'MA11'}],
                                  'prioridad_reparto': [3, 2, 1],
                                  })


def semanas_subgrupo(nombre_asignatura, subgrupo):
    # nombre_asignatura = 'potencia'
    # subgrupo = 'MA11-C'
    
    semana_inicial = asignaturas.loc[nombre_asignatura, 'semana_inicial']
    semana_subgrupo = ord(subgrupo[-1]) - 64
    num_sesiones = asignaturas.loc[nombre_asignatura, 'num_sesiones']
    
    semanas_sesiones = [sem for sem in range(num_sesiones) + semana_inicial + semana_subgrupo]

    return semanas_sesiones


l = []
lista_total = []
lista_estudiantes_subgrupos = pd.DataFrame()
for _, asignatura in asignaturas.iterrows():
    lista_todos_grupos_grado = []
    for _, grupo_grado in grupos_grado.iterrows():
        lista_grado = pd.read_excel(
            PATH_LISTAS / f'{asignatura.name} {grupo_grado.name}.xlsx', index_col='Email')[:-1]

        lista_grado['limitaciones_grupo_grado'] = None
        lista_grado['limitaciones_grupo_grado'] = lista_grado['limitaciones_grupo_grado'].apply(
            lambda x: grupo_grado['limitaciones'])

        lista_grado['prioridad_reparto_grupo_grado'] = grupo_grado['prioridad_reparto']

        lista_todos_grupos_grado.append(lista_grado)

    estudiantes_asignatura = pd.concat(lista_todos_grupos_grado)

    # genera e inicializa a 0 en la lista de estudiantes de la asignatura columna con el subgrupo
    estudiantes_asignatura[f'subgrupo_{asignatura.name}'] = '-'

    num_subgrupos_sesion = asignatura['num_subgrupos']

    lista_subgrupos_asignatura = list(map(lambda x: str(x[0]) + '-' + str(x[1]), product(
        asignatura['horario_sesiones'], string.ascii_uppercase[:num_subgrupos_sesion])))

    # baraja estudiantes
    estudiantes_asignatura = estudiantes_asignatura.sample(
        frac=1, random_state=SEMILLA_RND)

    # lista ordenando primero los estudiantes con mayor 'prioridad_reparto' según su grupo de grado
    lista_estudiantes_asignatura = estudiantes_asignatura.sort_values(
        by=['prioridad_reparto_grupo_grado'])

    for idx_correo, estudiante in lista_estudiantes_asignatura.iterrows():
        logging.info('\n\nEstudiante: %s , grupo: %s', idx_correo,
                     estudiante['Grupo matrícula'][-6:-1])
        # aleatoriza la lista de subgrupos para cada estudiante que se asigna
        # random.seed(SEMILLA_RND)
        # random.shuffle(lista_subgrupos_asignatura)
        
        for subgrupo_a_asignar in lista_subgrupos_asignatura:
            subgrupos_tamaños = lista_estudiantes_asignatura.groupby(
                f'subgrupo_{asignatura.name}').size()
            # cuenta plazas de cada subgrupo. La primera vez está vacío, por lo que se comprueba
            if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] < asignatura['plazas_sesion']:
                sesion_grupo_a_asignar = subgrupo_a_asignar.split('-')[0]
                # si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
                if idx_correo in lista_estudiantes_subgrupos.index:
                    subgrupos_ya_asignados = {
                        subgrupo:asignatura_subgrupo for asignatura_subgrupo, subgrupo in lista_estudiantes_subgrupos.loc[idx_correo].items() if subgrupo != '-' and pd.notna(subgrupo)}
                    logging.info('Asignatura %s - sesiones previamente asignadas %s',
                                 asignatura.name, subgrupos_ya_asignados)
                    sesiones_subgrupos_ya_asignados = [sesion.split(
                        '-')[0] for sesion in subgrupos_ya_asignados]
                else:
                    subgrupos_ya_asignados = []
                    sesiones_subgrupos_ya_asignados = []
                    
                # si el subgrupo coincide (sesión+subgrupo) hay que comprobar que no coincidan las semanas
                if subgrupo_a_asignar in subgrupos_ya_asignados:
                    asignatura_subgrupo_asignado = subgrupos_ya_asignados[subgrupo_a_asignar].split('_')[1]
                    semanas_subgrupos_ya_asignados = semanas_subgrupo(asignatura_subgrupo_asignado, subgrupo_a_asignar)
                    semanas_subgrupos_a_asignar = semanas_subgrupo(asignatura.name, subgrupo_a_asignar)
                    if set(semanas_subgrupos_ya_asignados).intersection(set(semanas_subgrupos_a_asignar)):
                        logging.error(
                            '%s no consigue asignar el subgrupo %s. Coinciden las semanas del subgrupo con otras anteriormente asignadas', asignatura.name, subgrupo_a_asignar)
                        continue
 
                lista_subgrupos_ya_asignados = [subgrupo for subgrupo in subgrupos_ya_asignados]
                
                if subgrupo_a_asignar not in lista_subgrupos_ya_asignados:
                    # estudiante sin restricciones
                    if estudiante['limitaciones_grupo_grado'] is None:
                        logging.info('%s asignado a %s',
                                     asignatura.name, subgrupo_a_asignar)
                        lista_estudiantes_asignatura.at[idx_correo,
                                                        f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
                        break
                    # estudiante con restricciones, se mira cual es su grupo horario. Si es compatible lo coge
                    elif sesion_grupo_a_asignar in estudiante['limitaciones_grupo_grado'][f'{asignatura.name}']:
                        logging.info('%s asignado a %s',
                                     asignatura.name, subgrupo_a_asignar)
                        lista_estudiantes_asignatura.at[idx_correo,
                                                        f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
                        break
                    else:
                        logging.error(
                            '%s no consigue asignar el subgrupo %s', asignatura.name, subgrupo_a_asignar)
                        continue
                    # si el subgrupo no es compatible, sale y busca en otro grupo
                # else:
                    # Nunca llega a este else, ya que los 2 ifs anteriores cubren todas las posibilidades
                    # continue
            else:
                logging.info('No hay plazas en el grupo %s',
                             subgrupo_a_asignar)

            # comprueba si no consigue asignar subgrupo
            if subgrupo_a_asignar == lista_subgrupos_asignatura[-1]:
                logging.error(
                    'Asignatura %s No hay subgrupo disponible para %s', asignatura.name, idx_correo)

    logging.error('\n\nLista estudiantes sin grupo de %s', asignatura.name)
    logging.error(
        lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-'])
    
    logging.info(lista_estudiantes_asignatura.groupby(by=f'subgrupo_{asignatura.name}').size())
    
    # lista_total.append(lista_estudiantes_asignatura)
    l.append(lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'])
    lista_estudiantes_subgrupos = pd.concat(l, axis=1)

# %%
# lista_total[0].join(lista_estudiantes_subgrupos, lsuffix='_').join(lista_total[1], lsuffix='-').to_excel('asdf.xlsx')
lista_estudiantes_subgrupos.to_excel('lista_subgrupos.xlsx')
