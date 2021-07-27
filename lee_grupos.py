# -*- coding: utf-8 -*-
"""
Created on Wed Jun  9 11:11:30 2021

@author: Ruben
"""
import string
from math import ceil
from itertools import product
import pandas as pd
from pathlib import Path
import random

PATH_LISTAS = Path('listas_apolo')

asignaturas = pd.DataFrame({'nombre': ['instrumentacion', 'potencia', 'robotica'],
                            'plazas_sesion': [8, 8, 20],
                            'num_sesiones': [3, 2, 3],
                            'horario_sesiones': [['MI09', 'MI11', 'JU09', 'JU11', 'VI11'], ['MI11', 'JU11', 'VI09'], ['MA11', 'MI09', 'MI11']]
                            })

grupos_grado = pd.DataFrame({'nombre': ['A302', 'A309', 'EE309'],
                       'limitaciones': [None, None, 'MI11'],
                       })

# asignatura = asignaturas.loc[0]
l = []
lista_total = []
lista_estudiantes_subgrupos = pd.DataFrame()
for _, asignatura in asignaturas.iterrows():
    lista_todos_grupos_grado = []
    for _, grupo_grado in grupos_grado.iterrows():
        lista_grado = pd.read_excel(
            PATH_LISTAS / f'{asignatura["nombre"]} {grupo_grado["nombre"]}.xlsx', index_col='Email')[:-1]
    
        lista_grado['limitaciones_grado'] = None
        lista_grado['limitaciones_grado'] = lista_grado['limitaciones_grado'].apply(
            lambda x: grupo_grado['limitaciones'])
    
        lista_todos_grupos_grado.append(lista_grado)
    
    estudiantes_asignatura = pd.concat(lista_todos_grupos_grado)
    
    # genera e inicializa a 0 en la lista de estudiantes de la asignatura columna con el subgrupo
    estudiantes_asignatura[f'subgrupo_{asignatura["nombre"]}'] = '-'
    
    num_grupos_asignatura = int(ceil(len(estudiantes_asignatura) / asignatura['plazas_sesion']))
    num_subgrupos_sesion = int(ceil(num_grupos_asignatura / len(asignatura['horario_sesiones'])))
    # OJO, se redondea a la baja para no crear subgrupos de más, pero podrían quedar alumnos sin subgrupo!
    
    lista_subgrupos_asignatura = list(map(lambda x: str(x[0]) + '-' + str(x[1]), product(asignatura['horario_sesiones'], string.ascii_uppercase[:num_subgrupos_sesion])))
    
    subgrupos_asignatura = pd.DataFrame(lista_subgrupos_asignatura, columns=['nombre'])
    
    # baraja estudiantes
    estudiantes_asignatura = estudiantes_asignatura.sample(frac=1, random_state=123)
    
    # lista ordenando primero los estudiantes con restricciones para priorizar el reparto
    lista_estudiantes_asignatura = estudiantes_asignatura.sort_values(by=['limitaciones_grado'])
    
    for idx_correo, estudiante in lista_estudiantes_asignatura.iterrows():
        # aleatoriza la lista de subgrupos para cada estudiante que se asigna
        random.seed(123)
        random.shuffle(lista_subgrupos_asignatura)
        for subgrupo_a_asignar in lista_subgrupos_asignatura:
            subgrupos_tamaños = lista_estudiantes_asignatura.groupby(f'subgrupo_{asignatura["nombre"]}').size()
            # cuenta plazas de cada subgrupo. La primera vez está vacío, por lo que se comprueba
            if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] < asignatura['plazas_sesion']:
                sesion_grupo_a_asignar = subgrupo_a_asignar.split('-')[0]
                # si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
                if idx_correo in lista_estudiantes_subgrupos.index:
                    sesiones_ya_asignadas = [subgrupo.split('-')[0] for subgrupo in lista_estudiantes_subgrupos.loc[idx_correo] if subgrupo != '-' and pd.notna(subgrupo)]
                    print('Sesiones YA asignadas', sesiones_ya_asignadas)
                else:
                    sesiones_ya_asignadas = []
                if sesion_grupo_a_asignar not in sesiones_ya_asignadas:
                    # estudiante sin restricciones
                    if estudiante['limitaciones_grado'] is None:
                        print(asignatura["nombre"], idx_correo, 'asignado a', subgrupo_a_asignar)
                        lista_estudiantes_asignatura.at[idx_correo, f'subgrupo_{asignatura["nombre"]}'] = subgrupo_a_asignar
                        break
                    # estudiante con restricciones, se mira cual es su grupo horario. Si es compatible lo coge
                    elif sesion_grupo_a_asignar in estudiante['limitaciones_grado']:
                        print(asignatura["nombre"], idx_correo, 'asignado a', subgrupo_a_asignar)
                        lista_estudiantes_asignatura.at[idx_correo, f'subgrupo_{asignatura["nombre"]}'] = subgrupo_a_asignar
                        break
                    else:
                        continue
                    # si el subgrupo no es compatible, sale y busca en otro grupo
                else:
                # si ya la sesión ya ha sido asignada en otra asignatura, sale y busca en otro grupo
                    continue
            else:
                print(asignatura["nombre"], idx_correo, 'No hay plazas en el grupo', subgrupo_a_asignar)
 
            # comprueba si no consigue asignar subgrupo
            # if subgrupo_a_asignar == lista_subgrupos_asignatura[-1]:
                # raise SystemError('No hay subgrupo disponible, los siguientes estudiantes no están asignados:', estudiantes_asignatura.loc[idx_correo:])
    
    lista_total.append(lista_estudiantes_asignatura)
    l.append(lista_estudiantes_asignatura[f'subgrupo_{asignatura["nombre"]}'])
    lista_estudiantes_subgrupos = pd.concat(l, axis=1)

#%%

# lista_total[0].join(lista_estudiantes_subgrupos, lsuffix='_').join(lista_total[1], lsuffix='-').to_excel('asdf.xlsx')
lista_estudiantes_subgrupos.to_excel('lista_subgrupos.xlsx')
