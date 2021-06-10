# -*- coding: utf-8 -*-
"""
Created on Wed Jun  9 11:11:30 2021

@author: Ruben
"""
import string
from math import ceil, floor
from itertools import product
import pandas as pd
from pathlib import Path

PATH_LISTAS = Path('listas_apolo')

asignaturas = pd.DataFrame({'nombre': ['instrumentacion', 'potencia'],
                            'plazas_sesion': [8, 8],
                            'num_sesiones': [3, 2],
                            'horario_sesiones': [['MI09', 'MI11', 'JU09', 'JU11', 'VI11'], ['MI11', 'JU11', 'VI09']]
                            })

grados = pd.DataFrame({'nombre': ['A302', 'A309', 'EE309'],
                       'limitaciones': [None, None, ['MA09', 'MI11', 'JU09', 'JU11', 'VI11']],
                       })

# asignatura = asignaturas.loc[0]
l = []
lista_total = []
lista_estudiantes_subgrupos = pd.DataFrame()
for _, asignatura in asignaturas.iterrows():
    lista_todos_grados = []
    for _, grado in grados.iterrows():
        lista_grado = pd.read_excel(
            PATH_LISTAS / f'{asignatura["nombre"]} {grado["nombre"]}.xlsx', index_col='Email')[:-1]
    
        lista_grado['limitaciones_grado'] = None
        lista_grado['limitaciones_grado'] = lista_grado['limitaciones_grado'].apply(
            lambda x: grado['limitaciones'])
    
        lista_todos_grados.append(lista_grado)
    
    estudiantes_asignatura = pd.concat(lista_todos_grados)
    
    # genera en la lista de estudiantes de la asignatura columna con el subgrupo
    estudiantes_asignatura[f'subgrupo_{asignatura["nombre"]}'] = None
    
    num_grupos_asignatura = int(ceil(len(estudiantes_asignatura) / asignatura['plazas_sesion']))
    num_subgrupos_dia = int(floor(num_grupos_asignatura / len(asignatura['horario_sesiones'])))
    # OJO, se redondea a la baja para no crear subgrupos de más, pero podrían quedar alumnos sin subgrupo!
    
    lista_subgrupos_asignatura = list(map(lambda x: str(x[0]) + '-' + str(x[1]), product(asignatura['horario_sesiones'], string.ascii_uppercase[:num_subgrupos_dia])))
    
    subgrupos_asignatura = pd.DataFrame(lista_subgrupos_asignatura, columns=['nombre'])
    subgrupos_asignatura['plazas_llenas'] = 0
    
    # baraja estudiantes
    estudiantes_asignatura = estudiantes_asignatura.sample(frac=1, random_state=123)
    
    # lista ordenando primero los estudiantes con restricciones
    lista_estudiantes_asignatura = estudiantes_asignatura.sort_values(by=['limitaciones_grado'])
    
    for idx_correo, estudiante in lista_estudiantes_asignatura.iterrows():
        for subgrupo_a_asignar in lista_subgrupos_asignatura:
            subgrupos_tamaños = lista_estudiantes_asignatura.groupby(f'subgrupo_{asignatura["nombre"]}').size()
            # cuenta plazas de cada subgrupo. La primera vez está vacío, por lo que se comprueba
            if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] < asignatura['plazas_sesion']:
                sesion_grupo_a_asignar = subgrupo_a_asignar.split('-')[0]
                # comprueba si la sesión ya ha sido asignada al estudiante anteriormente en otra asignatura
                if idx_correo in lista_estudiantes_subgrupos.index:
                    sesiones_ya_asignadas = [sesion.split('-')[0] for sesion in lista_estudiantes_subgrupos.loc[idx_correo].values if isinstance(sesion, str)]
                    print('Sesiones YA asignadas', sesiones_ya_asignadas)
                    # if idx_correo == 'd.hsanchez@alumnos.upm.es':
                        # raise SystemError
                    if sesion_grupo_a_asignar not in sesiones_ya_asignadas:
                        lista_estudiantes_asignatura.at[idx_correo, f'subgrupo_{asignatura["nombre"]}'] = subgrupo_a_asignar
                        break
                    else:
                    # si ya la sesión ya ha sido asignada en otra asignatura, sale y busca en otro grupo
                        continue
                # estudiante sin restricciones
                if estudiante['limitaciones_grado'] is None:
                    print(asignatura["nombre"], idx_correo, subgrupo_a_asignar)
                    lista_estudiantes_asignatura.at[idx_correo, f'subgrupo_{asignatura["nombre"]}'] = subgrupo_a_asignar
                    break
                # estudiante con restricciones, se mira cual es su grupo horario. Si es compatible lo coge
                elif sesion_grupo_a_asignar in estudiante['limitaciones_grado']:
                    print(asignatura["nombre"], idx_correo, subgrupo_a_asignar)
                    lista_estudiantes_asignatura.at[idx_correo, f'subgrupo_{asignatura["nombre"]}'] = subgrupo_a_asignar
                    break
                # si el subgrupo no es compatible no lo asigna y sigue buscando

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
lista_estudiantes_subgrupos.to_excel('lista_subgrupos_inst_pot.xlsx')
