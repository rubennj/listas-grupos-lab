# -*- coding: utf-8 -*-
"""
Created on Wed Jun  9 11:11:30 2021

@author: Ruben
"""

import numpy as np
import pandas as pd
from pathlib import Path

PATH_LISTAS = Path('listas_apolo')

asignaturas = pd.DataFrame({'nombre': ['instrumentacion', 'potencia'],
                            'plazas': [8, 8],
                            'num_sesiones': [3, 2],
                            'horario_sesiones': [['MI09', 'MI11', 'JU09', 'JU11', 'VI11'], ['MI11', 'JU11', 'VI09']]
                            })

grupos = pd.DataFrame({'nombre': ['A302', 'A309', 'EE309'],
                       'limitaciones': [[None], [None], ['MA09', 'MI11', 'JU09', 'JU11', 'VI11']],
                       })

# for _, asignatura in asignaturas.iterrows():
asignatura = 'instrumentacion'

lista_grupos = []
for _, grupo in grupos.iterrows():
    lista_inst_grupo = pd.read_excel(
        PATH_LISTAS / f'instrumentacion {grupo["nombre"]}.xlsx')[:-1]

    lista_inst_grupo['limitaciones_grupo'] = None
    lista_inst_grupo['limitaciones_grupo'] = lista_inst_grupo['limitaciones_grupo'].apply(
        lambda x: grupo['limitaciones'])

    lista_grupos.append(lista_inst_grupo)

lista_asignatura = pd.concat(lista_grupos).reset_index()

num_grupos = int(round(len(lista_asignatura) /
                       asignaturas[asignaturas['nombre'] == asignatura]['plazas']))

# fijar la semilla del random permite reproducir siempre el mismo patrón aleatorio
np.random.seed(123)

lista_inst_random = np.array(lista_asignatura)
# OJO hace el shuffle in_place!
np.random.shuffle(lista_inst_random)

for _, alumno in lista_asignatura.iterrows():
    if alumno['limitaciones_grupo'][0] is not None:
        print(alumno['Email'])

# df.at['C', 'x'] = 10

# reparte alumnos directamente, sin contar con limitaciones
# a = np.array_split(lista_inst_random, num_grupos)
