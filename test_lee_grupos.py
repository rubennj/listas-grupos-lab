# -*- coding: utf-8 -*-
"""
@author: Andrea Magro Canas

Descripción del código: Lee lista_subgrupos.xlsx y mira los alumnos que tienen "-" en los subgrupos, es decir,
los estudiantes que no están asignados. Después se imprime si la asignatura está bien o mal distribuida en función
del número de plazas por asignatura. 
NOTA: Se ha tenido en cuenta que en ciertas asignaturas el número de estudiantes es mayor que el número de plazas.
Que estos estudiantes no se asignen no significa que la distribución esté mal siempre y cuando se completen todas las plazas.

"""
import pandas as pd
from pathlib import Path

# Lee el excel con los alumnos con los grupos ya seleccionados
lista_subgrupos = pd.read_excel('lista_subgrupos.xlsx', dtype={'Nº Expediente en Centro':str})
lista_subgrupos.set_index('Nº Expediente en Centro', inplace=True)

# Asignaturas
asignaturas = ['instrumentacion', 'potencia', 'robotica', 'infind', 'automatizacion']
# "Plazas" = nº Alumnos - nº de plazas disponibles en cada asignatura (Se ha calculado a mano)
asignaturas_plazas = [165-160, 167-168, 147-180, 208-250, 189-240]

# Recorre las asignaturas
for index in range(len(asignaturas)):
    # Guarda los alumnos sin plaza
    aux = lista_subgrupos.loc[lista_subgrupos[f'subgrupo_{asignaturas[index]}'] == '-']
    # Si hay mas alumnos sin plaza que "Plazas" en los subgrupos entra
    if len(aux) > (asignaturas_plazas[index] if asignaturas_plazas[index] > 0 else 0):
        print(aux)
        print(f'{asignaturas[index]} esta mal.')
    else:
        # Muestra los alumnos que se han quedado sin plaza porque las plazas en la asignatura están llenas
        if len(aux) != 0:
            print(aux)
        print(f'{asignaturas[index]} esta bien.')