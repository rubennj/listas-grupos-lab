# -*- coding: utf-8 -*-
"""
Created on Fri Jul 30 17:56:26 2021

@author: Ruben
"""
import pandas as pd
from pathlib import Path

lista_subgrupos = pd.read_excel('lista_subgrupos.xlsx', dtype={'Nº Expediente en Centro':str})
lista_subgrupos.set_index('Nº Expediente en Centro', inplace=True)

asignaturas = ['instrumentacion', 'potencia', 'robotica', 'infind', 'automatizacion']
asignaturas_plazas = [165-160, 167-168, 147-180, 208-250, 189-240]

for index in range(len(asignaturas)):
    aux = lista_subgrupos.loc[lista_subgrupos[f'subgrupo_{asignaturas[index]}'] == '-']
    if len(aux) > (asignaturas_plazas[index] if asignaturas_plazas[index] > 0 else 0):
        print(aux)
        print(f'{asignaturas[index]} esta mal.')
    else:
        if len(aux) != 0:
            print(aux)
        print(f'{asignaturas[index]} esta bien.')