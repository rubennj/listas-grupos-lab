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

import lee_grupos

fin = False

while not fin:
    opcion = input('1: Mostrar excel tamaño subgrupos\n2: Mostrar excel no asignados\n3: Ejecutar asignar_grupos\n4: Escoge un estudiante\n5: Generar PDF\n6: Fin de programa\n')

    # Muestra los alumnos asignados a cada subgrupo de cada asignatura
    if opcion == '1':
        # Lee el excel con los alumnos con los grupos ya seleccionados
        lista_subgrupos = pd.read_excel('lista_subgrupos.xlsx', dtype={'Nº Expediente en Centro':str})
        lista_subgrupos.set_index('Nº Expediente en Centro', inplace=True)

        for col in lista_subgrupos.columns:
            if 'subgrupo_' in col:
                print(lista_subgrupos.groupby(by=[col], dropna=True).size())
        print()
    # Muestra los alumnos que no se han asignado de cada asignatura
    elif opcion == '2':
        # Lee el excel con los alumnos con los grupos ya seleccionados
        lista_subgrupos = pd.read_excel('lista_subgrupos.xlsx', dtype={'Nº Expediente en Centro':str})
        lista_subgrupos.set_index('Nº Expediente en Centro', inplace=True)

        count = 0
        for col in lista_subgrupos.columns:
            if 'subgrupo_' in col:
                print()
                print(col.split('_')[1].upper())
                if lista_subgrupos[lista_subgrupos[col] == '-'].empty:
                    print('DataFrame vacio')
                else:
                    print(lista_subgrupos[lista_subgrupos[col] == '-'])
                count += len(lista_subgrupos[lista_subgrupos[col] == '-'])
        print('\nHay', count, 'alumnos sin asignar.')
        print()
    # Ejecuta el programa principal ed lee_grupos
    elif opcion == '3':
        # Si imprime un codigo de error 0 va bien 
        cod_error, error = lee_grupos.asignar_grupos()
        print('Codigo de error:', cod_error, 'Error:', error)
        if cod_error == 0:
            lee_grupos.guardar_lista_grupos()
        print()
    # Muestra a un alumno con una matricula predeterminada
    elif opcion == '4':
        # Lee el excel con los alumnos con los grupos ya seleccionados
        lista_subgrupos = pd.read_excel('lista_subgrupos.xlsx', dtype={'Nº Expediente en Centro':str})
        lista_subgrupos.set_index('Nº Expediente en Centro', inplace=True)

        matricula = input('Dime un Nº de Matricula:')

        # Muestra por pantalla las asignaturas y el horario del alumno
        try:
            for col in lista_subgrupos.columns:
                if 'subgrupo' in col:
                    dia = lista_subgrupos.loc[matricula][col]
                    if pd.notna(dia):
                        print(dia[:2], dia[2:4], col.split('_')[1])
        except KeyError:
            print('Se ha introducido mal el Nº de Matricula')
        print()
    elif opcion == '5':
        # Crea los pdf a partir del archivo lista_subgrupos
        lee_grupos.crea_pdf_grupos_laboratorios()
        print()
    elif opcion == '6':
        fin = True
    else:
        print(opcion, 'opcion incorrecta.')