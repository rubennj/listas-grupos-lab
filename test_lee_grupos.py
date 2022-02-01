# -*- coding: utf-8 -*-
"""
@author: Andrea Magro Canas

Descripción del código: Se tienen 6 opciones.
La primera se utiliza para mostrar el tamaño de los subgrupos.
La segunda para mostrar los estudiantes no asignados.
La tercera ejecuta el programa principal lee_grupos
La cuarta saca los subgrupos de un estudiante metiendo su número de matrícula
La quinta genera un html con las listas de laboratorio
La sexta finaliza el programa

"""

import pandas as pd

import lee_grupos

fin = False

while not fin:
    opcion = input('1: Mostrar excel tamaño subgrupos\n2: Mostrar excel no asignados\n3: Ejecutar asignar_grupos\n4: Escoge un estudiante\n5: Generar HTML\n6: Calendario Alumno\n7: Calendario Profesor\n8: Fin de programa\n')

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
    # Ejecuta el programa principal: lee_grupos
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
        if matricula in lista_subgrupos.index:
            for col in lista_subgrupos.columns:
                if 'subgrupo' in col:
                    dia = lista_subgrupos.loc[matricula][col]
                    if pd.notna(dia):
                        print(dia[:2], dia[2:4], col.split('_')[1])
        else:
            print('Se ha introducido mal el Nº de Matricula')
        print()
    elif opcion == '5':

        pon_nombre = input('Generar HTML con nombres (y) o generarlos con los numero de matricula (n): ')

        # Crea los html a partir del archivo lista_subgrupos, por defecto se genera con los nombres de los estudiantes
        lee_grupos.crea_html_grupos_laboratorios(pon_nombre != 'n')
        print()
    elif opcion == '6':
        # Recoge el numero de matricula deseado
        matricula = input('Dime un Nº de Matricula: ')

        # Crea un calendario anual de un alumno
        codigo_error, error = lee_grupos.crea_calendario_anual_alumno(matricula)
        
        if codigo_error != 0:
            print(error)
        print()        
    elif opcion == '7':
        # Recoge el identificador deseado
        identificador = input('Dime el identificador del profesor: ')

        # Crea un calendario anual de un profesor
        codigo_error, error = lee_grupos.crea_calendario_anual_profesor(identificador)
        
        if codigo_error != 0:
            print(error)
        print()
    elif opcion == '8':
        fin = True
    else:
        print(opcion, 'opcion incorrecta.')