# -*- coding: utf-8 -*-
"""
@author: Andrea Magro Canas

Descripción del código: Algoritmo de asignación de estudiantes en grupos de laboratorio.

"""

import logging
import string
from itertools import product
import pandas as pd
from pathlib import Path
import json
import unidecode
import os

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(level='INFO', handlers=[
    logging.FileHandler("debug.log", mode='w+'),
    # logging.StreamHandler() # Comentado no me sale en la terminal, pero se me sigue guardando en el archivo
    ] 
)

# Pruebas
SEMILLA_RND = 123
# Path de los archivos
PATH_LISTAS = Path('listas_apolo')
PATH_PDF = Path('pdf')

# Variables globales
l = list()
lista_estudiantes_subgrupos = pd.DataFrame()

# Codigo de error
cod_error = 0

# Crea a mano una tabla con todos los subgrupos
grupos_grado = pd.DataFrame(index=['DM306', 'Q308', 'D307', 'A207', 'E208', 'A309', 'M306', 'A408', 'A204', 'E205', 'A302', 'M301', 'A404', 'EE208', 'EE309', 'EE403', 'EE507'],
                            data={'limitaciones_sesion' : [{}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}],
                                'prioridad_reparto' : [4, 4, 4, 3, 3, 3, 3, 3, 2, 2, 2, 2, 2, 1, 1, 1, 1],
                                })

# Nombre de las asignaturas simplificados
nombre_asignaturas = {
    'automatica' : 'automatica',
    'electronica' : 'electronica',
    'automatizacion' : 'automatizacion',
    'electronica de potencia' : 'potencia',
    'informatica industrial' : 'infind',
    'instrumentacion electronica' : 'instrumentacion',
    'robotica' : 'robotica',
    'electronica analogica' : 'analogica',
    'electronica digital' : 'digital',
    'regulacion automatica' : 'regulacion',
    'control' : 'control',
    'sed' : 'sed',
    'sei' : 'sei',
    'sii' : 'sii',
}

# Devuelve las asignaturas recogidas en un txt
def recoge_asignaturas_txt():

    global grupos_grado

    data = {'plazas_sesion' : [],
            'num_sesiones' : [],
            'horario_sesiones' : [],
            'num_subgrupos' : [],
            'semana_inicial' : [],
            }
    
    # Inicializa las limitaciones_sesion en todos los grupos
    for nombre_grupo, grupo in grupos_grado.iterrows():
        grupos_grado.at[nombre_grupo, 'limitaciones_sesion'] = dict()
    
    asignaturas = list()
    grupos_horarios = list()

    # Abre el fichero en modo lectura
    f = open('asignaturas.txt', 'r')
    # Lee todo el fichero y lo va guardando
    for asig in f:
        datos = asig.split('-')
        asignaturas.append(datos[0])
        horarios = json.loads(datos[3].replace('\'', '\"'))
        grupos_horarios.append(horarios)
        data['plazas_sesion'].append(int(datos[1]))
        data['num_sesiones'].append(int(datos[2]))
        # list(set()) elimina los duplicados
        # Con este for se recoge solo los horarios
        data['horario_sesiones'].append(list(set([horario.split('/')[1] for horario in horarios])))
        data['num_subgrupos'].append(int(datos[4]))
        data['semana_inicial'].append(int(datos[5]))
    f.close()

    # Recorre los grupos para añadirles las limitaciones por sesion
    for grupo, _ in grupos_grado.iterrows():
        for indice_asig in range(len(asignaturas)):
            for horario in grupos_horarios[indice_asig]:
                if horario.split('/')[0] == grupo:
                    # Si no esta creada la key en las limitaciones la crea
                    if asignaturas[indice_asig] not in grupos_grado.at[grupo, 'limitaciones_sesion']:
                        grupos_grado.at[grupo, 'limitaciones_sesion'][asignaturas[indice_asig]] = list()
                    grupos_grado.at[grupo, 'limitaciones_sesion'][asignaturas[indice_asig]].append(horario.split('/')[1])

    return pd.DataFrame(index=asignaturas, data=data)

# Devuelve una lista de las semanas que se impartira la asignatura del subgrupo dado
def semanas_subgrupo(asignatura, subgrupo):
    # Recoge la semana en la que empiezan las sesisones
    semana_inicial = asignatura['semana_inicial']
    # Recoge la letra del subgrupo A = 0, B = 1, C = 2 ...
    offset_subgrupo = ord(subgrupo[-1]) - 65
    # Recoge el numero de las sesiones
    num_sesiones = asignatura['num_sesiones']
    # Recoge el numero de cuantos subgrupos hay de la asignatura
    num_subgrupos = asignatura['num_subgrupos']

    # Dependiendo de la letra del subgrupo empezara en una semana u otra
    offset = semana_inicial+offset_subgrupo

    # Crea una lista de las semanas en las que se impartira la asignatura
    semanas_sesiones = [sem for sem in list(
        range(offset, num_subgrupos*num_sesiones+offset, num_subgrupos))]

    return semanas_sesiones

# Lee los excel de cada asignatura y sus grupos y devuelve los estudiantes con las limitaciones y prioridades de reparto
def lee_estudiantes_asignatura(asignatura):

    global grupos_grado, nombre_asignaturas, cod_error

    lista_todos_grupos_grado = []
    estudiantes_asignatura = pd.DataFrame()

    # Si no existe el excel con la asignatura salta la excepción
    try:
        # Abre los excel con los alumnos matriculados en cada grupo
        lista_grado = pd.read_excel(
            PATH_LISTAS / f'{nombre_asignaturas[asignatura.name]}.xlsx', dtype={'Nº Expediente en Centro':str})# [:-1]
        lista_grado.set_index('Nº Expediente en Centro', inplace=True)

        # Nombre de la columna del Laboratorio Anterior
        laboratorio = ''
        
        # Busca la columna del Laboratorio Anterior
        for col in list(lista_grado.columns):
            if 'Laboratorio' in col:
                laboratorio = col 
        
        # Solo coge los alumnos que no tenga laboratorio anterior
        lista_grado = lista_grado[lista_grado[laboratorio].isna()]

        # Añade la columna limitaciones_sesion_grupo_grado
        lista_grado['limitaciones_sesion_grupo_grado'] = None

        # Se asigna las limitaciones por sesion y la prioridad por reparto
        for idx_estudiante, estudiante in lista_grado.iterrows():
            grupo = estudiante['Grupo matrícula'].split('(')[1].split(')')[0]
            # Comprueba si se han introducido todos los grupos de laboratorio de la asignatura 
            if asignatura.name in grupos_grado.at[grupo, 'limitaciones_sesion'].keys():
                lista_grado.at[idx_estudiante, 'limitaciones_sesion_grupo_grado'] = [grupos_grado.at[grupo, 'limitaciones_sesion']]
            else:
                cod_error = 1
            # Se pone una prioridad de reparto a cada estudiante para cada grupo Optativas -> 4, Tarde -> 3, Mañana -> 2, Dobles Grados -> 1
            lista_grado.at[idx_estudiante, 'prioridad_reparto_grupo_grado'] = int(grupos_grado.at[grupo, 'prioridad_reparto'])

        lista_todos_grupos_grado.append(lista_grado)
    except FileNotFoundError as e:
        logging.error('No se ha encontrado el fichero %s', e.filename)
    except KeyError as e:
        logging.error('Formato del excel erroneo.')
        cod_error = 4

    # Si la lista no esta vacia se añade al dataFrame
    if lista_todos_grupos_grado:
        estudiantes_asignatura = pd.concat(lista_todos_grupos_grado)

    return estudiantes_asignatura

# Devuelve los grupos asignados de cada alumno
def lee_subgrupos_asignados_estudiante(asignatura, lista_estudiantes_subgrupos, idx_estudiante):
    # Si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
    if idx_estudiante in lista_estudiantes_subgrupos.index:
        subgrupos_ya_asignados = {
            subgrupo: asignatura_subgrupo for asignatura_subgrupo, subgrupo in lista_estudiantes_subgrupos.loc[idx_estudiante].items() if subgrupo != '-' and 'subgrupo_' in asignatura_subgrupo and pd.notna(subgrupo)
        }
        logging.info('Asignatura %s - sesiones previamente asignadas %s',
                    asignatura.name, subgrupos_ya_asignados)
        sesiones_subgrupos_ya_asignados = [sesion.split('-')[0] for sesion in subgrupos_ya_asignados]
    else:
        subgrupos_ya_asignados = []
        sesiones_subgrupos_ya_asignados = []

    return subgrupos_ya_asignados, sesiones_subgrupos_ya_asignados

# Comprueba si a un estudiante se le puede asignar un subgrupo incluso si coincide el mismo dia que otro subgrupo ya asignado
def comprueba_subgrupo_estudiante_semanas(asignatura, subgrupos_ya_asignados, subgrupo_a_asignar, sesiones_subgrupos_ya_asignados):
    
    # Comprueba que no estan en el mismo dia y la misma hora
    if subgrupo_a_asignar.split('-')[0] not in sesiones_subgrupos_ya_asignados:
        return True
    
    ultimo_subgrupo = list(subgrupos_ya_asignados.keys())[-1]
    sesion_subgrupo_a_asignar = subgrupo_a_asignar.split('-')[0]

    for subgrupo_ya_asignado in subgrupos_ya_asignados:
        # Guarda los subgrupos ya asignados en otras asignaturas
        asignatura_subgrupo_asignado = subgrupos_ya_asignados[subgrupo_ya_asignado].split('_')[1]
        # Guarda la semana del subgrupos ya asignado
        semanas_subgrupos_ya_asignados = semanas_subgrupo(
            asignatura, subgrupo_ya_asignado)
        # Guarda la semana de los subgrupos a asignar
        semanas_subgrupos_a_asignar = semanas_subgrupo(
            asignatura, subgrupo_a_asignar)
        
        sesion_subgrupo_ya_asignado = subgrupo_ya_asignado.split('-')[0]

        # Si coinciden (los subgrupos y) las semanas se guarda un mensaje de error como que el alumno no tiene grupo
        # Se añade una condición para que compruebe el subgrupo (faltaba) y las semanas
        if sesion_subgrupo_a_asignar == sesion_subgrupo_ya_asignado and any(semana in semanas_subgrupos_ya_asignados for semana in semanas_subgrupos_a_asignar):
            logging.error(
                '%s no consigue asignar el subgrupo %s con semanas %s. Coinciden alguna semana con las del subgrupo %s de %s: %s', asignatura.name, subgrupo_a_asignar, semanas_subgrupos_a_asignar, subgrupo_ya_asignado, asignatura_subgrupo_asignado, semanas_subgrupos_ya_asignados)
            return False
        # Si no coinciden, se asigna el subgrupo
        elif subgrupo_ya_asignado == ultimo_subgrupo:
            logging.info('%s asignado a %s (semanas %s)',
                         asignatura.name, subgrupo_a_asignar, semanas_subgrupos_a_asignar)
            return True

    return True

# Devuelve una lista ordenada de menor a mayor dependiendo de los alumnos en cada subgrupo
def ordenar_diccionario(diccionario):
    diccionario = dict(sorted(diccionario.items(), key=lambda x:x[1]))
    return list(diccionario.keys())

# Asigna a cada un subgrupo de la asignatura pasada por parametro
def asignar_subgrupos_estudiantes(lista_estudiantes_subgrupos, asignatura, lista_estudiantes_asignatura, diccionario_subgrupos, lista_subgrupo):

    # Recorre a los estudiantes
    for idx_estudiante, estudiante in lista_estudiantes_asignatura.iterrows():
        # Para hacer un reparto equitativo ordeno la lista de los subgrupos de menor a mayor para que los estudiantes sean asignados al grupo con más plazas
        lista_subgrupo = ordenar_diccionario(diccionario_subgrupos)

        estudiantes_sin_asignar = lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-']

        # Se sale del bucle porque ya estan todos los estudiantes ya asignados
        if len(estudiantes_sin_asignar) == 0:
            break
        # lista_estudiantes_asignatura['Grupo matrícula'] == estudiante['Grupo matrícula']
        # Si ya tiene asignado un subgrupo continua buscando
        elif lista_estudiantes_asignatura.at[idx_estudiante, f'subgrupo_{asignatura.name}'] != '-':
            continue

        # Recorre los subgrupos de la asignatura
        for subgrupo_a_asignar in lista_subgrupo:
            
            encontrado = False

            # Recoge el numero de estudiantes por cada asignatura de cada subgrupo. Cuenta los que hay en cada subgrupo
            subgrupos_tamaños = lista_estudiantes_asignatura.groupby(
                f'subgrupo_{asignatura.name}').size()

            # Cuenta plazas de cada subgrupo. La primera vez está vacío, por lo que se comprueba
            if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] < asignatura['plazas_sesion']:
                # Si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
                subgrupos_ya_asignados, sesiones_subgrupos_ya_asignados = lee_subgrupos_asignados_estudiante(
                    asignatura, lista_estudiantes_subgrupos, idx_estudiante)
                # Puede ser: MI11-A: Instru, MI11-A: Potencia, MA09-A: Info ya que no coinciden las semanas

                # Si el subgrupo coincide (sesión+subgrupo) hay que comprobar que no coincidan las semanas
                sesion_subgrupo_a_asignar = subgrupo_a_asignar.split('-')[0]
                
                # Comprueba las semanas Y que el subgrupo a asignar coincida con sus limitaciones_sesion_grupo_grado
                if comprueba_subgrupo_estudiante_semanas(asignatura, subgrupos_ya_asignados, subgrupo_a_asignar, sesiones_subgrupos_ya_asignados) and (
                    sesion_subgrupo_a_asignar in estudiante['limitaciones_sesion_grupo_grado'][0][f'{asignatura.name}']):
                    
                    # Si queda solo un estudiante sin asignar se le dará directamente el subgrupo
                    # O si al añadir el segundo estudiante el número de plazas es impar por lo que se asigna directamente el primero
                    if len(lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-']) == 1 or (
                        subgrupo_a_asignar in subgrupos_tamaños and subgrupos_tamaños[subgrupo_a_asignar] + 1 == asignatura['plazas_sesion']):
                        logging.info('%s asignado a %s',
                                            asignatura.name, subgrupo_a_asignar)
                        # Asigna el estudiante un subgrupo
                        lista_estudiantes_asignatura.at[idx_estudiante,
                            f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
                        break
                    elif estudiante['prioridad_reparto_grupo_grado'] != 1 and len(estudiantes_sin_asignar[estudiantes_sin_asignar['Grupo matrícula'] == estudiante['Grupo matrícula']]) == 1:
                        logging.info('%s asignado a %s',
                                            asignatura.name, subgrupo_a_asignar)
                        # Asigna el estudiante un subgrupo
                        lista_estudiantes_asignatura.at[idx_estudiante,
                            f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
                        break

                    # Recorre los estudiantes a partir del estudiante actual
                    for idx_est, est in lista_estudiantes_asignatura.loc[idx_estudiante:][1:].iterrows():

                        # Si ya tiene asignado un subgrupo continua buscando
                        if lista_estudiantes_asignatura.at[idx_est, f'subgrupo_{asignatura.name}'] != '-':
                            continue

                        # Cuenta plazas de cada subgrupo. La primera vez está vacío, por lo que se comprueba
                        if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] + 1 < asignatura['plazas_sesion']:
                            # Si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
                            subgrupos_ya_asignados, sesiones_subgrupos_ya_asignados = lee_subgrupos_asignados_estudiante(
                                asignatura, lista_estudiantes_subgrupos, idx_est)
                            # Puede ser: MI11-A: Instru, MI11-A: Potencia, MA09-A: Info ya que no coinciden las semanas

                            # Si el subgrupo coincide (sesión+subgrupo) hay que comprobar que no coincidan las semanas
                            sesion_subgrupo_a_asignar = subgrupo_a_asignar.split('-')[0]

                            # Comprueba las semanas Y que el subgrupo a asignar coincida con sus limitaciones_sesion_grupo_grado
                            if comprueba_subgrupo_estudiante_semanas(asignatura, subgrupos_ya_asignados, subgrupo_a_asignar, sesiones_subgrupos_ya_asignados) and (
                                sesion_subgrupo_a_asignar in est['limitaciones_sesion_grupo_grado'][0][f'{asignatura.name}']):
                                
                                logging.info('%s asignado a %s', asignatura.name, subgrupo_a_asignar)
                                # Asigna a los dos estudiantes a un subgrupo
                                lista_estudiantes_asignatura.at[idx_estudiante, f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
                                lista_estudiantes_asignatura.at[idx_est, f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
                                diccionario_subgrupos[subgrupo_a_asignar] = diccionario_subgrupos[subgrupo_a_asignar] + 2
                                encontrado = True
                                break

                    if encontrado:
                        break

    return lista_estudiantes_asignatura

# Asigna los grupos a los estudiantes
def asignar_grupos():

    global l, lista_estudiantes_subgrupos, cod_error
    
    # Pruebas
    # asignaturas = asignaturas.reindex(['automatizacion', 'potencia', 'infind', 'instrumentacion', 'robotica'])
    # asignaturas = asignaturas.reindex(['potencia', 'infind', 'automatizacion', 'robotica', 'instrumentacion'])
    # asignaturas = asignaturas.reindex(['robotica', 'instrumentacion', 'automatizacion', 'potencia', 'infind'])
    # asignaturas = asignaturas.reindex(['instrumentacion', 'robotica', 'infind', 'automatizacion', 'potencia'])

    # Pruebas
    # for index in range(20):
    l = []
    lista_estudiantes_subgrupos = pd.DataFrame()
    exito = False
    asignaturas = recoge_asignaturas_txt()
    cod_error = 0
    error = ''

    # Mientras que no se asigne correctamente los alumnos sigue buscando 
    while not exito and cod_error == 0:
        
        # Pruebas
        # print(index, asignaturas.index)
        
        # Se crea la variable exito comprobando si asignaturas esta vacia
        exito = asignaturas.empty

        # Recorre cada asignatura
        for _, asignatura in asignaturas.iterrows():
            # Recoge los estudiantes de los excel de cada asignatura y añade las limitaciones por grupo y sus prioridad
            estudiantes_asignatura = lee_estudiantes_asignatura(asignatura)

            # Si hay un error en el codigo se sale del bucle
            if cod_error == 1:
                error = f'Falta introducir algún grupo de laboratorio en la asignatura {asignatura.name}.'
                break
            elif cod_error == 4:
                error = f'Formato del excel erroneo, para saber los campos necesarios mirar ayuda.'
                break

            # Si no hay estudiantes en esta asignatura
            if estudiantes_asignatura.empty:
                logging.error('No hay estudiantes para la asignatura de %s', asignatura.name)
                asignaturas = asignaturas.drop(asignatura.name, axis=0)
                continue

            # Genera e inicializa a 0 en la lista de estudiantes de la asignatura una columna vacía con el subgrupo
            estudiantes_asignatura[f'subgrupo_{asignatura.name}'] = '-'

            # Se crea la lista con todos los horarios posibles por subgrupo Ej: MI09-A, MI09-B, MI09-C, MI09-D
            # Los horarios estan predefinidos en la variable 'asignaturas'
            lista_subgrupos_asignatura = list(map(lambda x: str(x[0]) + '-' + str(x[1]), product(
                asignatura['horario_sesiones'], string.ascii_uppercase[:asignatura['num_subgrupos']])))
            # Creo un diccionario para saber cuantos estudiantes hay en cada subgrupo
            diccionario_subgrupos_asignatura = {lista_subgrupos_asignatura[i]: 0 for i in range(len(lista_subgrupos_asignatura))}

            # Baraja los estudiantes
            estudiantes_asignatura = estudiantes_asignatura.sample(frac=1)

            # Crea una lista barajada pero ordenada de los estudiantes
            lista_estudiantes_asignatura = estudiantes_asignatura.sort_values(
                by=['prioridad_reparto_grupo_grado'])

            # Salta un Warning si hay mas alumnos que plazas en la asignatura
            if estudiantes_asignatura.shape[0] > int(asignatura['plazas_sesion']) * len(lista_subgrupos_asignatura):
                logging.warning('Asignatura %s tiene más alumnos (%s) que plazas (%s).',
                                asignatura.name, estudiantes_asignatura.shape[0], asignatura['plazas_sesion'] * len(lista_subgrupos_asignatura))
            
            # Crea una lista con los estudiantes asignados con sus respectivos grupos de laboratorio
            lista_estudiantes_asignatura = asignar_subgrupos_estudiantes(lista_estudiantes_subgrupos, asignatura, lista_estudiantes_asignatura, diccionario_subgrupos_asignatura, lista_subgrupos_asignatura)

            logging.error('\n\nLista estudiantes sin grupo de %s', asignatura.name)
            logging.error(lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-'])

            l.append(lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'])
            lista_estudiantes_subgrupos = pd.concat(l, axis=1)

            # Numero de plazas por cada asignatura
            num_plazas = int(asignatura['plazas_sesion']) * len(asignatura['horario_sesiones']) * asignatura['num_subgrupos']
            # Se hace la diferencia entre el número de alumnos y el número de plazas por asignatura
            # Si se queda negativo es porque hay plazas demás
            # Si se queda positivo es porque faltan plazas
            plazas_sin_asignar = estudiantes_asignatura.shape[0] - num_plazas
            # Lista de alumnos que no se les han asigando la asignatura (por falta de número de plazas, por alumnos impares o error del código)
            lista_alumnos_sin_asignar = lista_estudiantes_subgrupos[lista_estudiantes_subgrupos[f'subgrupo_{asignatura.name}'] == '-']

            # Comprueba que los alumnos no asignados son por falta de plazas o que no sean impares
            if len(lista_alumnos_sin_asignar) <= (plazas_sin_asignar if plazas_sin_asignar > 0 else 0) or len(estudiantes_asignatura) == 0:
                exito = True
            # Árbol de fuga: Si se recorre una asignatura y no se asigna de manera correcta se barajan las asignaturas
            # Se optimiza el programa porque de esta forma no es necesario recorrer todas las asignaturas si no se realiza una asignación eficiente
            else:
                
                encontrado = False

                # Recorre los estudiantes no asignados y busca si tienen plazas
                for idx_estudiante, _ in lista_alumnos_sin_asignar.iterrows():
                    for subgrupo, num_asignados in lista_estudiantes_asignatura.groupby(f'subgrupo_{asignatura.name}').size().items():
                        if subgrupo.split('-')[0] in estudiantes_asignatura.at[idx_estudiante, 'limitaciones_sesion_grupo_grado'][0][asignatura.name]:
                            if num_asignados < asignatura['plazas_sesion']:
                                encontrado = True
                                break
                    
                    if encontrado:
                        break
                
                # Si el estudiante no tiene plazas en sus respectivos grupos se genera un error
                if not encontrado:
                    cod_error = 2
                    error = f'No hay plazas suficientes para la asignatura de {asignatura.name} en los subgrupos {estudiantes_asignatura.at[list(lista_alumnos_sin_asignar.index)[0], "limitaciones_sesion_grupo_grado"][0][asignatura.name]}.'
                
                # Reinicia las variables globales
                asignaturas = asignaturas.sample(frac=1)
                exito = False
                l = []
                lista_estudiantes_subgrupos = pd.DataFrame()
                logging.error(
                    'En la asignatura %s no se ha asigando a todos los estudiantes', asignatura.name)
                logging.error(
                    'Estudiantes que no han podido ser asignados')
                logging.error(lista_alumnos_sin_asignar)
                break

        # Pruebas
        # exito = False
        # l = []
        # lista_estudiantes_subgrupos = pd.DataFrame()
        # asignaturas = asignaturas.sample(frac=1)

    return cod_error, error

# Guarda la lista en el excel
def guardar_lista_grupos():

    global lista_estudiantes_subgrupos, nombre_asignaturas, cod_error

    # Recoge las asignaturas del txt
    asignaturas = recoge_asignaturas_txt()

    # Combina subgrupos con datos de estudiantes
    archivos = PATH_LISTAS.glob('*.xlsx')

    df = pd.DataFrame()
    cod_error = 0
    error = ''

    try:
        # Une todos los estudiantes de los archivos excel
        for archivo in archivos:
            # Compara los nombres de los excels con las asignaturas
            if any(nombre_asignaturas[asignatura] in str(archivo) for asignatura in list(asignaturas.index)):
                lista_datos_grupo = pd.read_excel(archivo, dtype={'Nº Expediente en Centro':str})
                lista_datos_grupo.set_index('Nº Expediente en Centro', inplace=True)
                # Se agregan las columnas que se quieren visualizar en el excel
                df = pd.concat([df, lista_datos_grupo[['Apellidos', 'Nombre']]])

        # Borra los duplicados y luego los ordena segun su Nº de Expediente
        lista_estudiantes_datos = df[~df.index.duplicated()].sort_index()

        # Juntan en un lista a los estudiantes y los grupos asignados y lo ponen en el archivo excel
        lista_junta = pd.merge(lista_estudiantes_datos, lista_estudiantes_subgrupos, left_index=True, right_index=True)
        lista_junta.to_excel('lista_subgrupos.xlsx')
    
    except PermissionError as e:
        cod_error = 3
        error = f'El Excel {e.filename} esta abierto, si quiere realizar la operación debe cerrarlo.'
    
    return cod_error, error

# Traduce los meses del ingles al español
def traduce_meses(dia_mes):
    mes_espanol = {
        'January' : 'Enero',
        'February' : 'Febrero',
        'March' : 'Marzo',
        'April' : 'Abril',
        'May' : 'Mayo',
        'June' : 'Junio',
        'July' : 'Julio',
        'August' : 'Agosto',
        'September' : 'Septiembre',
        'October' : 'Octubre',
        'November' : 'Noviembre',
        'December' : 'Diciembre',
    }
    return dia_mes.split(' ')[0] + ' ' +  mes_espanol[dia_mes.split(' ')[1]]


# Crea un pdf con los grupos y horarios
def crea_pdf_grupos_laboratorios(pon_nombre):
    
    global cod_error
    error = ''

    # Se crea un diccionario para recorrer los dias y las horas
    dias = {'LU' : 'LUNES', 'MA' : 'MARTES', 'MI' : 'MIERCOLES', 'JU' : 'JUEVES', 'VI' : 'VIERNES'}
    horas = {'09' : '9:30', '11' : '11:30', '15' : '15:30', '17' : '17:30'}

    # Abre el excel con las listas de los subgrupos
    lista_datos_grupo = pd.read_excel('lista_subgrupos.xlsx', dtype={'Nº Expediente en Centro':str})
    lista_datos_grupo.set_index('Nº Expediente en Centro', inplace=True)

    # Abre el excel con las semanas de inicio
    semanas = pd.read_excel('Semanas.xlsx', index_col=0)

    asignaturas = list()
    estudiantes = pd.DataFrame()
    
    # Recoge las asignaturas del txt
    for col in lista_datos_grupo.columns:
        if 'subgrupo' in col:
            asignaturas.append(col.split('_')[1].lower())

    # Recorre las asignaturas
    for asignatura in asignaturas:
        
        # Abre el fichero en modo lectura
        f = open('asignaturas.txt', 'r')
        # Lee todo el fichero y guarda los datos de las asignaturas
        for asig in f:
            if asig.split('-')[0].lower() == asignatura.lower():
                num_sesiones   = int(asig.split('-')[2])
                horarios       = json.loads(asig.split('-')[3].replace('\'', '\"'))
                num_subgrupos  = int(asig.split('-')[4])
                semana_inicial = int(asig.split('-')[5])
        f.close()

        # Inicia el html, poniendo su encabezado y el estilo
        html = """
            <!DOCTYPE html>
            <html>
                <head>
                    <meta charset="utf-8"/>
                    <style type="text/css">
                        h1 {
                            color: blue;
                        }
                        table {
                            max-width: 50%;
                            font-family: arial, sans-serif;
                            border-collapse: collapse;
                            border-spacing: 50px;
                            vertical-align: top;
                            display: inline-block;
                            }
                        .tabla {
                            border: 1px solid black;
                            text-align: left;
                        }
                        table tr:last-child {
                            border: 0px;
                        }
                    </style>
                </head>
                <body>
        """

        # Añade un titulo con el nombre de la asignatura
        html += f'<h1> {asignatura.upper()} </h1>'

        # Recoge los subgrupos y su tamaño de la asignatura
        subgrupos = lista_datos_grupo[lista_datos_grupo[f'subgrupo_{asignatura}'] != '-'].groupby(f'subgrupo_{asignatura}').size()

        # Recoge el primer subgrupo y añade un encabezado a la tabla
        horario_anterior = list(subgrupos.index)[0].split('-')[0]        
        html_horarios = ''
        
        # Recorre los horarios para guardar los grupos que se han asignado
        for horario in horarios:
            if horario_anterior in horario:
                html_horarios += horario.split('/')[0] + ' '
        
        # Crea un titulo con los grupos y su horario
        html += f'<h2> GRUPO {html_horarios} </h2>'
        html += f'<h2 class="dia_hora">{dias[horario_anterior[:2]]} {horas[horario_anterior[2:]]}</h2><div>'

        # Recorre los subgrupos para introducir en el html la tabla de cada uno
        for subgrupo, _ in subgrupos.items():

            # Recoge el horario del subgrupo actual
            horario_actual = subgrupo.split('-')[0]
            
            # Si el horario ha cambiado entra
            if horario_actual != horario_anterior:
                # Cambia el horario anterior por el actual
                horario_anterior = horario_actual
                html_horarios = ''

                # Recorre los horarios para guardar los grupos que se han asignado
                for horario in horarios:
                    if horario_anterior in horario:
                        html_horarios += horario.split('/')[0] + ' '
                
                # Crea un titulo con los grupos y su horario
                html += f'</div><br><h2> GRUPO {html_horarios} </h2>'
                html += f'<h2 class="dia_hora">{dias[horario_actual[:2]]} {horas[horario_actual[2:]]}</h2><div>'
                

            # Inicializa la tabla en el html
            html += f"""
                    <table>
                        <tr>
                            <td class="tabla"><h3>Grupo {subgrupo.split('-')[1]}</h3></td>
                        </tr>
            """
            # Recoge los estudiantes de este subgrupo
            estudiantes = lista_datos_grupo.loc[lista_datos_grupo[f'subgrupo_{asignatura}'] == subgrupo, ['Apellidos', 'Nombre', f'subgrupo_{asignatura}']]
            
            # Si se ha elegido la opcion de Apellidos, Nombre se escribirá asi en el pdf
            if pon_nombre:
                # Recorre los estudiantes para añadirlos a la tabla
                for _, estudiante in estudiantes.iterrows():
                    html += f'<tr><td class="tabla">{unidecode.unidecode(estudiante["Apellidos"] + ", " + estudiante["Nombre"])}</td></tr>'
            else:
                # Recorre los estudiantes para añadir su numero de matricula a la tabla
                for matricula, _ in estudiantes.iterrows():
                    html += f'<tr><td class="tabla">{matricula}</td></tr>'

            # Recoge las semanas del horario actual
            dias_h = list(semanas[dias[horario_actual[:2]]])
            texto = ''
            
            # Recorre el numero de sesiones
            for i in range(num_sesiones):
                # Añade los dias en los que se va a impartir las clases
                # Traduce los meses y se realiza una abreviacion de 3 letras
                texto += '   ' + traduce_meses(dias_h[semana_inicial + (ord(subgrupo[-1]) - ord('A') + num_subgrupos * i) - 1].strftime('%d %B'))[:6]
            
            # Añade los horarios a la tabla
            html += f'<tr><td><br>{texto}</td></tr></table>'

        # Cierra el html
        html += '</body></html>'

        # Abre el fichero en modo escritura
        f = open(PATH_PDF / f'{asignatura}.html', 'w')
        # Añade el código html al fichero
        f.write(html)
        f.close()

        # Ejecuta el comando wkhtmltopdf para convertir el HTML a PDF
        os.system('wkhtmltopdf \"' + str(PATH_PDF / f'{asignatura}.html') + '\" \"' + str(PATH_PDF / f'{asignatura}.pdf\"'))
        
        # Borra los archivos html
        if os.path.exists(PATH_PDF / f'{asignatura}.html'):
            os.remove(PATH_PDF / f'{asignatura}.html')
    
    return cod_error, error