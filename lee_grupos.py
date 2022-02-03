# -*- coding: utf-8 -*-
"""
@author: Andrea Magro Canas

Descripción del código: Algoritmo de asignación de estudiantes en grupos de laboratorio. Se recogen los
estudiantes distribuidos y sus respectivos grupos de laboratorio en un Excel. También, se genera un HTML 
con las listas de las prácticas: asignatura, grupo, horario, subgrupo, los estudiantes que lo componen y fecha.

"""

import logging
import string
from itertools import product
import pandas as pd
from pathlib import Path
import json
from unidecode import unidecode
import configparser
import math
import datetime
import calendar
import io

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(level='INFO', handlers=[
    logging.FileHandler("debug.log", mode='w+'),
    # logging.StreamHandler() # Comentado no me sale en la terminal, pero se me sigue guardando en el archivo
    ] 
)

config = configparser.ConfigParser()
config.read('configuracion.ini', encoding='utf-8')

# Path de los archivos
PATH_LISTAS = Path(config.get('RUTAS', 'LISTAS'))
PATH_EXCEL = Path(config.get('RUTAS', 'EXCEL'))
PATH_HTML = Path(config.get('RUTAS', 'HTML'))
PATH_EXCELS = Path(config.get('RUTAS', 'EXCELS'))
PATH_ALUMNOS = Path(config.get('RUTAS', 'CALENDARIOS_ALUMNOS'))
PATH_PROFESORES = Path(config.get('RUTAS', 'CALENDARIOS_PROFESORES'))
PATH_CSS = Path(config.get('RUTAS', 'CSS'))

# Variables globales
l = list()
lista_estudiantes_subgrupos = pd.DataFrame()

# Codigo de error
cod_error = 0

# Crea a mano una tabla con todos los subgrupos
grupos_grado = pd.DataFrame(index=['DM306', 'Q308', 'D307', 'M301', 'A207', 'E208', 'A309', 'M306', 'A408', 'A204', 'E205', 'A302', 'A404', 'EE208', 'EE309', 'EE403', 'EE507'],
                            data={'limitaciones_sesion' : [{}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}],
                                'prioridad_reparto' : [4, 4, 4, 4, 3, 3, 3, 3, 3, 2, 2, 2, 2, 1, 1, 1, 1],
                                })

# Nombre de los ficheros Excel que se introducen con los alumnos
excel_asignaturas = {
    'Automatica' : 'automatica',
    'Electronica' : 'electronica',
    'Automatizacion' : 'automatizacion',
    'Electronica de Potencia' : 'potencia',
    'Informatica Industrial' : 'infind',
    'Instrumentacion Electronica' : 'instrumentacion',
    'Robotica' : 'robotica',
    'Electronica Analogica' : 'analogica',
    'Electronica Digital' : 'digital',
    'Regulacion Automatica' : 'regulacion',
    'Control' : 'control',
    'SED' : 'sed',
    'SEI' : 'sei',
    'SII' : 'sii',
}

# Nombre de las asignaturas simplificados
nombre_asignaturas = {
    'Automatica' : 'Automática',
    'Electronica' : 'Electrónica',
    'Automatizacion' : 'Automatización',
    'Electronica de Potencia' : 'Electrónica de Potencia',
    'Informatica Industrial' : 'Informática Industrial',
    'Instrumentacion Electronica' : 'Instrumentación Electrónica',
    'Robotica' : 'Robótica',
    'Electronica Analogica' : 'Electrónica Analógica',
    'Electronica Digital' : 'Electrónica Digital',
    'Regulacion Automatica' : 'Regulacion automatica',
    'Control' : 'Control',
    'SED' : 'SED',
    'SEI' : 'SEI',
    'SII' : 'SII',
}

# Devuelve las asignaturas recogidas en un txt
def recoge_asignaturas_txt():

    global grupos_grado

    data = {config.get('EXCEL', 'PLAZAS_SESION') : [],
            config.get('EXCEL', 'NUM_SESIONES') : [],
            config.get('EXCEL', 'HORARIO_SESIONES') : [],
            config.get('EXCEL', 'NUM_SUBGRUPOS') : [],
            config.get('EXCEL', 'SEMANA_INICIAL') : [],
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
        data[config.get('EXCEL', 'PLAZAS_SESION')].append(int(datos[1]))
        data[config.get('EXCEL', 'NUM_SESIONES')].append(int(datos[2]))
        # Con este for se recogen solo los horarios
        data[config.get('EXCEL', 'HORARIO_SESIONES')].append(list(set([horario.split('/')[1] for horario in horarios])))
        data[config.get('EXCEL', 'NUM_SUBGRUPOS')].append(int(datos[4]))
        data[config.get('EXCEL', 'SEMANA_INICIAL')].append(int(datos[5]))
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
    semana_inicial = asignatura[config.get('EXCEL', 'SEMANA_INICIAL')]
    # Recoge la letra del subgrupo A = 0, B = 1, C = 2 ...
    offset_subgrupo = ord(subgrupo[-1]) - 65
    # Recoge el numero de las sesiones
    num_sesiones = asignatura[config.get('EXCEL', 'NUM_SESIONES')]
    # Recoge el numero de cuantos subgrupos hay de la asignatura
    num_subgrupos = asignatura[config.get('EXCEL', 'NUM_SUBGRUPOS')]

    # Dependiendo de la letra del subgrupo empezara en una semana u otra
    offset = semana_inicial+offset_subgrupo

    # Crea una lista de las semanas en las que se impartira la asignatura
    semanas_sesiones = [sem for sem in list(
        range(offset, num_subgrupos*num_sesiones+offset, num_subgrupos))]

    return semanas_sesiones

# Lee los excel de cada asignatura y sus grupos y devuelve los estudiantes con las limitaciones y prioridades de reparto
def lee_estudiantes_asignatura(asignatura):

    global grupos_grado, excel_asignaturas, cod_error

    lista_todos_grupos_grado = []
    estudiantes_asignatura = pd.DataFrame()

    # Si no existe el excel con la asignatura salta la excepción
    try:
        # Abre los excel con los alumnos matriculados en cada grupo
        lista_grado = pd.read_excel(
            PATH_LISTAS / f'{excel_asignaturas[asignatura.name]}.xlsx', dtype={config.get('EXCEL', 'NUM_EXPEDIENTE'):str})# [:-1]
        lista_grado.set_index(config.get('EXCEL', 'NUM_EXPEDIENTE').replace('Â', ''), inplace=True)

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
            grupo = estudiante[config.get('EXCEL', 'GRUPO_MATRICULA')].split('(')[1].split(')')[0]
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

    # Recoge todas las asignaturas
    asignaturas = recoge_asignaturas_txt()

    for subgrupo_ya_asignado in subgrupos_ya_asignados:
        # Guarda los subgrupos ya asignados en otras asignaturas
        asignatura_subgrupo_asignado = subgrupos_ya_asignados[subgrupo_ya_asignado].split('_')[1]

        # Guarda la semana del subgrupos ya asignado
        semanas_subgrupos_ya_asignados = semanas_subgrupo(
            asignaturas.loc[asignatura_subgrupo_asignado], subgrupo_ya_asignado)
        # Guarda la semana de los subgrupos a asignar
        semanas_subgrupos_a_asignar = semanas_subgrupo(
            asignatura, subgrupo_a_asignar)
        
        sesion_subgrupo_ya_asignado = subgrupo_ya_asignado.split('-')[0]

        # Comprueba que el subgrupo a asignar no este ya asignado y que no coincidan las semanas
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

        # Se sale del bucle porque ya estan todos los estudiantes asignados
        if len(estudiantes_sin_asignar) == 0:
            break
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
            if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] < asignatura[config.get('EXCEL', 'PLAZAS_SESION')]:
                # Si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
                subgrupos_ya_asignados, sesiones_subgrupos_ya_asignados = lee_subgrupos_asignados_estudiante(
                    asignatura, lista_estudiantes_subgrupos, idx_estudiante)
                # Puede ser: MI11-A: Instru, MI11-A: Potencia, MA09-A: Info ya que no coinciden las semanas

                # Si el subgrupo coincide (sesión+subgrupo) hay que comprobar que no coincidan las semanas
                sesion_subgrupo_a_asignar = subgrupo_a_asignar.split('-')[0]
                
                # Comprueba las semanas y que el subgrupo a asignar coincida con sus limitaciones_sesion_grupo_grado
                if comprueba_subgrupo_estudiante_semanas(asignatura, subgrupos_ya_asignados, subgrupo_a_asignar, sesiones_subgrupos_ya_asignados) and (
                    sesion_subgrupo_a_asignar in estudiante['limitaciones_sesion_grupo_grado'][0][f'{asignatura.name}']):
                    
                    # Si queda solo un estudiante sin asignar se le dará directamente el subgrupo
                    # O si al añadir el segundo estudiante el número de plazas es impar por lo que se asigna directamente el primero
                    if len(lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-']) == 1 or (
                        subgrupo_a_asignar in subgrupos_tamaños and subgrupos_tamaños[subgrupo_a_asignar] + 1 == asignatura[config.get('EXCEL', 'PLAZAS_SESION')]):
                        logging.info('%s asignado a %s', asignatura.name, subgrupo_a_asignar)
                        # Asigna el estudiante un subgrupo
                        lista_estudiantes_asignatura.at[idx_estudiante,
                            f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
                        break
                    # Comprueba si solo queda un estudiante de ese grupo
                    elif len(estudiantes_sin_asignar[estudiantes_sin_asignar[config.get('EXCEL', 'GRUPO_MATRICULA')] == estudiante[config.get('EXCEL', 'GRUPO_MATRICULA')]]) == 1:
                        # Se añade el estudiante si no es del doble grado
                        # O si es del doble grado se añade si el horario no coincide con los demás grupos (MA09 in [JU09, JU11])
                        # Si es del doble grado y coincide el horario busca otro estudiante 
                        if estudiante['prioridad_reparto_grupo_grado'] != 1 or (
                            not (estudiante['prioridad_reparto_grupo_grado'] == 1 and (
                            any(sesion_subgrupo_a_asignar in lista_estudiantes_asignatura[lista_estudiantes_asignatura[config.get('EXCEL', 'GRUPO_MATRICULA')]==grupo].iloc[0]['limitaciones_sesion_grupo_grado'][0][asignatura.name] for grupo in list(lista_estudiantes_asignatura.groupby(config.get('EXCEL', 'GRUPO_MATRICULA')).groups.keys())[1:])))):
                            logging.info('%s asignado a %s', asignatura.name, subgrupo_a_asignar)
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
                        if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] + 1 < asignatura[config.get('EXCEL', 'PLAZAS_SESION')]:
                            # Si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
                            subgrupos_ya_asignados, sesiones_subgrupos_ya_asignados = lee_subgrupos_asignados_estudiante(
                                asignatura, lista_estudiantes_subgrupos, idx_est)
                            
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

# Comprueba que los estudiantes se hayan repartido en el menor numero de subgrupos
def comprueba_reparto_minimo(asignatura, lista_estudiantes):

    # Recoge el numero de estudiantes por cada asignatura de cada subgrupo. Cuenta los que hay en cada subgrupo
    subgrupos_tamaños = lista_estudiantes.groupby(f'subgrupo_{asignatura.name}').size()
    # Todos los horarios de la asignatura
    horarios = set(subgrupo.split('-')[0] for subgrupo in list(subgrupos_tamaños.keys()))
    # Crea un diccionario con los horarios y un tamaño inicial
    horarios_tamaños = dict.fromkeys(horarios, 0)

    # Añade el numero de alumnos por horario
    for subgrupo, tamaño in subgrupos_tamaños.items():
        horarios_tamaños[subgrupo.split('-')[0]] += tamaño
    
    # Recorre los horarios
    for subgrupo, tamaño in horarios_tamaños.items():
        # Redondea hacia arriba el tamaño de los horarios entre las plazas por sesion
        redondeo = math.ceil(tamaño / asignatura[config.get('EXCEL', 'PLAZAS_SESION')])
        # Si el redondeo es diferente al numero de subrupos introducidos por el usuario se procede a borrar los subgrupos sobrantes 
        if redondeo != asignatura[config.get('EXCEL', 'NUM_SUBGRUPOS')]:
            # Recorre los subgrupos sobrantes
            for i in range(asignatura[config.get('EXCEL', 'NUM_SUBGRUPOS')] - redondeo):
                # Recorre cada estudiante de los subgrupos sobrantes
                for idx_estudiante in list(lista_estudiantes.groupby(f'subgrupo_{asignatura.name}').groups[subgrupo + '-' + chr(ord('A') + i + redondeo)]):
                    # Se eliminan los subgrupos sobrantes y a los alumnos que pertenecian a estos grupos se les reparte de nuevo
                    lista_estudiantes.at[idx_estudiante, f'subgrupo_{asignatura.name}'] = '-'
    
    # Actualiza el numero de estudiantes por cada subgrupo
    subgrupos_tamaños = lista_estudiantes.groupby(f'subgrupo_{asignatura.name}').size()

    # Si hay estudiantes sin asignar entra
    if len(lista_estudiantes[lista_estudiantes[f'subgrupo_{asignatura.name}'] == '-']) != 0:
        subgrupos_tamaños.pop('-')

    return len(lista_estudiantes[lista_estudiantes[f'subgrupo_{asignatura.name}'] == '-']) != 0, list(dict(subgrupos_tamaños).keys()), dict(subgrupos_tamaños)

# Comprueba que los estudiantes se hayan repartido equitativamente
def comprueba_reparto_equitativo(asignatura, lista_estudiantes):

    # Recoge el numero de estudiantes por cada asignatura de cada subgrupo. Cuenta los que hay en cada subgrupo
    subgrupos_tamaños = lista_estudiantes.groupby(f'subgrupo_{asignatura.name}').size()
    # Todos los horarios de la asignatura
    horarios = set(subgrupo.split('-')[0] for subgrupo in list(subgrupos_tamaños.keys()))

    # Recorre los horarios
    for horario in horarios:
        tamaño_actual = -1
        for subgrupo, tamaño in subgrupos_tamaños.items():
            # Si el subgrupo esta en el horario actual
            if horario in subgrupo:
                # Si es el primer subgrupo lo guarda en el tamaño actual (A)
                if tamaño_actual == -1:
                    tamaño_actual = tamaño
                # Comprueba si la diferencia del tamaño entre subgrupos con el mismo horario es mayor o igual que 3
                elif abs(tamaño_actual - tamaño) >= 3:
                    return False

    return True

# Asigna los grupos a los estudiantes
def asignar_grupos():

    global l, lista_estudiantes_subgrupos, cod_error

    l = []
    lista_estudiantes_subgrupos = pd.DataFrame()
    exito = False
    asignaturas = recoge_asignaturas_txt()
    cod_error = 0
    error = ''

    # Mientras que no se asignen correctamente los alumnos sigue buscando 
    while not exito and cod_error == 0:
                
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

            # Si no hay estudiantes en esta asignatura, la borra y continua con el bucle
            if estudiantes_asignatura.empty:
                logging.error('No hay estudiantes para la asignatura de %s', asignatura.name)
                asignaturas = asignaturas.drop(asignatura.name, axis=0)
                continue

            # Genera e inicializa a 0 en la lista de estudiantes de la asignatura una columna vacía con el subgrupo
            estudiantes_asignatura[f'subgrupo_{asignatura.name}'] = '-'

            # Se crea la lista con todos los horarios posibles por subgrupo Ej: MI09-A, MI09-B, MI09-C, MI09-D
            # Los horarios estan predefinidos en la variable 'asignaturas'
            lista_subgrupos_asignatura = list(map(lambda x: str(x[0]) + '-' + str(x[1]), product(
                asignatura[config.get('EXCEL', 'HORARIO_SESIONES')], string.ascii_uppercase[:asignatura[config.get('EXCEL', 'NUM_SUBGRUPOS')]])))
            
            # Creo un diccionario para saber cuantos estudiantes hay en cada subgrupo
            diccionario_subgrupos_asignatura = {lista_subgrupos_asignatura[i]: 0 for i in range(len(lista_subgrupos_asignatura))}

            # Baraja los estudiantes
            estudiantes_asignatura = estudiantes_asignatura.sample(frac=1)

            # Crea una lista barajada pero ordenada de los estudiantes
            lista_estudiantes_asignatura = estudiantes_asignatura.sort_values(by=['prioridad_reparto_grupo_grado'])

            # Salta un Warning si hay mas alumnos que plazas en la asignatura
            if estudiantes_asignatura.shape[0] > int(asignatura[config.get('EXCEL', 'PLAZAS_SESION')]) * len(lista_subgrupos_asignatura):
                logging.warning('Asignatura %s tiene más alumnos (%s) que plazas (%s).',
                                asignatura.name, estudiantes_asignatura.shape[0], asignatura[config.get('EXCEL', 'PLAZAS_SESION')] * len(lista_subgrupos_asignatura))
            
            # Crea una lista con los estudiantes asignados con sus respectivos grupos de laboratorio
            lista_estudiantes_asignatura = asignar_subgrupos_estudiantes(lista_estudiantes_subgrupos, asignatura, lista_estudiantes_asignatura, diccionario_subgrupos_asignatura, lista_subgrupos_asignatura)

            logging.error('\n\nLista estudiantes sin grupo de %s', asignatura.name)
            logging.error(lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-'])

            # Numero de plazas por cada asignatura
            num_plazas = int(asignatura[config.get('EXCEL', 'PLAZAS_SESION')]) * len(asignatura[config.get('EXCEL', 'HORARIO_SESIONES')]) * asignatura[config.get('EXCEL', 'NUM_SUBGRUPOS')]
            # Se hace la diferencia entre el número de alumnos y el número de plazas por asignatura
            # Si se queda negativo es porque hay plazas demás
            # Si se queda positivo es porque faltan plazas
            plazas_sin_asignar = estudiantes_asignatura.shape[0] - num_plazas
            # Lista de alumnos que no se les han asigando la asignatura
            lista_alumnos_sin_asignar = lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-']

            # Comprueba que los alumnos no asignados son por falta de plazas o que no sean impares
            if len(lista_alumnos_sin_asignar) <= (plazas_sin_asignar if plazas_sin_asignar > 0 else 0):
                # Comprueba que los estudiantes se hayan repartido en el menor numero de subgrupos
                comprueba, lista_subgrupos_asignatura, diccionario_subgrupos_asignatura = comprueba_reparto_minimo(asignatura, lista_estudiantes_asignatura)

                # Si se necesita volver a repartir a los estudiantes por haber reducido el tamaño de los subgrupos entra
                if comprueba:
                    # Crea una lista con los estudiantes asignados con sus respectivos grupos de laboratorio
                    lista_estudiantes_asignatura = asignar_subgrupos_estudiantes(lista_estudiantes_subgrupos, asignatura, lista_estudiantes_asignatura, diccionario_subgrupos_asignatura, lista_subgrupos_asignatura)
            
            # Lista de alumnos que no se les han asigando la asignatura
            lista_alumnos_sin_asignar = lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-']

            l.append(lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'])
            lista_estudiantes_subgrupos = pd.concat(l, axis=1)

            # Comprueba que los alumnos no asignados son por falta de plazas o que no sean impares
            # Y Comprueba que los estudiantes se hayan repartido equitativamente
            if len(lista_alumnos_sin_asignar) <= (plazas_sin_asignar if plazas_sin_asignar > 0 else 0) and comprueba_reparto_equitativo(asignatura, lista_estudiantes_asignatura):
                exito = True
            # Via de escape: Si se recorre una asignatura y no se asigna de manera correcta se barajan las asignaturas
            # Se optimiza el programa porque de esta forma no es necesario recorrer todas las asignaturas si no se realiza una asignación eficiente
            else:
                
                encontrado = False

                # Recorre los estudiantes no asignados y busca si tienen plazas
                for idx_estudiante, _ in lista_alumnos_sin_asignar.iterrows():
                    for subgrupo, num_asignados in lista_estudiantes_asignatura.groupby(f'subgrupo_{asignatura.name}').size().items():
                        if subgrupo.split('-')[0] in estudiantes_asignatura.at[idx_estudiante, 'limitaciones_sesion_grupo_grado'][0][asignatura.name]:
                            if num_asignados < asignatura[config.get('EXCEL', 'PLAZAS_SESION')]:
                                encontrado = True
                                break
                    
                    if encontrado:
                        break
                
                # Si el estudiante no tiene plazas en sus respectivos grupos se genera un error
                if not encontrado and not lista_alumnos_sin_asignar.empty:
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

    # Si no ha habido errores, comprueba que la lista de estudiantes no este vacia
    if cod_error == 0 and lista_estudiantes_subgrupos.empty:
        cod_error = 5
        error = f'No se ha introducido los Excels para realizar el reparto.'

    return cod_error, error

# Guarda la lista en el excel
def guardar_lista_grupos():

    global lista_estudiantes_subgrupos, excel_asignaturas, cod_error

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
            if any(excel_asignaturas[asignatura] in str(archivo) for asignatura in list(asignaturas.index)):
                lista_datos_grupo = pd.read_excel(archivo, dtype={config.get('EXCEL', 'NUM_EXPEDIENTE'):str})
                lista_datos_grupo.set_index(config.get('EXCEL', 'NUM_EXPEDIENTE'), inplace=True)
                # Se agregan las columnas que se quieren visualizar en el excel
                df = pd.concat([df, lista_datos_grupo[[config.get('EXCEL', 'APELLIDOS'), config.get('EXCEL', 'NOMBRE')]]])

        # Borra los duplicados y luego los ordena segun su Nº de Expediente
        lista_estudiantes_datos = df[~df.index.duplicated()].sort_index()

        # Juntan en un lista a los estudiantes y los grupos asignados y lo ponen en el archivo excel
        lista_junta = pd.merge(lista_estudiantes_datos, lista_estudiantes_subgrupos, left_index=True, right_index=True)
        lista_junta.to_excel(PATH_EXCEL / config.get('ARCHIVOS', 'LISTA_EXCEL'))

        # Recorre las columnas de las asignaturas para crear un excel por cada asignatura
        for col in lista_junta.columns:
            if col.split('_')[0] == 'subgrupo':
                lista_junta[[config.get('EXCEL', 'APELLIDOS'), config.get('EXCEL', 'NOMBRE'), col]][lista_junta[col].notna()].to_excel(PATH_EXCELS / str(col.split('_')[1] + '_' + datetime.datetime.now().strftime('%d%m%Y') + '.xlsx'))
    
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


# Crea un html con los grupos y horarios
def crea_html_grupos_laboratorios(pon_nombre):
    
    global cod_error
    error = ''

    # Se crea un diccionario para recorrer los dias y las horas
    dias = {'LU' : 'LUNES', 'MA' : 'MARTES', 'MI' : 'MIERCOLES', 'JU' : 'JUEVES', 'VI' : 'VIERNES'}
    horas = {'09' : '9:30', '11' : '11:30', '15' : '15:30', '17' : '17:30'}

    # Abre el excel con las listas de los subgrupos
    lista_datos_grupo = pd.read_excel(config.get('ARCHIVOS', 'LISTA_EXCEL'), dtype={config.get('EXCEL', 'NUM_EXPEDIENTE'):str})
    lista_datos_grupo.set_index(config.get('EXCEL', 'NUM_EXPEDIENTE'), inplace=True)

    # Abre el excel con las semanas
    semanas = pd.read_excel(config.get('ARCHIVOS', 'SEMANAS'), index_col=0)

    asignaturas = list()
    estudiantes = pd.DataFrame()
    
    # Recoge las asignaturas del excel
    for col in lista_datos_grupo.columns:
        if 'subgrupo' in col:
            asignaturas.append(col.split('_')[1])

    # Recorre las asignaturas
    for asignatura in asignaturas:
        
        # Abre el fichero en modo lectura
        f = open('asignaturas.txt', 'r')
        # Lee todo el fichero y guarda los datos de las asignaturas
        for asig in f:
            if asig.split('-')[0] == asignatura:
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
        html += f'<h1> {nombre_asignaturas[asignatura].upper()} </h1>'

        # Recoge los subgrupos y su tamaño de la asignatura
        subgrupos = lista_datos_grupo[lista_datos_grupo[f'subgrupo_{asignatura}'] != '-'].groupby(f'subgrupo_{asignatura}').size()

        # Para recoger el primer grupo se inicializa vacio
        horario_anterior = list()

        # Recorre los subgrupos para introducir en el html la tabla de cada uno
        for subgrupo, _ in subgrupos.items():

            # Recoge el horario del subgrupo actual
            horario_actual = subgrupo.split('-')[0]
            
            # Si el horario anterior esta vacio o el horario ha cambiado entra
            if not horario_anterior or horario_actual != horario_anterior:
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
            estudiantes = lista_datos_grupo.loc[lista_datos_grupo[f'subgrupo_{asignatura}'] == subgrupo, [config.get('EXCEL', 'APELLIDOS'), config.get('EXCEL', 'NOMBRE'), f'subgrupo_{asignatura}']]
            
            # Si se ha elegido la opcion de Apellidos, Nombre se escribirá asi en el html
            if pon_nombre:
                # Recorre los estudiantes para añadirlos a la tabla
                for _, estudiante in estudiantes.iterrows():
                    html += f'<tr><td class="tabla">{unidecode(estudiante[config.get("EXCEL", "APELLIDOS")] + ", " + estudiante[config.get("EXCEL", "NOMBRE")])}</td></tr>'
            else:
                # Recorre los estudiantes para añadir su numero de matricula a la tabla
                for matricula, _ in estudiantes.iterrows():
                    html += f'<tr><td class="tabla">{matricula}</td></tr>'

            
            texto = ''
            dias_h = list()
            asignaturas_cuatrimestre = list()

            # Recorre los dos cuatrimestres
            for i in range(2):
                # El primer cuatrimestre será el impar
                if i == 0:
                    dias_h = list(semanas[dias[horario_actual[:2]]])[:int(config.get('SEMANAS', 'NUM_SEMANAS'))]
                    asignaturas_cuatrimestre = config.get('CUATRIMESTRE', 'IMPAR').split(',')
                else:
                    dias_h = list(semanas[dias[horario_actual[:2]]])[int(config.get('SEMANAS', 'NUM_SEMANAS')):]
                    asignaturas_cuatrimestre = config.get('CUATRIMESTRE', 'PAR').split(',')
                
                if asignatura in asignaturas_cuatrimestre:
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
        f = io.open(PATH_HTML / f'{asignatura}.html', 'w', encoding='utf8')
        # Añade el código html al fichero
        f.write(html)
        f.close()
    
    return cod_error, error

# Crea un calendario anual de los horarios en HTML
def crea_calendario_anual_alumno(num_matricula):
    
    global cod_error
    error = ''

    # Abre el excel con las listas de los subgrupos
    lista_subgrupos = pd.read_excel(config.get('ARCHIVOS', 'LISTA_EXCEL'), dtype={config.get('EXCEL', 'NUM_EXPEDIENTE'):str})
    lista_subgrupos.set_index(config.get('EXCEL', 'NUM_EXPEDIENTE'), inplace=True)

    # Inicializa el calendario
    cal = calendar.HTMLCalendar(firstweekday = 0)

    if num_matricula in lista_subgrupos.index:
        # Recoge las asignaturas
        asignaturas = recoge_asignaturas_txt()

        # Recoge al alumno de la lista
        alumno = lista_subgrupos.loc[num_matricula]
            
        # Abre el excel con las semanas de inicio
        semanas = pd.read_excel(config.get('ARCHIVOS', 'SEMANAS'), index_col=0)

        dias = {'LU' : 'LUNES', 'MA' : 'MARTES', 'MI' : 'MIERCOLES', 'JU' : 'JUEVES', 'VI' : 'VIERNES'}
        dias_de_la_semana = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']
        horas = {'09' : '9:30', '11' : '11:30', '15' : '15:30', '17' : '17:30'}
        # Meses que se van a representar
        meses = [9, 10, 11, 12, 2, 3, 4, 5]
        # Recoge el año de inicio del curso
        año = int(semanas.iloc[1, 1].strftime('%Y'))

        # Recoge las asignaturas del alumno
        asignaturas_alumnos = []
        for col in alumno.index:
            if 'subgrupo' in col:
                asignaturas_alumnos.append(col)


        dias_asignaturas = dict()

        # Recorre las asignaturas del alumno
        for nom_asignatura in asignaturas_alumnos:
            # Comprueba que exista el subgrupo para esta asignatura
            if not pd.isnull(alumno[nom_asignatura]):
                subgrupo = alumno[nom_asignatura]
                semanas_asignatura = list()

                # Recoge las semanas de la asignatura, especificando el cuatrimestre
                if nom_asignatura.split('_')[1] in config.get('CUATRIMESTRE', 'IMPAR').split(','):
                    semanas_asignatura = list(semanas[dias[subgrupo.split('-')[0][:2]]])[:int(config.get('SEMANAS', 'NUM_SEMANAS')) + 1]
                elif nom_asignatura.split('_')[1] in config.get('CUATRIMESTRE', 'PAR').split(','):
                    semanas_asignatura = list(semanas[dias[subgrupo.split('-')[0][:2]]])[int(config.get('SEMANAS', 'NUM_SEMANAS')) + 1:]

                dias_asignaturas[nom_asignatura] = list()
                
                asignatura = asignaturas.loc[nom_asignatura.split('_')[1]]
                semana_inicial = asignatura[config.get('EXCEL', 'SEMANA_INICIAL')]
                num_subgrupos = asignatura[config.get('EXCEL', 'NUM_SUBGRUPOS')]
                num_sesiones = asignatura[config.get('EXCEL', 'NUM_SESIONES')]

                # Recorre el numero de sesiones
                for i in range(num_sesiones):
                        # Guarda en cada asignatura el dia y el mes donde se impartiran los laboratorios
                    dias_asignaturas[nom_asignatura].append(semanas_asignatura[semana_inicial + (ord(subgrupo[-1]) - ord('A') + num_subgrupos * i) - 1].strftime('%d %m').split(' '))

        # Crea el calendario en tres columnas del año entero
        html_code = '<!DOCTYPE html><html><head><meta charset="utf-8"/>'
        html_code += f'<link rel="stylesheet" type="text/css" href="{PATH_CSS.absolute() / "calendar.css"}"/>'
        html_code += '</head><body>'
        html_code += f'<h1 class="nombre">{alumno[config.get("EXCEL", "NOMBRE")].upper()} {alumno[config.get("EXCEL", "APELLIDOS")].upper()}</h1>'
        html_code += '<table border="0" cellpadding="0" cellspacing="0" id="calendar">'

        # Recorre los dos años
        for i in range(2):
            
            html_code += '<tr class="month-row">'
            
            # Si es el primer año seran los meses 9, 10, 11, 12
            if i == 0:
                lista_meses = meses[:4]
            # Si es el segundo año seran los meses 2, 3, 4, 5
            else:
                lista_meses = meses[4:]

            # Recorre los meses
            for mes in lista_meses:
                html_code += '<td class="calendar-month">'
                # Devuelve en formato html el mes del año especificado
                html_mes = cal.formatmonth(año + i, mes)
                
                # Recorre las asignaturas
                for asignatura, lista_dia_mes in dias_asignaturas.items():
                    # Recorre los dias de laboratorio de cada asignatura
                    for dia_asignatura, mes_asignatura in lista_dia_mes:
                        # Si coincide que una asignatura imparte el laboratorio este mes
                        if mes == int(mes_asignatura):
                            dia_de_la_semana = calendar.weekday(año + i, mes, int(dia_asignatura))
                            # Si todavia no hay un laboratorio ese mismo dia se le pinta el background del color especifico de cada asignatura
                            if html_mes.find(f'class="{dias_de_la_semana[dia_de_la_semana]}">{int(dia_asignatura)}<') != -1:
                                html_mes = html_mes.replace(f'">{int(dia_asignatura)}<', f' {asignatura.split("_")[1].replace(" ", "_").lower()}"><b>{int(dia_asignatura)}</b><')
                            # Si ya hay un laboratorio ese dia se añadira un borde con el color de la asignatura
                            else:
                                html_mes = html_mes.replace(f'"><b>{int(dia_asignatura)}</b><', f' {asignatura.split("_")[1].replace(" ", "_").lower() + "-borde"}"><b>{int(dia_asignatura)}</b><')

                html_code += html_mes

                html_code += '</td>'
            
            html_code += '</tr>'

        html_code += '</table>'

        # Recorre las asignaturas para crear la leyenda
        for nom_asignatura in dias_asignaturas:
            asignatura = nom_asignatura.split('_')[1]
            subgrupo = alumno[nom_asignatura].split('-')[0]

            # Crea una caja con el color de la asignatura
            html_code += f'<div class="leyenda"><div class="caja-leyenda {asignatura.replace(" ", "_").lower()}"></div><div class="texto-leyenda"><h2>{nombre_asignaturas[asignatura]}</h2><h3>'
            html_code += f'{dias[subgrupo[:2]]} {horas[subgrupo[2:]]} Grupo {alumno[nom_asignatura].split("-")[1]}</h3></div></div><br>'

        html_code += '</body></html>'

        # Traduce los dias de la semana
        html_code = html_code.replace('Mon', 'L')
        html_code = html_code.replace('Tue', 'M')
        html_code = html_code.replace('Wed', 'X')
        html_code = html_code.replace('Thu', 'J')
        html_code = html_code.replace('Fri', 'V')
        html_code = html_code.replace('Sat', 'S')
        html_code = html_code.replace('Sun', 'D')
        
        # Traduce los meses
        html_code = html_code.replace('September', 'Septiembre')
        html_code = html_code.replace('October', 'Octubre')
        html_code = html_code.replace('November', 'Noviembre')
        html_code = html_code.replace('December', 'Diciembre')
        html_code = html_code.replace('February', 'Febrero')
        html_code = html_code.replace('March', 'Marzo')
        html_code = html_code.replace('April', 'Abril')
        html_code = html_code.replace('May', 'Mayo')

        # Crea el fichero HTML del calendario
        html = io.open(PATH_ALUMNOS / f'calendario_{num_matricula}.html', 'w', encoding='utf8')
        # Escribe el codigo HTML al fichero
        html.write(''.join(html_code))
        html.close()
    else:
        cod_error = 6
        error = 'Se ha introducido mal el número matrícula del alumno.'
    
    return cod_error, error


# Crea un calendario anual de los horarios en HTML
def crea_calendario_anual_profesor(identificador):
    
    global cod_error
    error = ''

    # Abre el excel con las listas de los subgrupos
    lista_subgrupos = pd.read_excel(config.get('ARCHIVOS', 'PROFESORES'), dtype={config.get('EXCEL', 'IDENTIFICADOR'):str})
    lista_subgrupos.set_index(config.get('EXCEL', 'IDENTIFICADOR'), inplace=True)

    # Inicializa el calendario
    cal = calendar.HTMLCalendar(firstweekday = 0)

    if identificador in lista_subgrupos.index:
        # Recoge las asignaturas
        asignaturas = recoge_asignaturas_txt()

        # Recoge al profesor de la lista
        profesor = lista_subgrupos.loc[identificador]

        # Abre el excel con las semanas de inicio
        semanas = pd.read_excel(config.get('ARCHIVOS', 'SEMANAS'), index_col=0)

        dias = {'LU' : 'LUNES', 'MA' : 'MARTES', 'MI' : 'MIERCOLES', 'JU' : 'JUEVES', 'VI' : 'VIERNES'}
        dias_de_la_semana = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']
        horas = {'09' : '9:30', '11' : '11:30', '15' : '15:30', '17' : '17:30'}
        # Meses que se van a representar
        meses = [9, 10, 11, 12, 2, 3, 4, 5]
        # Recoge el año de inicio del curso
        año = int(semanas.iloc[1, 1].strftime('%Y'))

        # Recoge las asignaturas del profesor
        asignaturas_profesor = list()
        for col in profesor.index:
            if 'subgrupo' in col:
                asignaturas_profesor.append(col)

        dias_asignaturas = dict()

        # Recorre las asignaturas del profesor
        for nom_asignatura in asignaturas_profesor:
            # Comprueba que exista un subgrupo para esta asignatura
            if not pd.isnull(profesor[nom_asignatura]):
                # Recoge los subgrupos de la asignatura
                subgrupos = profesor[nom_asignatura].split(',')

                # Recorre los subgrupos
                for subgrupo in subgrupos:
                    semanas_asignatura = list()
                    
                    # Recoge las semanas de la asignatura, especificando el cuatrimestre
                    if nom_asignatura.split('_')[1] in config.get('CUATRIMESTRE', 'IMPAR').split(','):
                        semanas_asignatura = list(semanas[dias[subgrupo.split('-')[0][:2]]])[:int(config.get('SEMANAS', 'NUM_SEMANAS')) + 1]
                    elif nom_asignatura.split('_')[1] in config.get('CUATRIMESTRE', 'PAR').split(','):
                        semanas_asignatura = list(semanas[dias[subgrupo.split('-')[0][:2]]])[int(config.get('SEMANAS', 'NUM_SEMANAS')) + 1:]
                    
                    # Sino existe la asignatura dentro de dias_asignaturas se crea
                    if not nom_asignatura in dias_asignaturas.keys(): 
                        dias_asignaturas[nom_asignatura] = list()

                    asignatura = asignaturas.loc[nom_asignatura.split('_')[1]]
                    semana_inicial = asignatura[config.get('EXCEL', 'SEMANA_INICIAL')]
                    num_subgrupos = asignatura[config.get('EXCEL', 'NUM_SUBGRUPOS')]
                    num_sesiones = asignatura[config.get('EXCEL', 'NUM_SESIONES')]

                    # Recorre el numero de sesiones
                    for i in range(num_sesiones):
                        # Guarda en cada asignatura el dia y el mes donde se impartiran los laboratorios
                        dias_asignaturas[nom_asignatura].append(semanas_asignatura[semana_inicial + (ord(subgrupo[-1]) - ord('A') + num_subgrupos * i) - 1].strftime('%d %m').split(' '))

        # Inicializa el calendario el calendario
        html_code = '<!DOCTYPE html><html><head><meta charset="utf-8"/>'
        html_code += f'<link rel="stylesheet" type="text/css" href="{PATH_CSS.absolute() / "calendar.css"}"/>'
        html_code += '</head><body>'
        html_code += f'<h1 class="nombre"> {profesor[config.get("EXCEL", "NOMBRE")].upper()} {profesor[config.get("EXCEL", "APELLIDOS")].upper()}</h1>'
        html_code += '<table border="0" cellpadding="0" cellspacing="0" id="calendar">'

        # Recorre los dos años
        for i in range(2):
            
            html_code += '<tr class="month-row">'
            
            # Si es el primer año seran los meses 9, 10, 11, 12
            if i == 0:
                lista_meses = meses[:4]
            # Si es el segundo año seran los meses 2, 3, 4, 5
            else:
                lista_meses = meses[4:]

            # Recorre los meses
            for mes in lista_meses:
                html_code += '<td class="calendar-month">'
                # Devuelve en formato html el mes del año especificado
                html_mes = cal.formatmonth(año + i, mes)
                
                # Recorre las asignaturas
                for asignatura, lista_dia_mes in dias_asignaturas.items():
                    # Recorre los dias de laboratorio de cada asignatura
                    for dia_asignatura, mes_asignatura in lista_dia_mes:
                        # Si coincide que una asignatura imparte el laboratorio este mes
                        if mes == int(mes_asignatura):
                            dia_de_la_semana = calendar.weekday(año + i, mes, int(dia_asignatura))
                            # Si todavia no hay un laboratorio ese mismo dia se le pinta el background del color especifico de cada asignatura
                            if html_mes.find(f'class="{dias_de_la_semana[dia_de_la_semana]}">{int(dia_asignatura)}<') != -1:
                                html_mes = html_mes.replace(f'">{int(dia_asignatura)}<', f' {asignatura.split("_")[1].replace(" ", "_").lower()}"><b>{int(dia_asignatura)}</b><')
                            # Si ya hay un laboratorio ese dia se añadira un borde con el color de la asignatura
                            else:
                                html_mes = html_mes.replace(f'"><b>{int(dia_asignatura)}</b><', f' {asignatura.split("_")[1].replace(" ", "_").lower() + "-borde"}"><b>{int(dia_asignatura)}</b><')
                
                html_code += html_mes + '</td>'
            
            html_code += '</tr>'

        html_code += '</table>'

        # Recorre las asignaturas para crear la leyenda
        for nom_asignatura in dias_asignaturas:
            asignatura = nom_asignatura.split('_')[1]
            subgrupos = profesor[nom_asignatura].split('-')[0]

            # Crea una caja con el color de la asignatura
            html_code += f'<div class="leyenda"><div class="caja-leyenda {asignatura.replace(" ", "_").lower()}"></div><div class="texto-leyenda"><h2>{nombre_asignaturas[asignatura]}</h2><h3>'
            html_code += f'{dias[subgrupos[:2]]} {horas[subgrupos[2:]]} Grupo {" Grupo ".join([subgrupo.split("-")[1] for subgrupo in profesor[nom_asignatura].split(",")])}</h3></div></div><br>'

        html_code += '</body></html>'

        # Traduce los dias de la semana
        html_code = html_code.replace('Mon', 'L')
        html_code = html_code.replace('Tue', 'M')
        html_code = html_code.replace('Wed', 'X')
        html_code = html_code.replace('Thu', 'J')
        html_code = html_code.replace('Fri', 'V')
        html_code = html_code.replace('Sat', 'S')
        html_code = html_code.replace('Sun', 'D')
        
        # Traduce los meses
        html_code = html_code.replace('September', 'Septiembre')
        html_code = html_code.replace('October', 'Octubre')
        html_code = html_code.replace('November', 'Noviembre')
        html_code = html_code.replace('December', 'Diciembre')
        html_code = html_code.replace('February', 'Febrero')
        html_code = html_code.replace('March', 'Marzo')
        html_code = html_code.replace('April', 'Abril')
        html_code = html_code.replace('May', 'Mayo')

        # Crea el fichero HTML del calendario
        html = io.open(PATH_PROFESORES / f'calendario_{identificador}.html', 'w', encoding='utf8')
        # Escribe el codigo HTML al fichero
        html.write(''.join(html_code))
        html.close()
    else:
        cod_error = 7
        error = 'Se ha introducido mal el identificador del profesor.'
    
    return cod_error, error