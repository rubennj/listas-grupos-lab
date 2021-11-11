"""
Created on Wed Jun  9 11:11:30 2021

@author: Ruben
"""

import logging
import string
from itertools import product
import pandas as pd
from pathlib import Path
from collections import deque

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(level='INFO', handlers=[
    logging.FileHandler("debug.log", mode='w+'),
    #logging.StreamHandler()   # Comentado no me sale en la terminal, pero se me sigue guardando en el archivo
    ] 
)

SEMILLA_RND = 123   # Como es una semilla, siempre baraja de la misma forma ¿no? ¿Para pruebas?
PATH_LISTAS = Path('listas_apolo')

# Crea a mano una tabla con todas las asignaturas 
asignaturas = pd.DataFrame(index=['instrumentacion', 'potencia', 'robotica', 'infind', 'automatizacion'],
                           data={'plazas_sesion': [8, 8, 20, 25, 30],
                                 'num_sesiones': [3, 2, 3, 6, 5],
                                 'horario_sesiones': [['MI09', 'MI11', 'JU09', 'JU11', 'VI11'], ['MI11', 'JU11', 'VI09'], ['MA11', 'MI09', 'MI11'], ['LU09', 'LU11', 'MA09', 'MA11', 'MI09'], ['MA09', 'MA11', 'JU09', 'JU11']],
                                 'num_subgrupos': [4, 7, 3, 2, 2],
                                 'semana_inicial': [3, 1, 3, 3, 5],
                                 })

grupos_grado = pd.DataFrame(index=['A302', 'A309', 'EE309'],
                            data={'limitaciones_sesion': [None, None, {'instrumentacion': 'MI11', 'potencia': 'MI11', 'robotica': 'MA11', 'infind': 'MA09', 'automatizacion': 'MA09'}],
                                  'prioridad_reparto': [3, 2, 1],
                                  })

# Devuelve una lista de las semanas que se impartira la asignatura del subgrupo dado
def semanas_subgrupo(nombre_asignatura, subgrupo):
    # Recoge la semana en la que empiezan las sesisones
    semana_inicial = asignaturas.loc[nombre_asignatura, 'semana_inicial']
    # Recoge la letra del subgrupo A = 0, B = 1, C = 2 ...
    offset_subgrupo = ord(subgrupo[-1]) - 65
    # Recoge el numero de las sesiones
    num_sesiones = asignaturas.loc[nombre_asignatura, 'num_sesiones']
    # Recoge el numero de cuantos subgrupos hay de la asignatura
    num_subgrupos = asignaturas.loc[nombre_asignatura, 'num_subgrupos']

    # Dependiendo de la letra del subgrupo empezara en una semana u otra
    offset = semana_inicial+offset_subgrupo

    # Crea una lista de las semanas en las que se impartira la asignatura
    semanas_sesiones = [sem for sem in list(
        range(offset, num_subgrupos*num_sesiones+offset, num_subgrupos))]

    return semanas_sesiones

# Lee los excel de cada asignatura y sus grupos y devuelve los estudiantes con las limitaciones y prioridades de reparto
def lee_estudiantes_asignatura(asignatura):
    lista_todos_grupos_grado = []
    for i, grupo_grado in grupos_grado.iterrows():
        # Abre los excel con los alumnos matriculados en cada grupo
        lista_grado = pd.read_excel(
            PATH_LISTAS / f'{asignatura.name} {grupo_grado.name}.xlsx', index_col='Email')[:-1]

        # Añade la columna limitaciones_sesion_grupo_grado 
        lista_grado['limitaciones_sesion_grupo_grado'] = None # Como algunos seran None pues se añade antes para q no haya problemas
        lista_grado['limitaciones_sesion_grupo_grado'] = lista_grado['limitaciones_sesion_grupo_grado'].apply(
            lambda x: grupo_grado['limitaciones_sesion'])

        lista_grado['prioridad_reparto_grupo_grado'] = grupo_grado['prioridad_reparto'] # Se pone una prioridad de reparto para cada grupo 'A302' -> 3, 'A309' -> 2, 'EE309' -> 1

        lista_todos_grupos_grado.append(lista_grado)

    estudiantes_asignatura = pd.concat(lista_todos_grupos_grado)

    return estudiantes_asignatura

# Devuelve los grupos asignados de cada alumno
def lee_subgrupos_asignados_estudiante(subgrupos_tamaños, asignatura, lista_estudiantes_subgrupos, subgrupo_a_asignar, idx_estudiante):
    # cuenta plazas de cada subgrupo. La primera vez está vacío, por lo que se comprueba
    if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] < asignatura['plazas_sesion']:
        # si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
        if idx_estudiante in lista_estudiantes_subgrupos.index:
            subgrupos_ya_asignados = {
                subgrupo: asignatura_subgrupo for asignatura_subgrupo, subgrupo in lista_estudiantes_subgrupos.loc[idx_estudiante].items() if subgrupo != '-' and 'subgrupo_' in asignatura_subgrupo and pd.notna(subgrupo)
            }
            logging.info('Asignatura %s - sesiones previamente asignadas %s',
                        asignatura.name, subgrupos_ya_asignados)
            sesiones_subgrupos_ya_asignados = [sesion.split(
                '-')[0] for sesion in subgrupos_ya_asignados]
        else:
            subgrupos_ya_asignados = []
            sesiones_subgrupos_ya_asignados = []

        return subgrupos_ya_asignados, sesiones_subgrupos_ya_asignados

# Asigna a un estudiante a un subgrupo
def asigna_subgrupo_estudiante_semanas(subgrupos_ya_asignados, lista_estudiantes_asignatura, subgrupo_a_asignar, idx_estudiante):
    ultimo_subgrupo = list(subgrupos_ya_asignados.keys())[-1]
    sesion_subgrupo_a_asignar = subgrupo_a_asignar.split('-')[0]

    for subgrupo_ya_asignado in subgrupos_ya_asignados:
        # Guarda los subgrupos ya asignados en otras asignaturas
        asignatura_subgrupo_asignado = subgrupos_ya_asignados[subgrupo_ya_asignado].split('_')[1]
        # Guarda la semana del subgrupos ya asignado
        semanas_subgrupos_ya_asignados = semanas_subgrupo(
            asignatura_subgrupo_asignado, subgrupo_ya_asignado)
        # Guarda la semana de los subgrupos a asignar
        semanas_subgrupos_a_asignar = semanas_subgrupo(
            asignatura.name, subgrupo_a_asignar)
        
        sesion_subgrupo_ya_asignado = subgrupo_ya_asignado.split('-')[0]

        # if (idx_estudiante == 'miguel.aspiroz.delacalle@alumnos.upm.es'):
        #     print(idx_estudiante)
        #     print(asignatura_subgrupo_asignado)
        #     print(subgrupo_ya_asignado, semanas_subgrupos_ya_asignados)
        #     print(subgrupo_a_asignar, semanas_subgrupos_a_asignar)
        #     input()

        # Si coinciden (los subgrupos y) las semanas se guarda un mensaje de error como que el alumno no tiene grupo
        # if any(semana in semanas_subgrupos_ya_asignados for semana in semanas_subgrupos_a_asignar):
        # Se añade una condición para que compruebe el subgrupo (faltaba) y las semanas
        if sesion_subgrupo_a_asignar == sesion_subgrupo_ya_asignado and any(semana in semanas_subgrupos_ya_asignados for semana in semanas_subgrupos_a_asignar):
            logging.error(
                '%s no consigue asignar el subgrupo %s con semanas %s. Coinciden alguna semana con las del subgrupo %s de %s: %s', asignatura.name, subgrupo_a_asignar, semanas_subgrupos_a_asignar, subgrupo_ya_asignado, asignatura_subgrupo_asignado, semanas_subgrupos_ya_asignados)
            return False
        # Si no coinciden, se asigna el subgrupo
        elif subgrupo_ya_asignado == ultimo_subgrupo:
            logging.info('%s asignado a %s (semanas %s)',
                         asignatura.name, subgrupo_a_asignar, semanas_subgrupos_a_asignar)
            lista_estudiantes_asignatura.at[idx_estudiante,
                                            f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
            return True

# Empieza el programa
l = []
lista_total = []    # No se usa
lista_estudiantes_subgrupos = pd.DataFrame()

# Recorre cada asignatura
for _, asignatura in asignaturas.iterrows():

    # Recoge los estudiantes de los excel de cada asignatura y añade las limitaciones por grupo y sus prioridad
    estudiantes_asignatura = lee_estudiantes_asignatura(asignatura)
    
    # Genera e inicializa a 0 en la lista de estudiantes de la asignatura una columna vacía con el subgrupo
    estudiantes_asignatura[f'subgrupo_{asignatura.name}'] = '-'

    # Guarda el numero de los subgrupos por asignatura
    num_subgrupos_sesion = asignatura['num_subgrupos']

    # Se crea la lista con todos los horarios posibles por subgrupo Ej: MI09-A, MI09-B, MI09-C, MI09-D
    # Los horarios estan predefinidos en la variable 'asignaturas'
    lista_subgrupos_asignatura = list(map(lambda x: str(x[0]) + '-' + str(x[1]), product(
        asignatura['horario_sesiones'], string.ascii_uppercase[:num_subgrupos_sesion])))
    # Crea una cola con la lista_subgrupos_asignatura
    ciclo_subgrupos_asignatura = deque(lista_subgrupos_asignatura)

    # Baraja los estudiantes de momento siempre igual
    estudiantes_asignatura = estudiantes_asignatura.sample(
        frac=1, random_state=SEMILLA_RND)

    # Crea una lista barajada pero ordenada de los estudiantes
    lista_estudiantes_asignatura = estudiantes_asignatura.sort_values(
        by=['prioridad_reparto_grupo_grado'])

    # Recorre todos los estudiantes
    for idx_estudiante, estudiante in lista_estudiantes_asignatura.iterrows():

        # Guarda la informacion del logeo de todos los estudiantes
        logging.info('\n\nEstudiante: %s , grupo: %s', idx_estudiante,
                     estudiante['Grupo matrícula'][-5:-1])
        
        # Recorre los subgrupos de cada uno de los alumnos
        for subgrupo_a_asignar in list(ciclo_subgrupos_asignatura):
            ciclo_subgrupos_asignatura.rotate(1)
            
            logging.info('\n\nAsignatura %s Subgrupo: %s',
                         asignatura.name, subgrupo_a_asignar)

            # Recoge el numero de estudiantes por cada asignatura de cada subgrupo. Cuenta los que hay en cada subgrupo
            subgrupos_tamaños = lista_estudiantes_asignatura.groupby(
                f'subgrupo_{asignatura.name}').size()
            # if(asignatura.name == 'automatizacion'):
            #     print(subgrupos_tamaños)

            # Cuenta plazas de cada subgrupo. La primera vez está vacío, por lo que se comprueba
            if subgrupo_a_asignar not in subgrupos_tamaños or subgrupos_tamaños[subgrupo_a_asignar] < asignatura['plazas_sesion']:
                # Si el estudiante ya tiene un subgrupo de una asignación previa, se obtiene
                subgrupos_ya_asignados, sesiones_subgrupos_ya_asignados = lee_subgrupos_asignados_estudiante(
                    subgrupos_tamaños, asignatura, lista_estudiantes_subgrupos, subgrupo_a_asignar, idx_estudiante)
                # Miguel MI11-A: Instru, MI11-A: Potencia, MA09-A: Info (con semanas)

                # Si el subgrupo coincide (sesión+subgrupo) hay que comprobar que no coincidan las semanas
                # lista_subgrupos_ya_asignados = [subgrupo for subgrupo in subgrupos_ya_asignados]
                sesion_subgrupo_a_asignar = subgrupo_a_asignar.split('-')[0] # Me quedo con la parte izquierda de -. MI09-A

                # Si esta en el mismo dia y hora pues entra para saber si coincide las semanas
                if sesion_subgrupo_a_asignar in sesiones_subgrupos_ya_asignados:

                    # print(idx_estudiante)
                    # print(sesion_subgrupo_a_asignar)
                    # print(sesiones_subgrupos_ya_asignados)
                    # input()
                    # Se asigna el subgrupo teniendo en cuenta las semanas mediante asigna_subgrupo_estudiante_semanas() (BREAK, deja de buscar más subgrupos), si no sigue buscando
                    if asigna_subgrupo_estudiante_semanas(
                            subgrupos_ya_asignados, lista_estudiantes_asignatura, subgrupo_a_asignar, idx_estudiante):
                        break
                else:
                    # estudiante sin restricciones: limitaciones_sesion está vacía
                    # estudiante con restricciones: se mira si no está vacía la lista de limitaciones_sesion y cual es su grupo horario. Si es compatible lo coge
                    if not estudiante['limitaciones_sesion_grupo_grado'] or (sesion_subgrupo_a_asignar in estudiante['limitaciones_sesion_grupo_grado'][f'{asignatura.name}']):
                        logging.info('%s asignado a %s',
                                    asignatura.name, subgrupo_a_asignar)
                        lista_estudiantes_asignatura.at[idx_estudiante,
                                                        f'subgrupo_{asignatura.name}'] = subgrupo_a_asignar
                        break
                    # si el subgrupo no es compatible, sale y busca en otro grupo
                    else:
                        logging.error(
                            '%s no consigue asignar el subgrupo %s. Hay restriccion %s', asignatura.name, subgrupo_a_asignar, estudiante['limitaciones_sesion_grupo_grado'][f'{asignatura.name}'])
                        continue
            else:
                logging.info('No hay plazas en el grupo %s',
                            subgrupo_a_asignar)

            # comprueba si no consigue asignar subgrupo
            if subgrupo_a_asignar == ciclo_subgrupos_asignatura[-1]:
                logging.error(
                    'Asignatura %s No hay subgrupo disponible para %s', asignatura.name, idx_estudiante)

    logging.error('\n\nLista estudiantes sin grupo de %s', asignatura.name)

    logging.error(
        lista_estudiantes_asignatura[lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'] == '-'])

    logging.info(lista_estudiantes_asignatura.groupby(
        by=f'subgrupo_{asignatura.name}').size())
        
    # lista_total.append(lista_estudiantes_asignatura)

    l.append(lista_estudiantes_asignatura[f'subgrupo_{asignatura.name}'])
    lista_estudiantes_subgrupos = pd.concat(l, axis=1)


# %% Combina subgrupos con datos de estudiantes
archivos = PATH_LISTAS.glob('*.xlsx')

df = pd.DataFrame()
# Une todos los estudiantes de los archivos excel
for archivo in archivos:
    lista_datos_grupo = pd.read_excel(archivo, index_col='Email').drop(columns=['Cód. Asignatura', 'Nombre Asignatura', 'Nº Orden'])[:-1]
    df = pd.concat([df, lista_datos_grupo])

# Borra los duplicados y luego los ordena segun su email
lista_estudiantes_datos = df[~df.index.duplicated()].sort_index()

# Juntan en un lista a los estudiantes y los grupos asignados y lo ponen en el archivo excel
lista_junta = pd.merge(lista_estudiantes_datos, lista_estudiantes_subgrupos, left_index=True, right_index=True)
lista_junta.to_excel('lista_subgrupos.xlsx')