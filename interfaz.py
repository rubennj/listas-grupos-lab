# -*- coding: utf-8 -*-
"""
@author: Andrea Magro Canas

Descripción del código: Programación del entorno gráfico.

"""

import sys
import os
from PyQt5 import uic, QtCore, QtGui
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import QPrintPreviewDialog
from PyQt5.QtGui import (QColor)
import pandas as pd
import json
from unidecode import unidecode

import lee_grupos

borrar = False
# Grupos de los laboratorios [Mañana, Tarde, Ambos]
grupos = [[], [], []]

# Crea los grupos de los laboratorios
for i, lista_grupos in enumerate(lee_grupos.config.options('GRUPOS')):
    grupos[i] = lee_grupos.config.get('GRUPOS', lista_grupos).split(',')

class GUI(QMainWindow):
    
    def __init__(self):
        super().__init__()
        # Carga la interfaz en nuestra clase
        uic.loadUi("interfaz.ui", self)
        # Variables
        self.generadosExcel = False
        # Eventos de los botones
        self.setWindowIcon(QtGui.QIcon('./img/icono.png'))
        self.BtnAcercaDeDisponibilidad.setIcon(QtGui.QIcon('./img/informacion.png'))
        self.BtnAcercaDeDisponibilidad.setIconSize(QSize(12, 12))
        self.BtnAcercaDeDisponibilidad.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAcercaDeDisponibilidad.setToolTip('Acerca De')
        self.BtnAcercaDeDisponibilidad.clicked.connect(self.fn_acerca_de)
        self.BtnAcercaDeAgregarLabs.setIcon(QtGui.QIcon('./img/informacion.png'))
        self.BtnAcercaDeAgregarLabs.setIconSize(QSize(12, 12))
        self.BtnAcercaDeAgregarLabs.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAcercaDeAgregarLabs.setToolTip('Acerca De')
        self.BtnAcercaDeAgregarLabs.clicked.connect(self.fn_acerca_de)
        self.BtnAcercaDeAsignacion.setIcon(QtGui.QIcon('./img/informacion.png'))
        self.BtnAcercaDeAsignacion.setIconSize(QSize(12, 12))
        self.BtnAcercaDeAsignacion.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAcercaDeAsignacion.setToolTip('Acerca De')
        self.BtnAcercaDeAsignacion.clicked.connect(self.fn_acerca_de)
        self.BtnAcercaDeCalendario.setIcon(QtGui.QIcon('./img/informacion.png'))
        self.BtnAcercaDeCalendario.setIconSize(QSize(12, 12))
        self.BtnAcercaDeCalendario.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAcercaDeCalendario.setToolTip('Acerca De')
        self.BtnAcercaDeCalendario.clicked.connect(self.fn_acerca_de)
        self.BtnAyudaDisponibilidad.setIcon(QtGui.QIcon('./img/ayuda.png'))
        self.BtnAyudaDisponibilidad.setIconSize(QSize(13, 13))
        self.BtnAyudaDisponibilidad.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAyudaDisponibilidad.setToolTip('Botón de ayuda')
        self.BtnAyudaDisponibilidad.clicked.connect(self.fn_ayuda_disponibilidad)
        self.BtnAyudaAgregarLabs.setIcon(QtGui.QIcon('./img/ayuda.png'))
        self.BtnAyudaAgregarLabs.setIconSize(QSize(13, 13))
        self.BtnAyudaAgregarLabs.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAyudaAgregarLabs.setToolTip('Botón de ayuda')
        self.BtnAyudaAgregarLabs.clicked.connect(self.fn_ayuda_agregarlabs)
        self.BtnAyudaAsignacion.setIcon(QtGui.QIcon('./img/ayuda.png'))
        self.BtnAyudaAsignacion.setIconSize(QSize(13, 13))
        self.BtnAyudaAsignacion.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAyudaAsignacion.setToolTip('Botón de ayuda')
        self.BtnAyudaAsignacion.clicked.connect(self.fn_ayuda_asignacion)
        self.BtnAyudaCalendario.setIcon(QtGui.QIcon('./img/ayuda.png'))
        self.BtnAyudaCalendario.setIconSize(QSize(13, 13))
        self.BtnAyudaCalendario.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAyudaCalendario.setToolTip('Botón de ayuda')
        self.BtnAyudaCalendario.clicked.connect(self.fn_ayuda_calendario)
        self.BtnAsignacion.setIcon(QtGui.QIcon('./img/asignar.png'))
        self.BtnAsignacion.setIconSize(QSize(13, 13))
        self.BtnAsignacion.setToolTip('Realizar el reparto de estudiantes')
        self.BtnAsignacion.clicked.connect(self.fn_asignar_grupos)
        self.BtnGuardarExcel.setEnabled(False)
        self.BtnGuardarExcel.setIcon(QtGui.QIcon('./img/excelDes.png'))
        self.BtnGuardarExcel.setIconSize(QSize(15, 15))
        self.BtnGuardarExcel.setToolTip('Generar Excel con el reparto')
        self.BtnGuardarExcel.clicked.connect(self.fn_guarda_excel)
        self.BtnCrearHTML.setEnabled(False)
        self.BtnCrearHTML.setIcon(QtGui.QIcon('./img/htmlDes.png'))
        self.BtnCrearHTML.setIconSize(QSize(15, 15))
        self.BtnCrearHTML.setToolTip('Generar las listas de laboratorio en HTML')
        self.BtnCrearHTML.clicked.connect(self.fn_guarda_html)
        self.BtnAceptar.setIcon(QtGui.QIcon('./img/aceptar.png'))
        self.BtnAceptar.setIconSize(QSize(11, 11))
        self.BtnAceptar.setToolTip('Guardar los datos del laboratorio')
        self.BtnAceptar.clicked.connect(self.fn_guardar_asignatura)
        self.BtnBorrarLabs.setIcon(QtGui.QIcon('./img/borrarTodo.png'))
        self.BtnBorrarLabs.setIconSize(QSize(12, 12))
        self.BtnBorrarLabs.clicked.connect(self.fn_borrar_laboratorios)
        self.BtnBorrarAulas.setIcon(QtGui.QIcon('./img/borrar.png'))
        self.BtnBorrarAulas.setIconSize(QSize(11, 11)) 
        self.BtnBorrarAulas.clicked.connect(self.fn_borrar_aulas)
        self.BtnBorrarHorario.setIcon(QtGui.QIcon('./img/borrarTodo.png'))
        self.BtnBorrarHorario.setIconSize(QSize(12, 12))
        self.BtnBorrarHorarioSeleccionado.setToolTip('Borra los horarios')
        self.BtnBorrarHorario.clicked.connect(self.fn_borrar_horarios)
        self.BtnBorrarHorarioSeleccionado.setIcon(QtGui.QIcon('./img/borrar.png'))
        self.BtnBorrarHorarioSeleccionado.setIconSize(QSize(12, 12))
        self.BtnBorrarHorarioSeleccionado.setToolTip('Borra los horarios seleccionados')
        self.BtnBorrarHorarioSeleccionado.clicked.connect(self.fn_borrar_horario_seleccionado)
        self.BtnGuardaHorario.setIcon(QtGui.QIcon('./img/guardar.png'))
        self.BtnGuardaHorario.setIconSize(QSize(10, 10))
        self.BtnGuardaHorario.setToolTip('Guardar horarios')
        self.BtnGuardaHorario.clicked.connect(self.fn_guardar_horarios)
        self.BtnGuardarAulas.setIcon(QtGui.QIcon('./img/guardar.png'))
        self.BtnGuardarAulas.setIconSize(QSize(10, 10))
        self.BtnGuardarAulas.setToolTip('Guardar las aulas compartidas')
        self.BtnGuardarAulas.clicked.connect(self.fn_guardar_aulas)
        self.BtnAceptarAlumno.setIcon(QtGui.QIcon('./img/horario.png'))
        self.BtnAceptarAlumno.setIconSize(QSize(14, 14))
        self.BtnAceptarAlumno.setToolTip('Generar el horario del alumno')
        self.BtnAceptarAlumno.clicked.connect(self.fn_carga_calendario)
        self.BtnCalendarioAlumno.setIcon(QtGui.QIcon('./img/htmlAct.png'))
        self.BtnCalendarioAlumno.setIconSize(QSize(13, 13))
        self.BtnCalendarioAlumno.setToolTip('Guardar en HTML el calendario anual del alumno')
        self.BtnCalendarioAlumno.clicked.connect(self.fn_calendario_anual_alumno)
        self.BtnCalendarioProfesor.setIcon(QtGui.QIcon('./img/htmlAct.png'))
        self.BtnCalendarioProfesor.setIconSize(QSize(13, 13))
        self.BtnCalendarioProfesor.setToolTip('Guardar en HTML el calendario anual del profesor')
        self.BtnCalendarioProfesor.clicked.connect(self.fn_calendario_anual_profesor)
        # Evento de las pestañas
        self.tabWidget.currentChanged.connect(self.fn_reinicia_pestanas)
        self.lblSeleccionFormato.setStyleSheet("color: gray")
        # Eventos TreeView
        self.ArbolAsignaturas.doubleClicked.connect(self.fn_selecciona_asignatura)
        # Eventos ComboBox
        self.ComboBoxAsignatura.textActivated.connect(self.fn_anadir_horarios)
        self.ComboBoxGrupos.textActivated.connect(self.fn_anadir_horas)
        # Evento de la Tabla
        self.TablaHorario.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.TablaHorario.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # Ruta de los archivos
        self.lblRuta.setText('RUTA DE LOS ARCHIVOS')
        self.lblRutaArchivosExcel.setText('Ruta de lista_subgrupos:  ' + str(lee_grupos.PATH_EXCEL.absolute()))
        self.lblRutaArchivosExcels.setText('Ruta de los Excels:  ' + str(lee_grupos.PATH_EXCELS.absolute()))
        self.lblRutaArchivosHTML.setText('Ruta de los HTMLs:   ' + str(lee_grupos.PATH_HTML.absolute()))
    
    # Carga una alerta en el botón de ayuda de la pestaña de Disponibilidad
    def fn_ayuda_disponibilidad(self):
        mensaje_alerta('Que', 'En esta pestaña se define la disponibilidad de las aulas proporcionadas por secretaria donde se van a cursas las diferentes asignaturas.\n'+
                        '- Para seleccionar las asignaturas de ambos cuatrimestres se pulsa doble click.\n'+
                        '- Esta elección habilita los desplegables.\n'+
                        '- Botón "Guardar": almacena la selección.\n'+
                        '- Botón "Borrar Horarios": elimina todos los horarios seleccionados.\n'+
                        '- Botón "Borrar Seleccionados": elimina el horario escogido.\n'+
                        '- Visualizador "Horarios": muestra las selecciones realizadas.')

    # Carga una alerta en el botón de ayuda de la pestaña de Agregar Laboratorios
    def fn_ayuda_agregarlabs(self):
        mensaje_alerta('Que', 'En esta pestaña se agregan las asignaturas de los laboratorios que se quieren asignar.\n'+
                        '- Desplegable Asignaturas: permite seleccionar la materia a asignar.\n'+
                        '- ComboBox: recoge los datos que se introducirán en el algoritmo.\n'+
                        '- Desplegable Aulas: permite seleccionar las asignaturas que comparten aula.\n'+
                        '- Botón "Aceptar": almacena la selección.\n'+
                        '- Botón "Guardar Aulas": almacena las aulas escogidas.\n'+
                        '- Botón "Borrar Aulas": elimina las aulas seleccionadas.\n'+
                        '- Botón "Borrar Laboratorios": elimina los laboratorios elegidos en "Horarios".\n'+
                        '- Visualizador "Comparten Aulas": muestra las aulas que han sido guardadas.\n'+
                        '- Visualizador "Horarios": muestra los horarios disponibles de la materia seleccionada y permite elegir los grupos de práticas.')

    # Carga una alerta en el botón de ayuda de la pestaña de Asignación
    def fn_ayuda_asignacion(self):
        mensaje_alerta('Que', 'En esta pestaña se asignan los estudiantes y se habilita guardar este reparto en varios formatos.\n'+
                        '- Botón "Asignación": realiza el reparto de estudiantes.\n'+
                        '- Botón "Guardar en Excel": crea varios Excel con el reparto.\n'+
                        '- Botón "Crear HTML": genera varios HTMLs con las listas de laboratorio.\n'+
                        '- RadioButton: permite elegir si se quiere el HTML con nombres o con números de matrícula.')

    # Carga una alerta en el botón de ayuda de la pestaña de Calendarios
    def fn_ayuda_calendario(self):
        mensaje_alerta('Que', 'En esta pestaña se crea el calendario de prácticas de un alumno. Incluye el horario y las semanas que tiene que asistir a los laboratorios\n'+
                        '- Campo Nº matrícula: sirve para introducir el número de matrícula del alumno.\n'+
                        '- Botón "Horario": completa la tabla con el horario semanal de un alumno.\n'+
                        '- Botón "Cal. Anual": genera un HTML con el calendario anual del estudiante o del profesor.')
    
    # Carga un mensaje al pulsar el botón Acerca De
    def fn_acerca_de(self):
        mensaje_alerta('Que', 'Autor: Andrea Magro Canas.\n'+
                        '- Licencia: Copyright (C) 2022 Andrea Magro Canas, Universidad Politécnica de Madrid.\n'+
                        '- Versión: v1.0.0.')

    # Carga los datos
    def fn_asignar_grupos(self):
        cod_error, error = lee_grupos.asignar_grupos()

        if cod_error == 0:
            self.BtnGuardarExcel.setEnabled(True)
            self.BtnGuardarExcel.setIcon(QtGui.QIcon('./img/excelAct.png'))
            mensaje_alerta('Inf', 'Ha terminado de asignar a los estudiantes.')
        else:
            mensaje_alerta('Err', error)
    
    # Guarda a los estudiantes en el excel
    def fn_guarda_excel(self):
        cod_error, error = lee_grupos.guardar_lista_grupos()

        if cod_error == 0:
            self.BtnCrearHTML.setEnabled(True)
            self.lblSeleccionFormato.setEnabled(True)
            self.lblSeleccionFormato.setStyleSheet("color: white")
            self.radioBtnMatricula.setEnabled(True)
            self.radioBtnNombre.setEnabled(True)
            self.BtnCrearHTML.setEnabled(True)
            self.BtnCrearHTML.setIcon(QtGui.QIcon('./img/htmlAct.png'))
            self.generadosExcel = True
            mensaje_alerta('Inf', 'Ha terminado de guardar a los estudiantes en el Excel.')
        else:
            mensaje_alerta('Err', error)
    
    # Guarda a los estudiantes en el HTML
    def fn_guarda_html(self):
        cod_error, error = lee_grupos.crea_html_grupos_laboratorios(self.radioBtnNombre.isChecked())

        if cod_error == 0:
            mensaje_alerta('Inf', 'Ha terminado de guardar a los estudiantes en el HTML.')
        else:
            mensaje_alerta('Err', error)

    # Variables de las asignaturas
    def fn_guardar_asignatura(self):
        # Recoge las variables de la interfaz
        asignatura = self.ComboBoxAsignatura.currentText()
        plazas = self.PlazasText.value()
        num_sesiones = self.NumSesionesText.value()
        horario = coger_horarios(self.AreaHorarios, asignatura)
        num_subgrupos = self.NumSubgruposText.value()
        semana_inicial = self.SemanaInicialText.value()
        
        # Comprueba que las variables esten correctamente introducidas por las interfaz y las añade en el txt
        if asignatura != '':
            if horario != []:
                if plazas != 0 and num_sesiones != 0 and num_subgrupos != 0 and semana_inicial != 0:
                    asignaturas = list()

                    # Abre el fichero en modo lectura
                    f = open('asignaturas.txt', 'r')
                    # Lee todo el fichero y lo guarda en la lista
                    for asig in f:
                        asignaturas.append(asig.strip('\n'))
                    f.close()

                    # Comprueba si la asignatura introducida esta en la lista
                    encontrado = False
                    i = 0
                    while not encontrado and i < len(asignaturas):
                        # Si encuentra la asignatura introducida como esta repetida se sobreescribe
                        if asignaturas[i].split('-')[0] == unidecode(asignatura):
                            asignaturas[i] = unidecode(asignatura) + '-' + str(plazas) + '-' + str(num_sesiones) + '-' + str(horario) + '-' + str(num_subgrupos) + '-' + str(semana_inicial)
                            encontrado = True
                        i += 1
                    
                    # Si no se encuentra una misma asignatura se añade a las asignaturas ya existentes
                    if not encontrado:
                        texto = unidecode(asignatura) + '-' + str(plazas) + '-' + str(num_sesiones) + '-' + str(horario) + '-' + str(num_subgrupos) + '-' + str(semana_inicial)
                        asignaturas.append(texto)
                    
                    # Abre el fichero en modo escritura
                    f = open('asignaturas.txt', 'w')
                    # Escribe en el fichero las asignaturas
                    for txt in asignaturas:
                        f.write(txt + '\n')
                    f.close()
            
                    mensaje_alerta('Inf', 'Añadido correctamente.')

        # Si se han introducido mal los datos salta un error
                else:
                    mensaje_alerta('Err', 'No se han introducido bien los datos.')
            else:
                mensaje_alerta('Err', 'No se ha marcado ningun horario.')
        else:
            mensaje_alerta('Err', 'No se ha asignado la asignatura.')

    # Guarda las aulas
    def fn_guardar_aulas(self):

        asignaturaCompartida1 = self.ComboBoxAsignaturaComparten1.currentText()
        asignaturaCompartida2 = self.ComboBoxAsignaturaComparten2.currentText()

        # Comprueba que dos asignaturas no sean iguales o esten vacias y si coinciden los horarios
        if comprobarAsignaturas(unidecode(asignaturaCompartida1), unidecode(asignaturaCompartida2)):
            
            aulaCompartida = list()
            encontrado = False

            # Abre el fichero en modo escritura
            f = open('compartenAula.txt', 'r')
            # Escribe en el fichero las asignaturas
            for aula in f:
                aulaCompartida.append(aula.strip('\n'))
            f.close()

            # Abre el fichero en modo escritura
            f = open('compartenAula.txt', 'w')
            # Escribe en el fichero las asignaturas
            for aula in aulaCompartida:
                if unidecode(asignaturaCompartida1) in aula.split('/') and unidecode(asignaturaCompartida2) in aula.split('/'):
                    mensaje_alerta('War', 'Las asignaturas con aulas compartidas ya estan en el txt.')
                    encontrado = True
                f.write(aula + '\n')
            
            # Si no se ha encontrado que las asignaturasCompartidas estan ya en el txt se añaden
            if not encontrado and not (asignaturaCompartida1 == '' or asignaturaCompartida2 == ''):
                f.write(unidecode(asignaturaCompartida1) + '/' + unidecode(asignaturaCompartida2) + '\n')
                aulaCompartida.append(unidecode(asignaturaCompartida1) + '/' + unidecode(asignaturaCompartida2))

            f.close()

            self.TablaCompartenAula.clearContents()
            self.TablaCompartenAula.setRowCount(len(aulaCompartida))

            for index, aula in enumerate(aulaCompartida):
                # Añade el texto en la tabla
                asignaturas_compartidas = aula.strip('\n').split('/')
                self.TablaCompartenAula.setItem(index, 0, QTableWidgetItem(lee_grupos.nombre_asignaturas[asignaturas_compartidas[0]] + '/' + lee_grupos.nombre_asignaturas[asignaturas_compartidas[1]]))

    # Manda un mensaje de confirmacion para borrar los laboratorios
    def fn_borrar_laboratorios(self):

        global borrar

        borrar = False

        alerta = QMessageBox(QMessageBox.Question, 'Alerta', '¿Estas seguro de querer borrar permanetemente la lista de los laboratorios?')
        alerta.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        alerta.buttonClicked.connect(borrar_laboratorios)
        alerta.exec()

        # Si se borran las asignaturas se reiniciaran las variables de interfaz 
        if borrar:
            self.ComboBoxAsignatura.clear()
            self.fn_cargar_asignaturas()
            self.ComboBoxAsignatura.setCurrentIndex(-1)
            self.AreaHorarios.setWidget(QWidget())
            self.PlazasText.setValue(0)
            self.NumSesionesText.setValue(0)
            self.NumSubgruposText.setValue(0)
            self.SemanaInicialText.setValue(0)
    
    # Borrar las aulas Compartidas
    def fn_borrar_aulas(self):

        # Guarda todos los horarios seleccionados
        if self.TablaCompartenAula.selectedItems():
            aulasCompartidas = list()

            # Guarda todos los horarios seleccionados
            for i in range(self.TablaCompartenAula.rowCount()):
                if not self.TablaCompartenAula.item(i, 0).isSelected():
                    aulasCompartidas.append(unidecode(self.TablaCompartenAula.item(i, 0).text()))
            
            self.TablaCompartenAula.clearContents()
            self.TablaCompartenAula.setRowCount(len(aulasCompartidas))
            
            # Abre el fichero en modo escritura
            f = open('compartenAula.txt', 'w')
            # Escribe en el fichero las aulas compartidas
            for index, aula in enumerate(aulasCompartidas):
                # Añade el texto en la tabla
                asignaturas_compartidas = aula.strip('\n').split('/')
                self.TablaCompartenAula.setItem(index, 0, QTableWidgetItem(lee_grupos.nombre_asignaturas[asignaturas_compartidas[0]] + '/' + lee_grupos.nombre_asignaturas[asignaturas_compartidas[1]]))
                f.write(aula)
            f.close()

        else:
            mensaje_alerta('War', 'No se ha seleccionado ningúna aula compartida.')

    # Selecciona la asignatura que se ha seleccionado en el TreeView
    def fn_selecciona_asignatura(self, index):
        # Recoge la asignatura seleccionada en el TreeView
        asignatura = self.ArbolAsignaturas.selectedIndexes()[0].data(0)
        # Si selecciona una opcion diferente a un 'Cuatrimestre' se añade la asignatura al label
        if asignatura.find('Cuatrimestre') == -1:
            self.lblAsignaturaAsignada.setText(asignatura)
            self.ComboBoxHoras.clear()
            self.ComboBoxDias.setCurrentIndex(-1)

            inserta_horarios_tabla(self.TablaHorarios, asignatura)
            inserta_grupos(self.ComboBoxGrupos, asignatura)
    
    # Añade los horarios
    def fn_anadir_horas(self):
        
        global grupos

        # Grupo asignado en el comboBox
        grupo = self.ComboBoxGrupos.currentText()
        self.ComboBoxHoras.clear()
        self.ComboBoxDias.setCurrentIndex(-1)

        horas = list()

        for fila in range(len(grupos)):
            for col in range(len(grupos[fila])):
                if grupo == grupos[fila][col]:
                    if fila == 0 or fila == 2:
                        horas.append('09:30')
                        horas.append('11:30')
                    if fila == 1 or fila == 2:
                        horas.append('15:30')
                        horas.append('17:30')

        self.ComboBoxHoras.addItems(horas)
        # La interfaz no carga ninguna opcion del comboBox (comienza vacio)
        self.ComboBoxHoras.setCurrentIndex(-1)

    # Borra los horarios del txt de la asignatura seleccionada
    def fn_borrar_horarios(self):
        # Label donde esta guardada la asignatura
        asignatura = self.lblAsignaturaAsignada.text()
        # Si no se ha seleccionado una asignatura salta un mensaje de error
        if asignatura != '...':
            
            asignaturas1 = list()
            asignaturas2 = list()
            encontrado = False

            # Abre el fichero en modo lectura
            f = open('horarios.txt', 'r')
            # Lee todo el fichero y lo guarda en la lista
            for asig in f:
                # Añade todas las asignaturas a la lista menos la que se quiere borrar
                if asig.split('-')[0] != unidecode(asignatura):
                    asignaturas1.append(asig.strip('\n'))
                else:
                    encontrado = True
            f.close()

            # Comprueba si se ha encontrado la asignatura en el fichero
            if encontrado:
                # Abre el fichero en modo lectura
                f = open('asignaturas.txt', 'r')
                # Lee todo el fichero y lo guarda en la lista
                for asig in f:
                    # Añade todas las asignaturas a la lista menos la que se quiere borrar
                    if asig.split('-')[0] != unidecode(asignatura):
                        asignaturas2.append(asig.strip('\n'))
                f.close()

                # Abre el fichero en modo escritura
                f = open('horarios.txt', 'w')
                # Escribe en el fichero las asignaturas
                for txt in asignaturas1:
                    f.write(txt + '\n')
                f.close()

                # Abre el fichero en modo escritura
                f = open('asignaturas.txt', 'w')
                # Escribe en el fichero las asignaturas
                for txt in asignaturas2:
                    f.write(txt + '\n')
                f.close()
                
                self.TablaHorarios.clearContents()
                self.TablaHorarios.setRowCount(0)
                
                mensaje_alerta('Inf', 'Se ha borrado corectamente.')
            else:
                mensaje_alerta('Inf', 'La asignatura no se encuentra en el fichero.')
        else:
            mensaje_alerta('Err', 'No se ha asignado la asignatura.')

    # Borra el horario seleccionado
    def fn_borrar_horario_seleccionado(self):
        # Label donde esta guardada la asignatura
        asignatura = self.lblAsignaturaAsignada.text()

        # Si no se ha seleccionado una asignatura salta un mensaje de error
        if asignatura != '...':
            
            asignaturas = list()
            horarios = list()
            txt = ''

            # Guarda todos los horarios seleccionados
            for item in self.TablaHorarios.selectedItems():
                horarios.append(item.text())

            # Si se han seleccionado algun horario entra
            if horarios:
                # Abre el fichero en modo lectura
                f = open('horarios.txt', 'r')
                # Lee todo el fichero y lo guarda en la lista
                for asig in f:
                    # Añade todas las asignaturas a la lista menos la que se quiere borrar
                    if asig.split('-')[0] == unidecode(asignatura):
                        aux = json.loads(asig.split('-')[1].replace('\'','\"'))

                        # Quita todos los horarios seleccionados
                        for horario in horarios:
                            aux.remove(horario)
                        
                        # Si la lista no esta vacia se añadira al txt 
                        if aux:
                            txt = asig.split('-')[0] + '-' + str(aux)
                        else:
                            txt = ''
                    else:
                        txt = asig.strip('\n')

                    if txt != '':
                        asignaturas.append(txt)

                f.close()

                # Abre el fichero en modo escritura
                f = open('horarios.txt', 'w')
                # Escribe en el fichero las asignaturas
                for txt in asignaturas:
                    f.write(txt + '\n')
                f.close()

                # Inserta los horarios en la tabla
                inserta_horarios_tabla(self.TablaHorarios, asignatura)
                
                mensaje_alerta('Inf', 'Se ha borrado corectamente.')
            else:
                mensaje_alerta('War', 'No se ha seleccionado ningún horario.')

    # Guarda los horarios de las asignaturas
    def fn_guardar_horarios(self):
        # Label donde esta guardada la asignatura
        asignatura = self.lblAsignaturaAsignada.text()
        # Si no se ha seleccionado una asignatura salta un mensaje de error
        if asignatura != '...':
            # Si no se ha seleccionado un dia o una hora salta un mensaje de error
            if self.ComboBoxGrupos.currentIndex() != -1 and self.ComboBoxDias.currentIndex() != -1 and self.ComboBoxHoras.currentIndex() != -1:
                # Guarda el grupo, el dia y la hora
                grupo = self.ComboBoxGrupos.currentText().upper()
                dia = self.ComboBoxDias.currentText().upper()[:2]
                hora = self.ComboBoxHoras.currentText().split(':')[0]

                asignaturas = list()

                # Abre el fichero en modo lectura
                f = open('horarios.txt', 'r')
                # Lee todo el fichero y lo guarda en la lista
                for asig in f:
                    asignaturas.append(asig.strip('\n'))
                f.close()

                # Comprueba si la asignatura instroducida esta en la lista
                encontrado = False
                repetido = False
                i = 0
                while not encontrado and i < len(asignaturas):
                    # Si la asignatura introducida esta repetida se sobreescribe
                    if asignaturas[i].split('-')[0] == unidecode(asignatura):
                        # Se recoge los horarios de las asignaturas del fichero
                        # Como la cadena esta en formato string se traduce a formato lista
                        horarios = json.loads(asignaturas[i].split('-')[1].replace('\'','\"'))

                        # Comprueba si el horario introducido esta en la lista
                        j = 0
                        while not encontrado and j < len(horarios):
                            # Si el horario introducido esta repetido salta un aviso
                            if horarios[j] == (grupo + '/' + dia + hora):
                                repetido = True
                                encontrado = True
                            j += 1
                        
                        # Si el horario introducido no esta en la lista se añade
                        if not encontrado:
                            horarios.append(grupo + '/' + dia + hora)
                            horarios = ordenar_horarios(horarios)
                            asignaturas[i] = unidecode(asignatura) + '-' + str(horarios)
                            encontrado = True
                    i += 1
                
                # Si no se encuentra una misma asignatura se añade a las asignaturas ya existentes
                if not encontrado:
                    # Se traduce la asignatura a tipo string  
                    texto = unidecode(asignatura) + '-[\'' + (grupo + '/' + dia + hora) + '\']'
                    asignaturas.append(texto)

                # Abre el fichero en modo escritura
                f = open('horarios.txt', 'w')
                # Escribe en el fichero las asignaturas
                for txt in asignaturas:
                    f.write(txt + '\n')
                f.close()

                if repetido:
                    mensaje_alerta('Inf', 'Ya se ha asignado este horario a esta asignatura.')
                else:
                    inserta_horarios_tabla(self.TablaHorarios, asignatura)
                    mensaje_alerta('Inf', 'Añadido correctamente.')
            else:
                mensaje_alerta('Err', 'No se ha asignado el grupo, el día o la hora.')
        else:
            mensaje_alerta('Err', 'No se ha asignado la asignatura.')

    # Reinicia los valores de cada pestaña
    def fn_reinicia_pestanas(self, index):
        nombreTab = self.tabWidget.tabText(index)

        # Pestaña Asignación
        if nombreTab == 'Asignación':
            if not self.generadosExcel:
                self.BtnGuardarExcel.setEnabled(False)
                self.BtnCrearHTML.setIcon(QtGui.QIcon('./img/excelDes.png'))
                self.BtnCrearHTML.setEnabled(False)
                self.BtnCrearHTML.setIcon(QtGui.QIcon('./img/htmlDes.png'))
                self.radioBtnNombre.setChecked(True)
                self.lblSeleccionFormato.setEnabled(False)
                self.radioBtnMatricula.setEnabled(False)
                self.radioBtnNombre.setEnabled(False)
        # Pestaña Disponibilidad
        elif nombreTab == 'Disponibilidad':
            self.ComboBoxGrupos.setCurrentIndex(-1)
            self.ComboBoxGrupos.clear()
            self.ComboBoxDias.setCurrentIndex(-1)
            self.ComboBoxHoras.setCurrentIndex(-1)
            self.ComboBoxHoras.clear()
            self.ArbolAsignaturas.collapseAll()
            self.lblAsignaturaAsignada.setText('...')
            self.TablaHorarios.clearContents()
            self.TablaHorarios.setRowCount(0)
        # Pestaña Agregar Laboratorios
        elif nombreTab == 'Agregar Labs':
            self.ComboBoxAsignatura.clear()
            self.ComboBoxAsignaturaComparten1.clear()
            self.ComboBoxAsignaturaComparten2.clear()
            self.TablaCompartenAula.clearContents()
            self.fn_cargar_asignaturas()
            self.ComboBoxAsignatura.setCurrentIndex(-1)
            self.ComboBoxAsignaturaComparten1.setCurrentIndex(-1)
            self.ComboBoxAsignaturaComparten2.setCurrentIndex(-1)
            self.AreaHorarios.setWidget(QWidget())
            self.PlazasText.setValue(0)
            self.NumSesionesText.setValue(0)
            self.NumSubgruposText.setValue(0)
            self.SemanaInicialText.setValue(0)
        # Pestaña Agregar Grupos
        elif nombreTab == 'Agregar Grupos':
            pass
        # Pestaña Calendario
        elif nombreTab == 'Calendarios':
            self.TxtNumMatricula.setPlainText('')
            self.TxtIdentificador.setPlainText('')
            self.TablaHorario.clearContents()
            self.lblHorarios.setText('')
    
    # Carga los horarios de las asignaturas en el comboBox
    def fn_cargar_asignaturas(self):
        asignaturas = list()
        aulaCompartida = list()

        # Abre el fichero en modo lectura
        f = open('horarios.txt', 'r')
        # Lee todo el fichero y lo guarda en la lista
        for asignatura in f:
            asignaturas.append(lee_grupos.nombre_asignaturas[asignatura.strip('\n').split('-')[0]])
        f.close()

        # Abre el fichero en modo lectura
        f = open('compartenAula.txt', 'r')
        # Lee todo el fichero y lo guarda en la lista
        for aula in f:
            aulaCompartida.append(aula)
        f.close()

        self.ComboBoxAsignatura.insertItems(0, asignaturas)
        self.ComboBoxAsignaturaComparten1.insertItems(0, asignaturas)
        self.ComboBoxAsignaturaComparten2.insertItems(0, asignaturas)
        
        self.TablaCompartenAula.setRowCount(len(aulaCompartida))

        for index, aula in enumerate(aulaCompartida):
            asignaturas_compartidas = aula.strip('\n').split('/')
            self.TablaCompartenAula.setItem(index, 0, QTableWidgetItem(lee_grupos.nombre_asignaturas[asignaturas_compartidas[0]] + '/' + lee_grupos.nombre_asignaturas[asignaturas_compartidas[1]]))
    
    # Añade los horarios al ScrollArea
    def fn_anadir_horarios(self, index):
        # Recoge la asignatura seleccionada en el ComboBox
        asignatura = self.ComboBoxAsignatura.currentText()
        
        if asignatura != '':
            
            horarios = list()
            horario_asignatura = list()
            plazas_sesion = 0
            num_sesiones = 0
            num_subgrupos = 0
            semana_inicial = 0

            # Abre el fichero en modo lectura
            f = open('horarios.txt', 'r')
            # Lee todo el fichero y lo guarda en la lista
            for asig in f:
                if asig.strip('\n').split('-')[0] == unidecode(asignatura):
                    # Se recoge los horarios de las asignaturas del fichero
                    # Como la cadena esta en formato string se traduce a formato lista
                    horarios = json.loads(asig.split('-')[1].replace('\'','\"'))
            f.close()

            # Crea checkbox con los horarios de la asignatura seleccionada
            widget = QWidget()
            layout = QVBoxLayout(widget)
            self.AreaHorarios.setWidget(widget)

            # Abre el fichero en modo lectura
            f = open('asignaturas.txt', 'r')
            # Lee todo el fichero y guarda los horarios de las asignaturas
            for asig in f:
                if asig.split('-')[0] == unidecode(asignatura):
                    plazas_sesion = int(asig.split('-')[1])
                    num_sesiones = int(asig.split('-')[2])
                    horario_asignatura = json.loads(asig.split('-')[3].replace('\'','\"'))
                    num_subgrupos = int(asig.split('-')[4])
                    semana_inicial = int(asig.split('-')[5])
            f.close()

            for i, horario in enumerate(horarios):
                # Si hay algun horario repetido no se añade al layout
                if not any(elem.split('/')[1] == horarios[i].split('/')[1] for elem in horarios[:i]):
                    checkBox = QCheckBox(asignatura + ' ' + horario.split('/')[1])
                    # Si el horario ya esta en el txt se checkea
                    if any(horario.split('/')[1] == aux.split('/')[1] for aux in horario_asignatura):
                        checkBox.setCheckState(2)
                    layout.addWidget(checkBox)
            
            self.PlazasText.setValue(plazas_sesion)
            self.NumSesionesText.setValue(num_sesiones)
            self.NumSubgruposText.setValue(num_subgrupos)
            self.SemanaInicialText.setValue(semana_inicial)

    # Obtiene la ruta del archivo pdf que se quiere cargar
    def fn_buscar_archivos(self):
        nombre_fichero = QFileDialog.getOpenFileName(self, 'Abrir Fichero', '', 'PDF (*.pdf)')
        self.TxtFicheroPath.setText(nombre_fichero[0])

    
    # Carga los laboratorios que tiene el alumno en tablaHorario
    def fn_carga_calendario(self):
        # Recoge el numero de matricula metido en PlainText
        matricula = self.TxtNumMatricula.toPlainText()
        # Lee los numeros de matricula del excel con los grupos de laboratorio asignados
        lista_subgrupos = pd.read_excel('lista_subgrupos.xlsx', dtype={'Nº Expediente en Centro':str})
        lista_subgrupos.set_index('Nº Expediente en Centro', inplace=True)
        self.TablaHorario.clearContents()
        
        # Abre el excel con las semanas de inicio
        semanas = pd.read_excel(lee_grupos.config.get('ARCHIVOS', 'SEMANAS'), index_col=0)

        # Se crea un diccionario para recorrer los dias y las horas
        dias = {'LU':0, 'MA':1, 'MI':2, 'JU':3, 'VI':4}
        dia_semana = {'LU' : 'LUNES', 'MA' : 'MARTES', 'MI' : 'MIERCOLES', 'JU' : 'JUEVES', 'VI' : 'VIERNES'}
        horas = {'09':0, '11':1, '15':2, '17':3}

        texto = ''

        # Comprueba que exista el subgrupo para esta asignatura
        if matricula in lista_subgrupos.index:
            for col in lista_subgrupos.columns:
                if 'subgrupo' in col:
                    
                    dia = lista_subgrupos.loc[matricula][col]

                    # Se descartan las celdas vacias y con '-' quedando solo los subgrupos compuestos dias+horas (MI09->MI+09)
                    if pd.notna(dia) and dia != '-':
                        num_sesiones = -1
                        num_subgrupos = -1
                        semana_inicial = -1
                        semanas_asignatura = list()
                        
                        if col.split('_')[1] in lee_grupos.config.get('CUATRIMESTRE', 'IMPAR').split(','):
                            semanas_asignatura = list(semanas[dia_semana[dia.split('-')[0][:2]]])[:int(lee_grupos.config.get('SEMANAS', 'NUM_SEMANAS')) + 1]
                        else:
                            semanas_asignatura = list(semanas[dia_semana[dia.split('-')[0][:2]]])[int(lee_grupos.config.get('SEMANAS', 'NUM_SEMANAS')) + 1:]

                        # Abre el fichero en modo lectura
                        f = open('asignaturas.txt', 'r')
                        # Lee todo el fichero y guarda los horarios de las asignaturas
                        for asig in f:
                            if asig.split('-')[0] == col.split('_')[1]:
                                num_sesiones   = int(asig.split('-')[2])
                                num_subgrupos  = int(asig.split('-')[4])
                                semana_inicial = int(asig.split('-')[5])
                        f.close()
                    
                        # Se comprueba que la celda donde se va a insertar la asignatura este vacia y se añade
                        if self.TablaHorario.item(horas[dia[2:4]], dias[dia[:2]]) == None:
                            self.TablaHorario.setItem(horas[dia[2:4]], dias[dia[:2]], QTableWidgetItem(lee_grupos.nombre_asignaturas[col.split('_')[1]].upper()))
                            self.TablaHorario.item(horas[dia[2:4]], dias[dia[:2]]).setBackground(QColor(17, 59, 228, 255)) # Color de fondo de la celda
                        # Si ya se ha agregado una asignatura previamente, se guarda junto la que se quiere introducir y se añaden las dos
                        else:
                            text = lee_grupos.nombre_asignaturas[col.split('_')[1]].upper() + ' / ' + self.TablaHorario.item(horas[dia[2:4]], dias[dia[:2]]).text().upper()
                            self.TablaHorario.setItem(horas[dia[2:4]], dias[dia[:2]], QTableWidgetItem(text))
                            self.TablaHorario.item(horas[dia[2:4]], dias[dia[:2]]).setBackground(QColor(239, 108, 0, 255)) #Números del calendario
                        # Se comprueba que exista la asignatura en el txt                            
                        if num_sesiones != -1:
                            texto += lee_grupos.nombre_asignaturas[col.split('_')[1]].upper() + ':'
                            for i in range(num_sesiones):
                                texto += '   ' + lee_grupos.traduce_meses(semanas_asignatura[semana_inicial + (ord(dia[-1]) - ord('A') + num_subgrupos * i) - 1].strftime('%d %B'))
                            
                            texto += '\n'
                        else:
                            texto = ''
                            break
        
            self.lblHorarios.setText(texto)
            
            # Comprueba que se hayan encontrado todas las asignaturas en el fichero asignaturas.txt
            if texto == '':
                self.TablaHorario.clearContents()
                self.lblHorarios.setText(texto)
                mensaje_alerta('Err', 'No se encuentran todas las asignaturas del alumno en el fichero asignaturas.txt')
        else:
            self.TablaHorario.clearContents()
            self.lblHorarios.setText(texto)
            mensaje_alerta('Err', 'No existe este Nº matricula')
    
    # Crea el calendario anual del alumno
    def fn_calendario_anual_alumno(self):
        # Recoge el numero de matricula metido en PlainText
        matricula = self.TxtNumMatricula.toPlainText()

        if matricula != '':
            cod_error, error = lee_grupos.crea_calendario_anual_alumno(matricula)
            
            if cod_error == 0:
                mensaje_alerta('Inf', 'Se ha creado el calendario anual del alumno')
            else:
                mensaje_alerta('Err', error)
        else:
            mensaje_alerta('Err', 'Se ha introducido el númerod de matrícula del alumno')

    # Crea el calendario anual del profesor
    def fn_calendario_anual_profesor(self):
        # Recoge el identificador metido en PlainText
        identificador = self.TxtIdentificador.toPlainText()

        if identificador != '':
            cod_error, error = lee_grupos.crea_calendario_anual_profesor(identificador)
            
            if cod_error == 0:
                mensaje_alerta('Inf', 'Se ha creado el calendario anual del profesor')
            else:
                mensaje_alerta('Err', error)
        else:
            mensaje_alerta('Err', 'Se ha introducido el identificador del profesor')

# Crea una alerta
def mensaje_alerta(icono, texto):
    mensaje_icono = {
        'Que' : QMessageBox.Question,
        'Inf' : QMessageBox.Information,
        'Err' : QMessageBox.Critical,
        'War' : QMessageBox.Warning
    }
    alerta = QMessageBox(mensaje_icono[icono], 'Alerta', texto)
    alerta.exec()

# Ordena los horarios
def ordenar_horarios(horarios):
    aux = list()
    horarios.sort()

    dias = ['LU', 'MA', 'MI', 'JU', 'VI']

    for dia in dias:
        for horario in horarios:
            if horario.split('/')[1][:2] == dia:
                aux.append(horario)

    return aux

# Recoge los horarios del ScrollArea
def coger_horarios(AreaHorarios, asignatura):
    
    lista_horarios = list()

    if asignatura != '':
        horarios = list()
        layout = AreaHorarios.widget().layout()

        # Abre el fichero en modo lectura
        f = open('horarios.txt', 'r')
        # Lee todo el fichero y lo guarda en la lista
        for asig in f:
            if asig.split('-')[0] == unidecode(asignatura):
                # Se recoge los horarios de las asignaturas del fichero
                # Como la cadena esta en formato string se traduce a formato lista
                horarios = json.loads(asig.split('-')[1].replace('\'','\"'))
        f.close()

        # Obtiene los horarios que se han seleccionado
        if layout != None:
            for index in range(layout.count()):
                widget = layout.itemAt(index).widget()
                if widget.isChecked():
                    for horario in horarios:
                        if widget.text().split(' ')[-1] in horario:
                            # Coge el horario del laboratorio asignado
                            lista_horarios.append(horario)

    return lista_horarios

# Borra lista de los laboratorios
def borrar_laboratorios(opcion):
    global borrar

    if opcion.text() == 'OK':
        f = open('asignaturas.txt', 'w')
        f.close()
        borrar = True

# Inserta los horarios en la tabla de la asignatura asignada
def inserta_horarios_tabla(tabla_horarios, asignatura):
    # Borra el contenido de la TablaHorarios
    tabla_horarios.clearContents()
    tabla_horarios.setRowCount(0)

    # Abre el fichero en modo lectura
    f = open('horarios.txt', 'r')

    # Lee todo el fichero
    for asig in f:
        # Comprueba si existe la asignatura en el fichero
        if asig.split('-')[0] == unidecode(asignatura):
            # Guarda los horarios de las asignaturas
            # Como la cadena esta en formato string se traduce a formato lista
            horarios = json.loads(asig.split('-')[1].replace('\'','\"'))
            # Añade los huecos para que entren los horarios
            tabla_horarios.setRowCount(len(horarios))
            for i, horario in enumerate(horarios):
                # Crea el texto y lo centra
                # texto = QTableWidgetItem(horario)
                # texto.setTextAlignment(Qt.AlignCenter)
                # Añade el texto en la tabla
                tabla_horarios.setItem(0, i, QTableWidgetItem(horario))
            break
    f.close()

# Inserta los grupos en el comboBox de la asignatura asignada
def inserta_grupos(comboBox_grupos, asignatura):
    
    global grupos

    asignaturas = dict()
    asignatura = unidecode(asignatura)

    # Recorre las asignaturas del archivo de configuracion
    for asig in lee_grupos.config.options('ASIGNATURAS'):
        # Inicializa el diccionario
        asignaturas[asig.replace('_', ' ')] = [[False, False, False, False, False], [False, False, False, False, False], [False, False, False, False, False, False, False, False]]
        # Recorre los grupos de cada asignatura
        for grupo in lee_grupos.config.get('ASIGNATURAS', asig).split(','):
            for i in range(len(asignaturas[asig.replace('_', ' ')])):
                for j in range(len(asignaturas[asig.replace('_', ' ')][i])):
                    # Compara cada grupo de la asignatura con la variable grupos
                    if grupos[i][j] == grupo:
                        asignaturas[asig.replace('_', ' ')][i][j] = True
                        
    # Borra los grupos anteriores
    comboBox_grupos.clear()

    # Recorre el vector de la asignatura seleccionada
    for i in range(len(asignaturas[asignatura.lower()])):
        for j in range(len(asignaturas[asignatura.lower()][i])):
            # Si el grupo es True se agrega al comboBox
            if asignaturas[asignatura.lower()][i][j]:
                comboBox_grupos.addItem(grupos[i][j])
    
    # La interfaz no carga ninguna opcion del comboBox (comienza vacio)
    comboBox_grupos.setCurrentIndex(-1)

# Comprueba que dos asignaturas no sean iguales o esten vacias y si coinciden los horarios
def comprobarAsignaturas(asignatura1, asignatura2):

    # Comprueba que las dos asignaturas se hayan seleccionado
    if asignatura1 != '' and asignatura2 != '':
        # Comprueba que las dos asignaturas no sean iguales
        if asignatura1 == asignatura2:
            mensaje_alerta('Err', 'Se han introducido las mismas asignaturas.')
            return False
        # Si se han añadido correctamente las asignaturas comprueba si coinciden los horarios
        else:
            horario1 = ''
            horario2 = ''

            # Abre el fichero en modo lectura
            f = open('asignaturas.txt', 'r')
            # Lee todo el fichero y lo guarda los horarios de las asignaturas
            for asig in f:
                if asignatura1 == asig.split('-')[0]:
                    horario1 = json.loads(asig.split('-')[3].replace('\'','\"'))
                elif asignatura2 == asig.split('-')[0]:
                    horario2 = json.loads(asig.split('-')[3].replace('\'','\"'))
            
            f.close()

            # Quita los grupos de matricula de los horarios
            horario1 = list(set([horario.split('/')[1] for horario in horario1]))
            horario2 = list(set([horario.split('/')[1] for horario in horario2]))

            # Comprueba si algun horario coincide
            if any(horario in horario2 for horario in horario1):
                mensaje_alerta('Err', 'Coinciden los horarios de las asignaturas que comparten las aulas.')
                return False

    else:
        # Comprueba si alguna de las asignaturas no esta vacia
        if asignatura1 != '' or asignatura2 != '':
            mensaje_alerta('War', 'Solo se ha seleccionado una opción en "Asignaturas que comparten aula".')

    return True

# Comprueba si estan creadas las carpetas y los archivos y sino estan creados se crean 
def comprobar_archivos(config):

    # Cromprueba que todas las carpetas esten creadas y sino las crea
    for carpeta in lee_grupos.config.options('RUTAS'):
        if not os.path.isdir(config.get('RUTAS', carpeta)):
            os.mkdir(config.get('RUTAS', carpeta))

    # Cromprueba que los archivos esten creados y sino las crea
    if not os.path.isfile('asignaturas.txt'):
        f = open('asignaturas.txt', 'w')
        f.close()
    if not os.path.isfile('horarios.txt'):
        f = open('horarios.txt', 'w')
        f.close()
    if not os.path.isfile('compartenAula.txt'):
        f = open('compartenAula.txt', 'w')
        f.close()

if __name__ == '__main__':

    # Comprueba si estan creadas las carpetas y archivos
    comprobar_archivos(lee_grupos.config)

    app = QApplication(sys.argv)
    stream = QtCore.QFile('DarkStyle.qss')
    stream.open(QtCore.QIODevice.ReadOnly)
    app.setStyleSheet(QtCore.QTextStream(stream).readAll())
    interfaz = GUI()
    interfaz.show()
    sys.exit(app.exec_())