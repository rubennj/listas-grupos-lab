# -*- coding: utf-8 -*-
"""
@author: Andrea Magro Canas

Descripción del código: Programación del entorno gráfico.

"""

import sys
from PyQt5 import uic, QtCore, QtGui
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import QPrintPreviewDialog
from PyQt5.QtGui import (QColor)
import pandas as pd
import json

import lee_grupos

# Grupos de los laboratorios
grupos = [
    # Mañana
    ['A207', 'E208', 'A309', 'M306', 'A408'],
    # Tarde
    ['A204', 'E205', 'A302', 'M301', 'A404'],
    # Dobles Grados y Optativas
    ['EE208', 'EE309', 'EE403', 'EE507', 'D307', 'DM306', 'Q308']
]

class GUI(QMainWindow):
    
    def __init__(self):
        super().__init__()
        # Carga la interfaz en nuestra clase
        uic.loadUi("interfaz.ui", self)
        # Eventos de los botones
        self.setWindowIcon(QtGui.QIcon('./img/icono.png'))
        self.BtnAyudaDisponibilidad.clicked.connect(self.fn_ayuda)
        self.BtnAyudaDisponibilidad.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAyudaAgregarLabs.clicked.connect(self.fn_ayuda)
        self.BtnAyudaAgregarLabs.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAyudaAsignacion.clicked.connect(self.fn_ayuda)
        self.BtnAyudaAsignacion.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAyudaCalendarioAlumnos.clicked.connect(self.fn_ayuda)
        self.BtnAyudaCalendarioAlumnos.setStyleSheet("QPushButton{border-color: rgba(40, 135, 199, 70%);}")
        self.BtnAsignacion.clicked.connect(self.fn_asignar_grupos)
        self.BtnGuardarExcel.setEnabled(False)
        self.BtnGuardarExcel.clicked.connect(self.fn_guarda_excel)
        self.BtnCrearPDF.setEnabled(False)
        self.BtnCrearPDF.clicked.connect(self.fn_guarda_pdf)
        self.BtnAceptar.clicked.connect(self.fn_guardar_asignatura)
        self.BtnBorrarLabs.clicked.connect(self.fn_borrar_laboratorios)
        self.BtnBorrarAulas.clicked.connect(self.fn_borrar_aulas)
        self.BtnBorrarHorario.clicked.connect(self.fn_borrar_horarios)
        self.BtnGuardaHorario.clicked.connect(self.fn_guardar_horarios)
        self.BtnGuardaAsignaturas.clicked.connect(self.fn_carga_asignaturas)
        self.BtnExportarPDF.clicked.connect(self.fn_exportar_PDF)
        # Evento de las pestañas
        self.tabWidget.currentChanged.connect(self.fn_reinicia_pestanas)
        # Eventos TreeView
        self.ArbolAsignaturas.doubleClicked.connect(self.fn_selecciona_asignatura)
        # Eventos ComboBox
        self.ComboBoxAsignatura.textActivated.connect(self.fn_anadir_horarios)
        self.ComboBoxGrupos.textActivated.connect(self.fn_anadir_horas)
        # Evento de la Tabla
        self.TablaHorario.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.TablaHorario.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
    
    # Carga una alerta en el botón de ayuda
    def fn_ayuda(self):
        mensaje_alerta('Que', 'PESTAÑA 1: \nSe define la disponibilidad de las aulas proporcionadas por secretaria donde se van a cursas las diferentes asignaturas.\n'+
                        'PESTAÑA 2: \nSe agregan los laboratorios introduciendo la asignatura de la que se va a cursar la práctica y su información.\n'+
                        'PESTANA 3: \nSe asignan los estudiantes. Se genera el Excel con las listas de los alumnos y sus laboratorios. También se puede crear un PDF similar a las listas que se publican cada año en la ETSIDI\n'+
                        'PESTAÑA 4: \nSe crea el calendario de prácticas de un alumno a través de su Nºmatrícula. Incluye el horario y las semanas que tiene que asistir a los laboratorios.\n'+
                        'EXCEL: \n- Localización: listas_apolo \n- Extensión: .xlsx \n- Formato:\n  1. Se deben respetar los nombres de las siguientes columnas: "Grupo de matrícula", "Apellidos", "Nombre" y "Nº Expediente en Centro".c' +
                        '  2. En estos campos no pueden haber celdas vacías \n  3. La tabla con los estudiantes debe estar bien delimitada, no deben existir bordes de más.listas de los laboratorios\n'+
                        '  4. El único Excel que no se coloca en listas_apolo es Semanas.xlsx que está en la carpeta general\n'+
                        '  Por tanto, los Excel que se tienen que introducir como entradas son los de las asignaturas (listas_apolo) y Semanas.xlsx')
    # Carga los datos
    def fn_asignar_grupos(self):
        cod_error, error = lee_grupos.asignar_grupos()
        if cod_error == 0:
            mensaje_alerta('Inf', 'Ha terminado de asignar a los estudiantes.')
            self.BtnGuardarExcel.setEnabled(True)
        else:
            mensaje_alerta('Err', error)
    
    # Guarda a los estudiantes en el excel
    def fn_guarda_excel(self):
        cod_error, error = lee_grupos.guardar_lista_grupos()
        if cod_error == 0:
            mensaje_alerta('Inf', 'Ha terminado de guardar a los estudiantes en el Excel.')
            self.BtnCrearPDF.setEnabled(True)
        else:
            mensaje_alerta('Err', error)
    
    # Guarda a los estudiantes en el excel
    def fn_guarda_pdf(self):
        cod_error, error = lee_grupos.crea_pdf_grupos_laboratorios(self.radioBtnNombre.isChecked())
        if cod_error == 0:
            mensaje_alerta('Inf', 'Ha terminado de guardar a los estudiantes en el PDF.')
        else:
            mensaje_alerta('Err', error)

    # Variables de las asginaturas
    def fn_guardar_asignatura(self):
        # Recoge las variables de la interfaz
        asignatura = self.ComboBoxAsignatura.currentText().lower()
        plazas = self.PlazasText.value()
        num_sesiones = self.NumSesionesText.value()
        horario = coger_horarios(self.AreaHorarios, asignatura)
        num_subgrupos = self.NumSubgruposText.value()
        semana_inicial = self.SemanaInicialText.value()
        asignaturaCompartida1 = self.ComboBoxAsignaturaComparten1.currentText().lower()
        asignaturaCompartida2 = self.ComboBoxAsignaturaComparten2.currentText().lower()
        
        # Comprueba que las variables esten correctamente y las introduce en el txt
        if asignatura != '':
            if horario != []:
                if plazas != 0 and num_sesiones != 0 and num_subgrupos != 0 and semana_inicial != 0:
                    if comprobarAsignaturas(asignatura, horario, asignaturaCompartida1, asignaturaCompartida2):
                        asignaturas = list()
                        aulaCompartida = list()

                        # Abre el fichero en modo lectura
                        f = open('asignaturas.txt', 'r')
                        # Lee todo el fichero y lo guarda en la lista
                        for asig in f:
                            asignaturas.append(asig.strip('\n'))
                        f.close()
                        
                        # Abre el fichero en modo escritura
                        f = open('compartenAula.txt', 'r')
                        # Escribe en el fichero las asignaturas
                        for aula in f:
                            aulaCompartida.append(aula)
                        f.close()

                        # Comprueba si la asignatura instroducida esta en la lista
                        encontrado = False
                        i = 0
                        while not encontrado and i < len(asignaturas):
                            # Si encuentra la asignatura introducida esta repetida se sobreescribe
                            if asignatura == asignaturas[i].split('-')[0]:
                                asignaturas[i] = asignatura + '-' + str(plazas) + '-' + str(num_sesiones) + '-' + str(horario) + '-' + str(num_subgrupos) + '-' + str(semana_inicial)
                                encontrado = True
                            i += 1
                        
                        # Si no se encuentra una misma asignatura se añade a las asignaturas ya existentes
                        if not encontrado:
                            texto = asignatura + '-' + str(plazas) + '-' + str(num_sesiones) + '-' + str(horario) + '-' + str(num_subgrupos) + '-' + str(semana_inicial)
                            asignaturas.append(texto)
                        
                        # Abre el fichero en modo escritura
                        f = open('asignaturas.txt', 'w')
                        # Escribe en el fichero las asignaturas
                        for txt in asignaturas:
                            f.write(txt + '\n')
                        f.close()

                        encontrado = False

                        # Abre el fichero en modo escritura
                        f = open('compartenAula.txt', 'w')
                        # Escribe en el fichero las asignaturas
                        for aula in aulaCompartida:
                            if asignaturaCompartida1 in aula.strip('\n').split('/') and asignaturaCompartida2 in aula.strip('\n').split('/'):
                                mensaje_alerta('War', 'Las asignaturas con aulas compartidas ya estan en el txt.')
                                encontrado = True
                            f.write(aula)
                        
                        # Si no se ha encontrado que las asignaturasCompartidas estan ya en el txt se añaden
                        if not encontrado and not (asignaturaCompartida1 == '' or asignaturaCompartida2 == ''):
                            f.write(asignaturaCompartida1 + '/' + asignaturaCompartida2 + '\n')
                            aulaCompartida.append(asignaturaCompartida1 + '/' + asignaturaCompartida2)

                        f.close()
                        
                        self.TablaCompartenAula.setRowCount(len(aulaCompartida))

                        for index, aula in enumerate(aulaCompartida):
                            self.TablaCompartenAula.setItem(index, 0, QTableWidgetItem(aula.strip('\n')))
                
                        mensaje_alerta('Inf', 'Añadido correctamente.')

        # Si se han introducido mal los datos salta un error
                else:
                    mensaje_alerta('Err', 'No se han introducido bien los datos.')
            else:
                mensaje_alerta('Err', 'No se ha marcado ningun horario.')
        else:
            mensaje_alerta('Err', 'No se ha asignado la asignatura.')

    # Manda un mensaje de confirmacion para borrar los laboratorios
    def fn_borrar_laboratorios(self):
        alerta = QMessageBox(QMessageBox.Question, 'Alerta', '¿Estas seguro de querer borrar permanetemente la lista de los laboratorios?')
        alerta.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        alerta.buttonClicked.connect(borrar_laboratorios)
        alerta.exec()
    
    # Manda un mensaje de confirmacion para borrar los laboratorios
    def fn_borrar_aulas(self):
        self.TablaCompartenAula.clearContents()
        self.TablaCompartenAula.setRowCount(0)
        f = open('compartenAula.txt', 'w')
        f.close()

    # Selecciona la asignatura que se ha seleccionado en el TreeView
    def fn_selecciona_asignatura(self, index):
        # Recoge la asigantura seleccionada en el TreeView
        asignatura = self.ArbolAsignaturas.selectedIndexes()[0].data(0)
        # Si selecciona una opocion diferente a un 'Cuatrimestre' se añade la asignatura al label
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
            
            asignaturas = list()
            encontrado = False

            # Abre el fichero en modo lectura
            f = open('horarios.txt', 'r')
            # Lee todo el fichero y lo guarda en la lista
            for asig in f:
                # Añade todas las asignaturas a la lista menos la que se quiere borrar
                if asignatura.lower() != asig.split('-')[0].lower():
                    asignaturas.append(asig.strip('\n'))
                else:
                    encontrado = True
            f.close()

            # Comprueba si se ha encontrado la asignatura en el fichero
            if encontrado:                    
                # Abre el fichero en modo escritura
                f = open('horarios.txt', 'w')
                # Escribe en el fichero las asignaturas
                for txt in asignaturas:
                    f.write(txt + '\n')
                f.close()
                
                self.TablaHorarios.clearContents()
                self.TablaHorarios.setRowCount(0)
                
                mensaje_alerta('Inf', 'Se ha borrado corectamente.')
            else:
                mensaje_alerta('Inf', 'La asignatura no se encuentra en el fichero.')
        else:
            mensaje_alerta('Err', 'No se ha asignado la asignatura.')

        

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
                    if asignatura == asignaturas[i].split('-')[0]:
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
                            asignaturas[i] = asignatura + '-' + str(horarios)
                            encontrado = True
                    i += 1
                
                # Si no se encuentra una misma asignatura se añade a las asignaturas ya existentes
                if not encontrado:
                    # Se traduce la asignatura a tipo string  
                    texto = asignatura + '-[\'' + (grupo + '/' + dia + hora) + '\']'
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
                    mensaje_alerta('Inf', 'Añadido correctamente.')
                    inserta_horarios_tabla(self.TablaHorarios, asignatura)
            else:
                mensaje_alerta('Err', 'No se ha asignado el grupo, el dia o la hora.')
        else:
            mensaje_alerta('Err', 'No se ha asignado la asignatura.')

    # Reinicia los valores de cada pestaña
    def fn_reinicia_pestanas(self, index):
        nombreTab = self.tabWidget.tabText(index)

        # Pestaña Asignación
        if nombreTab == 'Asignación':
            self.BtnGuardarExcel.setEnabled(False)
            self.BtnCrearPDF.setEnabled(False)
            self.radioBtnNombre.setChecked(True)
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
        # Pestaña Calendario Alumnos
        elif nombreTab == 'Calendario Alumnos':
            self.TxtNumMatricula.setPlainText('')
            self.TablaHorario.clearContents()
    
    # Carga los horarios de las asignaturas en el comboBox
    def fn_cargar_asignaturas(self):
        asignaturas = list()
        aulaCompartida = list()

        # Abre el fichero en modo lectura
        f = open('horarios.txt', 'r')
        # Lee todo el fichero y lo guarda en la lista
        for asignatura in f:
            asignaturas.append(asignatura.strip('\n').split('-')[0])
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
            self.TablaCompartenAula.setItem(index, 0, QTableWidgetItem(aula))
    
    # Añade los horarios al ScrollArea
    def fn_anadir_horarios(self, index):
        # Recoge la asigantura seleccionada en el ComboBox
        asignatura = self.ComboBoxAsignatura.currentText()
        if asignatura != '':
            # Abre el fichero en modo lectura
            f = open('horarios.txt', 'r')
            # Lee todo el fichero y lo guarda en la lista
            for asig in f:
                if asig.strip('\n').split('-')[0] == asignatura:
                    # Se recoge los horarios de las asignaturas del fichero
                    # Como la cadena esta en formato string se traduce a formato lista
                    horarios = json.loads(asig.split('-')[1].replace('\'','\"'))
            f.close()

            # Crea checkbox con los horarios de la asignatura seleccionada
            widget = QWidget()
            layout = QVBoxLayout(widget)
            self.AreaHorarios.setWidget(widget)

            horario_asignatura = list()

            # Abre el fichero en modo lectura
            f = open('asignaturas.txt', 'r')
            # Lee todo el fichero y guarda los horarios de las asignaturas
            for asig in f:
                if asig.split('-')[0].lower() == asignatura.lower():
                    horario_asignatura = json.loads(asig.split('-')[3].replace('\'','\"'))
            f.close()

            for i, horario in enumerate(horarios):
                # Si hay algun horario repetido no se añade al layout
                if not any(elem.split('/')[1] == horarios[i].split('/')[1] for elem in horarios[:i]):
                    checkBox = QCheckBox(asignatura + ' ' + horario.split('/')[1])
                    # Si el horario ya esta en el txt se checkea
                    if any(horario.split('/')[1] == aux.split('/')[1] for aux in horario_asignatura):
                        checkBox.setCheckState(2)
                    layout.addWidget(checkBox)

    # Obtiene la ruta del archivo pdf que se quiere cargar
    def fn_buscar_archivos(self):
        nombre_fichero = QFileDialog.getOpenFileName(self, 'Abrir Fichero', '', 'PDF (*.pdf)')
        self.TxtFicheroPath.setText(nombre_fichero[0])

    
    # Carga los laboratorios que tiene el alumno en tablaHorario
    def fn_carga_asignaturas(self):
        # Recoge el numero de matricula metido en PlainText
        matricula = self.TxtNumMatricula.toPlainText()
        # Lee los numeros de matricula del Excel con los grupos de laboratorio asignados
        lista_subgrupos = pd.read_excel(
                'lista_subgrupos.xlsx', dtype={'Nº Expediente en Centro':str})
        lista_subgrupos.set_index('Nº Expediente en Centro', inplace=True)
        self.TablaHorario.clearContents()
        
        # Abre el excel con las semanas de inicio
        semanas = pd.read_excel('Semanas.xlsx', index_col=0)

        # Se crea un diccionario para recorrer los dias y las horas
        dias = {'LU':0, 'MA':1, 'MI':2, 'JU':3, 'VI':4}
        dia_semana = {'LU' : 'LUNES', 'MA' : 'MARTES', 'MI' : 'MIERCOLES', 'JU' : 'JUEVES', 'VI' : 'VIERNES'}
        horas = {'09':0, '11':1, '15':2, '17':3}

        texto = ''

        try:
            for col in lista_subgrupos.columns:
                if 'subgrupo' in col:
                    dia = lista_subgrupos.loc[matricula][col]
                    num_sesiones = -1
                    num_subgrupos = -1
                    semana_inicial = -1

                    # Abre el fichero en modo lectura
                    f = open('asignaturas.txt', 'r')
                    # Lee todo el fichero y guarda los horarios de las asignaturas
                    for asig in f:
                        if asig.split('-')[0].lower() == col.split('_')[1].lower():
                            num_sesiones   = int(asig.split('-')[2])
                            num_subgrupos  = int(asig.split('-')[4])
                            semana_inicial = int(asig.split('-')[5])
                    f.close()
                    
                    # Se descartan las celdas vacias y con '-' quedando solo los subgrupos compuestos dias+horas (MI09->MI+09)
                    if pd.notna(dia) and dia != '-':
                        # Se comprueba que la celda donde se va a insertar la asignatura este vacia y se añade
                        if self.TablaHorario.item(horas[dia[2:4]], dias[dia[:2]]) == None:
                            self.TablaHorario.setItem(horas[dia[2:4]], dias[dia[:2]], QTableWidgetItem(col.split('_')[1].upper()))
                            self.TablaHorario.item(horas[dia[2:4]], dias[dia[:2]]).setBackground(QColor(17, 59, 228, 255)) # Color de fondo de la celda
                            #self.tablaHorario.item(horas[dia[2:4]], dias[dia[:2]]).setForeground(QColor(12, 12, 13, 255)) # Color de los números de la celda
                        # Si ya se ha agregado una asignatura previamente, se guarda junto la que se quiere introducir y se añaden las dos
                        else:
                            text = col.split('_')[1].upper() + ' / ' + self.TablaHorario.item(horas[dia[2:4]], dias[dia[:2]]).text().upper()
                            self.TablaHorario.setItem(horas[dia[2:4]], dias[dia[:2]], QTableWidgetItem(text))
                            self.TablaHorario.item(horas[dia[2:4]], dias[dia[:2]]).setBackground(QColor(239, 108, 0, 255)) #Números del calendario
                            #self.TablaHorario.item(horas[dia[2:4]], dias[dia[:2]]).setForeground(QColor(5, 5, 5, 255)) #Números del calendario
                        
                        if num_sesiones != -1:
                	        # Recoge las semanas del horario
                            dias_h = list(semanas[dia_semana[lista_subgrupos.loc[matricula][col][:2]]])
                            texto += col.split('_')[1].upper() + ':'
                            for i in range(num_sesiones):
                                texto += '   ' + lee_grupos.traduce_meses(dias_h[semana_inicial + (ord(dia[-1]) - ord('A') + num_subgrupos * i) - 1].strftime('%d %B'))
                            
                            texto += '\n'
        except KeyError:
            self.TablaHorario.clearContents()
            mensaje_alerta('Err', 'No existe este Nº matricula')
        
        self.lblHorarios.setText(texto)

    # Exportar a pdf los 
    def fn_exportar_PDF(self):
        if self.lblHorarios.text() != '':
            dialog = QPrintPreviewDialog()  
            dialog.paintRequested.connect(self.handlePaintRequest)
            dialog.exec_()
        else:
            mensaje_alerta('Err', 'No se ha seleccionado ningún estudiante.')
    
    def handlePaintRequest(self, printer):
        document = QtGui.QTextDocument()
        cursor = QtGui.QTextCursor(document)

        table = cursor.insertTable(
            self.TablaHorario.rowCount()+1, self.TablaHorario.columnCount()+1)
        cursor.movePosition(QtGui.QTextCursor.NextCell)
        
        # Obtengo el encabezado de la tabla
        for col in range(table.columns()-1):
            cursor.insertText(self.TablaHorario.horizontalHeaderItem(col).text())
            cursor.movePosition(QtGui.QTextCursor.NextCell)
        for row in range(table.rows()-1):
            cursor.insertText(self.TablaHorario.verticalHeaderItem(row).text())
            cursor.movePosition(QtGui.QTextCursor.NextCell)
            # Obtengo el contenido de la tabla
            for column in range(table.columns()-1):
                texto = ''
                if self.TablaHorario.item(row, column) != None:
                    texto = self.TablaHorario.item(row, column).text()
                cursor.insertText(texto)
                cursor.movePosition(QtGui.QTextCursor.NextCell)
        
        cursor.movePosition(QtGui.QTextCursor.NextBlock)
        cursor.insertText('\n\n' + self.lblHorarios.text())
        document.print_(printer)

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

# Borra lista de los laboratorios
def borrar_laboratorios(opcion):
    if opcion.text() == 'OK':
        f = open('asignaturas.txt', 'w')
        f.close()

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
            if asig.split('-')[0].lower() == asignatura.lower():
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
        if asig.split('-')[0].lower() == asignatura.lower():
            # Guarda los horarios de las asignaturas
            # Como la cadena esta en formato string se traduce a formato lista
            horarios = json.loads(asig.split('-')[1].replace('\'','\"'))
            # Añade los huecos para que entren los horarios
            tabla_horarios.setRowCount(len(horarios))
            for i, horario in enumerate(horarios):
                # Crea el texto y lo centra
                texto = QTableWidgetItem(horario)
                texto.setTextAlignment(Qt.AlignCenter)
                # Añade el texto en la tabla
                tabla_horarios.setItem(0, i, texto)
            break
    f.close()

# Inserta los grupos en el comboBox de la asignatura asignada
def inserta_grupos(comboBox_grupos, asignatura):
    
    global grupos

    # Asignaturas
    asignaturas = {
        # Asignatura : [Mañana, Tarde, Dobles Grados]
        'Automatica': [[True, True, False, True, False], [True, True, False, True, False], [True, False, False, False, True, True, True]],
        'Electronica': [[True, True, False, False, False], [True, True, False, False, False], [True, False, False, False, True, True, True]],
        'Automatizacion': [[False, False, True, False, False], [False, False, True, False, False], [False, True, False, False, False, False, False]],
        'Electronica de Potencia': [[False, False, True, False, False], [False, False, True, False, False], [False, True, False, False, False, False, False]],
        'Informatica Industrial': [[False, False, True, False, False], [False, False, True, False, False], [False, True, False, False, False, False, False]],
        'Instrumentacion electronica': [[False, False, True, False, False], [False, False, True, False, False], [False, True, False, False, False, False, False]],
        'Robotica': [[False, False, True, False, False], [False, False, True, False, False], [False, False, True, False, False, False, False]],
        'Electronica Analogica': [[False, False, True, False, False], [False, False, True, False, False], [False, True, False, False, False, False, False]],
        'Electronica Digital': [[False, False, True, False, False], [False, False, True, False, False], [False, True, False, False, False, False, False]],
        'Regulacion Automatica': [[False, False, True, False, False], [False, False, True, False, False], [False, True, False, False, False, False, False]],
        'Control': [[False, False, False, False, True], [False, False, False, False, True], [False, False, True, False, False, False, False]],
        'SED': [[False, False, False, False, True], [False, False, False, False, True], [False, False, True, False, False, False, False]],
        'SEI': [[False, False, False, False, True], [False, False, False, False, True], [False, False, False, True, False, False, False]],
        'SII': [[False, False, False, False, True], [False, False, False, False, True], [False, False, True, False, False, False, False]]
    }

    # Borra los grupos anteriores
    comboBox_grupos.clear()

    # Recorre el vector de la asignatura seleccionada
    for i in range(len(asignaturas[asignatura])):
        for j in range(len(asignaturas[asignatura][i])):
            # Si el grupo es True se agrega al comboBox
            if asignaturas[asignatura][i][j]:
                comboBox_grupos.addItem(grupos[i][j])
    
    # La interfaz no carga ninguna opcion del comboBox (comienza vacio)
    comboBox_grupos.setCurrentIndex(-1)

# Comprueba que dos asignaturas no sean iguales o esten vacias y si coinciden los horarios
# Se va a comprobar solo el grupo porque multiplicando el numero de sesiones con el de subgrupos no hay el suficiente tiempo para que no se solapen
# los dos laboratorio. Se ha tenido en cuenta que el numero minimo de sesiones son 3
def comprobarAsignaturas(asignatura_actual, horario_actual, asignatura1, asignatura2):

    # Comprueba que las dos asignaturas se hayan seleccionado
    if asignatura1 != '' and asignatura2 != '':
        # Comprueba que las dos asignaturas no sean iguales
        if asignatura1 == asignatura2:
            mensaje_alerta('Err', 'No se ha introducido bien las asignaturas con aulas compartidas.')
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

            # Si la asignatura actual es alguna de las asignaturas que comparten aula se remplaza los horarios
            if asignatura1 == asignatura_actual:
                horario1 = horario_actual
            elif asignatura2 == asignatura_actual:
                horario2 = horario_actual 
            
            # Quita los grupos de matricula de los horarios
            horario1 = list(set([horario.split('/')[1] for horario in horario1]))
            horario2 = list(set([horario.split('/')[1] for horario in horario2]))

            # Comprueba si algun horario coincide
            if any(horario in horario2 for horario in horario1):
                mensaje_alerta('Err', 'Coinciden los horarios de las asignaturas que comparten las aulas.')
                return False

            f.close()
    else:
        # Comprueba si alguna de las asignaturas no esta vacia
        if asignatura1 != '' or asignatura2 != '':
            mensaje_alerta('War', 'Solo se ha seleccionado una opción en "Asignaturas que comparten aula".')

    return True

if __name__ == '__main__':
    app = QApplication(sys.argv)
    stream = QtCore.QFile('DarkStyle.qss')
    stream.open(QtCore.QIODevice.ReadOnly)
    app.setStyleSheet(QtCore.QTextStream(stream).readAll())
    interfaz = GUI()
    interfaz.show()
    sys.exit(app.exec_())