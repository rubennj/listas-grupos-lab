import sys
from PyQt5 import uic, QtCore
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

import lee_grupos

class GUI(QMainWindow):
    
    def __init__(self):
        super().__init__()
        # Carga la interfaz en nuestra clase
        uic.loadUi("interfaz.ui", self)
        # Eventos de los botones
        self.BtnPrincipal.clicked.connect(self.fn_asignar_grupos)
        self.BtnSecundario.setEnabled(False)
        self.BtnSecundario.clicked.connect(self.fn_guarda_lista)
        self.BtnAceptar.clicked.connect(self.fn_guardar_asignatura)
        self.BtnVaciarTxt.clicked.connect(self.fn_borrar_laboratorios)
        self.BtnGuardaHorario.clicked.connect(self.fn_guarda_horarios)
        # Evento de las pestañas
        self.tabWidget.currentChanged.connect(self.fn_reinicia_pestanas)
        # Eventos TreeView
        self.AsignaturasTree.doubleClicked.connect(self.fn_selecciona_asignatura)
        # Eventos ComboBox
        self.ComboBoxAsignatura.currentIndexChanged.connect(self.fn_anadir_horarios)
    
    # Carga los datos
    def fn_asignar_grupos(self):
        lee_grupos.asignar_grupos()
        mensaje_alerta('Inf', 'Ha terminado de asignar a los estudiantes.')
        self.BtnSecundario.setEnabled(True)
    
    # Guarda a los estudiantes en el excel
    def fn_guarda_lista(self):
        lee_grupos.guardar_lista_grupos()
        mensaje_alerta('Inf', 'Ha terminado de guardar a los estudiantes.')
    
    # Variables de las asginaturas
    def fn_guardar_asignatura(self):
        # Recoge las variables de la interfaz
        asignatura = self.ComboBoxAsignatura.currentText().lower()
        plazas = self.PlazasText.value()
        num_sesiones = self.NumSesionesText.value()
        horario = horarios(self.TablaHorarios)
        num_subgrupos = self.NumSubgruposText.value()
        semana_inicial = self.SemanaInicialText.value()
        
        # Comprueba que las variables esten correctamente y las introduce en el txt
        if asignatura != '' and plazas != 0 and num_sesiones != 0 and horario != [] and num_subgrupos != 0 and semana_inicial != 0:
            asignaturas = list()
            # Abre el fichero en modo lectura
            f = open('asignaturas.txt', 'r')
            # Lee todo el fichero y lo guarda en la lista
            for asig in f:
                asignaturas.append(asig.strip('\n'))
            f.close()
            # Abre el fichero en modo escritura
            f = open('asignaturas.txt', 'w')
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
            
            # Escribe en el fichero las asignaturas
            for txt in asignaturas:
                f.write(txt + '\n')
            f.close()
            
            mensaje_alerta('Inf', 'Añadido correctamente.')

        # Si se han introducido mal los datos salta un error
        else:
            mensaje_alerta('Err', 'No se han introducido bien los datos.')

    # Manda un mensaje de confirmacion para borrar los laboratorios
    def fn_borrar_laboratorios(self):
        alerta = QMessageBox(QMessageBox.Question, 'Alerta', '¿Estas seguro de querer borrar permanetemente la lista de los laboratorios?')
        alerta.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        alerta.buttonClicked.connect(borrar_laboratorios)
        alerta.exec()

    # Selecciona la asignatura que se ha seleccionado en el TreeView
    def fn_selecciona_asignatura(self, index):
        # Recoge la asigantura seleccionada en el TreeView
        asignatura = self.AsignaturasTree.selectedIndexes()[0].data(0)
        # Si selecciona una opocion diferente a un 'Cuatrimestre' se añade la asignatura al label
        if asignatura.find('Cuatrimestre') == -1:
            self.lblAsignaturaAsignada.setText(asignatura)

    # Guarda los horarios de las asignaturas
    def fn_guarda_horarios(self):
        # Label donde esta guardada la asignatura
        asignatura = self.lblAsignaturaAsignada.text()
        # Si no se ha seleccionado una asignatura salta un mensaje de error
        if asignatura != '...':
            # Si no se ha seleccionado un dia o una hora salta un mensaje de error
            if self.ComboBoxDias.currentIndex() != -1 and self.ComboBoxHoras.currentIndex() != -1:
                # Guarda el dia y la hora
                dia = self.ComboBoxDias.currentText().upper()[:2]
                hora = self.ComboBoxHoras.currentText().split(':')[0]

                asignaturas = list()

                # Abre el fichero en modo lectura
                f = open('horarios.txt', 'r')
                # Lee todo el fichero y lo guarda en la lista
                for asig in f:
                    asignaturas.append(asig.strip('\n'))
                f.close()

                # Abre el fichero en modo escritura
                f = open('horarios.txt', 'w')
                # Comprueba si la asignatura instroducida esta en la lista
                encontrado = False
                repetido = False
                i = 0
                while not encontrado and i < len(asignaturas):
                    # Si la asignatura introducida esta repetida se sobreescribe
                    if asignatura == asignaturas[i].split('-')[0]:
                        # Se recoge los horarios de las asignaturas del fichero
                        aux = asignaturas[i].split('-')[1]
                        # Como se convierte en un string se recogen los horarios eliminando los corchetes generados al convertirse al txt
                        horarios = aux.strip('[\'').strip('\']').split('\', \'')
                        # Comprueba si el horario introducido esta en la lista
                        j = 0
                        while not encontrado and j < len(horarios):
                            # Si el horario introducido esta repetido salta un aviso
                            if horarios[j] == (dia + hora):
                                repetido = True
                                encontrado = True
                            j += 1
                        # Si el horario introducido no esta en la lista se añade
                        if not encontrado:
                            horarios.append(dia + hora)
                            horarios = ordenar_horarios(horarios)
                            asignaturas[i] = asignatura + '-' + str(horarios)
                            encontrado = True
                    i += 1
                
                # Si no se encuentra una misma asignatura se añade a las asignaturas ya existentes
                if not encontrado:
                    texto = asignatura + '-' + (dia + hora)
                    asignaturas.append(texto)
                
                # Escribe en el fichero las asignaturas
                for txt in asignaturas:
                    f.write(txt + '\n')
                f.close()

                if repetido:
                    mensaje_alerta('Inf', 'Ya se ha asignado este horario a esta asignatura.')
                else:
                    mensaje_alerta('Inf', 'Añadido correctamente.')
            else:
                mensaje_alerta('Err', 'No se ha asignado el dia o la hora.')
        else:
            mensaje_alerta('Err', 'No se ha asignado la asignatura.')

    # Reinicia los valores de cada pestaña
    def fn_reinicia_pestanas(self, index):
        if index == 0:
            self.BtnSecundario.setEnabled(False)
        elif index == 1:
            self.ComboBoxDias.setCurrentIndex(-1)
            self.ComboBoxHoras.setCurrentIndex(-1)
            self.AsignaturasTree.collapseAll()
            self.lblAsignaturaAsignada.setText('...')
        elif index == 2:
            self.ComboBoxAsignatura.clear()
            self.fn_cargar_asignaturas()
            self.ComboBoxAsignatura.setCurrentIndex(-1)
            self.scrollArea.setWidget(QWidget())
            self.PlazasText.setValue(0)
            self.NumSesionesText.setValue(0)
            self.NumSubgruposText.setValue(0)
            self.SemanaInicialText.setValue(0)
    
    # Carga los horarios de las asignaturas en el comboBox
    def fn_cargar_asignaturas(self):
        asignaturas = list()

        # Abre el fichero en modo lectura
        f = open('horarios.txt', 'r')
        # Lee todo el fichero y lo guarda en la lista
        for asignatura in f:
            asignaturas.append(asignatura.strip('\n').split('-')[0])
        f.close()

        self.ComboBoxAsignatura.insertItems(0, asignaturas)
    
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
                    aux = asig.split('-')[1]
                    # Como se convierte en un string se recogen los horarios eliminando los corchetes generados al convertirse al txt
                    horarios = aux.strip('[\'').strip('\']\n').split('\', \'')
            f.close()

            # Crea checkbox con los horarios de la asignatura seleccionada
            base_widget = QWidget()
            base_layout = QVBoxLayout(base_widget)
            self.scrollArea.setWidget(base_widget)

            for i, horario in enumerate(horarios):
                checkBox = QCheckBox(asignatura + ' ' + horario)
                base_layout.addWidget(checkBox)

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
            if horario[:2] == dia:
                aux.append(horario)

    return aux

if __name__ == '__main__':
    app = QApplication(sys.argv)
    stream = QtCore.QFile('DarkStyle.qss')
    stream.open(QtCore.QIODevice.ReadOnly)
    app.setStyleSheet(QtCore.QTextStream(stream).readAll())
    interfaz = GUI()
    interfaz.show()
    sys.exit(app.exec_())