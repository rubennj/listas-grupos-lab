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
        #Ajuste del tamaño de la tabla
        self.tablaHorarios.setColumnWidth(0,160)
    
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
        asignatura = self.asignaturaText.toPlainText().lower()
        plazas = self.plazasText.value()
        num_sesiones = self.numSesionesText.value()
        horario = horarios(self.tablaHorarios)
        num_subgrupos = self.numSubgruposText.value()
        semana_inicial = self.semanaInicialText.value()
        
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

    # 
    def fn_borrar_laboratorios(self):
        alerta = QMessageBox(QMessageBox.Question, 'Alerta', '¿Estas seguro de querer borrar permanetemente la lista de los laboratorios?')
        alerta.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        alerta.buttonClicked.connect(borrar_laboratorios)
        alerta.exec()

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

# Devuelve los horarios de la tabla con el formato adecuado 
def horarios(tablaHorarios):
    dias = {
        'LUNES' : 'LU',
        'MARTES' : 'MA',
        'MIERCOLES' : 'MI',
        'MIÉRCOLES' : 'MI',
        'JUEVES' : 'JU',
        'VIERNES' : 'VI'
    }

    try:
        tabla = list()
        
        # Recorre la tabla
        for i in range(0, tablaHorarios.rowCount()):
            # Comprueba que no esten vacios los campos
            if tablaHorarios.item(i, 0).text() != '' and tablaHorarios.item(i, 1).text() != '':
                # Guarda el dia y la hora
                dia = dias[tablaHorarios.item(i, 0).text().upper()]
                hora = tablaHorarios.item(i, 1).text().split(':')[0].zfill(2)

                # Comprueba que la hora se añada correctamente
                if len(hora) != 2 or not hora.isnumeric():
                    raise Exception()
                else:
                    # Junta el dia y la hora para obtener el horario
                    tabla.append(dia + hora)
            # Si la fila entera esta vacia no entra y si en una fila hay una celda vacia y otra llena da un error
            elif not (tablaHorarios.item(i, 0).text() == '' and tablaHorarios.item(i, 1).text() == ''):
                raise Exception()
        
        if len(tabla) == 0:
            mensaje_alerta('Err', 'No se ha añadido nada.')
    except:
        # Se vacia si ha dado un error
        tabla = list()
    finally:
        return tabla

# Borra lista de los laboratorios
def borrar_laboratorios(opcion):
    if opcion.text() == 'OK':
        f = open('asignaturas.txt', 'w')
        f.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    stream = QtCore.QFile('DarkStyle.qss')
    stream.open(QtCore.QIODevice.ReadOnly)
    app.setStyleSheet(QtCore.QTextStream(stream).readAll())
    interfaz = GUI()
    interfaz.show()
    sys.exit(app.exec_())