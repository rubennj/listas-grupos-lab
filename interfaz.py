import sys
from PyQt5 import uic
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox

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
    
    # Carga los datos
    def fn_asignar_grupos(self):
        lee_grupos.asignar_grupos()
        mensaje_alerta('Ha terminado de asignar a los estudiantes')
        self.BtnSecundario.setEnabled(True)
    
    # Guarda a los estudiantes en el excel
    def fn_guarda_lista(self):
        lee_grupos.guardar_lista_grupos()
        mensaje_alerta('Ha terminado de guardar a los estudiantes')

# Crea una alerta
def mensaje_alerta(texto):
    alerta = QMessageBox()
    alerta.setText(texto)
    alerta.exec()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    interfaz = GUI()
    interfaz.show()
    sys.exit(app.exec_())