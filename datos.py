import datetime
import socket
import re
import uuid
import subprocess
import wmi
import platform
from PySide6 import QtWidgets, QtGui, QtCore
from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextEdit, QPushButton, QApplication, QLabel, QWidget, QHBoxLayout
from ftplib import FTP

# Función para obtener la dirección MAC
def obtener_direccion_mac():
    try:
        if platform.system() == "Windows":
            # Verificar el sistema operativo para manejar Windows 7 y Windows 10
            if platform.release() == "7":
                mac = ':'.join(re.findall('..', '%012X' % uuid.UUID(int=uuid.getnode()).hex[-12:]))
            else:
                mac = ':'.join(re.findall('..', '%012X' % uuid.getnode()))
            return mac
        else:
            return ""
    except Exception as e:
        print("Error al obtener la dirección MAC:", str(e))
        return ""

# Función para obtener la máscara de subred del adaptador de red habilitado
def obtener_mascara_subred():
    adapters = wmi.WMI().Win32_NetworkAdapterConfiguration(IPEnabled=True)
    if adapters:
        return adapters[0].IPSubnet[0] if adapters[0].IPSubnet else ""
    return ""

def obtener_puerta_enlace():
    adapters = wmi.WMI().Win32_NetworkAdapterConfiguration(IPEnabled=True)
    for adapter in adapters:
        if adapter.DefaultIPGateway:
            return adapter.DefaultIPGateway[0]
    return ""

class CamposWindow(QDialog):
    def __init__(self, campos_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PLANTILLA REMEDY")
        self.setGeometry(100, 100, 800, 600)

        layout = QVBoxLayout()
        self.campos_text = QTextEdit()
        self.campos_text.setReadOnly(True)
        layout.addWidget(self.campos_text)

        self.campos_data = campos_data
        self.actualizar_campos_text()

        # Botón para copiar el contenido de la plantilla
        copy_button = QPushButton("Copiar al Portapapeles")
        copy_button.clicked.connect(self.copiar_contenido)
        layout.addWidget(copy_button)

        # Botón para guardar en un archivo y en el servidor FTP
        save_button = QPushButton("Guardar en Archivo y FTP")
        save_button.clicked.connect(self.guardar_en_archivo_y_ftp)
        layout.addWidget(save_button)

        self.setLayout(layout)

    def actualizar_campos_text(self):
        plantilla = """ 
            A data de avui: {FECHA} 

            Amb relació amb la incidència: {INCIDENCIA}, hem fet la visita al centre: {CENTRO}
            a la planta: {PLANTA} , consulta: {CONSULTA}, on hem retirat el ordinador antic amb etiqueta LT2B:{LT2B PC ANTIGUO}, 
            i hem fet instal·lació del nou ordinador amb etiqueta LT2B:  {LT2B PC NUEVO}

            Es revisa lloc de treball i podem observar que disposa del següent material:

                • Pantalla amb etiqueta: {LT2B/PANTALLA}
                • Impressora amb etiqueta : {LT2B/IMPRESSORA} 
                • Roseta xarxa ordinador: {ROSETA/PC}

            Dades del equip nou

            Nom del equip: {NOMBRE DEL EQUIPO}
            IP: {IPV4}
            Mac: {MAC}
            Gw: {GW}
            Marca/Model: {MODELO/EQUIPO}
            SN Equip: {NÚMERO DE SERIE DEL EQUIPO}
            Mascara: {MÁSCARA}
            Versió Maqueta: {VERSIÓN DE LA MAQUETA}
            Sistema Operatiu: {SISTEMA OPERATIVO}"""
        texto_final = plantilla.format(**self.campos_data)
        self.campos_text.setPlainText(texto_final)

    def copiar_contenido(self):
        # Obtén el contenido actual de la plantilla
        contenido = self.campos_text.toPlainText()

        # Copia el contenido al portapapeles
        clipboard = QApplication.clipboard()
        clipboard.setText(contenido)

    def guardar_en_archivo_y_ftp(self):
        # Obtén el contenido actual de la plantilla
        contenido = self.campos_text.toPlainText()

        # Obtén el nombre del equipo
        nombre_pc = socket.gethostname()

        # Guarda el contenido en un archivo de texto con el nombre del PC
        ruta_txt = f'C:/Temp/{nombre_pc}.txt'
        with open(ruta_txt, 'w') as archivo:
            archivo.write(contenido)

        # Cargar el archivo en el servidor FTP
        try:
            ftp = FTP('10.85.252.18')
            ftp.login()  # Iniciar sesión en el servidor FTP (puede requerir credenciales)
            with open(ruta_txt, 'rb') as archivo_local:
                ftp.storbinary(f'STOR /METROSUDFTP/00.Datospc/{nombre_pc}.txt', archivo_local)
            ftp.quit()
            print("Archivo guardado en el servidor FTP con éxito.")
        except Exception as e:
            print("Error al cargar el archivo al servidor FTP:", str(e))

# Función para obtener los datos y generar la plantilla Remedy
def generar_plantilla_remedy():
    # Obtener los datos ingresados en la interfaz gráfica
    nueva_incidencia = incidencia_entry.text()
    nuevo_centro = centro_entry.text()
    nueva_planta = planta_entry.text()
    nueva_consulta = consulta_entry.text()
    lt2b_ordinador_antic = lt2b_ordinador_antic_entry.text()
    nuevo_lt2b_ordinador_nuevo = lt2b_ordinador_nuevo_entry.text()
    nueva_lt2b_pantalla = lt2b_pantalla_entry.text()
    nueva_lt2b_impresora = lt2b_impresora_entry.text()
    nueva_roseta = roseta_entry.text()
    mascara_subred = obtener_mascara_subred()
    puerta_enlace = obtener_puerta_enlace()

    # Obtener el nombre del sistema operativo
    sistema_operativo = platform.system()

    # Obtener la versión del sistema operativo
    version_sistema_operativo = platform.release()

    # Obtener la arquitectura de bits (32 o 64 bits)
    arquitectura_bits = platform.architecture()[0] + "s"

    # Obtener información del sistema operativo utilizando platform
    sistema_info = f"{sistema_operativo} {version_sistema_operativo} {arquitectura_bits}"

    # Crear un diccionario solo con los datos ingresados
    datos = {
        'FECHA': datetime.datetime.now().strftime("%d-%m-%Y" " " "%H:%M:%S"),
        'INCIDENCIA': nueva_incidencia,
        'CENTRO': nuevo_centro,
        'PLANTA': nueva_planta,
        'CONSULTA': nueva_consulta,
        'LT2B PC ANTIGUO': lt2b_ordinador_antic,
        'LT2B PC NUEVO': nuevo_lt2b_ordinador_nuevo,
        'LT2B/PANTALLA': nueva_lt2b_pantalla,
        'LT2B/IMPRESSORA': nueva_lt2b_impresora,
        'ROSETA/PC': nueva_roseta,
        'NOMBRE DEL EQUIPO': socket.gethostname(),
        'IPV4': socket.gethostbyname(socket.gethostname()),
        'MÁSCARA': mascara_subred,
        'GW': puerta_enlace,
        'MAC': obtener_direccion_mac(),
        'MODELO/EQUIPO': subprocess.check_output('wmic csproduct get vendor,name').decode("utf-8").split('Vendor')[1].strip(),
        'NÚMERO DE SERIE DEL EQUIPO': subprocess.check_output('wmic bios get serialnumber').decode('utf-8').split('SerialNumber')[1].strip(),
        'VERSIÓN DE LA MAQUETA': subprocess.check_output('Reg query HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\VersionMaster').decode('utf-8').split('REG_SZ')[2].split('    ')[1].strip(),
        'SISTEMA OPERATIVO': sistema_info,
    }

    # Creación de la plantilla Remedy
    campos_window = CamposWindow(datos)
    campos_window.exec()

# Configuración de la interfaz gráfica
app = QtWidgets.QApplication([])

app.setStyle("Fusion")  # Utiliza el estilo Fusion para personalizar la apariencia
palette = QtGui.QPalette()
palette.setColor(QtGui.QPalette.Window, QtGui.QColor(255, 255, 255))  # Establece el color de fondo en blanco
app.setPalette(palette)

root = QtWidgets.QWidget()
root.setWindowTitle("DATOS PARA RELLENAR INC/WO REMEDY")

layout = QVBoxLayout(root)

frame = QtWidgets.QWidget()
frame_layout = QHBoxLayout(frame)

# Agregar una imagen a la ventana de la aplicación con PySide6
image_label = QLabel()
image = QtGui.QImage("c:/temp/datospc/images/pc.jpg")  # Reemplaza "pc.jpg" con la ruta de tu imagen
image = image.scaled(QtCore.QSize(300, 300))
image_label.setPixmap(QtGui.QPixmap.fromImage(image))
image_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)

cabeceras = [
    ("INCIDENCIA:", ""),
    ("CENTRE:", ""),
    ("PLANTA:", ""),
    ("CONSULTA:", ""),
    ("LT2B ANTIC ORDINADOR:", ""),
    ("LT2B NOU ORDINADOR:", ""),
    ("LT2B MONITOR:", ""),
    ("LT2B IMPRESSORA:", ""),
    ("ROSETA XARXA ORDINADOR:", "")
]

# Declarar las variables de entrada como globales
global incidencia_entry
global centro_entry
global planta_entry
global consulta_entry
global lt2b_ordinador_antic_entry
global lt2b_ordinador_nuevo_entry
global lt2b_pantalla_entry
global lt2b_impresora_entry
global roseta_entry

input_frame = QtWidgets.QWidget()
input_layout = QtWidgets.QFormLayout(input_frame)

for label_text, default_value in cabeceras:
    label = QtWidgets.QLabel(label_text)
    entry = QtWidgets.QLineEdit()
    entry.setText(default_value)

    # Asignar las variables globales
    if label_text == "INCIDENCIA:":
        incidencia_entry = entry
    elif label_text == "CENTRE:":
        centro_entry = entry
    elif label_text == "PLANTA:":
        planta_entry = entry
    elif label_text == "CONSULTA:":
        consulta_entry = entry
    elif label_text == "LT2B ANTIC ORDINADOR:":
        lt2b_ordinador_antic_entry = entry
    elif label_text == "LT2B NOU ORDINADOR:":
        lt2b_ordinador_nuevo_entry = entry
    elif label_text == "LT2B MONITOR:":
        lt2b_pantalla_entry = entry
    elif label_text == "LT2B IMPRESSORA:":
        lt2b_impresora_entry = entry
    elif label_text == "ROSETA XARXA ORDINADOR:":
        roseta_entry = entry

    input_layout.addRow(label, entry)

frame_layout.addWidget(image_label)
frame_layout.addWidget(input_frame)

frame.setLayout(frame_layout)
layout.addWidget(frame)

guardar_button = QtWidgets.QPushButton("Mostrar Plantilla Remedy")
guardar_button.clicked.connect(generar_plantilla_remedy)
layout.addWidget(guardar_button)

root.setLayout(layout)
root.show()

app.exec()
