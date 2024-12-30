import os
import sys
import subprocess
import pandas as pd
import mysql.connector
import pyperclip
from datetime import datetime
from PyQt5.QtWidgets import QComboBox, QStyledItemDelegate, QGridLayout
import json
import shutil  # Para manejar copias y renombrar archivos de manera segura
import glob

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTreeView, QFileDialog, QMenu, QAction, QMessageBox,
    QTabWidget, QLabel, QLineEdit, QDialog, QInputDialog, QMenuBar,
    QHeaderView, QListWidget, QListWidgetItem, QSplitter, QTreeWidget, QTreeWidgetItem, QCheckBox, QTableView, QFrame
)
from functools import partial
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QFont, QColor, QDoubleValidator, QIntValidator, QRegExpValidator
from PyQt5.QtCore import Qt, QRegExp
from PyQt5.QtWidgets import QDateEdit
from PyQt5.QtCore import QDate

# Función para conectar a MySQL
def conectar_mysql():
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="tu_usuario",
            password="tu_contraseña",
            database="nombre_base_datos"
        )
        cursor = conn.cursor()
        print("Conexión a MySQL exitosa")
        return conn, cursor
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return None, None

archivos_abiertos = []
cotizaciones = []

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cell-Andrés: Software Contable")
        self.setGeometry(100, 100, 1200, 800)
        self.df = pd.DataFrame()  # DataFrame para productos
        self.df_clientes = pd.DataFrame()  # DataFrame para clientes
    
        # Crear el menú superior
        self.create_menus()
    
        # Crear el panel principal
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
    
        # Crear el menú lateral
        self.setup_menu_lateral()
    
        # Dividir el espacio entre el menú lateral y el área de trabajo
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.menu_lateral)
    
        # Crear un widget que contenga el notebook y el panel de negocios
        notebook_and_panel_widget = QWidget()
        notebook_and_panel_layout = QVBoxLayout(notebook_and_panel_widget)
        notebook_and_panel_layout.setContentsMargins(0, 0, 0, 0)
        notebook_and_panel_layout.setSpacing(0)
    
        # Área de pestañas
        self.notebook = QTabWidget()
        self.notebook.setTabsClosable(True)
        self.notebook.tabCloseRequested.connect(self.close_tab)
        notebook_and_panel_layout.addWidget(self.notebook)
    
        # Crear el panel de Negocios Abiertos
        self.panel_negocios = QWidget()
        self.panel_negocios.setStyleSheet("background-color: lightblue;")  # Establecer color azul claro
        panel_layout = QVBoxLayout(self.panel_negocios)
        panel_layout.setContentsMargins(10, 10, 10, 10)
        panel_layout.setSpacing(5)
    
    
        # Separador
        separador = QFrame()
        separador.setFrameShape(QFrame.HLine)
        separador.setFrameShadow(QFrame.Sunken)
        panel_layout.addWidget(separador)
    
        # Contenedor para la lista de negocios y archivos
        self.lista_negocios_widget = QWidget()
        self.lista_negocios_layout = QVBoxLayout(self.lista_negocios_widget)
        self.lista_negocios_layout.setContentsMargins(0, 0, 0, 0)
        self.lista_negocios_layout.setSpacing(5)
        panel_layout.addWidget(self.lista_negocios_widget)
    
        # Agregar el panel de Negocios Abiertos al layout debajo del notebook
        notebook_and_panel_layout.addWidget(self.panel_negocios)
    
        # Añadir el widget que contiene el notebook y el panel al splitter
        splitter.addWidget(notebook_and_panel_widget)
        splitter.setStretchFactor(1, 1)  # Hacer que el notebook ocupe el espacio restante
    
        # Añadir el splitter al layout principal
        main_layout.addWidget(splitter)
    
        # Cargar cotizaciones y negocios guardados
        self.cargar_cotizaciones_guardadas()
        self.cargar_negocios()
    
        # Diccionario para mantener seguimiento de los negocios y archivos abiertos
        self.negocios_abiertos = {}
        
        self.menu_lateral.setContextMenuPolicy(Qt.CustomContextMenu)
        self.menu_lateral.customContextMenuRequested.connect(self.mostrar_menu_contextual_menu_lateral)
        self.negocios = []
        # Cargar la configuración de negocios
        self.cargar_configuracion_negocios()
        # Resto de la inicialización...



    def agregar_nuevo_negocio(self):
        """
        Abre un diálogo para agregar un nuevo negocio con nuevas bases de datos.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Agregar Nuevo Negocio")
        dialog.resize(400, 500)
        layout = QVBoxLayout(dialog)
    
        # Campos para ingresar nombre y path del negocio
        nombre_label = QLabel("Nombre del Negocio:")
        nombre_input = QLineEdit()
        layout.addWidget(nombre_label)
        layout.addWidget(nombre_input)
    
        path_label = QLabel("Ruta de la Carpeta del Negocio:")
        path_input = QLineEdit()
        path_browse_btn = QPushButton("Buscar")
        path_browse_btn.clicked.connect(lambda: self.browse_folder(path_input))
        path_layout = QHBoxLayout()
        path_layout.addWidget(path_input)
        path_layout.addWidget(path_browse_btn)
        layout.addLayout(path_layout)
    
        # Botones
        btn_layout = QHBoxLayout()
        btn_agregar = QPushButton("Agregar")
        btn_agregar.clicked.connect(lambda: self.confirmar_agregar_nuevo_negocio(nombre_input.text(), path_input.text(), dialog))
        btn_cancelar = QPushButton("Cancelar")
        btn_cancelar.clicked.connect(dialog.reject)
        btn_layout.addWidget(btn_agregar)
        btn_layout.addWidget(btn_cancelar)
        layout.addLayout(btn_layout)
    
        dialog.exec_()
    
    def browse_folder(self, line_edit):
        """
        Abre un diálogo para seleccionar una carpeta y coloca la ruta en el QLineEdit.
        """
        folder_path = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta")
        if folder_path:
            line_edit.setText(folder_path)
    
    def confirmar_agregar_nuevo_negocio(self, nombre, path, dialog):
        """
        Confirma y agrega un nuevo negocio con nuevas bases de datos.
        """
        if not nombre.strip():
            QMessageBox.warning(self, "Entrada Inválida", "El nombre del negocio no puede estar vacío.")
            return
        if not path.strip():
            QMessageBox.warning(self, "Entrada Inválida", "La ruta de la carpeta del negocio no puede estar vacía.")
            return
        if any(n['nombre'] == nombre for n in self.negocios):
            QMessageBox.warning(self, "Error", f"Ya existe un negocio con el nombre '{nombre}'.")
            return
        if not os.path.exists(path):
            try:
                os.makedirs(path)
                print(f"Carpeta '{path}' creada para el negocio.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo crear la carpeta: {e}")
                return
    
        # Definir las bases de datos por defecto
        bases_de_datos = {
            "productos": "productos.xlsx",
            "clientes": "clientes.xlsx",
            "ventas": "ventas.xlsx",
            "contratos": "contratos.xlsx",
            "gastos": "gastos.xlsx",
            "prorrogas": "prorrogas.xlsx"
        }
    
        # Crear archivos Excel vacíos para cada base de datos
        for db, filename in bases_de_datos.items():
            file_path = os.path.join(path, filename)
            if not os.path.exists(file_path):
                df = pd.DataFrame()  # Crear DataFrame vacío
                df.to_excel(file_path, index=False)
                print(f"Archivo '{file_path}' creado.")
    
        # Agregar el negocio a la lista
        nuevo_negocio = {
            "nombre": nombre,
            "path": path,
            "bases_de_datos": bases_de_datos
        }
        self.negocios.append(nuevo_negocio)
        self.guardar_negocios()
    
        QMessageBox.information(self, "Negocio Agregado", f"Negocio '{nombre}' agregado exitosamente.")
        dialog.accept()

    def modificar_negocio(self):
        """
        Abre un diálogo para modificar la configuración de un negocio existente.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Modificar Negocio")
        dialog.resize(500, 400)
        layout = QVBoxLayout(dialog)
    
        # Lista de negocios
        self.modificar_negocios_model = QStandardItemModel()
        self.modificar_negocios_model.setHorizontalHeaderLabels(["Nombre del Negocio", "Path"])
    
        for negocio in self.negocios:
            nombre_item = QStandardItem(negocio['nombre'])
            path_item = QStandardItem(negocio['path'])
            self.modificar_negocios_model.appendRow([nombre_item, path_item])
    
        tree_view = QTreeView()
        tree_view.setModel(self.modificar_negocios_model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
        tree_view.setSelectionMode(QTreeView.SingleSelection)
        tree_view.setHeaderHidden(True)
        layout.addWidget(tree_view)
    
        # Botones para modificar
        btn_layout = QHBoxLayout()
        btn_modificar = QPushButton("Modificar")
        btn_modificar.clicked.connect(lambda: self.confirmar_modificar_negocio(tree_view, dialog))
        btn_cancelar = QPushButton("Cancelar")
        btn_cancelar.clicked.connect(dialog.reject)
        btn_layout.addWidget(btn_modificar)
        btn_layout.addWidget(btn_cancelar)
        layout.addLayout(btn_layout)
    
        dialog.exec_()
    
    def confirmar_modificar_negocio(self, tree_view, dialog):
        """
        Confirma y aplica las modificaciones al negocio seleccionado.
        """
        indexes = tree_view.selectedIndexes()
        if indexes:
            selected_row = indexes[0].row()
            negocio = self.negocios[selected_row]
    
            # Abrir un nuevo diálogo para modificar los detalles
            modificar_dialog = QDialog(self)
            modificar_dialog.setWindowTitle("Modificar Negocio")
            modificar_dialog.resize(400, 300)
            layout = QVBoxLayout(modificar_dialog)
    
            # Campos para modificar nombre y path
            nombre_label = QLabel("Nombre del Negocio:")
            nombre_input = QLineEdit(negocio['nombre'])
            layout.addWidget(nombre_label)
            layout.addWidget(nombre_input)
    
            path_label = QLabel("Ruta de la Carpeta del Negocio:")
            path_input = QLineEdit(negocio['path'])
            path_browse_btn = QPushButton("Buscar")
            path_browse_btn.clicked.connect(lambda: self.browse_folder(path_input))
            path_layout = QHBoxLayout()
            path_layout.addWidget(path_input)
            path_layout.addWidget(path_browse_btn)
            layout.addLayout(path_layout)
    
            # Botones
            btn_layout = QHBoxLayout()
            btn_guardar = QPushButton("Guardar Cambios")
            btn_guardar.clicked.connect(lambda: self.aplicar_modificar_negocio(selected_row, nombre_input.text(), path_input.text(), modificar_dialog))
            btn_cancelar_modificar = QPushButton("Cancelar")
            btn_cancelar_modificar.clicked.connect(modificar_dialog.reject)
            btn_layout.addWidget(btn_guardar)
            btn_layout.addWidget(btn_cancelar_modificar)
            layout.addLayout(btn_layout)
    
            modificar_dialog.exec_()
        else:
            QMessageBox.warning(self, "Modificar Negocio", "Seleccione un negocio para modificar.")
    
    def aplicar_modificar_negocio(self, index, nuevo_nombre, nuevo_path, dialog):
        """
        Aplica las modificaciones al negocio seleccionado.
        """
        if not nuevo_nombre.strip():
            QMessageBox.warning(self, "Entrada Inválida", "El nombre del negocio no puede estar vacío.")
            return
        if not nuevo_path.strip():
            QMessageBox.warning(self, "Entrada Inválida", "La ruta de la carpeta del negocio no puede estar vacía.")
            return
        if any(n['nombre'] == nuevo_nombre for i, n in enumerate(self.negocios) if i != index):
            QMessageBox.warning(self, "Error", f"Ya existe un negocio con el nombre '{nuevo_nombre}'.")
            return
        if not os.path.exists(nuevo_path):
            QMessageBox.warning(self, "Error", f"La carpeta '{nuevo_path}' no existe.")
            return
    
        # Actualizar el negocio en la lista
        negocio = self.negocios[index]
        negocio['nombre'] = nuevo_nombre
        negocio['path'] = nuevo_path
        # Asumiendo que las bases de datos no cambian, si se desea permitir cambiar, agregar lógica adicional
    
        # Actualizar el modelo de la lista de negocios
        self.modificar_negocios_model.setItem(index, 0, QStandardItem(nuevo_nombre))
        self.modificar_negocios_model.setItem(index, 1, QStandardItem(nuevo_path))
    
        # Guardar las configuraciones
        self.guardar_negocios()
    
        QMessageBox.information(self, "Negocio Modificado", f"Negocio '{nuevo_nombre}' modificado exitosamente.")
        dialog.accept()
    
        
    

    def mostrar_menu_contextual_menu_lateral(self, position):
        index = self.menu_lateral.indexAt(position)
        if not index.isValid():
            return
    
        item = self.menu_lateral.itemFromIndex(index)
        menu = QMenu()
    
        parent = item.parent()
        if parent and parent == self.negocios_abiertos_item:
            # Este es un negocio
            cerrar_negocio_action = QAction("Cerrar Negocio", self)
            cerrar_negocio_action.triggered.connect(lambda: self.cerrar_negocio(item.text(0)))
            menu.addAction(cerrar_negocio_action)
        elif parent and parent.parent() and parent.parent() == self.negocios_abiertos_item:
            # Este es un archivo dentro de un negocio
            nombre_negocio = parent.text(0)
            nombre_archivo = item.text(0)
            cerrar_archivo_action = QAction("Cerrar Archivo", self)
            cerrar_archivo_action.triggered.connect(lambda: self.cerrar_archivo(nombre_negocio, nombre_archivo))
            menu.addAction(cerrar_archivo_action)
    
        if not menu.isEmpty():
            menu.exec_(self.menu_lateral.viewport().mapToGlobal(position))
    


    def confirmar_agregar_negocio_existente(self, nombre, path, dialog):
        """
        Confirma y agrega un negocio con bases de datos existentes.
        """
        if not nombre.strip():
            QMessageBox.warning(self, "Entrada Inválida", "El nombre del negocio no puede estar vacío.")
            return
        if not path.strip():
            QMessageBox.warning(self, "Entrada Inválida", "La ruta de la carpeta del negocio no puede estar vacía.")
            return
        if any(n['nombre'] == nombre for n in self.negocios):
            QMessageBox.warning(self, "Error", f"Ya existe un negocio con el nombre '{nombre}'.")
            return
        if not os.path.exists(path):
            QMessageBox.warning(self, "Error", f"La carpeta '{path}' no existe.")
            return
    
        # Definir las bases de datos por defecto
        bases_de_datos = {
            "productos": "productos.xlsx",
            "clientes": "clientes.xlsx",
            "ventas": "ventas.xlsx",
            "contratos": "contratos.xlsx",
            "gastos": "gastos.xlsx",
            "prorrogas": "prorrogas.xlsx"
        }
    
        # Verificar que todos los archivos de bases de datos existan
        archivos_faltantes = []
        for db, filename in bases_de_datos.items():
            file_path = os.path.join(path, filename)
            if not os.path.exists(file_path):
                archivos_faltantes.append(filename)
    
        if archivos_faltantes:
            QMessageBox.warning(self, "Archivos Faltantes",
                                f"No se encontraron los siguientes archivos en '{path}':\n" +
                                "\n".join(archivos_faltantes))
            return
    
        # Agregar el negocio a la lista
        nuevo_negocio = {
            "nombre": nombre,
            "path": path,
            "bases_de_datos": bases_de_datos
        }
        self.negocios.append(nuevo_negocio)
        self.guardar_negocios()
    
        QMessageBox.information(self, "Negocio Agregado", f"Negocio '{nombre}' agregado exitosamente.")
        dialog.accept()

    def consultar_bases_datos_negocios(self):
        """
        Abre un diálogo para consultar y seleccionar negocios y abrir sus bases de datos.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Consultar Bases de Datos de Negocios")
        dialog.resize(600, 400)
        layout = QVBoxLayout(dialog)
    
        # Lista de negocios
        self.negocios_model = QStandardItemModel()
        self.negocios_model.setHorizontalHeaderLabels(["Nombre del Negocio", "Path"])
    
        for negocio in self.negocios:
            nombre_item = QStandardItem(negocio['nombre'])
            path_item = QStandardItem(negocio['path'])
            self.negocios_model.appendRow([nombre_item, path_item])
    
        tree_view = QTreeView()
        tree_view.setModel(self.negocios_model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
        tree_view.setSelectionMode(QTreeView.SingleSelection)
        tree_view.setHeaderHidden(True)
        layout.addWidget(tree_view)
    
        # Botones
        btn_layout = QHBoxLayout()
        btn_abrir = QPushButton("Abrir Negocio")
        btn_abrir.clicked.connect(lambda: self.abrir_negocio_seleccionado(tree_view, dialog))
        btn_cancelar = QPushButton("Cancelar")
        btn_cancelar.clicked.connect(dialog.reject)
        btn_layout.addWidget(btn_abrir)
        btn_layout.addWidget(btn_cancelar)
        layout.addLayout(btn_layout)
    
        dialog.exec_()
    
        # Función para crear la carpeta de clientes
    def crear_carpeta_clientes(self):
        """
        Crea la carpeta 'clientes' si no existe.
        """
        clientes_folder = 'clientes'
        if not os.path.exists(clientes_folder):
            os.makedirs(clientes_folder)
            print(f"Carpeta '{clientes_folder}' creada.")
        else:
            print(f"Carpeta '{clientes_folder}' ya existe.")
                
    def abrir_negocio_seleccionado(self, tree_view, dialog):
        """
        Abre las bases de datos del negocio seleccionado.
        """
        indexes = tree_view.selectedIndexes()
        if indexes:
            selected_row = indexes[0].row()
            negocio = self.negocios[selected_row]
            path_negocio = negocio['path']
            bases_de_datos = negocio['bases_de_datos']
    
            # Verificar que la carpeta del negocio exista
            if not os.path.exists(path_negocio):
                QMessageBox.critical(self, "Error", f"La carpeta del negocio '{negocio['nombre']}' no existe en '{path_negocio}'.")
                return
    
            # Abrir cada base de datos
            for db, filename in bases_de_datos.items():
                file_path = os.path.join(path_negocio, filename)
                if os.path.exists(file_path):
                    if db == "productos":
                        self.cargar_excel(file_path, tipo="productos")
                    elif db == "clientes":
                        self.cargar_excel_clientes(file_path, file_name=filename)
                    elif db == "ventas":
                        self.cargar_excel_ventas(file_path, tipo="ventas")
                    elif db == "contratos":
                        self.cargar_excel_prorroga(file_path, tipo="contratos")
                    elif db == "gastos":
                        self.cargar_excel_gasto(file_path, tipo="gastos")
                    elif db == "prorrogas":
                        self.cargar_excel_contrato(file_path, tipo="prorrogas")
                    # Agrega más condiciones si hay más bases de datos
                else:
                    QMessageBox.warning(self, "Archivo No Encontrado", f"No se encontró el archivo '{filename}' para '{db}' en '{path_negocio}'.")
    
            QMessageBox.information(self, "Negocio Abierto", f"Negocio '{negocio['nombre']}' abierto exitosamente.")
            dialog.accept()
        else:
            QMessageBox.warning(self, "Seleccionar Negocio", "Seleccione un negocio de la lista.")

    def actualizar_vista_negocio(self, tipo_dato, file_path):
        """
        Actualiza la vista de un negocio específico después de modificar su base de datos.
    
        :param tipo_dato: Tipo de base de datos (e.g., 'productos').
        :param file_path: Ruta completa al archivo Excel de la base de datos.
        """
        try:
            # Leer el DataFrame actualizado desde Excel
            df_actualizado = pd.read_excel(file_path)
            print(f"DataFrame para '{tipo_dato}' recargado correctamente desde '{file_path}'.")
    
            # Buscar si la pestaña ya está abierta
            tab_index = -1
            for i in range(self.notebook.count()):
                pestaña = self.notebook.widget(i)
                if hasattr(pestaña, 'file_path') and pestaña.file_path == file_path:
                    tab_index = i
                    break
    
            # Si la pestaña existe, cerrarla
            if tab_index != -1:
                self.notebook.removeTab(tab_index)
                print(f"Pestaña existente para '{file_path}' cerrada.")
    
            # Crear y agregar la pestaña actualizada
            self.crear_pestaña_negocio(tipo_dato, file_path)
            print(f"Pestaña actualizada para '{file_path}' creada y mostrada.")
    
        except Exception as e:
            QMessageBox.critical(self, "Error al Actualizar Vista",
                                 f"No se pudo actualizar la vista de '{tipo_dato}'.\nError: {e}")
            print(f"Error al actualizar la vista de '{tipo_dato}': {e}")




    def setup_menu_lateral(self):
        self.menu_lateral = QTreeWidget()
        self.menu_lateral.setHeaderHidden(True)
        self.menu_lateral.setMaximumWidth(200)
        self.menu_lateral.itemClicked.connect(self.handle_menu_selection)
    
        # Establecer estilo para el menú lateral
        self.menu_lateral.setStyleSheet("""
            QTreeWidget {
                background-color: #ADD8E6;  /* Azul claro */
            }
            QTreeWidget::item {
                font-size: 14px;
            }
            QTreeWidget::item:selected {
                background-color: #ADD8E6;  /* Azul un poco más oscuro */
                color: white;
            }
        """)
    
        # Añadir opciones al menú lateral
        opciones = [
            ("Bases de datos", ["Consultar Bases de Datos de Negocios"]),
            ("Productos", ["Buscar Producto", "Agregar Producto", "Modificar Producto"]),
            ("Clientes", ["Buscar Cliente", "Agregar Cliente", "Modificar Cliente"]),
            ("Cotizaciones", ["Ver Cotizaciones", "Generar Factura Electrónica"]),
            ("Opciones Contables", ["Cuadre de Caja", "Visualización de Ventas", "Generar Factura Normal"]),
            ("Negocios Abiertos", [])  # Añadir Negocios Abiertos como la última categoría
        ]
    
        for categoria, subopciones in opciones:
            categoria_item = QTreeWidgetItem([categoria])
            categoria_item.setFlags(Qt.ItemIsEnabled)
            # Establecer color de fondo para las categorías (azul un poco más oscuro)
            categoria_item.setBackground(0, QColor("#87CEEB"))
            categoria_item.setFont(0, QFont("Arial", 12, QFont.Bold))
            self.menu_lateral.addTopLevelItem(categoria_item)
    
            if categoria == "Negocios Abiertos":
                # Guardar referencia al item de Negocios Abiertos
                self.negocios_abiertos_item = categoria_item
                # No expandir por defecto
                categoria_item.setExpanded(False)
                continue  # No añadir subopciones ahora
    
            for sub in subopciones:
                sub_item = QTreeWidgetItem([sub])
                sub_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                categoria_item.addChild(sub_item)
    
            # Colapsar las categorías por defecto
            categoria_item.setExpanded(False)



    def handle_menu_selection(self, item, column):
        opcion = item.text(0)
        parent = item.parent()
        if parent:
            categoria = parent.text(0)
            if categoria == "Bases de datos":
                if opcion == "Consultar Bases de Datos de Negocios":
                    self.gestionar_negocios()
            elif categoria == "Productos":
                if opcion == "Agregar Producto":
                    self.agregar_producto()
                elif opcion == "Buscar Producto":
                    self.buscar_producto()
                elif opcion == "Modificar Producto":
                    self.modificar_producto()
            elif categoria == "Clientes":
                if opcion == "Agregar Cliente":
                    self.agregar_cliente()
                elif opcion == "Buscar Cliente":
                    self.buscar_cliente()
                elif opcion == "Modificar Cliente":
                    self.modificar_cliente()
            elif categoria == "Cotizaciones":
                if opcion == "Ver Cotizaciones":
                    self.ver_cotizaciones()
                elif opcion == "Generar Factura Electrónica":
                    self.generar_factura_electronica()
            elif categoria == "Opciones Contables":
                if opcion == "Cuadre de Caja":
                    self.cuadre_de_caja()
                elif opcion == "Visualización de Ventas":
                    self.visualizacion_de_ventas()
                elif opcion == "Generar Factura Normal":
                    self.generar_factura_normal()
            elif categoria == "Negocios Abiertos":
                # Este es un negocio abierto
                # Expandir o colapsar el negocio al hacer clic en él
                item.setExpanded(not item.isExpanded())
            else:
                # Puede ser un archivo dentro de un negocio
                grandparent = parent.parent()
                if grandparent and grandparent.text(0) == "Negocios Abiertos":
                    # Este es un archivo dentro de un negocio abierto
                    nombre_negocio = parent.text(0)
                    nombre_archivo = opcion
                    # Navegar a la pestaña correspondiente
                    self.ir_a_pestaña_archivo(nombre_negocio, nombre_archivo)
                else:
                    QMessageBox.information(self, "Opción seleccionada", f"Opción: {opcion} seleccionada")
        else:
            categoria = item.text(0)
            if categoria == "Negocios Abiertos":
                # Expandir o colapsar "Negocios Abiertos"
                item.setExpanded(not item.isExpanded())
            else:
                # Expandir o colapsar otras categorías de nivel superior
                item.setExpanded(not item.isExpanded())

    def recargar_dataframe_pestaña_actual(self):
            pestaña = self.notebook.currentWidget()
            if hasattr(pestaña, 'file_path'):
                file_path = pestaña.file_path
                try:
                    pestaña.df = pd.read_excel(file_path)
                    print(f"DataFrame recargado desde {file_path}")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo recargar el archivo: {e}")
            else:
                QMessageBox.warning(self, "Advertencia", "No se encontró la ruta del archivo en la pestaña actual.")

    def agregar_producto(self):
        """
        Abre un diálogo para agregar un nuevo producto al archivo Excel del negocio actual.
        Incluye dinámicamente todos los campos del Excel excepto 'No.Producto'.
        """
        pestaña = self.notebook.currentWidget()
        if not hasattr(pestaña, 'df') or not hasattr(pestaña, 'file_path'):
            QMessageBox.warning(self, "Advertencia", "No hay un archivo de productos abierto actualmente.")
            return
    
        # Verificar si las columnas necesarias están presentes
        required_columns = ["Precio", "Cantidad Disponible"]
        if not all(col in pestaña.df.columns for col in required_columns):
            QMessageBox.warning(self, "Advertencia", "La pestaña actual no corresponde a un archivo de productos válido.")
            return
    
        agregar_dialog = QDialog(self)
        agregar_dialog.setWindowTitle("Agregar Nuevo Producto")
        agregar_dialog.resize(400, 600)  # Ajusta el tamaño según sea necesario
        layout = QVBoxLayout(agregar_dialog)
    
        # Definir columnas para productos (excluyendo 'No.Producto' y otras si es necesario)
        excluded_columns = ["No.Producto", "Impuesto Cargo", "Impuesto Retención", "Valor Total"]
        columns = [col for col in pestaña.df.columns if col not in excluded_columns]
    
        # Diccionario para almacenar los widgets de entrada
        input_widgets = {}
    
        for col in columns:
            label = QLabel(f"{col}:")
            
            # Detectar si la columna es de tipo fecha
            if "fecha" in col.lower():
                date_edit = QDateEdit()
                date_edit.setCalendarPopup(True)
                date_edit.setDate(QDate.currentDate())  # Establecer fecha actual por defecto
                layout.addWidget(label)
                layout.addWidget(date_edit)
                input_widgets[col] = date_edit
            else:
                line_edit = QLineEdit()
                layout.addWidget(label)
                layout.addWidget(line_edit)
                input_widgets[col] = line_edit
    
        # Botón para guardar el nuevo producto
        btn_guardar = QPushButton("Guardar Producto")
        btn_guardar.clicked.connect(lambda: self.guardar_producto_generico(
            input_widgets,
            agregar_dialog,
            pestaña
        ))
        layout.addWidget(btn_guardar)
    
        agregar_dialog.exec_()

    
    def confirmar_agregar_producto(self, dialog, inputs, producto_tab):
        # Crear un diccionario con los datos ingresados
        nuevo_producto = {}
        for col, input_widget in inputs.items():
            valor = input_widget.text()
            nuevo_producto[col] = valor
    
        # Validar que campos esenciales no estén vacíos (ajusta según tus necesidades)
        if not nuevo_producto.get('No.Producto') or not nuevo_producto.get('Nombre de Producto'):
            QMessageBox.warning(self, "Advertencia", "Los campos 'No.Producto' y 'Nombre de Producto' son obligatorios.")
            return
    
        # Agregar el nuevo producto al DataFrame
        producto_tab.df = producto_tab.df.append(nuevo_producto, ignore_index=True)
        print("Nuevo producto agregado al DataFrame.")
    
        # Guardar el DataFrame en el archivo Excel
        producto_tab.df.to_excel(producto_tab.file_path, index=False)
        print("Archivo Excel de productos actualizado.")
    
        # Actualizar la vista en la pestaña
        self.actualizar_pestaña_excel()
    
        # Cerrar el diálogo
        dialog.accept()
    
        QMessageBox.information(self, "Éxito", "Producto agregado exitosamente.")


    def gestionar_negocios(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Gestión de Negocios")
        dialog.resize(600, 500)
        layout = QVBoxLayout(dialog)
    
        # Lista de negocios existentes
        negocios_list = QListWidget()
        negocios = self.cargar_configuracion_negocios()
        for negocio in negocios:
            negocios_list.addItem(negocio['nombre'])
        layout.addWidget(negocios_list)
    
        # Botones
        btn_layout = QHBoxLayout()
        btn_nuevo = QPushButton("Agregar Nuevo Negocio")
        btn_nuevo.clicked.connect(lambda: self.agregar_negocio_nuevo(dialog, negocios_list))
        btn_existente = QPushButton("Agregar Negocio Existente")
        btn_existente.clicked.connect(lambda: self.agregar_negocio_existente(dialog, negocios_list))
        btn_editar = QPushButton("Editar Negocio")
        btn_editar.clicked.connect(lambda: self.editar_negocio(negocios_list.currentItem()))
        btn_abrir = QPushButton("Abrir Negocio")
        btn_abrir.clicked.connect(lambda: self.abrir_negocio(negocios_list.currentItem()))
        btn_layout.addWidget(btn_nuevo)
        btn_layout.addWidget(btn_existente)
        btn_layout.addWidget(btn_editar)
        btn_layout.addWidget(btn_abrir)
        layout.addLayout(btn_layout)
    
        dialog.exec_()



        
    def cargar_configuracion_negocios(self):
        """
        Carga la configuración de negocios desde 'negocios_config.json' y asigna la lista a self.negocios.
        Siempre retorna una lista de negocios, incluso si está vacía.
        """
        try:
            with open('negocios_config.json', 'r') as f:
                negocios = json.load(f)
            self.negocios = negocios
            print(f"{len(self.negocios)} negocios cargados.")
            return self.negocios
        except FileNotFoundError:
            self.negocios = []
            print("No se encontró el archivo 'negocios_config.json'. Se inicializa una lista vacía de negocios.")
            return self.negocios
        except json.JSONDecodeError:
            QMessageBox.critical(self, "Error", "El archivo 'negocios_config.json' está corrupto.")
            self.negocios = []
            return self.negocios
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cargar los negocios: {e}")
            self.negocios = []
            return self.negocios




    def guardar_configuracion_negocios(self, negocios):
        try:
            with open('negocios_config.json', 'w') as f:
                json.dump(negocios, f, indent=4)
            print("Configuración de negocios guardada correctamente.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo guardar la configuración de negocios.\nError: {e}")
            print(f"Error al guardar la configuración de negocios: {e}")



    def agregar_negocio_nuevo(self, parent_dialog, negocios_list):
        dialog = QDialog(self)
        dialog.setWindowTitle("Agregar Nuevo Negocio")
        dialog.resize(400, 300)
        layout = QVBoxLayout(dialog)
    
        # Nombre del negocio
        nombre_label = QLabel("Nombre del Negocio:")
        nombre_input = QLineEdit()
        layout.addWidget(nombre_label)
        layout.addWidget(nombre_input)
    
        # Información comunicativa
        info_label = QLabel("Se crearán nuevas bases de datos asociadas con este negocio automáticamente.")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
    
        # Botón para crear el negocio
        btn_guardar = QPushButton("Crear Negocio")
        btn_guardar.clicked.connect(lambda: self.guardar_negocio_nuevo(dialog, nombre_input.text()))
        layout.addWidget(btn_guardar)
    
        dialog.exec_()
    
        # Actualizar la lista de negocios después de cerrar el diálogo
        negocios = self.cargar_configuracion_negocios()
        negocios_list.clear()
        for negocio in negocios:
            negocios_list.addItem(negocio['nombre'])

    def agregar_columna_a_lista(self, nueva_columna_input, lista_columnas_extra):
        """
        Agrega una nueva columna a la lista de columnas adicionales.
        """
        nombre_columna = nueva_columna_input.text().strip()
        if nombre_columna:
            lista_columnas_extra.addItem(nombre_columna)
            nueva_columna_input.clear()
        else:
            QMessageBox.warning(self, "Advertencia", "El nombre de la columna no puede estar vacío.")






    def guardar_negocio_existente(self, dialog, nombre_negocio, archivos_seleccionados):
        if not nombre_negocio.strip():
            QMessageBox.warning(self, "Error", "Debe ingresar un nombre para el negocio.")
            return
    
        negocios = self.cargar_configuracion_negocios()
        nombres_existentes = [negocio['nombre'] for negocio in negocios]
        if nombre_negocio in nombres_existentes:
            QMessageBox.warning(self, "Error", f"Ya existe un negocio con el nombre '{nombre_negocio}'. Por favor, elija otro nombre.")
            return
    
        # Crear carpeta para el negocio
        negocio_folder = os.path.join('negocios', nombre_negocio)
        if not os.path.exists(negocio_folder):
            os.makedirs(negocio_folder)
        else:
            QMessageBox.warning(self, "Error", f"La carpeta para el negocio '{nombre_negocio}' ya existe.")
            return
    
        # Definir las columnas por defecto según tipo_dato
        columnas_defecto = {
            'productos': ['No.Producto', 'Nombre de Producto', 'Precio', 'Cantidad Disponible', 'Fecha'],
            'clientes': ['Nombre1', 'Nombre2', 'Apellido1', 'Apellido2', 'Identificacion'],
            'ventas': ['No.Venta', 'Fecha', 'Cliente', 'Precio Costo', 'Precio Venta', 'Valor Total'],
            'prorroga': ['No.Prórroga', 'Fecha', 'Contrato', 'Detalle'],
            'gasto': ['No.Gasto', 'Fecha', 'Descripción', 'Monto'],
            'contrato': ['No.Contrato', 'Fecha', 'Cliente', 'Descripción']
        }
    
        # Copiar los archivos Excel seleccionados al directorio del negocio
        for tipo, file_path in archivos_seleccionados.items():
            if tipo in columnas_defecto:
                # Verificar si el archivo existe
                if not os.path.exists(file_path):
                    QMessageBox.warning(self, "Error", f"Archivo de {tipo} no encontrado: {file_path}")
                    continue
    
                destino = os.path.join(negocio_folder, f"{tipo}_{nombre_negocio}.xlsx")
                try:
                    df = pd.read_excel(file_path)
                    # Añadir columnas por defecto si faltan
                    for col in columnas_defecto[tipo]:
                        if col not in df.columns:
                            df[col] = ""
                    df.to_excel(destino, index=False)
                    print(f"Archivo '{tipo}_{nombre_negocio}.xlsx' guardado correctamente.")
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"No se pudo procesar el archivo de {tipo}: {e}")
                    continue
    
        # Agregar el nuevo negocio a la configuración
        nuevo_negocio = {
            'nombre': nombre_negocio,
            'ruta': negocio_folder
        }
        negocios.append(nuevo_negocio)
        self.guardar_configuracion_negocios(negocios)
        QMessageBox.information(self, "Negocio Agregado", f"El negocio '{nombre_negocio}' ha sido agregado con sus bases de datos.")
        dialog.accept()
    
    def agregar_negocio_existente(self, parent_dialog, negocios_list):
        dialog = QDialog(self)
        dialog.setWindowTitle("Agregar Negocio Existente")
        dialog.resize(500, 400)
        layout = QVBoxLayout(dialog)
    
        # Nombre del negocio
        nombre_label = QLabel("Nombre del Negocio:")
        nombre_input = QLineEdit()
        layout.addWidget(nombre_label)
        layout.addWidget(nombre_input)
    
        # Información comunicativa
        info_label = QLabel("Seleccione las bases de datos existentes para este negocio.")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
    
        # Selección de bases de datos
        tipos_datos = ['productos', 'clientes', 'ventas', 'prorroga', 'gasto', 'contrato']
        checkboxes = {}
        for tipo in tipos_datos:
            checkbox = QCheckBox(tipo.capitalize())
            checkboxes[tipo] = checkbox
            layout.addWidget(checkbox)
    
        # Botón para seleccionar archivos Excel
        archivos_seleccionados = {}
        btn_seleccionar = QPushButton("Seleccionar Archivos Excel")
        def seleccionar_archivos():
            for tipo, checkbox in checkboxes.items():
                if checkbox.isChecked():
                    file_path, _ = QFileDialog.getOpenFileName(self, f"Seleccionar Archivo Excel de {tipo.capitalize()}", "", "Excel Files (*.xlsx)")
                    if file_path:
                        archivos_seleccionados[tipo] = file_path
            QMessageBox.information(self, "Archivos Seleccionados", "Archivos Excel seleccionados correctamente.")
        btn_seleccionar.clicked.connect(seleccionar_archivos)
        layout.addWidget(btn_seleccionar)
    
        # Botón para guardar el negocio
        def guardar_negocio():
            self.guardar_negocio_existente(dialog, nombre_input.text(), archivos_seleccionados)
    
        btn_guardar = QPushButton("Guardar Negocio Existente")
        btn_guardar.clicked.connect(guardar_negocio)
        layout.addWidget(btn_guardar)
    
        dialog.exec_()
    
        # Actualizar la lista de negocios después de cerrar el diálogo
        negocios = self.cargar_configuracion_negocios()
        negocios_list.clear()
        for negocio in negocios:
            negocios_list.addItem(negocio['nombre'])
    
    def cargar_excel(self, file_path=None, file_name=None):
        """
        Carga un archivo Excel de productos y crea una nueva pestaña en la interfaz.
    
        Args:
            file_path (str, optional): Ruta completa del archivo Excel de productos. Si no se proporciona, se abre un diálogo para seleccionar el archivo.
            file_name (str, optional): Nombre del archivo Excel de productos. Necesario si se proporciona file_path.
        """
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(self, "Abrir Archivo Excel", "", "Excel Files (*.xlsx)")
            if not file_path:
                print("No se seleccionó un archivo.")
                return
    
        if not file_name:
            file_name = os.path.basename(file_path)
    
        # Verificar si el archivo ya está abierto
        for i in range(self.notebook.count()):
            if self.notebook.tabText(i) == file_name:
                # Actualizar la pestaña existente
                self.notebook.setCurrentIndex(i)
                pestaña = self.notebook.widget(i)
                try:
                    df = pd.read_excel(file_path)
                    if 'productos' in file_name.lower() and 'Fecha' not in df.columns:
                        df['Fecha'] = pd.to_datetime('today').strftime('%Y-%m-%d')
                        df.to_excel(file_path, index=False)
                        print("'Fecha' agregada al DataFrame de productos.")
                    pestaña.df = df
                    pestaña.file_path = file_path  # Almacenar la ruta completa
                    self.df_productos = df  # Asignar a self.df_productos
                    self.actualizar_pestaña_excel()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel: {e}")
                return
        # Si no está abierto, abrir una nueva pestaña
        try:
            df = pd.read_excel(file_path)
            if 'productos' in file_name.lower() and 'Fecha' not in df.columns:
                df['Fecha'] = pd.to_datetime('today').strftime('%Y-%m-%d')
                df.to_excel(file_path, index=False)
                print("'Fecha' agregada al DataFrame de productos.")
            self.df_productos = df  # Asignar a self.df_productos
            self.crear_pestaña_excel(file_path, file_name)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel: {e}")



    
    def crear_pestaña_excel(self, file_path, file_name):
        """
        Crea una nueva pestaña en la interfaz para mostrar el contenido del archivo Excel.
    
        Args:
            file_path (str): Ruta completa del archivo Excel.
            file_name (str): Nombre del archivo Excel.
        """
        nueva_pestaña = QWidget()
        layout = QVBoxLayout(nueva_pestaña)
    
        # Leer el DataFrame
        try:
            df = pd.read_excel(file_path)
            if 'productos' in file_name.lower() and 'Fecha' not in df.columns:
                df['Fecha'] = pd.to_datetime('today').strftime('%Y-%m-%d')
                df.to_excel(file_path, index=False)
                print("'Fecha' agregada al DataFrame de productos.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel: {e}")
            return
    
        # Crear modelo y vista para el DataFrame
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(df.columns)
    
        tree_view = QTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setSortingEnabled(False)
    
        for index, row in df.iterrows():
            items = []
            for field in row:
                item = QStandardItem()
                item.setEditable(False)
                item.setData(field, Qt.DisplayRole)
                items.append(item)
            model.appendRow(items)
    
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionsClickable(False)
        tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos(pos, tree_view))
    
        layout.addWidget(tree_view)
        nueva_pestaña.tree_view = tree_view
        nueva_pestaña.df = df
        nueva_pestaña.file_path = file_path  # Almacenar la ruta completa
    
        # Agregar la pestaña con el nombre adecuado
        self.notebook.addTab(nueva_pestaña, file_name)  # Usa file_name para el título de la pestaña
        self.notebook.setCurrentWidget(nueva_pestaña)
    

    
    
    def ver_cotizaciones(self):
        print("Abriendo diálogo de cotizaciones...")
        cotizaciones_dialog = QDialog(self)
        cotizaciones_dialog.setWindowTitle("Cotizaciones")
        cotizaciones_dialog.resize(800, 600)
        layout = QVBoxLayout(cotizaciones_dialog)
    
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["ID", "Nombre", "Negocio", "Estado", "Fecha Última Actualización", "Cliente"])

        # Al agregar las cotizaciones al modelo:
        for cotizacion in cotizaciones:
            items = [
                QStandardItem(str(cotizacion['id'])),
                QStandardItem(cotizacion['nombre']),
                QStandardItem(cotizacion.get('negocio', '')),
                QStandardItem(cotizacion['estado']),
                QStandardItem(cotizacion['fecha']),
                QStandardItem(cotizacion.get('cliente', ''))
            ]
            items[3].setEditable(True)  # 'Estado' es editable
            model.appendRow(items)
    
        tree_view = QTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setSortingEnabled(True)
        tree_view.setEditTriggers(QTreeView.SelectedClicked)
        model.itemChanged.connect(self.actualizar_cotizacion)
        self.cotizaciones_tree_view = tree_view
    
        # Asignar el delegado para la columna 'Estado'
        estado_delegate = ComboBoxDelegate(["Pendiente", "Aceptada", "Rechazada", "Cerrada"], self)
        tree_view.setItemDelegateForColumn(3, estado_delegate)
    
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(tree_view)
    
        # Botones
        btn_layout = QHBoxLayout()
        btn_agregar = QPushButton("Agregar Nueva Cotización")
        btn_agregar.clicked.connect(lambda: self.agregar_nueva_cotizacion(model))
        btn_layout.addWidget(btn_agregar)
        btn_eliminar = QPushButton("Eliminar Cotización")
        btn_eliminar.clicked.connect(lambda: self.eliminar_cotizacion(tree_view))
        btn_layout.addWidget(btn_eliminar)
        layout.addLayout(btn_layout)
    
        # Barra de búsqueda
        search_layout = QHBoxLayout()
        search_label = QLabel("Buscar Cotización:")
        search_input = QLineEdit()
        search_button = QPushButton("Buscar")
        search_button.clicked.connect(lambda: self.buscar_cotizacion(model, search_input.text()))
        search_layout.addWidget(search_label)
        search_layout.addWidget(search_input)
        search_layout.addWidget(search_button)
        layout.addLayout(search_layout)
    
        # Conectar doble clic para abrir cotización
        tree_view.doubleClicked.connect(lambda index: self.abrir_pestaña_cotizacion_by_id(
            index.model().item(index.row(), 0).text()))
    
        print("Mostrando diálogo de cotizaciones.")
        cotizaciones_dialog.exec_()

    
    def actualizar_pestaña_excel(self):
        """
        Actualiza la vista del Excel en la pestaña actual para reflejar los cambios en el DataFrame.
        """
        pestaña = self.notebook.currentWidget()
        if hasattr(pestaña, 'file_path') and hasattr(pestaña, 'df'):
            tree_view = pestaña.tree_view
            df = pestaña.df
            file_path = pestaña.file_path
            file_name = os.path.basename(file_path)
    
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(df.columns)
    
            for index, row in df.iterrows():
                items = []
                for field in row:
                    item = QStandardItem()
                    item.setEditable(False)
                    item.setData(field, Qt.DisplayRole)
                    items.append(item)
                model.appendRow(items)
    
            tree_view.setModel(model)
            tree_view.setEditTriggers(QTreeView.NoEditTriggers)
    
            # Configurar el header
            header = tree_view.header()
            header.setSectionResizeMode(QHeaderView.Stretch)
            header.setSectionsClickable(False)
    
            # Mantener el menú contextual
            tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
            if 'clientes' in file_name.lower():
                tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos_cliente_en_dialog(pos, tree_view))
            else:
                tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos(pos, tree_view))
    
            print(f"Vista de '{file_name}' actualizada.")
        else:
            print("No se encontró la ruta del archivo o el DataFrame en la pestaña actual.")
            
    def guardar_negocio_nuevo(self, dialog, nombre_negocio):
        if not nombre_negocio.strip():
            QMessageBox.warning(self, "Error", "Debe ingresar un nombre para el negocio.")
            return
    
        negocios = self.cargar_configuracion_negocios()
        nombres_existentes = [negocio['nombre'] for negocio in negocios]
        if nombre_negocio in nombres_existentes:
            QMessageBox.warning(self, "Error", f"Ya existe un negocio con el nombre '{nombre_negocio}'. Por favor, elija otro nombre.")
            return
    
        # Crear carpeta para el negocio
        negocio_folder = os.path.join('negocios', nombre_negocio)
        if not os.path.exists(negocio_folder):
            os.makedirs(negocio_folder)
            print(f"Carpeta '{negocio_folder}' creada.")
        else:
            QMessageBox.warning(self, "Error", f"La carpeta para el negocio '{nombre_negocio}' ya existe.")
            return
    
        # Definir las bases de datos y sus columnas por defecto
        tipos_datos = {
            'productos': ['No.Producto', 'Nombre de Producto', 'Precio', 'Cantidad Disponible'],
            'clientes': ['Nombre1', 'Nombre2', 'Apellido1', 'Apellido2', 'Identificacion'],
            'ventas': ['No.Venta', 'Fecha', 'Cliente', 'Total'],
            'prorroga': ['No.Prórroga', 'Fecha', 'Contrato', 'Detalle'],
            'gasto': ['No.Gasto', 'Fecha', 'Descripción', 'Monto'],
            'contrato': ['No.Contrato', 'Fecha', 'Cliente', 'Descripción']
        }
    
        for tipo, columnas_defecto in tipos_datos.items():
            # Dialogo para agregar columnas adicionales
            dialog_columnas = QDialog(self)
            dialog_columnas.setWindowTitle(f"Agregar Columnas a {tipo.capitalize()}")
            dialog_columnas.resize(400, 300)
            layout = QVBoxLayout(dialog_columnas)
    
            # Información de columnas por defecto
            info_label = QLabel(f"Columnas por defecto para {tipo.capitalize()}: {', '.join(columnas_defecto)}")
            info_label.setWordWrap(True)
            layout.addWidget(info_label)
    
            # Input para nueva columna
            agregar_columna_layout = QHBoxLayout()
            nueva_columna_input = QLineEdit()
            nueva_columna_input.setPlaceholderText("Nombre de nueva columna")
            btn_agregar_columna = QPushButton("Agregar")
            btn_agregar_columna.clicked.connect(lambda: self.agregar_columna_a_lista(nueva_columna_input, lista_columnas_extra))
            agregar_columna_layout.addWidget(nueva_columna_input)
            agregar_columna_layout.addWidget(btn_agregar_columna)
            layout.addLayout(agregar_columna_layout)
    
            # Lista de columnas adicionales
            lista_columnas_extra = QListWidget()
            layout.addWidget(lista_columnas_extra)
    
            # Botón para aceptar
            btn_aceptar = QPushButton("Aceptar")
            btn_aceptar.clicked.connect(dialog_columnas.accept)
            layout.addWidget(btn_aceptar)
    
            dialog_columnas.exec_()
    
            # Obtener las columnas adicionales
            columnas_adicionales = [lista_columnas_extra.item(i).text() for i in range(lista_columnas_extra.count())]
            columnas_totales = columnas_defecto + columnas_adicionales
    
            # Crear el DataFrame con las columnas totales
            df = pd.DataFrame(columns=columnas_totales)
    
            # Guardar el DataFrame en un archivo Excel
            file_name = f"{tipo}_{nombre_negocio}.xlsx"
            file_path = os.path.join(negocio_folder, file_name)
            df.to_excel(file_path, index=False)
            print(f"Archivo '{file_name}' creado en '{negocio_folder}'.")
    
        # Agregar el nuevo negocio a la configuración
        nuevo_negocio = {
            'nombre': nombre_negocio,
            'ruta': negocio_folder
        }
        negocios.append(nuevo_negocio)
        self.guardar_configuracion_negocios(negocios)
        QMessageBox.information(self, "Negocio Creado", f"El negocio '{nombre_negocio}' ha sido creado con sus bases de datos.")
        dialog.accept()
    



    def seleccionar_excel_base_datos_existente(self, tipo_dato, nombre_negocio, archivos_seleccionados):
        if not nombre_negocio:
            QMessageBox.warning(self, "Error", "Debe ingresar un nombre para el negocio antes de agregar bases de datos.")
            return
    
        file_path, _ = QFileDialog.getOpenFileName(self, f"Seleccionar Archivo Excel de {tipo_dato.capitalize()}", "", "Excel Files (*.xlsx)")
        if file_path:
            archivos_seleccionados[tipo_dato] = file_path
            QMessageBox.information(self, "Archivo Seleccionado", f"Archivo de {tipo_dato} seleccionado: {file_path}")
        else:
            print(f"No se seleccionó un archivo de {tipo_dato}")
            
    

    def seleccionar_excel_negocio_existente(self, nombre_negocio, checkboxes):
        tipos_seleccionados = [tipo for tipo, cb in checkboxes.items() if cb.isChecked()]
        archivos_seleccionados = {}
        for tipo in tipos_seleccionados:
            file_path, _ = QFileDialog.getOpenFileName(self, f"Seleccionar Archivo Excel de {tipo.capitalize()}", "", "Excel Files (*.xlsx)")
            if file_path:
                archivos_seleccionados[tipo] = file_path
        return archivos_seleccionados


    def editar_negocio(self, item_negocio):
        if item_negocio:
            nombre_negocio = item_negocio.text()
            negocio = next((n for n in self.cargar_configuracion_negocios() if n['nombre'] == nombre_negocio), None)
            if negocio:
                dialog = QDialog(self)
                dialog.setWindowTitle(f"Editar Negocio: {nombre_negocio}")
                dialog.resize(500, 600)
                layout = QVBoxLayout(dialog)
    
                # Lista de tipos de datos
                tipos_datos = ['productos', 'clientes', 'ventas', 'prorroga', 'gasto', 'contrato']
                for tipo in tipos_datos:
                    # Construir el patrón de búsqueda
                    file_pattern = f"{tipo}_*.xlsx"
                    search_path = os.path.join(negocio['ruta'], file_pattern)
    
                    # Buscar archivos que coincidan con el patrón
                    matching_files = glob.glob(search_path)
    
                    # Depuración: Imprimir los archivos encontrados
                    print(f"Buscando archivos para '{tipo}' en '{negocio['ruta']}': {matching_files}")
    
                    if matching_files:
                        for file_path in matching_files:
                            file_name = os.path.basename(file_path)
                            btn_editar = QPushButton(f"Agregar Columnas a {file_name}")
                            btn_editar.clicked.connect(partial(self.agregar_columnas_base_datos, tipo, file_path))
                            layout.addWidget(btn_editar)
                            print(f"Botón para agregar columnas a '{tipo}' añadido para el archivo '{file_name}'.")
                    else:
                        # Si no existe ninguna base de datos para este tipo, mostrar botón para crear una nueva
                        btn_agregar_base_datos = QPushButton(f"Agregar Base de Datos '{tipo.capitalize()}'")
                        btn_agregar_base_datos.clicked.connect(partial(self.agregar_nueva_base_datos_a_negocio, tipo, negocio))
                        layout.addWidget(btn_agregar_base_datos)
                        print(f"Botón para agregar nueva base de datos '{tipo}' añadido.")
    
                dialog.exec_()
        else:
            QMessageBox.warning(self, "Error", "Seleccione un negocio para editar.")
            
        self.recargar_dataframe_pestaña_actual()
        self.actualizar_pestaña_excel()   



    

    def agregar_nueva_base_datos_a_negocio(self, tipo_dato, negocio):
        """
        Agrega una nueva base de datos a un negocio existente.
        
        :param tipo_dato: Tipo de base de datos a agregar (e.g., 'productos').
        :param negocio: Diccionario con la información del negocio.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Agregar Base de Datos '{tipo_dato.capitalize()}' al Negocio '{negocio['nombre']}'")
        dialog.resize(400, 300)
        layout = QVBoxLayout(dialog)
    
        # Información de columnas por defecto
        columnas_defecto = {
            'productos': ['No.Producto', 'Nombre de Producto', 'Precio', 'Cantidad Disponible', 'Fecha'],
            'clientes': ['Nombre1', 'Nombre2', 'Apellido1', 'Apellido2', 'Identificacion'],
            'ventas': ['No.Venta', 'Fecha', 'Cliente', 'Precio Costo', 'Precio Venta', 'Valor Total'],
            'prorroga': ['No.Prórroga', 'Fecha', 'Contrato', 'Detalle'],
            'gasto': ['No.Gasto', 'Fecha', 'Descripción', 'Monto'],
            'contrato': ['No.Contrato', 'Fecha', 'Cliente', 'Descripción']
        }
    
        info_label = QLabel(f"Columnas por defecto para {tipo_dato.capitalize()}: {', '.join(columnas_defecto.get(tipo_dato, []))}")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
    
        # Input para nueva columna
        agregar_columna_layout = QHBoxLayout()
        nueva_columna_input = QLineEdit()
        nueva_columna_input.setPlaceholderText("Nombre de nueva columna")
        lista_columnas_extra = QListWidget()
        btn_agregar_columna = QPushButton("Agregar")
        btn_agregar_columna.clicked.connect(lambda: self.agregar_columna_a_lista(nueva_columna_input, lista_columnas_extra))
        agregar_columna_layout.addWidget(nueva_columna_input)
        agregar_columna_layout.addWidget(btn_agregar_columna)
        layout.addLayout(agregar_columna_layout)
    
        # Lista de columnas adicionales
        layout.addWidget(lista_columnas_extra)
    
        # Botón para crear la base de datos
        btn_crear = QPushButton("Crear Base de Datos")
        btn_crear.clicked.connect(lambda: self.crear_base_datos(
            tipo_dato, 
            negocio, 
            columnas_defecto.get(tipo_dato, []), 
            [lista_columnas_extra.item(i).text() for i in range(lista_columnas_extra.count())], 
            dialog
        ))
        layout.addWidget(btn_crear)
    
        dialog.exec_()




    

    def agregar_columnas_base_datos(self, tipo_dato, file_path):
        """
        Agrega una nueva columna a una base de datos existente y actualiza la vista en tiempo real.
    
        :param tipo_dato: Tipo de base de datos (e.g., 'productos').
        :param file_path: Ruta completa al archivo Excel de la base de datos.
        """
        # Crear un diálogo para ingresar el nombre de la nueva columna
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Agregar Columna a {tipo_dato.capitalize()}")
        dialog.resize(300, 150)
        layout = QVBoxLayout(dialog)
    
        label = QLabel("Ingrese el nombre de la nueva columna:")
        layout.addWidget(label)
    
        nueva_columna_input = QLineEdit()
        layout.addWidget(nueva_columna_input)
    
        btn_agregar = QPushButton("Agregar")
        btn_agregar.clicked.connect(partial(self.confirmar_agregar_columna, tipo_dato, file_path, dialog, nueva_columna_input))
        layout.addWidget(btn_agregar)
    
        dialog.exec_()

    
    def confirmar_agregar_columna(self, tipo_dato, file_path, dialog, nueva_columna_input):
        """
        Confirma la adición de la nueva columna al DataFrame y actualiza la vista.
    
        :param tipo_dato: Tipo de base de datos (e.g., 'productos').
        :param file_path: Ruta completa al archivo Excel de la base de datos.
        :param dialog: Diálogo actual a cerrar tras la confirmación.
        :param nueva_columna_input: QLineEdit donde se ingresa el nombre de la nueva columna.
        """
        nueva_columna = nueva_columna_input.text().strip()
        if not nueva_columna:
            QMessageBox.warning(self, "Error", "El nombre de la columna no puede estar vacío.")
            print("Intento de agregar columna vacía.")
            return
    
        try:
            # Leer el DataFrame existente
            df = pd.read_excel(file_path)
            print(f"DataFrame cargado desde '{file_path}' para agregar la columna '{nueva_columna}'.")
    
            if nueva_columna in df.columns:
                QMessageBox.warning(self, "Error", f"La columna '{nueva_columna}' ya existe en la base de datos.")
                print(f"La columna '{nueva_columna}' ya existe en '{file_path}'.")
                return
    
            # Agregar la nueva columna con valores predeterminados (por ejemplo, vacíos)
            df[nueva_columna] = ""
            print(f"Columna '{nueva_columna}' agregada al DataFrame.")
    
            # Guardar el DataFrame actualizado en Excel
            df.to_excel(file_path, index=False)
            print(f"DataFrame guardado con la nueva columna en '{file_path}'.")
    
            # Actualizar la vista en tiempo real
            self.actualizar_vista_negocio(tipo_dato, file_path)
            print(f"Vista de '{tipo_dato}' actualizada en tiempo real.")
    
            QMessageBox.information(self, "Éxito", f"Columna '{nueva_columna}' agregada exitosamente.")
            dialog.accept()
    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo agregar la columna.\nError: {e}")
            print(f"Error al agregar la columna '{nueva_columna}': {e}")



    
    def crear_pestaña_negocio(self, tipo_dato, file_path):
        """
        Crea una nueva pestaña en el notebook para una base de datos específica.
    
        :param tipo_dato: Tipo de base de datos (e.g., 'productos').
        :param file_path: Ruta completa al archivo Excel de la base de datos.
        """
        nueva_pestaña = QWidget()
        layout = QVBoxLayout(nueva_pestaña)
    
        # Leer el DataFrame
        try:
            df = pd.read_excel(file_path)
            print(f"DataFrame leído desde '{file_path}'.")
        except Exception as e:
            QMessageBox.critical(self, "Error",
                                 f"No se pudo leer el archivo '{file_path}'.\nError: {e}")
            print(f"Error al leer el archivo '{file_path}': {e}")
            return
    
        # Crear el modelo y la vista
        modelo = QStandardItemModel()
        modelo.setHorizontalHeaderLabels(df.columns.tolist())
    
        for index, row in df.iterrows():
            items = []
            for field in row:
                item = QStandardItem(str(field))
                item.setEditable(False)
                items.append(item)
            modelo.appendRow(items)
    
        tree_view = QTreeView()
        tree_view.setModel(modelo)
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
        tree_view.header().setSectionResizeMode(QHeaderView.Stretch)
    
        layout.addWidget(tree_view)
    
        # Asignar atributos para identificación
        nueva_pestaña.tipo_dato = tipo_dato
        nueva_pestaña.file_path = file_path
        nueva_pestaña.tree_view = tree_view
    
        # Agregar la pestaña al notebook con el nombre del archivo
        file_basename = os.path.basename(file_path)
        self.notebook.addTab(nueva_pestaña, f"{file_basename}")
        self.notebook.setCurrentWidget(nueva_pestaña)
        print(f"Pestaña creada para '{tipo_dato}' con archivo '{file_basename}'.")






    def crear_base_datos(self, tipo_dato, negocio, columnas_defecto, columnas_adicionales, dialog):
        """
        Crea la base de datos Excel con las columnas definidas y abre una nueva pestaña para ella.
        
        :param tipo_dato: Tipo de base de datos (e.g., 'productos').
        :param negocio: Diccionario con la información del negocio.
        :param columnas_defecto: Lista de columnas por defecto.
        :param columnas_adicionales: Lista de columnas adicionales.
        :param dialog: Diálogo a cerrar tras la creación.
        """
        if not columnas_defecto:
            QMessageBox.warning(self, "Error", f"No se han definido columnas por defecto para '{tipo_dato}'.")
            print(f"No se han definido columnas por defecto para '{tipo_dato}'.")
            return
    
        # Combinar columnas por defecto y adicionales
        columnas_totales = columnas_defecto + columnas_adicionales
    
        # Crear el DataFrame con las columnas totales
        df = pd.DataFrame(columns=columnas_totales)
    
        # Definir la ruta del archivo Excel con el patrón 'tipo_negocioNombre.xlsx'
        nombre_negocio = negocio['nombre']
        # Reemplazar espacios y caracteres especiales en el nombre del negocio si es necesario
        nombre_negocio_sanitizado = "_".join(nombre_negocio.split()).lower()
        file_name = f"{tipo_dato}_{nombre_negocio_sanitizado}.xlsx"
        file_path = os.path.join(negocio['ruta'], file_name)
        
        # Mensaje de depuración antes de crear el archivo
        print(f"Intentando crear la base de datos en: {file_path} con columnas: {columnas_totales}")
    
        try:
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "Éxito", f"Base de datos '{tipo_dato.capitalize()}' creada exitosamente en '{negocio['ruta']}'.")
            print(f"Archivo '{file_name}' creado en '{negocio['ruta']}'.")
    
            # Crear y agregar la pestaña correspondiente
            print(f"Pestaña para '{tipo_dato}' creada y mostrada.")
    
            # Cerrar el diálogo
            dialog.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo crear la base de datos.\nError: {e}")
            print(f"Error al crear la base de datos '{file_name}': {e}")
        self.abrir_archivo_negocio(negocio['nombre'], tipo_dato, file_path)
        
    def archivo_ya_abierto(self, file_path):
        """
        Verifica si un archivo específico ya está abierto en la aplicación.
    
        Args:
            file_path (str): Ruta completa del archivo a verificar.
    
        Returns:
            bool: True si el archivo está abierto, False en caso contrario.
        """
        for i in range(self.notebook.count()):
            pestaña = self.notebook.widget(i)
            if hasattr(pestaña, 'file_path') and pestaña.file_path == file_path:
                return True
        return False

    def abrir_archivo_negocio(self, nombre_negocio, tipo, file_path):
        """
        Abre una base de datos específica dentro de un negocio.
    
        Args:
            nombre_negocio (str): Nombre del negocio.
            tipo (str): Tipo de base de datos (e.g., 'productos', 'ventas').
            file_path (str): Ruta completa del archivo a abrir.
        """
        nombre_archivo = os.path.basename(file_path)
        
        if self.archivo_ya_abierto(file_path):
            QMessageBox.information(self, "Información", f"La base de datos '{nombre_archivo}' ya está abierta.")
            return  # Evitar abrirlo de nuevo
    
        # Abrir el archivo y crear la pestaña correspondiente
        if tipo == 'productos':
            self.cargar_excel(file_path)  # Solo pasa file_path
        elif tipo == 'clientes':
            self.cargar_excel_clientes(file_path, nombre_archivo)
        elif tipo == 'ventas':
            self.cargar_excel_ventas(file_path, nombre_archivo)  # Pasa ambos argumentos
        elif tipo == 'prorroga':
            self.cargar_excel_prorroga(file_path, nombre_archivo)
        elif tipo == 'gasto':
            self.cargar_excel_gasto(file_path, nombre_archivo)
        elif tipo == 'contrato':
            self.cargar_excel_contrato(file_path, nombre_archivo)
        else:
            QMessageBox.warning(self, "Error", f"No se puede abrir el tipo de base de datos '{tipo}'.")
            return
    
        # Agregar el archivo al panel de "Negocios Abiertos"
        self.agregar_archivo_al_panel(nombre_negocio, nombre_archivo)


        

    def guardar_nuevas_columnas(self, tipo_dato, file_path, lista_columnas_extra):
        nuevas_columnas = [lista_columnas_extra.item(i).text() for i in range(lista_columnas_extra.count())]
        if not nuevas_columnas:
            QMessageBox.warning(self, "Error", "Debe agregar al menos una nueva columna.")
            return
    
        try:
            df = pd.read_excel(file_path)
            columnas_existentes = df.columns.tolist()
            for columna in nuevas_columnas:
                if columna not in columnas_existentes:
                    df[columna] = ""  # Asignar valores por defecto vacíos
                    print(f"Columna '{columna}' agregada a '{file_path}'.")
                else:
                    QMessageBox.warning(self, "Error", f"La columna '{columna}' ya existe en '{tipo_dato.capitalize()}'.")
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "Éxito", f"Las nuevas columnas han sido agregadas a '{tipo_dato.capitalize()}'.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo agregar las columnas.\nError: {e}")


            
            
    def guardar_cambios_negocio(self, dialog, nombre_negocio, checkboxes):
        negocio_folder = os.path.join('negocios', nombre_negocio)
        if not os.path.exists(negocio_folder):
            QMessageBox.critical(self, "Error", f"La carpeta del negocio '{nombre_negocio}' no existe.")
            return
    
        tipos_datos = ['productos', 'clientes', 'ventas', 'prorroga', 'gasto', 'contrato']
        for tipo in tipos_datos:
            if tipo in checkboxes and checkboxes[tipo].isChecked():
                file_path, _ = QFileDialog.getOpenFileName(self, f"Seleccionar Archivo Excel de {tipo.capitalize()}", "", "Excel Files (*.xlsx)")
                if file_path:
                    destino = os.path.join(negocio_folder, f"{tipo}_{nombre_negocio}.xlsx")
                    shutil.copy(file_path, destino)
                    print(f"Archivo de {tipo} copiado a {destino}")
                else:
                    QMessageBox.warning(self, "Error", f"No se seleccionó un archivo para {tipo}.")
        
        QMessageBox.information(self, "Cambios Guardados", f"Las bases de datos seleccionadas han sido agregadas al negocio '{nombre_negocio}'.")
        dialog.accept()

    

            
    def seleccionar_excel_base_datos(self, tipo_dato, nombre_negocio):
        if not nombre_negocio:
            QMessageBox.warning(self, "Error", "Debe ingresar un nombre para el negocio antes de agregar bases de datos.")
            return
    
        file_path, _ = QFileDialog.getOpenFileName(self, f"Seleccionar Archivo Excel de {tipo_dato.capitalize()}", "", "Excel Files (*.xlsx)")
        if file_path:
            # Crear carpeta para el negocio si no existe
            negocio_folder = os.path.join('negocios', nombre_negocio)
            if not os.path.exists(negocio_folder):
                os.makedirs(negocio_folder)
    
            # Copiar el archivo Excel al directorio del negocio
            destino = os.path.join(negocio_folder, f"{tipo_dato}.xlsx")
            shutil.copy(file_path, destino)
            print(f"Archivo de {tipo_dato} copiado a {destino}")
        else:
            print(f"No se seleccionó un archivo de {tipo_dato}")
            
    def guardar_negocio(self, dialog, nombre_negocio):
        if not nombre_negocio:
            QMessageBox.warning(self, "Error", "Debe ingresar un nombre para el negocio.")
            return
    
        negocios = self.cargar_configuracion_negocios()
        nombres_existentes = [negocio['nombre'] for negocio in negocios]
        if nombre_negocio in nombres_existentes:
            QMessageBox.warning(self, "Error", f"Ya existe un negocio con el nombre '{nombre_negocio}'. Por favor, elija otro nombre.")
            return
    
        # Crear entrada para el nuevo negocio
        nuevo_negocio = {
            'nombre': nombre_negocio,
            'ruta': os.path.join('negocios', nombre_negocio)
        }
        negocios.append(nuevo_negocio)
        self.guardar_configuracion_negocios(negocios)
        QMessageBox.information(self, "Negocio Guardado", f"El negocio '{nombre_negocio}' ha sido guardado.")
        dialog.accept()
    
    def abrir_negocio(self, item_negocio):
        """
        Abre un negocio seleccionando las bases de datos que se desean abrir.
    
        Args:
            item_negocio (QListWidgetItem): Elemento seleccionado de la lista de negocios.
        """
        if item_negocio:
            nombre_negocio = item_negocio.text()
            negocios = self.cargar_configuracion_negocios()
            negocio = next((n for n in negocios if n['nombre'] == nombre_negocio), None)
            if negocio:
                ruta_negocio = negocio['ruta']
                tipos_datos = ['productos', 'clientes', 'ventas', 'prorroga', 'gasto', 'contrato']  # Excluyendo 'clientes'
    
                # Diálogo para seleccionar qué bases de datos abrir
                dialog = QDialog(self)
                dialog.setWindowTitle(f"Abrir Bases de Datos del Negocio '{nombre_negocio}'")
                dialog.resize(400, 300)
                layout = QVBoxLayout(dialog)
    
                checkboxes = {}
                for tipo in tipos_datos:
                    checkbox = QCheckBox(tipo.capitalize())
                    # Verificar si el archivo existe
                    file_name = f"{tipo}_{nombre_negocio}.xlsx"
                    file_path = os.path.join(ruta_negocio, file_name)
                    if os.path.exists(file_path):
                        # Verificar si ya está abierto
                        if not self.archivo_ya_abierto(file_path):
                            checkbox.setChecked(True)
                        else:
                            checkbox.setChecked(False)
                            checkbox.setEnabled(False)  # Deshabilitar si ya está abierto
                            checkbox.setToolTip("Este archivo ya está abierto.")
                    else:
                        checkbox.setEnabled(False)  # Deshabilitar si no existe
                    checkboxes[tipo] = checkbox
                    layout.addWidget(checkbox)
    
                btn_abrir = QPushButton("Abrir Seleccionadas")
                btn_abrir.clicked.connect(dialog.accept)
                layout.addWidget(btn_abrir)
    
                dialog.exec_()
    
                # Abrir los archivos seleccionados
                for tipo, checkbox in checkboxes.items():
                    if checkbox.isChecked():
                        file_name = f"{tipo}_{nombre_negocio}.xlsx"
                        file_path = os.path.join(ruta_negocio, file_name)
                        if os.path.exists(file_path):
                            self.abrir_archivo_negocio(nombre_negocio, tipo, file_path)
                        else:
                            print(f"Archivo {file_name} no encontrado en {ruta_negocio}")
            else:
                QMessageBox.warning(self, "Error", f"No se encontró el negocio '{item_negocio.text()}'.")
        else:
            QMessageBox.warning(self, "Error", "Seleccione un negocio para abrir.")





    def agregar_negocio_al_panel(self, nombre_negocio):
        # Verificar si el negocio ya está en el panel
        if nombre_negocio in self.negocios_abiertos:
            return  # Ya agregado
    
        # Crear un QTreeWidgetItem para el negocio
        negocio_item = QTreeWidgetItem([nombre_negocio])
        negocio_item.setFlags(Qt.ItemIsEnabled)
        negocio_item.setFont(0, QFont("Arial", 11, QFont.Bold))
        self.negocios_abiertos_item.addChild(negocio_item)
    
        # Expandir Negocios Abiertos
        self.negocios_abiertos_item.setExpanded(True)
    
        # Mantener seguimiento del negocio y sus archivos
        self.negocios_abiertos[nombre_negocio] = {'item': negocio_item, 'files': {}}
    
    def agregar_archivo_al_panel(self, nombre_negocio, nombre_archivo):
        if nombre_negocio not in self.negocios_abiertos:
            self.agregar_negocio_al_panel(nombre_negocio)
    
        negocio_data = self.negocios_abiertos[nombre_negocio]
        negocio_item = negocio_data['item']
    
        # Verificar si el archivo ya está agregado
        if nombre_archivo in negocio_data['files']:
            return
    
        # Crear un QTreeWidgetItem para el archivo
        archivo_item = QTreeWidgetItem([nombre_archivo])
        archivo_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        negocio_item.addChild(archivo_item)
    
        # Expandir el negocio para mostrar los archivos
        negocio_item.setExpanded(True)
    
        # Mantener seguimiento del archivo
        negocio_data['files'][nombre_archivo] = archivo_item

    
    
    def ir_a_pestaña_archivo(self, nombre_negocio, nombre_archivo):
        for i in range(self.notebook.count()):
            if self.notebook.tabText(i) == nombre_archivo:
                self.notebook.setCurrentIndex(i)
                break

    
    def cerrar_negocio(self, nombre_negocio):
        if nombre_negocio in self.negocios_abiertos:
            negocio_data = self.negocios_abiertos[nombre_negocio]
            # Cerrar todas las pestañas de los archivos del negocio
            for nombre_archivo in list(negocio_data['files'].keys()):
                self.cerrar_archivo(nombre_negocio, nombre_archivo)
    
            # Eliminar el item del negocio del QTreeWidget
            self.negocios_abiertos_item.removeChild(negocio_data['item'])
    
            # Eliminar del diccionario
            del self.negocios_abiertos[nombre_negocio]
    
    def cerrar_archivo(self, nombre_negocio, nombre_archivo):
        if nombre_negocio in self.negocios_abiertos:
            negocio_data = self.negocios_abiertos[nombre_negocio]
            if nombre_archivo in negocio_data['files']:
                # Cerrar la pestaña correspondiente
                for i in range(self.notebook.count()):
                    if self.notebook.tabText(i) == nombre_archivo:
                        self.notebook.removeTab(i)
                        break
    
                # Eliminar el item del archivo del QTreeWidget
                archivo_item = negocio_data['files'][nombre_archivo]
                negocio_data['item'].removeChild(archivo_item)
    
                # Eliminar del diccionario
                del negocio_data['files'][nombre_archivo]
    
                # Si ya no hay archivos bajo el negocio, eliminar el negocio
                if not negocio_data['files']:
                    self.cerrar_negocio(nombre_negocio)




    def cargar_excel_ventas(self, file_path, file_name):
        """
        Carga el archivo Excel de ventas y crea una nueva pestaña en la interfaz.
    
        Args:
            file_path (str): Ruta completa del archivo Excel de ventas.
            file_name (str): Nombre del archivo Excel de ventas.
        """
        nueva_pestaña = QWidget()
        layout = QVBoxLayout(nueva_pestaña)
    
        try:
            df = pd.read_excel(file_path, dtype={'No.Venta': int})
        except ValueError:
            df = pd.read_excel(file_path)
            df['No.Venta'] = pd.to_numeric(df['No.Venta'], errors='coerce').fillna(0).astype(int)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel de ventas: {e}")
            return
    
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(df.columns)
    
        tree_view = QTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setSortingEnabled(False)
    
        for index, row in df.iterrows():
            items = []
            for field in row:
                item = QStandardItem()
                item.setEditable(False)
                item.setData(field, Qt.DisplayRole)
                items.append(item)
            model.appendRow(items)
    
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionsClickable(False)
        tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos(pos, tree_view))
    
        layout.addWidget(tree_view)
        nueva_pestaña.tree_view = tree_view
        nueva_pestaña.df = df
        nueva_pestaña.file_path = file_path
    
        self.notebook.addTab(nueva_pestaña, file_name)
        self.notebook.setCurrentWidget(nueva_pestaña)

    
    def cargar_excel_prorroga(self, file_path, file_name):
        """
        Carga el archivo Excel de prórroga y crea una nueva pestaña en la interfaz.
    
        Args:
            file_path (str): Ruta completa del archivo Excel de prórroga.
            file_name (str): Nombre del archivo Excel de prórroga.
        """
        nueva_pestaña = QWidget()
        layout = QVBoxLayout(nueva_pestaña)
    
        try:
            df = pd.read_excel(file_path, dtype={'No.Prórroga': int})
        except ValueError:
            df = pd.read_excel(file_path)
            df['No.Prórroga'] = pd.to_numeric(df['No.Prórroga'], errors='coerce').fillna(0).astype(int)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel de prórrogas: {e}")
            return
    
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(df.columns)
    
        tree_view = QTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setSortingEnabled(False)
    
        for index, row in df.iterrows():
            items = []
            for field in row:
                item = QStandardItem()
                item.setEditable(False)
                item.setData(field, Qt.DisplayRole)
                items.append(item)
            model.appendRow(items)
    
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionsClickable(False)
        tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos(pos, tree_view))
    
        layout.addWidget(tree_view)
        nueva_pestaña.tree_view = tree_view
        nueva_pestaña.df = df
        nueva_pestaña.file_path = file_path
    
        self.notebook.addTab(nueva_pestaña, file_name)
        self.notebook.setCurrentWidget(nueva_pestaña)

    
    def cargar_excel_gasto(self, file_path, file_name):
        """
        Carga el archivo Excel de gastos y crea una nueva pestaña en la interfaz.
    
        Args:
            file_path (str): Ruta completa del archivo Excel de gastos.
            file_name (str): Nombre del archivo Excel de gastos.
        """
        nueva_pestaña = QWidget()
        layout = QVBoxLayout(nueva_pestaña)
    
        try:
            df = pd.read_excel(file_path, dtype={'No.Gasto': int})
        except ValueError:
            df = pd.read_excel(file_path)
            df['No.Gasto'] = pd.to_numeric(df['No.Gasto'], errors='coerce').fillna(0).astype(int)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel de gastos: {e}")
            return
    
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(df.columns)
    
        tree_view = QTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setSortingEnabled(False)
    
        for index, row in df.iterrows():
            items = []
            for field in row:
                item = QStandardItem()
                item.setEditable(False)
                item.setData(field, Qt.DisplayRole)
                items.append(item)
            model.appendRow(items)
    
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionsClickable(False)
        tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos(pos, tree_view))
    
        layout.addWidget(tree_view)
        nueva_pestaña.tree_view = tree_view
        nueva_pestaña.df = df
        nueva_pestaña.file_path = file_path
    
        self.notebook.addTab(nueva_pestaña, file_name)
        self.notebook.setCurrentWidget(nueva_pestaña)

    
    def cargar_excel_contrato(self, file_path, file_name):
        """
        Carga el archivo Excel de contratos y crea una nueva pestaña en la interfaz.
    
        Args:
            file_path (str): Ruta completa del archivo Excel de contratos.
            file_name (str): Nombre del archivo Excel de contratos.
        """
        nueva_pestaña = QWidget()
        layout = QVBoxLayout(nueva_pestaña)
    
        try:
            df = pd.read_excel(file_path, dtype={'No.Contrato': int})
        except ValueError:
            df = pd.read_excel(file_path)
            df['No.Contrato'] = pd.to_numeric(df['No.Contrato'], errors='coerce').fillna(0).astype(int)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel de contratos: {e}")
            return
    
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(df.columns)
    
        tree_view = QTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setSortingEnabled(False)
    
        for index, row in df.iterrows():
            items = []
            for field in row:
                item = QStandardItem()
                item.setEditable(False)
                item.setData(field, Qt.DisplayRole)
                items.append(item)
            model.appendRow(items)
    
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionsClickable(False)
        tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos(pos, tree_view))
    
        layout.addWidget(tree_view)
        nueva_pestaña.tree_view = tree_view
        nueva_pestaña.df = df
        nueva_pestaña.file_path = file_path
    
        self.notebook.addTab(nueva_pestaña, file_name)
        self.notebook.setCurrentWidget(nueva_pestaña)





    

    def create_menus(self):
        menubar = self.menuBar()

        # Menú Archivo
        file_menu = menubar.addMenu("Archivo")
        nuevo_action = QAction("Nuevo Archivo", self)
        nuevo_action.triggered.connect(self.nuevo_archivo)
        file_menu.addAction(nuevo_action)

        abrir_productos_action = QAction("Abrir Archivo de Productos", self)
        abrir_productos_action.triggered.connect(self.cargar_excel)
        file_menu.addAction(abrir_productos_action)

        abrir_clientes_action = QAction("Abrir Archivo de Clientes", self)
        abrir_clientes_action.triggered.connect(self.cargar_excel_clientes)
        file_menu.addAction(abrir_clientes_action)
        
        bases_datos_menu = menubar.addMenu("Bases de Datos")

        # Acción para consultar bases de datos de negocios
        consultar_bd_action = QAction("Consultar Bases de Datos de Negocios", self)
        consultar_bd_action.triggered.connect(self.consultar_bases_datos_negocios)
        bases_datos_menu.addAction(consultar_bd_action)
    
        # Acción para agregar un nuevo negocio
        agregar_negocio_action = QAction("Agregar Nuevo Negocio", self)
        agregar_negocio_action.triggered.connect(self.agregar_nuevo_negocio)
        bases_datos_menu.addAction(agregar_negocio_action)
    
        # Acción para agregar un negocio con bases de datos existentes
        agregar_negocio_existente_action = QAction("Agregar Negocio con Bases de Datos Existentes", self)
        agregar_negocio_existente_action.triggered.connect(self.agregar_negocio_existente)
        bases_datos_menu.addAction(agregar_negocio_existente_action)
    
        # Acción para modificar un negocio
        modificar_negocio_action = QAction("Modificar Negocio", self)
        modificar_negocio_action.triggered.connect(self.modificar_negocio)
        bases_datos_menu.addAction(modificar_negocio_action)

        guardar_action = QAction("Guardar", self)
        guardar_action.triggered.connect(self.guardar_archivo_actual)
        file_menu.addAction(guardar_action)

        guardar_como_action = QAction("Guardar Como...", self)
        guardar_como_action.triggered.connect(self.guardar_archivo_como)
        file_menu.addAction(guardar_como_action)

        importar_action = QAction("Importar", self)
        importar_action.triggered.connect(lambda: print("Funcionalidad de importar pendiente"))
        file_menu.addAction(importar_action)

        exportar_action = QAction("Exportar", self)
        exportar_action.triggered.connect(lambda: print("Funcionalidad de exportar pendiente"))
        file_menu.addAction(exportar_action)

        conectar_mysql_action = QAction("Conectar a MySQL", self)
        conectar_mysql_action.triggered.connect(conectar_mysql)
        file_menu.addAction(conectar_mysql_action)

        salir_action = QAction("Salir", self)
        salir_action.triggered.connect(self.close)
        file_menu.addAction(salir_action)

        # Menú Herramientas
        herramientas_menu = menubar.addMenu("Herramientas")
        calculadora_action = QAction("Calculadora", self)
        calculadora_action.triggered.connect(lambda: subprocess.Popen('calc' if os.name == 'nt' else 'gnome-calculator'))
        herramientas_menu.addAction(calculadora_action)

        importar_datos_action = QAction("Importar Datos", self)
        importar_datos_action.triggered.connect(lambda: print("Importar Datos"))
        herramientas_menu.addAction(importar_datos_action)

        exportar_datos_action = QAction("Exportar Datos", self)
        exportar_datos_action.triggered.connect(lambda: print("Exportar Datos"))
        herramientas_menu.addAction(exportar_datos_action)

        opciones_avanzadas_action = QAction("Opciones Avanzadas", self)
        opciones_avanzadas_action.triggered.connect(lambda: print("Opciones Avanzadas"))
        herramientas_menu.addAction(opciones_avanzadas_action)

        # Menú Ayuda
        ayuda_menu = menubar.addMenu("Ayuda")
        documentacion_action = QAction("Documentación", self)
        documentacion_action.triggered.connect(lambda: print("Abrir Documentación"))
        ayuda_menu.addAction(documentacion_action)

        soporte_action = QAction("Soporte Técnico", self)
        soporte_action.triggered.connect(lambda: print("Soporte Técnico"))
        ayuda_menu.addAction(soporte_action)

        acerca_de_action = QAction("Acerca de", self)
        acerca_de_action.triggered.connect(lambda: QMessageBox.information(self, "Acerca de", "Información del software"))
        ayuda_menu.addAction(acerca_de_action)
        
        
    def cargar_negocios(self):
        """
        Carga los negocios desde un archivo JSON o una base de datos y los asigna a self.negocios.
        """
        try:
            with open('negocios_data.json', 'r') as f:
                self.negocios = json.load(f)
                print(f"{len(self.negocios)} negocios cargados.")
        except FileNotFoundError:
            self.negocios = []
            print("No se encontró el archivo 'negocios_data.json'. Se inicializa una lista vacía de negocios.")
        except Exception as e:
            QMessageBox.critical(self, "Error al Cargar Negocios", f"No se pudo cargar los negocios.\nError: {e}")
            self.negocios = []

    def guardar_negocios(self):
        """
        Guarda las configuraciones de negocios en 'negocios.json'.
        """
        try:
            with open(self.negocios_data_file, 'w') as f:
                json.dump({'negocios': self.negocios}, f, indent=4)
            print(f"Negocios guardados correctamente en '{self.negocios_data_file}'.")
        except Exception as e:
            QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar los negocios.\nError: {e}")
            print(f"Error al guardar los negocios: {e}")

    def mostrar_negocios(self):
        """
        Muestra la lista de negocios en la interfaz.
        """
        layout = QVBoxLayout()

        for negocio in self.negocios:
            btn_negocio = QPushButton(negocio['nombre'])
            btn_negocio.clicked.connect(lambda checked, n=negocio: self.editar_negocio(n))
            layout.addWidget(btn_negocio)
            print(f"Botón para negocio '{negocio['nombre']}' añadido.")

        self.layout.addLayout(layout)
        
    def close_tab(self, index):
        widget = self.notebook.widget(index)
        nombre_archivo = self.notebook.tabText(index)
        self.notebook.removeTab(index)
        widget.deleteLater()
        print(f"Pestaña '{nombre_archivo}' cerrada.")
    
        # Buscar y eliminar el archivo del panel
        for nombre_negocio, negocio_data in self.negocios_abiertos.items():
            if nombre_archivo in negocio_data['files']:
                archivo_item = negocio_data['files'][nombre_archivo]
                negocio_data['item'].removeChild(archivo_item)
                del negocio_data['files'][nombre_archivo]
                if not negocio_data['files']:
                    self.cerrar_negocio(nombre_negocio)
                break



    def cargar_clientes_guardados(self):
            clientes_folder = 'clientes'  # Carpeta donde se almacenan los archivos de clientes
            if not os.path.exists(clientes_folder):
                os.makedirs(clientes_folder)
                print(f"Carpeta '{clientes_folder}' creada.")
    
            archivos_clientes = [f for f in os.listdir(clientes_folder) if f.endswith('.xlsx')]
            for archivo in archivos_clientes:
                file_path = os.path.join(clientes_folder, archivo)
                self.cargar_excel_clientes(file_path, archivo)

    def cargar_excel_clientes(self, file_path=None, file_name=None):
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(self, "Abrir Archivo Excel de Clientes", "", "Excel Files (*.xlsx)")
            if not file_path:
                print("No se seleccionó un archivo de clientes.")
                return
            file_name = os.path.basename(file_path)
        else:
            if not file_name:
                file_name = os.path.basename(file_path)
    
        # Verificar si el archivo ya está abierto
        for i in range(self.notebook.count()):
            if self.notebook.tabText(i) == file_name:
                self.notebook.setCurrentIndex(i)
                pestaña = self.notebook.widget(i)
                try:
                    # Leer el Excel especificando que 'Identificacion' es string
                    self.df_clientes = pd.read_excel(file_path, dtype={'Identificacion': str})
                    if 'id_cliente' not in self.df_clientes.columns:
                        self.df_clientes.insert(0, 'id_cliente', self.df_clientes.index + 1)
                        self.df_clientes.to_excel(file_path, index=False)
                        print("'id_cliente' agregado al DataFrame de clientes.")
                    pestaña.df = self.df_clientes
                    pestaña.file_path = file_path  # Almacenar la ruta completa
                    self.actualizar_pestaña_cliente_excel()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel de clientes: {e}")
                return
    
        # Si no está abierto, abrir una nueva pestaña
        try:
            self.df_clientes = pd.read_excel(file_path, dtype={'Identificacion': str})
            if 'id_cliente' not in self.df_clientes.columns:
                self.df_clientes.insert(0, 'id_cliente', self.df_clientes.index + 1)
                self.df_clientes.to_excel(file_path, index=False)
                print("'id_cliente' agregado al DataFrame de clientes.")
            self.crear_pestaña_cliente_excel(file_path, file_name)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel de clientes: {e}")




    def crear_pestaña_cliente_excel(self, file_path, file_name):
        nueva_pestaña = QWidget()
        layout = QVBoxLayout(nueva_pestaña)
    
        # Leer el DataFrame especificando que 'id_cliente' es int
        try:
            df = pd.read_excel(file_path, dtype={'id_cliente': int, 'Identificacion': str})
        except ValueError:
            df = pd.read_excel(file_path)
            df['id_cliente'] = pd.to_numeric(df['id_cliente'], errors='coerce').fillna(0).astype(int)
            df['Identificacion'] = df['Identificacion'].astype(str)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al leer el archivo Excel de clientes: {e}")
            return
    
        # Crear modelo y vista para el DataFrame
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(df.columns)
    
        tree_view = QTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setSortingEnabled(False)
    
        for index, row in df.iterrows():
            items = []
            for field in row:
                item = QStandardItem()
                item.setEditable(False)
                item.setData(field, Qt.DisplayRole)
                items.append(item)
            model.appendRow(items)
    
        # Deshabilitar la edición a través de la vista
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
    
        # Configurar el header
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionsClickable(False)
    
        # Conectar menú contextual
        tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos_cliente_en_dialog(pos, tree_view))
    
        layout.addWidget(tree_view)
        nueva_pestaña.tree_view = tree_view  # Guardar referencia
    
        # Guardar el DataFrame y la ruta del archivo en la pestaña
        nueva_pestaña.df = df
        nueva_pestaña.file_path = file_path  # Almacenar la ruta completa
    
        self.notebook.addTab(nueva_pestaña, os.path.basename(file_path))
        self.notebook.setCurrentWidget(nueva_pestaña)


    def actualizar_pestaña_cliente_excel(self):
        """
        Actualiza la vista del Excel de clientes en la pestaña actual para reflejar los cambios en el DataFrame.
        """
        pestaña = self.notebook.currentWidget()
        tree_view = pestaña.tree_view
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(self.df_clientes.columns)

        # Deshabilitar el ordenamiento antes de cargar los datos
        tree_view.setSortingEnabled(False)

        for index, row in self.df_clientes.iterrows():
            items = []
            for field in row:
                item = QStandardItem()
                item.setEditable(False)
                item.setData(field, Qt.DisplayRole)
                items.append(item)
            model.appendRow(items)

        tree_view.setModel(model)
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)  # Deshabilitar edición

        # Mantener el menú contextual
        tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos_cliente(pos, tree_view))

        tree_view.setSortingEnabled(False)
        tree_view.header().setSectionResizeMode(QHeaderView.Stretch)
        print("Vista de Clientes actualizada.")

    def agregar_cliente(self):
        """
        Abre un diálogo para agregar un nuevo cliente al archivo Excel.
        Incluye dinámicamente todos los campos del Excel.
        """
        agregar_dialog = QDialog(self)
        agregar_dialog.setWindowTitle("Agregar Nuevo Cliente")
        agregar_dialog.resize(400, 300)  # Ajusta el tamaño según sea necesario
        layout = QVBoxLayout(agregar_dialog)
    
        # Obtener dinámicamente todas las columnas excepto 'id_cliente'
        columnas = [col for col in self.df_clientes.columns if col != 'id_cliente']
    
        # Diccionario para almacenar los widgets de entrada
        input_widgets = {}
    
        for col in columnas:
            label = QLabel(f"{col}:")
            line_edit = QLineEdit()
            layout.addWidget(label)
            layout.addWidget(line_edit)
            input_widgets[col] = line_edit
    
        # Botón para guardar el nuevo cliente
        btn_guardar = QPushButton("Guardar Cliente")
        btn_guardar.clicked.connect(lambda: self.guardar_cliente_generico(
            input_widgets,
            agregar_dialog
        ))
        layout.addWidget(btn_guardar)
    
        agregar_dialog.exec_()

        

    def guardar_cliente_generico(self, input_widgets, dialog):
        """
        Guarda un nuevo cliente en el DataFrame y actualiza el archivo Excel.
        """
        try:
            # Crear un diccionario para almacenar los nuevos datos del cliente
            nuevo_cliente = {}
            for col in self.df_clientes.columns:
                if col == "id_cliente":
                    # Asignar el siguiente ID automáticamente
                    if not self.df_clientes.empty:
                        siguiente_id = self.df_clientes['id_cliente'].max() + 1
                    else:
                        siguiente_id = 1  # Si no hay clientes, iniciar en 1
                    nuevo_cliente[col] = siguiente_id
                elif col in input_widgets:
                    valor = input_widgets[col].text().strip()
                    if col == "Identificacion" and not valor:
                        QMessageBox.critical(self, "Error de Validación", f"El campo '{col}' no puede estar vacío.")
                        return
                    nuevo_cliente[col] = valor
                else:
                    # Asignar valor predeterminado para columnas adicionales
                    nuevo_cliente[col] = ""
    
            # Validar que 'Identificacion' sea única
            if nuevo_cliente['Identificacion'] in self.df_clientes['Identificacion'].values:
                QMessageBox.critical(self, "Error de Validación", f"Ya existe un cliente con la Identificación '{nuevo_cliente['Identificacion']}'.")
                return
    
            # Crear un DataFrame para el nuevo cliente
            nuevo_cliente_df = pd.DataFrame([nuevo_cliente])
    
            # Concatenar el nuevo cliente al DataFrame existente
            self.df_clientes = pd.concat([self.df_clientes, nuevo_cliente_df], ignore_index=True)
            # Asegurar que 'id_cliente' es int antes de guardar
            self.df_clientes['id_cliente'] = pd.to_numeric(self.df_clientes['id_cliente'], errors='coerce').fillna(0).astype(int)
    
            # Obtener la pestaña actual y su ruta de archivo
            pestaña = self.notebook.currentWidget()
            file_path = pestaña.file_path  # Utilizar la ruta completa almacenada
    
            try:
                self.df_clientes.to_excel(file_path, index=False)
                print(f"Archivo Excel de clientes guardado correctamente en: {file_path}")
                self.actualizar_pestaña_cliente_excel()
                dialog.accept()
                QMessageBox.information(self, "Agregar Cliente", f"Cliente '{nuevo_cliente.get('Nombre1', 'N/A')}' agregado exitosamente con ID {siguiente_id}.")
            except Exception as e:
                print(f"Error al guardar el archivo Excel de clientes: {e}")
                QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {e}")
    
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al guardar el nuevo cliente: {ex}")


    def modificar_cliente(self):
        """
        Abre un diálogo para modificar un cliente seleccionado en el archivo Excel.
        Incluye dinámicamente todos los campos del Excel.
        """
        pestaña = self.notebook.currentWidget()
        tree_view = pestaña.tree_view
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            model = tree_view.model()
    
            # Obtener todas las columnas del DataFrame
            columns = self.df_clientes.columns.tolist()
    
            # Crear un diccionario con los datos actuales del cliente
            data = {}
            for col in columns:
                item = model.item(row, columns.index(col))
                data[col] = item.text()
    
            modificar_dialog = QDialog(self)
            modificar_dialog.setWindowTitle(f"Modificar Cliente ID: {data.get('id_cliente', 'N/A')}")
            modificar_dialog.resize(400, 300)  # Ajusta el tamaño según sea necesario
            layout = QVBoxLayout(modificar_dialog)
    
            # Diccionario para almacenar los widgets de entrada
            input_widgets = {}
    
            for col in columns:
                if col == "id_cliente":
                    # Mostrar 'id_cliente' como QLabel en lugar de QLineEdit
                    label = QLabel(f"{col}: {data.get(col, '')}")
                    layout.addWidget(label)
                else:
                    label = QLabel(f"{col}:")
                    line_edit = QLineEdit()
                    line_edit.setText(str(data.get(col, "")))
                    layout.addWidget(label)
                    layout.addWidget(line_edit)
                    input_widgets[col] = line_edit
    
            # Botón para guardar los cambios
            btn_guardar = QPushButton("Guardar Cambios")
            btn_guardar.clicked.connect(lambda: self.guardar_modificaciones_cliente_genericas(
                input_widgets,
                row,
                modificar_dialog
            ))
            layout.addWidget(btn_guardar)
    
            modificar_dialog.exec_()
        else:
            QMessageBox.warning(self, "Advertencia", "No se ha seleccionado ningún cliente.")


    
    def guardar_modificaciones_cliente_genericas(self, input_widgets, row, dialog):
        """
        Guarda las modificaciones realizadas al cliente en el DataFrame y actualiza el archivo Excel.
        """
        try:
            # Obtener la Identificacion original antes de modificar
            original_identificacion = self.df_clientes.at[row, 'Identificacion']
    
            # Iterar sobre cada columna y actualizar el DataFrame
            for col in self.df_clientes.columns:
                if col == "id_cliente":
                    continue  # No actualizar 'id_cliente'
                elif col in input_widgets:
                    valor = input_widgets[col].text().strip()
                    if col == "Identificacion" and not valor:
                        QMessageBox.critical(self, "Error de Validación", f"El campo '{col}' no puede estar vacío.")
                        return
                    self.df_clientes.at[row, col] = valor
                else:
                    # Si hay columnas adicionales, asignar un valor predeterminado si es necesario
                    self.df_clientes.at[row, col] = self.df_clientes.at[row, col]
    
            # Validar que 'Identificacion' sea única
            nueva_identificacion = self.df_clientes.at[row, 'Identificacion']
            if nueva_identificacion != original_identificacion:
                if nueva_identificacion in self.df_clientes['Identificacion'].values:
                    QMessageBox.critical(self, "Error de Validación", f"Ya existe un cliente con la Identificación '{nueva_identificacion}'.")
                    # Revertir el cambio
                    self.df_clientes.at[row, 'Identificacion'] = original_identificacion
                    return
    
            # Asegurar que 'id_cliente' es int antes de guardar
            self.df_clientes['id_cliente'] = pd.to_numeric(self.df_clientes['id_cliente'], errors='coerce').fillna(0).astype(int)
    
            # Obtener la pestaña actual y su ruta de archivo
            pestaña = self.notebook.currentWidget()
            file_path = pestaña.file_path  # Utilizar la ruta completa almacenada
    
            try:
                self.df_clientes.to_excel(file_path, index=False)
                print(f"Archivo Excel de clientes guardado correctamente en: {file_path}")
                self.actualizar_pestaña_cliente_excel()
                dialog.accept()
                QMessageBox.information(self, "Guardar Cambios", "Cambios guardados exitosamente.")
            except Exception as e:
                print(f"Error al guardar el archivo Excel de clientes: {e}")
                QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {e}")
    
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al guardar las modificaciones: {ex}")

    def buscar_cliente(self):
        campos = self.df_clientes.columns.tolist()
        campo, ok = QInputDialog.getItem(self, "Buscar Cliente", "Seleccione el campo de búsqueda:", campos, 0, False)
        if ok and campo:
            search_term, ok = QInputDialog.getText(self, "Buscar Cliente", f"Ingrese el término de búsqueda para '{campo}':")
            if ok and search_term:
                encontrado = False
                for i in range(self.notebook.count()):
                    pestaña = self.notebook.widget(i)
                    if hasattr(pestaña, 'tree_view'):
                        tree_view = pestaña.tree_view
                        model = tree_view.model()
                        col_index = self.df_clientes.columns.get_loc(campo)
                        for row in range(model.rowCount()):
                            item = model.item(row, col_index)
                            if search_term.lower() in str(item.data(Qt.DisplayRole)).lower():
                                index = model.indexFromItem(item)
                                tree_view.scrollTo(index, QTreeView.PositionAtCenter)
                                tree_view.setCurrentIndex(index)
                                print(f"Cliente encontrado en fila {row + 1}, columna '{campo}'.")
                                encontrado = True
                                break
                    if encontrado:
                        break
                if not encontrado:
                    QMessageBox.information(self, "Buscar Cliente", f"No se encontró ningún cliente que coincida con: {search_term} en el campo '{campo}'.")

    def mostrar_menu_datos_cliente_en_dialog(self, position, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            menu = QMenu()
            copiar_action = QAction("Copiar", self)
            copiar_action.triggered.connect(lambda: self.copiar_datos(tree_view))
            menu.addAction(copiar_action)
    
            detalles_action = QAction("Detalles", self)
            detalles_action.triggered.connect(lambda: self.mostrar_detalles_cliente_en_dialog(tree_view))
            menu.addAction(detalles_action)
    
            enviar_info_action = QAction("Enviar información del cliente", self)
            enviar_info_action.triggered.connect(lambda: self.enviar_informacion_cliente(tree_view))
            menu.addAction(enviar_info_action)
    
            menu.exec_(tree_view.viewport().mapToGlobal(position))
        else:
            print("Clic fuera de una fila válida.")


            
    def mostrar_detalles_cliente_en_dialog(self, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            data = [tree_view.model().item(row, col).data(Qt.DisplayRole) for col in range(tree_view.model().columnCount())]
            detalles = {self.df_clientes.columns[i]: data[i] for i in range(len(data))}
    
            # Crear un diálogo personalizado para mostrar los detalles
            detalles_dialog = QDialog(self)
            detalles_dialog.setWindowTitle(f"Detalles del Cliente ID: {detalles.get('id_cliente', 'N/A')}")
            detalles_dialog.resize(400, 300)
            layout = QVBoxLayout(detalles_dialog)
    
            # Usar un layout de grid para una mejor presentación
            grid = QGridLayout()
    
            for i, (clave, valor) in enumerate(detalles.items()):
                label_clave = QLabel(f"<b>{clave}:</b>")
                label_valor = QLabel(str(valor))
                grid.addWidget(label_clave, i, 0, Qt.AlignRight)
                grid.addWidget(label_valor, i, 1, Qt.AlignLeft)
    
            layout.addLayout(grid)
    
            # Botón para cerrar el diálogo
            btn_cerrar = QPushButton("Cerrar")
            btn_cerrar.clicked.connect(detalles_dialog.accept)
            layout.addWidget(btn_cerrar, alignment=Qt.AlignCenter)
    
            detalles_dialog.exec_()
        else:
            QMessageBox.warning(self, "Detalles Cliente", "Seleccione un cliente para ver sus detalles.")

    def copiar_datos(self, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            selected_data = [index.data(Qt.DisplayRole) for index in indexes]
            pyperclip.copy("\t".join(map(str, selected_data)))
            print("Datos copiados:", selected_data)

    def mostrar_detalles_cliente(self, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            data = [tree_view.model().item(row, col).data(Qt.DisplayRole) for col in range(tree_view.model().columnCount())]
            detalles = {self.df_clientes.columns[i]: data[i] for i in range(len(data))}

            # Crear un diálogo personalizado para mostrar los detalles
            detalles_dialog = QDialog(self)
            detalles_dialog.setWindowTitle(f"Detalles del Cliente ID: {detalles.get('id_cliente', 'N/A')}")
            detalles_dialog.resize(400, 300)
            layout = QVBoxLayout(detalles_dialog)

            # Usar un layout de grid para una mejor presentación
            grid = QGridLayout()

            for i, (clave, valor) in enumerate(detalles.items()):
                label_clave = QLabel(f"<b>{clave}:</b>")
                label_valor = QLabel(str(valor))
                grid.addWidget(label_clave, i, 0, Qt.AlignRight)
                grid.addWidget(label_valor, i, 1, Qt.AlignLeft)

            layout.addLayout(grid)

            # Botón para cerrar el diálogo
            btn_cerrar = QPushButton("Cerrar")
            btn_cerrar.clicked.connect(detalles_dialog.accept)
            layout.addWidget(btn_cerrar, alignment=Qt.AlignCenter)

            detalles_dialog.exec_()

    def enviar_informacion_cliente(self, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            data = [tree_view.model().item(index.row(), index.column()).data(Qt.DisplayRole) for index in indexes]
            cliente_info = "\n".join([f"{self.df_clientes.columns[i]}: {data[i]}" for i in range(len(data))])
            # Aquí puedes implementar la lógica para enviar la información del cliente
            # Por ejemplo, copiar al portapapeles o enviar por correo electrónico
            pyperclip.copy(cliente_info)
            QMessageBox.information(self, "Enviar Información", "Información del cliente copiada al portapapeles.")
            print(f"Información del cliente enviada:\n{cliente_info}")


    
    
    
    # Función corregida para guardar un nuevo producto sin sobrescribir el archivo existente
    def guardar_producto(self, input_widgets, dialog):
        """
        Guarda un nuevo producto en el DataFrame y actualiza el archivo Excel.
        """
        try:
            print("Iniciando proceso de guardado de producto...")
            # Crear un diccionario para almacenar los nuevos datos del producto
            nuevo_producto = {}
            for col, widget in input_widgets.items():
                valor = widget.text().strip()
                print(f"Procesando columna: {col}, Valor: {valor}")
                if col == "No.Producto" and not valor:
                    print("Error: El campo 'No.Producto' no puede estar vacío.")
                    QMessageBox.critical(self, "Error de Validación", f"El campo '{col}' no puede estar vacío.")
                    return
                nuevo_producto[col] = valor
    
            # Validar que 'No.Producto' sea único
            if nuevo_producto['No.Producto'] in self.df['No.Producto'].values:
                print(f"Error: Ya existe un producto con el Número '{nuevo_producto['No.Producto']}'")
                QMessageBox.critical(self, "Error de Validación", f"Ya existe un producto con el Número '{nuevo_producto['No.Producto']}'.")
                return
    
            # Generar la fecha actual
            nuevo_producto['Fecha'] = pd.to_datetime('today').strftime('%Y-%m-%d')
            print(f"Fecha generada para el nuevo producto: {nuevo_producto['Fecha']}")
    
            # Crear un DataFrame para el nuevo producto
            nuevo_producto_df = pd.DataFrame([nuevo_producto])
    
            # Concatenar el nuevo producto al DataFrame existente
            self.df = pd.concat([self.df, nuevo_producto_df], ignore_index=True)
            print("Nuevo producto añadido al DataFrame.")
    
            # Obtener la pestaña actual y su ruta de archivo
            pestaña = self.notebook.currentWidget()
            file_path = pestaña.file_path  # Utilizar la ruta completa almacenada
            print(f"Ruta del archivo a guardar: {file_path}")
    
            try:
                self.df.to_excel(file_path, index=False)
                print(f"Archivo Excel de productos guardado correctamente en: {file_path}")
                self.actualizar_pestaña_excel()
                dialog.accept()
                QMessageBox.information(self, "Agregar Producto", f"Producto '{nuevo_producto.get('Nombre de Producto', 'N/A')}' agregado exitosamente.")
            except Exception as e:
                print(f"Error al guardar el archivo Excel de productos: {e}")
                QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {e}")
    
        except Exception as ex:
            print(f"Error inesperado al guardar el producto: {ex}")
            QMessageBox.critical(self, "Error", f"Ocurrió un error al guardar el nuevo producto: {ex}")


    def guardar_producto_generico(self, input_widgets, dialog, pestaña):
        """
        Guarda un nuevo producto en el DataFrame de la pestaña actual y actualiza el archivo Excel.
        """
        try:
            # Crear un diccionario para almacenar los nuevos datos del producto
            nuevo_producto = {}
            for col, widget in input_widgets.items():
                if isinstance(widget, QDateEdit):
                    # Obtener la fecha seleccionada y formatearla como string
                    valor = widget.date().toString("yyyy-MM-dd")
                else:
                    valor = widget.text().strip()
                    # Convertir a tipo adecuado si es necesario
                    if col == "Precio":
                        if valor:
                            try:
                                valor = float(valor)
                            except ValueError:
                                QMessageBox.critical(self, "Error de Validación", f"El valor de '{col}' debe ser un número válido.")
                                return
                        else:
                            valor = 0.0
                    elif col == "Cantidad Disponible":
                        if valor:
                            try:
                                valor = int(valor)
                            except ValueError:
                                QMessageBox.critical(self, "Error de Validación", f"El valor de '{col}' debe ser un número entero válido.")
                                return
                        else:
                            valor = 0
                    # Mantener como string para otros tipos de datos
                    # Puedes agregar más condiciones si tienes más tipos de datos específicos
    
                nuevo_producto[col] = valor
    
            # Generar el siguiente id_producto
            siguiente_id = self.generar_id_producto(pestaña.df)
            nuevo_producto['No.Producto'] = siguiente_id
    
            # Validar que 'No.Producto' sea único (aunque debería serlo automáticamente)
            if nuevo_producto['No.Producto'] in pestaña.df['No.Producto'].values:
                QMessageBox.critical(self, "Error de Validación", f"Ya existe un producto con el Número '{nuevo_producto['No.Producto']}'.")
                return
    
            # Crear un DataFrame para el nuevo producto
            nuevo_producto_df = pd.DataFrame([nuevo_producto])
    
            # Concatenar el nuevo producto al DataFrame existente
            pestaña.df = pd.concat([pestaña.df, nuevo_producto_df], ignore_index=True)
    
            # Guardar el DataFrame actualizado en el archivo Excel
            file_path = pestaña.file_path
            try:
                pestaña.df.to_excel(file_path, index=False)
                print(f"Archivo Excel guardado correctamente en: {file_path}")
                self.actualizar_pestaña_excel()
                dialog.accept()
                QMessageBox.information(self, "Agregar Producto", f"Producto agregado exitosamente con ID {siguiente_id}.")
            except Exception as e:
                print(f"Error al guardar el archivo Excel: {e}")
                QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {e}")
    
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al guardar el nuevo producto: {ex}")


    def modificar_producto(self):
        """
        Abre un diálogo para modificar un producto seleccionado en el archivo Excel del negocio actual.
        Incluye dinámicamente todos los campos del Excel excepto algunos no modificables.
        """
        pestaña = self.notebook.currentWidget()
        if not hasattr(pestaña, 'df') or not hasattr(pestaña, 'file_path'):
            QMessageBox.warning(self, "Advertencia", "No hay un archivo de productos abierto actualmente.")
            return
    
        # Verificar si las columnas necesarias están presentes
        required_columns = ["Precio", "Cantidad Disponible"]
        if not all(col in pestaña.df.columns for col in required_columns):
            QMessageBox.warning(self, "Advertencia", "La pestaña actual no corresponde a un archivo de productos válido.")
            return
    
        tree_view = pestaña.tree_view
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            model = tree_view.model()
    
            # Obtener todas las columnas del DataFrame
            columns = pestaña.df.columns.tolist()
    
            # Definir columnas a excluir
            excluded_columns = ["Impuesto Cargo", "Impuesto Retención", "Valor Total"]
    
            # Crear un diccionario con los datos actuales del producto
            data = {}
            for col in columns:
                if col in excluded_columns:
                    continue  # Excluir las columnas no modificables
                item = model.item(row, columns.index(col))
                if col == "No.Producto":
                    try:
                        data[col] = int(item.text())
                    except ValueError:
                        QMessageBox.critical(self, "Error de Datos", f"El 'No.Producto' debe ser un número entero válido.")
                        return
                else:
                    data[col] = item.text()
    
            modificar_dialog = QDialog(self)
            modificar_dialog.setWindowTitle(f"Modificar Producto No: {data.get('No.Producto', 'N/A')}")
            modificar_dialog.resize(400, 600)  # Ajusta el tamaño según sea necesario
            layout = QVBoxLayout(modificar_dialog)
    
            # Diccionario para almacenar los widgets de entrada
            input_widgets = {}
    
            for col in columns:
                if col in excluded_columns:
                    continue  # Saltar las columnas excluidas
    
                label = QLabel(f"{col}:")
                line_edit = QLineEdit()
                line_edit.setText(str(data.get(col, "")))
    
                # Aplicar validadores según el tipo de dato
                if col == "Precio":
                    if pestaña.df[col].dtype in ['int64', 'float64']:
                        validator = QDoubleValidator(0.00, 1000000.00, 2, self)
                        line_edit.setValidator(validator)
                elif col == "Cantidad Disponible":
                    if pestaña.df[col].dtype == 'int64':
                        validator = QIntValidator(0, 1000000, self)
                        line_edit.setValidator(validator)
                elif "Fecha" in col:
                    # Validar formato de fecha si es necesario
                    regex = QRegExp(r'\d{4}-\d{2}-\d{2}')  # Formato YYYY-MM-DD
                    fecha_validator = QRegExpValidator(regex, self)
                    line_edit.setValidator(fecha_validator)
                else:
                    pass  # Permitir cualquier entrada para otras columnas
    
                # Si la columna es 'No.Producto', hacerla de solo lectura
                if col == "No.Producto":
                    line_edit.setReadOnly(True)
    
                layout.addWidget(label)
                layout.addWidget(line_edit)
                input_widgets[col] = line_edit
    
            # Botón para guardar los cambios
            btn_guardar = QPushButton("Guardar Cambios")
            btn_guardar.clicked.connect(lambda: self.guardar_modificaciones_genericas(
                input_widgets,
                row,
                modificar_dialog,
                pestaña
            ))
            layout.addWidget(btn_guardar)
    
            modificar_dialog.exec_()
        else:
            QMessageBox.warning(self, "Advertencia", "No se ha seleccionado ningún producto.")
    


    def guardar_modificaciones_genericas(self, input_widgets, row, dialog, pestaña):
        """
        Guarda las modificaciones realizadas al producto en el DataFrame de la pestaña actual y actualiza el archivo Excel.
        """
        try:
            # Iterar sobre cada columna y actualizar el DataFrame
            for col, widget in input_widgets.items():
                valor = widget.text().strip()
                if col == "Precio":
                    if valor:
                        try:
                            valor = float(valor)
                        except ValueError:
                            QMessageBox.critical(self, "Error de Validación", f"El valor de '{col}' debe ser un número válido.")
                            return
                    else:
                        valor = 0.0
                elif col == "Cantidad Disponible":
                    if valor:
                        try:
                            valor = int(valor)
                        except ValueError:
                            QMessageBox.critical(self, "Error de Validación", f"El valor de '{col}' debe ser un número entero válido.")
                            return
                    else:
                        valor = 0
                elif col == "No.Producto":
                    # No modificar 'No.Producto', ya que es solo de lectura
                    continue
                # Mantener como string para otros tipos de datos
                pestaña.df.at[row, col] = valor
    
            # Guardar el DataFrame actualizado en el archivo Excel
            file_path = pestaña.file_path
            try:
                # Asegurarse de que 'No.Producto' es int antes de guardar
                pestaña.df['No.Producto'] = pestaña.df['No.Producto'].astype(int)
    
                pestaña.df.to_excel(file_path, index=False)
                print(f"Archivo Excel guardado correctamente en: {file_path}")
                self.actualizar_pestaña_excel()
                dialog.accept()
                QMessageBox.information(self, "Guardar Cambios", "Cambios guardados exitosamente.")
            except Exception as e:
                print(f"Error al guardar el archivo Excel: {e}")
                QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {e}")
    
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al guardar las modificaciones: {ex}")

    def on_tab_changed(self, index):
        """
        Habilita o deshabilita los botones de Agregar y Modificar Producto según la pestaña actual.
        """
        pestaña = self.notebook.widget(index)
        if hasattr(pestaña, 'df') and hasattr(pestaña, 'file_path'):
            df = pestaña.df
            # Verificar si las columnas necesarias están presentes
            required_columns = ["Precio", "Cantidad Disponible"]
            if all(col in df.columns for col in required_columns):
                # Habilitar los botones
                self.agregar_producto_button.setEnabled(True)
                self.modificar_producto_button.setEnabled(True)
            else:
                # Deshabilitar los botones
                self.agregar_producto_button.setEnabled(False)
                self.modificar_producto_button.setEnabled(False)
        else:
            # Deshabilitar los botones
            self.agregar_producto_button.setEnabled(False)
            self.modificar_producto_button.setEnabled(False)
    
        
    def generar_id_producto(self, df):
        """
        Genera un ID único para un nuevo producto basado en el máximo ID existente.
        
        :param df: DataFrame de productos.
        :return: ID único para el nuevo producto.
        """
        if df.empty:
            return 1
        else:
            return df["No.Producto"].max() + 1
    
        
    
    def procesar_factura_electronica(self, tree_view, dialog):
        indexes = tree_view.selectionModel().selectedRows()
        if indexes:
            productos_factura = []
            for index in indexes:
                row = index.row()
                no_producto = int(tree_view.model().item(row, 0).text())
                nombre_producto = tree_view.model().item(row, 1).text()
                precio = float(tree_view.model().item(row, 2).text())
                cantidad_disponible = int(tree_view.model().item(row, 3).text())
                
    
                # Solicitar cantidad a facturar
                cantidad, ok = QInputDialog.getInt(self, "Cantidad a Facturar", f"Ingrese la cantidad para '{nombre_producto}':", 1, 1, cantidad_disponible)
                if ok:
                    if cantidad > cantidad_disponible:
                        QMessageBox.warning(self, "Cantidad Excedida", f"La cantidad solicitada excede la disponible para '{nombre_producto}'.")
                        continue
                    else:
                        productos_factura.append({
                            "No.Producto": no_producto,
                            "Nombre de Producto": nombre_producto,
                            "Precio": precio,
                            "Cantidad": cantidad
                        })
                else:
                    continue
    
            if productos_factura:
                # Mostrar información de la factura
                print("Factura Electrónica Generada:")
                for producto in productos_factura:
                    print(producto)
    
                # Actualizar cantidades disponibles en self.df y guardar en Excel
                for producto in productos_factura:
                    index = self.df[self.df["No.Producto"] == producto["No.Producto"]].index
                    if not index.empty:
                        self.df.loc[index, "Cantidad Disponible"] -= producto["Cantidad"]
                        # Recalcular el Valor Total
                        precio = self.df.loc[index, "Precio"].values[0]
                        cantidad = self.df.loc[index, "Cantidad Disponible"].values[0]
                        impuesto_cargo = self.df.loc[index, "Impuesto Cargo"].values[0]
                        impuesto_retencion = self.df.loc[index, "Impuesto Retención"].values[0]
                        valor_total = precio * cantidad * (1 + impuesto_cargo/100 - impuesto_retencion/100)
                        self.df.loc[index, "Valor Total"] = valor_total
                    else:
                        print(f"Producto {producto['No.Producto']} no encontrado en DataFrame.")
    
                # Guardar cambios en el Excel
                file_path = self.notebook.tabText(self.notebook.currentIndex())
                try:
                    self.df.to_excel(file_path, index=False)
                    print(f"Archivo Excel actualizado correctamente en: {file_path}")
                    self.actualizar_pestaña_excel()
                except Exception as e:
                    print(f"Error al guardar el archivo Excel: {e}")
                    QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {e}")
    
                # Guardar la venta en ventas_data.json
                total_venta = sum(p['Precio'] * p['Cantidad'] for p in productos_factura)
                venta = {
                    'fecha': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'productos': productos_factura,
                    'total': total_venta
                }
    
                try:
                    with open('ventas_data.json', 'r') as f:
                        ventas = json.load(f)
                except FileNotFoundError:
                    ventas = []
    
                ventas.append(venta)
    
                with open('ventas_data.json', 'w') as f:
                    json.dump(ventas, f, indent=4)
    
                QMessageBox.information(self, "Factura Electrónica", "Factura Electrónica generada con éxito.")
                dialog.accept()
            else:
                QMessageBox.warning(self, "Sin Productos", "No se seleccionaron productos para facturar.")
        else:
            QMessageBox.warning(self, "Sin Selección", "Seleccione al menos un producto para generar la factura.")
    
    def cuadre_de_caja(self):
        # Implementación pendiente para realizar cuadre de caja
        print("Realizando Cuadre de Caja (Funcionalidad pendiente)")
        QMessageBox.information(self, "Cuadre de Caja", "Funcionalidad pendiente")
    
    def visualizacion_de_ventas(self):
        # Implementación básica para visualizar ventas
        ventas_dialog = QDialog(self)
        ventas_dialog.setWindowTitle("Visualización de Ventas")
        layout = QVBoxLayout(ventas_dialog)
    
        # Cargar ventas desde ventas_data.json
        try:
            with open('ventas_data.json', 'r') as f:
                ventas = json.load(f)
        except FileNotFoundError:
            ventas = []
    
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["No. Venta", "Fecha", "Total"])
    
        for idx, venta in enumerate(ventas):
            items = [
                QStandardItem(str(idx + 1)),
                QStandardItem(venta['fecha']),
                QStandardItem(f"{venta['total']:.2f}")
            ]
            model.appendRow(items)
    
        tree_view = QTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(tree_view)
    
        ventas_dialog.exec_()
    
    def generar_factura_normal(self):
        # Implementación para generar factura normal
        print("Generando Factura Normal (Funcionalidad pendiente)")
        QMessageBox.information(self, "Generar Factura Normal", "Funcionalidad pendiente")


    def nuevo_archivo(self):
        # Funcionalidad para crear un nuevo archivo (implementación pendiente)
        print("Nuevo archivo creado (Funcionalidad pendiente).")

    def guardar_archivo_actual(self):
        pestaña = self.notebook.currentWidget()
        if hasattr(pestaña, 'file_path') and hasattr(pestaña, 'df'):
            file_path = pestaña.file_path
            try:
                if 'clientes' in os.path.basename(file_path).lower():
                    pestaña.df.to_excel(file_path, index=False)
                    print(f"Archivo de clientes guardado correctamente en: {file_path}")
                else:
                    pestaña.df.to_excel(file_path, index=False)
                    print(f"Archivo guardado correctamente en: {file_path}")
                QMessageBox.information(self, "Guardar Archivo", "Archivo guardado exitosamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {e}")
        else:
            self.guardar_archivo_como()


    def guardar_archivo_como(self):
        # Funcionalidad para "Guardar Como"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Archivo Como", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                self.df.to_excel(file_path, index=False)
                print(f"Archivo guardado correctamente en: {file_path}")
                # Actualizar el nombre de la pestaña
                self.notebook.setTabText(self.notebook.currentIndex(), os.path.basename(file_path))
            except Exception as e:
                print(f"Error al guardar el archivo: {e}")
                QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {e}")




    def insertar_producto_en_tree(self, tree_view, nuevo_producto, cotizacion):
        """
        Inserta un nuevo producto en el QTreeView de la cotización abierta.
    
        :param tree_view: QTreeView de la cotización donde se agregará el producto.
        :param nuevo_producto: Lista con los datos del nuevo producto.
        :param cotizacion: Diccionario de la cotización actual.
        """
        model = tree_view.model()
        if model is None:
            QMessageBox.critical(self, "Error", "El modelo de la vista no está disponible.")
            return
    
        # Crear los items para cada columna
        items = [QStandardItem(str(field)) for field in nuevo_producto]
    
        # Configurar la editabilidad de cada columna
        for col, item in enumerate(items):
            if col in [0, 1, 6, 7, 8]:  # No. Producto, Descripción, Impuestos, Valor Total
                item.setEditable(False)
            else:
                item.setEditable(True)
    
        # Añadir la fila al modelo
        model.appendRow(items)
    
        # Opcional: Actualizar el QTreeView para reflejar el cambio
        tree_view.expandAll()
        tree_view.scrollToBottom()



    def seleccionar_cotizacion_para_agregar(self, dialog, tree_view_cotizaciones, producto_data, precio_venta):
        indexes = tree_view_cotizaciones.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            cotizacion_id = int(tree_view_cotizaciones.model().item(row, 0).text())
            cotizacion = next((c for c in cotizaciones if c['id'] == cotizacion_id), None)
            if not cotizacion:
                QMessageBox.warning(self, "Error", "Cotización no encontrada.")
                return
    
            # Agregar el producto a la cotización seleccionada
            nuevo_producto = [
                int(producto_data[0]),  # No.Producto
                producto_data[1],       # Descripción
                1,                      # Cantidad inicial
                float(producto_data[2]),# Precio Costo
                float(precio_venta),    # Precio Venta
                0.0,                    # % Desc.
                19.0,                   # Impuesto Cargo
                2.0,                    # Impuesto Retención
                0.0                     # Valor Total (se calculará)
            ]
            cotizacion['productos'].append(nuevo_producto)
            cotizacion['fecha'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"Producto {producto_data[1]} agregado a la Cotización '{cotizacion['nombre']}'")
    
            # Actualizar en tiempo real si la cotización ya está abierta
            for i in range(self.notebook.count()):
                if self.notebook.tabText(i) == cotizacion['nombre']:
                    pestaña = self.notebook.widget(i)
                    tree_productos = pestaña.tree_productos
                    self.insertar_producto_en_tree(tree_productos, nuevo_producto, cotizacion)
                    break
    
            # Guardar la cotización actualizada
            self.guardar_cotizacion_excel(cotizacion)
    
            # Cerrar el diálogo
            dialog.accept()
        else:
            QMessageBox.warning(self, "Error", "Seleccione una cotización.")






    def on_product_data_changed(self, model, topLeft, bottomRight, cotizacion):
        # Asegurarse de que el cambio se hizo en una sola celda
        if topLeft.row() != bottomRight.row() or topLeft.column() != bottomRight.column():
            return
    
        row = topLeft.row()
        column = topLeft.column()
    
        # Indices de las columnas relevantes
        CANTIDAD_COL = 2
        PRECIO_COSTO_COL = 3
        PRECIO_VENTA_COL = 4
        DESCUENTO_COL = 5
        IMPUESTO_CARGO_COL = 6
        IMPUESTO_RETENCION_COL = 7
        VALOR_TOTAL_COL = 8
    
        # Obtener los datos actuales de la fila
        try:
            cantidad = float(model.item(row, CANTIDAD_COL).text())
            precio_costo = float(model.item(row, PRECIO_COSTO_COL).text())
            precio_venta = float(model.item(row, PRECIO_VENTA_COL).text())
            descuento = float(model.item(row, DESCUENTO_COL).text())
            impuesto_cargo = float(model.item(row, IMPUESTO_CARGO_COL).text())
            impuesto_retencion = float(model.item(row, IMPUESTO_RETENCION_COL).text())
        except ValueError:
            QMessageBox.warning(self, "Error de Entrada", "Por favor, ingrese valores numéricos válidos.")
            return
    
        # Cálculos
        precio_con_descuento = precio_venta * (1 - descuento / 100)
        subtotal = precio_con_descuento * cantidad
        total_impuestos_cargo = subtotal * (impuesto_cargo / 100)
        total_impuestos_retencion = subtotal * (impuesto_retencion / 100)
        valor_total = subtotal + total_impuestos_cargo - total_impuestos_retencion
    
        # Actualizar el "Valor Total" en el modelo
        valor_total_item = model.item(row, VALOR_TOTAL_COL)
        valor_total_item.setText(f"{valor_total:.2f}")
    
        # Actualizar los datos en 'cotizacion', asegurando 'No.Producto' es int
        try:
            no_producto = int(model.item(row, 0).text())
        except ValueError:
            QMessageBox.critical(self, "Error de Datos", f"El 'No.Producto' debe ser un número entero válido.")
            return
    
        cotizacion['productos'][row] = [
            no_producto,          # No. Producto (int)
            model.item(row, 1).text(),  # Descripción
            cantidad,             # Cantidad
            precio_costo,         # Precio Costo
            precio_venta,         # Precio Venta
            descuento,            # % Desc.
            impuesto_cargo,       # Impuesto Cargo
            impuesto_retencion,   # Impuesto Retención
            valor_total           # Valor Total
        ]
    
        # Actualizar la fecha de última modificación
        cotizacion['fecha'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
        # Guardar los cambios en el archivo Excel
        self.guardar_cotizacion_excel(cotizacion)



    def mostrar_menu_datos(self, position, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            menu = QMenu()
            copiar_action = QAction("Copiar", self)
            copiar_action.triggered.connect(lambda: self.copiar_datos(tree_view))
            menu.addAction(copiar_action)

            detalles_action = QAction("Detalles", self)
            detalles_action.triggered.connect(lambda: self.mostrar_detalles_producto(tree_view))
            menu.addAction(detalles_action)

            enviar_info_action = QAction("Enviar información del producto", self)
            enviar_info_action.triggered.connect(lambda: self.enviar_informacion_producto(tree_view))
            menu.addAction(enviar_info_action)

            agregar_cotizacion_action = QAction("Agregar a Cotización", self)
            agregar_cotizacion_action.triggered.connect(lambda: self.agregar_a_cotizacion(tree_view))
            menu.addAction(agregar_cotizacion_action)

            menu.exec_(tree_view.viewport().mapToGlobal(position))
        else:
            print("Clic fuera de una fila válida.")

    def copiar_datos(self, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            selected_data = [index.data(Qt.DisplayRole) for index in indexes]
            pyperclip.copy("\t".join(map(str, selected_data)))
            print("Datos copiados:", selected_data)

    def mostrar_detalles_producto(self, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            data = [tree_view.model().item(row, col).data(Qt.DisplayRole) for col in range(tree_view.model().columnCount())]
            detalles = {self.df.columns[i]: data[i] for i in range(len(data))}
    
            # Crear un diálogo personalizado para mostrar los detalles
            detalles_dialog = QDialog(self)
            detalles_dialog.setWindowTitle(f"Detalles del Producto ID: {detalles.get('No.Producto', 'N/A')}")
            detalles_dialog.resize(400, 300)
            layout = QVBoxLayout(detalles_dialog)
    
            # Usar un layout de grid para una mejor presentación
            grid = QGridLayout()
    
            for i, (clave, valor) in enumerate(detalles.items()):
                label_clave = QLabel(f"<b>{clave}:</b>")
                label_valor = QLabel(str(valor))
                grid.addWidget(label_clave, i, 0, Qt.AlignRight)
                grid.addWidget(label_valor, i, 1, Qt.AlignLeft)
    
            layout.addLayout(grid)
    
            # Botón para cerrar el diálogo
            btn_cerrar = QPushButton("Cerrar")
            btn_cerrar.clicked.connect(detalles_dialog.accept)
            layout.addWidget(btn_cerrar, alignment=Qt.AlignCenter)
    
            detalles_dialog.exec_()


    def enviar_informacion_producto(self, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            data = indexes[0].model().itemFromIndex(indexes[0]).data(Qt.DisplayRole)
            print(f"Enviar información del producto: {data} (Funcionalidad pendiente)")

    def agregar_a_cotizacion(self, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            data = [tree_view.model().item(row, col).data(Qt.DisplayRole) for col in range(tree_view.model().columnCount())]
    
            # Obtener el negocio actual de la cotización
            pestaña = self.notebook.currentWidget()
            if not hasattr(pestaña, 'cotizacion'):
                QMessageBox.warning(self, "Error", "No se pudo determinar la cotización actual.")
                return
            cotizacion = pestaña.cotizacion
            negocio_nombre = cotizacion.get('negocio', '')
            if not negocio_nombre:
                QMessageBox.warning(self, "Error", "La cotización no tiene un negocio asociado.")
                return
    
            cotizacion_dialog = QDialog(self)
            cotizacion_dialog.setWindowTitle("Seleccionar Cotización")
            layout = QVBoxLayout(cotizacion_dialog)
    
            # Filtrar las cotizaciones por el negocio actual
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(["ID", "Nombre", "Negocio", "Estado"])
    
            for cot in cotizaciones:
                if cot.get('negocio') == negocio_nombre:
                    items = [
                        QStandardItem(str(cot['id'])),
                        QStandardItem(cot.get('nombre', f"Cotización {cot['id']}")),
                        QStandardItem(cot.get('negocio', '')),
                        QStandardItem(cot['estado'])
                    ]
                    model.appendRow(items)
    
            tree_view_cotizaciones = QTreeView()
            tree_view_cotizaciones.setModel(model)
            tree_view_cotizaciones.setAlternatingRowColors(True)
            tree_view_cotizaciones.setSortingEnabled(True)
            tree_view_cotizaciones.setEditTriggers(QTreeView.NoEditTriggers)
            header = tree_view_cotizaciones.header()
            header.setSectionResizeMode(QHeaderView.Stretch)
            layout.addWidget(tree_view_cotizaciones)
    
            # Campos adicionales
            precio_label = QLabel("Precio Venta:")
            precio_input = QLineEdit()
            precio_input.setText(str(data[2]))  # Precio inicial igual al de costo
            layout.addWidget(precio_label)
            layout.addWidget(precio_input)
    
            btn_seleccionar = QPushButton("Seleccionar")
            btn_seleccionar.clicked.connect(lambda: self.seleccionar_cotizacion_para_agregar(
                cotizacion_dialog, tree_view_cotizaciones, data, precio_input.text()
            ))
            layout.addWidget(btn_seleccionar)
    
            cotizacion_dialog.exec_()
        else:
            QMessageBox.warning(self, "Error", "Seleccione un producto para agregar a la cotización.")

    def buscar_producto(self):
        """
        Busca un producto en el DataFrame de productos basado en un campo y término de búsqueda.
        """
        if not hasattr(self, 'df_productos'):
            QMessageBox.warning(self, "Buscar Producto", "No hay datos de productos cargados.")
            return
    
        campos = self.df_productos.columns.tolist()
        campo, ok = QInputDialog.getItem(self, "Buscar Producto", "Seleccione el campo de búsqueda:", campos, 0, False)
        if ok and campo:
            search_term, ok = QInputDialog.getText(self, "Buscar Producto", f"Ingrese el término de búsqueda para '{campo}':")
            if ok and search_term:
                encontrado = False
                for i in range(self.notebook.count()):
                    pestaña = self.notebook.widget(i)
                    # Verificar que la pestaña corresponda a productos
                    if hasattr(pestaña, 'tree_view') and 'productos' in self.notebook.tabText(i).lower():
                        tree_view = pestaña.tree_view
                        model = tree_view.model()
                        try:
                            col_index = self.df_productos.columns.get_loc(campo)
                        except KeyError:
                            QMessageBox.warning(self, "Error", f"El campo '{campo}' no existe en los datos de productos.")
                            return
    
                        for row in range(model.rowCount()):
                            item = model.item(row, col_index)
                            if search_term.lower() in str(item.data(Qt.DisplayRole)).lower():
                                index = model.indexFromItem(item)
                                tree_view.scrollTo(index, QTreeView.PositionAtCenter)
                                tree_view.setCurrentIndex(index)
                                print(f"Producto encontrado en fila {row + 1}, columna '{campo}'.")
                                encontrado = True
                                break
                    if encontrado:
                        break
                if not encontrado:
                    QMessageBox.information(self, "Buscar Producto", f"No se encontró ningún producto que coincida con: {search_term} en el campo '{campo}'.")

    def seleccionar_producto(self):
        tree_view = self.sender()
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            data = [tree_view.model().item(row, col).data(Qt.DisplayRole) for col in range(tree_view.model().columnCount())]

            detalles_dialog = QDialog(self)
            detalles_dialog.setWindowTitle(f"Detalles del Producto ID: {data[0]}")
            layout = QVBoxLayout(detalles_dialog)

            for i, value in enumerate(data):
                label = QLabel(f"{self.df.columns[i]}: {value}")
                layout.addWidget(label)

            # No añadimos botones para modificar o agregar
            detalles_dialog.exec_()






    def calcular_valor_total_producto(self, model, row, cotizacion):
        # Indices de las columnas relevantes
        CANTIDAD_COL = 2
        PRECIO_COSTO_COL = 3
        PRECIO_VENTA_COL = 4
        DESCUENTO_COL = 5
        IMPUESTO_CARGO_COL = 6
        IMPUESTO_RETENCION_COL = 7
        VALOR_TOTAL_COL = 8
    
        # Obtener los datos actuales de la fila
        try:
            cantidad = float(model.item(row, CANTIDAD_COL).text())
            precio_costo = float(model.item(row, PRECIO_COSTO_COL).text())
            precio_venta = float(model.item(row, PRECIO_VENTA_COL).text())
            descuento = float(model.item(row, DESCUENTO_COL).text())
            impuesto_cargo = float(model.item(row, IMPUESTO_CARGO_COL).text())
            impuesto_retencion = float(model.item(row, IMPUESTO_RETENCION_COL).text())
        except ValueError:
            QMessageBox.warning(self, "Error de Entrada", "Por favor, ingrese valores numéricos válidos.")
            return
    
        # Cálculos
        precio_con_descuento = precio_venta * (1 - descuento / 100)
        subtotal = precio_con_descuento * cantidad
        total_impuestos_cargo = subtotal * (impuesto_cargo / 100)
        total_impuestos_retencion = subtotal * (impuesto_retencion / 100)
        valor_total = subtotal + total_impuestos_cargo - total_impuestos_retencion
    
        # Actualizar el "Valor Total" en el modelo
        valor_total_item = model.item(row, VALOR_TOTAL_COL)
        valor_total_item.setText(f"{valor_total:.2f}")
    
        # Actualizar los datos en 'cotizacion'
        cotizacion['productos'][row] = [
            model.item(row, 0).text(),        # No. Producto
            model.item(row, 1).text(),        # Descripción
            cantidad,                         # Cantidad
            precio_costo,                     # Precio Costo
            precio_venta,                     # Precio Venta
            descuento,                        # % Desc.
            impuesto_cargo,                   # Impuesto Cargo
            impuesto_retencion,               # Impuesto Retención
            valor_total                       # Valor Total
        ]
    
        # Actualizar la fecha de última modificación
        cotizacion['fecha'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
        # Guardar los cambios en el archivo Excel
        self.guardar_cotizacion_excel(cotizacion)




    

        
    def agregar_nueva_cotizacion(self, model):
        """
        Crea una nueva cotización y permite seleccionar el negocio al que estará asociada,
        sin necesidad de tener una pestaña del negocio abierta.
        """
        nombre_cotizacion, ok = QInputDialog.getText(self, "Nombre de Cotización", "Ingrese el nombre de la cotización:")
        if ok and nombre_cotizacion:
            nombres_existentes = [c['nombre'] for c in cotizaciones]
            if nombre_cotizacion in nombres_existentes:
                QMessageBox.warning(
                    self, "Error", f"Ya existe una cotización con el nombre '{nombre_cotizacion}'. Por favor, elija otro nombre.")
                return
    
            # Verificar que self.negocios esté cargado
            if not self.negocios:
                QMessageBox.warning(self, "Error", "No hay negocios disponibles. Por favor, agregue un negocio primero.")
                return
    
            # Mostrar diálogo para seleccionar el negocio
            negocios_nombres = [negocio['nombre'] for negocio in self.negocios]
            negocio_seleccionado, ok = QInputDialog.getItem(self, "Seleccionar Negocio", "Seleccione el negocio para la cotización:", negocios_nombres, 0, False)
            if not ok or not negocio_seleccionado:
                QMessageBox.warning(self, "Advertencia", "Debe seleccionar un negocio para la cotización.")
                return
    
            # Asignar un ID único incrementando el máximo ID existente
            if cotizaciones:
                new_id = max(c['id'] for c in cotizaciones) + 1
            else:
                new_id = 1
    
            # Inicializar campos adicionales
            nueva_cotizacion = {
                'id': new_id,
                'nombre': nombre_cotizacion,
        
                'negocio': negocio_seleccionado,  # Asociar el negocio seleccionado
                'estado': 'Pendiente',
                'fecha': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'productos': [],
                'cliente': '',
                'contacto': '',
                'responsable': '',
                'centro_costo': '',
                'tipo': 'C-1 Cotización',
                'fecha_elaboracion': datetime.now().strftime("%d/%m/%Y"),
                'moneda': 'COP - Peso colombiano',
                'encabezado': ''
            }
            cotizaciones.append(nueva_cotizacion)
            items = [
                QStandardItem(str(nueva_cotizacion['id'])),
                QStandardItem(nueva_cotizacion['nombre']),
                
                QStandardItem(nueva_cotizacion['negocio']),  # Mostrar el negocio en la tabla
                QStandardItem(nueva_cotizacion['estado']),
                QStandardItem(nueva_cotizacion['fecha']),
                QStandardItem(nueva_cotizacion['cliente'])
            ]
            items[3].setEditable(True)  # 'Estado' es editable
            model.appendRow(items)
            print("Nueva cotización agregada:", nueva_cotizacion)
            # Guardar cambios en JSON
            self.guardar_cotizaciones()
        else:
            QMessageBox.warning(self, "Advertencia", "Debe ingresar un nombre para la cotización.")


    def obtener_negocio_actual(self):
        # Suponiendo que la pestaña actual corresponde al negocio activo
        pestaña = self.notebook.currentWidget()
        if hasattr(pestaña, 'negocio'):
            negocio_nombre = pestaña.negocio
            negocio = next((n for n in self.negocios if n['nombre'] == negocio_nombre), None)
            return negocio
        else:
            # Si no se puede determinar el negocio actual
            return None
    
    
    
    
        def buscar_cotizacion(self, model, search_text):
            found = False
            for row in range(model.rowCount()):
                item_nombre = model.item(row, 1)
                if search_text.lower() in item_nombre.text().lower():
                    # Seleccionar y desplazar hasta el elemento encontrado
                    index = model.indexFromItem(item_nombre)
                    self.cotizaciones_tree_view.setCurrentIndex(index)
                    self.cotizaciones_tree_view.scrollTo(index)
                    found = True
                    break
            if not found:
                QMessageBox.information(self, "Buscar Cotización", f"No se encontró ninguna cotización con nombre: {search_text}")
        
    

    def actualizar_detalle_cotizacion(self, cotizacion, campo, valor):
        cotizacion[campo] = valor
        cotizacion['fecha'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Guardar los cambios en el archivo JSON
        with open('cotizaciones_data.json', 'w') as f:
            json.dump(cotizaciones, f, indent=4)


    
    def abrir_pestaña_cotizacion_by_id(self, cotizacion_id):
        """
        Abre una pestaña para la cotización especificada por ID, considerando el negocio asociado.
        """
        print(f"Intentando abrir la cotización con ID: {cotizacion_id}")
        try:
            cotizacion_id = int(cotizacion_id)
        except ValueError:
            print(f"Error: ID de cotización inválido: {cotizacion_id}")
            QMessageBox.warning(self, "Error", f"ID de cotización inválido: {cotizacion_id}.")
            return
    
        cotizacion = next((c for c in cotizaciones if c['id'] == cotizacion_id), None)
        if not cotizacion:
            print(f"Error: No se encontró la cotización con ID {cotizacion_id}")
            QMessageBox.warning(self, "Error", f"No se encontró la cotización con ID {cotizacion_id}.")
            return
    
        cotizacion_nombre = cotizacion['nombre']
        negocio_nombre = cotizacion.get('negocio', '')
        if not negocio_nombre:
            QMessageBox.warning(self, "Error", "La cotización no tiene un negocio asociado.")
            return
    
        # Verificar si la pestaña ya está abierta
        for i in range(self.notebook.count()):
            if self.notebook.tabText(i) == cotizacion_nombre:
                print(f"La pestaña para la cotización '{cotizacion_nombre}' ya está abierta. Enfocando pestaña existente.")
                self.notebook.setCurrentIndex(i)
                return
    
        # Crear la nueva pestaña
        print(f"Creando una nueva pestaña para la cotización '{cotizacion_nombre}'...")
        nueva_pestaña = QWidget()
        layout = QVBoxLayout(nueva_pestaña)
    
        # Guardar la cotización y el negocio asociado en la pestaña
        nueva_pestaña.cotizacion = cotizacion
        nueva_pestaña.negocio = negocio_nombre

    
        # Sección de detalles de la cotización
        form_layout = QGridLayout()
        tipo_label = QLabel("Tipo:")
        tipo_combo = QComboBox()
        tipo_combo.addItems(["C-1 Cotización", "C-2 Otro tipo"])
        tipo_combo.setCurrentText(cotizacion.get('tipo', 'C-1 Cotización'))
    
        cliente_label = QLabel("Cliente:")
        cliente_input = QLineEdit(cotizacion.get('cliente', ''))
        cliente_buscar_btn = QPushButton("🔍")
        cliente_buscar_btn.clicked.connect(lambda: self.buscar_o_agregar_cliente_para_cotizacion(cotizacion, cliente_input))
    
        contacto_label = QLabel("Contacto:")
        contacto_combo = QComboBox()
        contacto_combo.addItems(["Seleccionar", "Contacto 1", "Contacto 2"])
        contacto_combo.setCurrentText(cotizacion.get('contacto', 'Seleccionar'))
    
        responsable_label = QLabel("Responsable de la cotización:")
        responsable_input = QLineEdit(cotizacion.get('responsable', ''))
        responsable_buscar_btn = QPushButton("🔍")
    
        centro_label = QLabel("Centro de costo:")
        centro_input = QLineEdit(cotizacion.get('centro_costo', ''))
        centro_buscar_btn = QPushButton("🔍")
    
        numero_label = QLabel("Número:")
        numero_input = QLineEdit("Numeración Automática")
        numero_input.setReadOnly(True)
    
        fecha_label = QLabel("Fecha de elaboración:")
        fecha_input = QLineEdit(cotizacion.get('fecha_elaboracion', datetime.now().strftime("%d/%m/%Y")))
        fecha_calendario_btn = QPushButton("📅")
    
        moneda_label = QLabel("Moneda:")
        moneda_combo = QComboBox()
        moneda_combo.addItems(["COP - Peso colombiano", "USD - Dólar estadounidense"])
        moneda_combo.setCurrentText(cotizacion.get('moneda', 'COP - Peso colombiano'))
    
        # Agregar los widgets al grid layout
        form_layout.addWidget(tipo_label, 0, 0)
        form_layout.addWidget(tipo_combo, 0, 1)
        form_layout.addWidget(cliente_label, 1, 0)
        form_layout.addWidget(cliente_input, 1, 1)
        form_layout.addWidget(cliente_buscar_btn, 1, 2)
        form_layout.addWidget(contacto_label, 2, 0)
        form_layout.addWidget(contacto_combo, 2, 1)
        form_layout.addWidget(responsable_label, 3, 0)
        form_layout.addWidget(responsable_input, 3, 1)
        form_layout.addWidget(responsable_buscar_btn, 3, 2)
        form_layout.addWidget(centro_label, 4, 0)
        form_layout.addWidget(centro_input, 4, 1)
        form_layout.addWidget(centro_buscar_btn, 4, 2)
    
        form_layout.addWidget(numero_label, 0, 3)
        form_layout.addWidget(numero_input, 0, 4)
        form_layout.addWidget(fecha_label, 1, 3)
        form_layout.addWidget(fecha_input, 1, 4)
        form_layout.addWidget(fecha_calendario_btn, 1, 5)
        form_layout.addWidget(moneda_label, 2, 3)
        form_layout.addWidget(moneda_combo, 2, 4)
    
        layout.addLayout(form_layout)
    
        # Encabezado
        encabezado_label = QLabel("Encabezado:")
        encabezado_input = QLineEdit(cotizacion.get('encabezado', ''))
        layout.addWidget(encabezado_label)
        layout.addWidget(encabezado_input)
    
        # Tabla de productos
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels([
            "No. Producto", "Descripción", "Cantidad", "Precio Costo", "Precio Venta",
            "% Desc.", "Impuesto Cargo", "Impuesto Retención", "Valor Total"
        ])
    
        # Cargar los productos de la cotización
        for producto in cotizacion['productos']:
            while len(producto) < 9:
                producto.append(0)
            items = [QStandardItem(str(field)) for field in producto]
            # Configurar la editabilidad de cada columna
            for col, item in enumerate(items):
                if col in [0, 1, 6, 7, 8]:  # No. Producto, Descripción, Impuestos, Valor Total
                    item.setEditable(False)
                else:
                    item.setEditable(True)
            model.appendRow(items)
    
        tree_productos = QTreeView()
        tree_productos.setModel(model)
        tree_productos.setAlternatingRowColors(True)
        tree_productos.setSortingEnabled(False)
        tree_productos.setEditTriggers(QTreeView.DoubleClicked | QTreeView.SelectedClicked)
        header = tree_productos.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(tree_productos)
        nueva_pestaña.tree_productos = tree_productos
    
        # Conectar el evento dataChanged
        model.dataChanged.connect(lambda topLeft, bottomRight, roles: self.on_product_data_changed(
            model, topLeft, bottomRight, cotizacion))
    
        # Añadir botones para acciones adicionales
        btn_layout = QHBoxLayout()
        btn_eliminar_producto = QPushButton("Eliminar Producto")
        btn_eliminar_producto.clicked.connect(lambda: self.eliminar_producto_cotizacion(tree_productos, cotizacion, cotizacion['id']))
        btn_layout.addWidget(btn_eliminar_producto)
    
        btn_guardar_pdf = QPushButton("Guardar como PDF")
        btn_guardar_pdf.clicked.connect(lambda: self.guardar_cotizacion_pdf(cotizacion['id']))
        btn_layout.addWidget(btn_guardar_pdf)
    
        btn_enviar_cotizacion = QPushButton("Enviar Cotización")
        btn_enviar_cotizacion.clicked.connect(lambda: self.enviar_cotizacion(cotizacion['id']))
        btn_layout.addWidget(btn_enviar_cotizacion)
    
        layout.addLayout(btn_layout)
    
        # Guardar referencias a los widgets para poder guardar los cambios
        nueva_pestaña.tipo_combo = tipo_combo
        nueva_pestaña.cliente_input = cliente_input
        nueva_pestaña.contacto_combo = contacto_combo
        nueva_pestaña.responsable_input = responsable_input
        nueva_pestaña.centro_input = centro_input
        nueva_pestaña.fecha_input = fecha_input
        nueva_pestaña.moneda_combo = moneda_combo
        nueva_pestaña.encabezado_input = encabezado_input
    
        # Conectar señales para guardar cambios
        tipo_combo.currentTextChanged.connect(lambda value: self.actualizar_detalle_cotizacion(cotizacion, 'tipo', value))
        cliente_input.textChanged.connect(lambda value: self.actualizar_detalle_cotizacion(cotizacion, 'cliente', value))
        contacto_combo.currentTextChanged.connect(lambda value: self.actualizar_detalle_cotizacion(cotizacion, 'contacto', value))
        responsable_input.textChanged.connect(lambda value: self.actualizar_detalle_cotizacion(cotizacion, 'responsable', value))
        centro_input.textChanged.connect(lambda value: self.actualizar_detalle_cotizacion(cotizacion, 'centro_costo', value))
        fecha_input.textChanged.connect(lambda value: self.actualizar_detalle_cotizacion(cotizacion, 'fecha_elaboracion', value))
        moneda_combo.currentTextChanged.connect(lambda value: self.actualizar_detalle_cotizacion(cotizacion, 'moneda', value))
        encabezado_input.textChanged.connect(lambda value: self.actualizar_detalle_cotizacion(cotizacion, 'encabezado', value))
    
        # Añadir la pestaña a la interfaz
        self.notebook.addTab(nueva_pestaña, cotizacion_nombre)
        self.notebook.setCurrentWidget(nueva_pestaña)
        print(f"Pestaña de cotización '{cotizacion_nombre}' abierta correctamente.")


    def crear_tree_view(self, model, editable_columns=[], ocultar_columnas=[]):
        """
        Crea y configura un CustomTreeView con las propiedades deseadas.
    
        :param model: QStandardItemModel para el QTreeView.
        :param editable_columns: Lista de índices de columnas que serán editables.
        :param ocultar_columnas: Lista de índices de columnas que serán ocultadas.
        :return: Configurado CustomTreeView.
        """
        tree_view = CustomTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setSortingEnabled(False)  # Deshabilitar ordenamiento
        tree_view.setEditTriggers(QTreeView.DoubleClicked | QTreeView.SelectedClicked)
        
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionsClickable(False)  # Hacer que los encabezados no sean clicables
    
        # Ocultar columnas especificadas
        for col in ocultar_columnas:
            tree_view.hideColumn(col)
    
        # Configurar la editabilidad de columnas específicas
        for row in range(model.rowCount()):
            for col in editable_columns:
                item = model.item(row, col)
                if item:
                    item.setEditable(True)
    
        return tree_view



    def buscar_o_agregar_cliente_para_cotizacion(self, cotizacion, cliente_input):
        dialog = QDialog(self)
        dialog.setWindowTitle("Seleccionar o Agregar Cliente")
        dialog.resize(700, 500)
        layout = QVBoxLayout(dialog)
        negocio_nombre = cotizacion.get('negocio', '')
        if not negocio_nombre:
            QMessageBox.warning(self, "Error", "La cotización no tiene un negocio asociado.")
            return
    
        # Cargar los clientes del negocio correspondiente
        clientes_df = self.cargar_clientes_del_negocio(negocio_nombre)
        if clientes_df is None:
            QMessageBox.warning(self, "Error", f"No se encontraron clientes para el negocio '{negocio_nombre}'.")
            return

    
        # Sección para seleccionar el tipo de búsqueda
        search_type_layout = QHBoxLayout()
        search_type_label = QLabel("Buscar por:")
        search_type_combo = QComboBox()
        search_type_combo.addItems(["Identificación", "Nombre Completo"])
        search_type_layout.addWidget(search_type_label)
        search_type_layout.addWidget(search_type_combo)
        layout.addLayout(search_type_layout)
    
        # Campo de búsqueda
        search_layout = QHBoxLayout()
        search_label = QLabel("Término de búsqueda:")
        search_input = QLineEdit()
        search_button = QPushButton("Buscar")
        search_button.clicked.connect(lambda: self.buscar_cliente_en_dialog(
            search_type_combo.currentText(),
            search_input.text(),
            tree_view
        ))
        search_layout.addWidget(search_label)
        search_layout.addWidget(search_input)
        search_layout.addWidget(search_button)
        layout.addLayout(search_layout)
    
        # Vista de clientes
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(clientes_df.columns)
    
        tree_view = QTreeView()
        tree_view.setModel(model)
        tree_view.setAlternatingRowColors(True)
        tree_view.setSortingEnabled(True)
        tree_view.setEditTriggers(QTreeView.NoEditTriggers)
        header = tree_view.header()
        header.setSectionResizeMode(QHeaderView.Stretch)
    
        # Cargar clientes en el modelo
        for index, row in clientes_df.iterrows():
            items = []
            for field in row:
                item = QStandardItem(str(field))
                item.setEditable(False)
                items.append(item)
            model.appendRow(items)
    
        # Context menu para copiar o ver detalles
        tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        tree_view.customContextMenuRequested.connect(lambda pos: self.mostrar_menu_datos_cliente_en_dialog(pos, tree_view))
    
        layout.addWidget(tree_view)
    
        # Botones
        btn_layout = QHBoxLayout()
        btn_seleccionar = QPushButton("Seleccionar Cliente")
        btn_seleccionar.clicked.connect(lambda: self.seleccionar_cliente_en_dialog(tree_view, cotizacion, cliente_input, dialog))
        btn_agregar = QPushButton("Agregar Nuevo Cliente")
        btn_agregar.clicked.connect(lambda: self.agregar_cliente_desde_dialog(dialog))
        btn_layout.addWidget(btn_seleccionar)
        btn_layout.addWidget(btn_agregar)
        layout.addLayout(btn_layout)
    
        dialog.exec_()
    


    def cargar_clientes_del_negocio(self, negocio_nombre):
        # Buscar el negocio en self.negocios
        negocio = next((n for n in self.negocios if n['nombre'] == negocio_nombre), None)
        if negocio:
            file_path = negocio.get('ruta_clientes')
            if file_path and os.path.exists(file_path):
                try:
                    df = pd.read_excel(file_path, dtype={'Identificacion': str})
                    return df
                except Exception as e:
                    print(f"Error al cargar clientes del negocio '{negocio_nombre}': {e}")
                    return None
            else:
                print(f"No se encontró el archivo de clientes para el negocio '{negocio_nombre}'.")
                return None
        else:
            print(f"Negocio '{negocio_nombre}' no encontrado.")
            return None

        
    def cargar_productos_del_negocio(self, negocio_nombre):
        # Buscar el negocio en self.negocios
        negocio = next((n for n in self.negocios if n['nombre'] == negocio_nombre), None)
        if negocio:
            file_path = negocio.get('ruta_productos')
            if file_path and os.path.exists(file_path):
                try:
                    df = pd.read_excel(file_path)
                    return df
                except Exception as e:
                    print(f"Error al cargar productos del negocio '{negocio_nombre}': {e}")
                    return None
            else:
                print(f"No se encontró el archivo de productos para el negocio '{negocio_nombre}'.")
                return None
        else:
            print(f"Negocio '{negocio_nombre}' no encontrado.")
            return None

        
    def buscar_cliente_en_dialog(self, search_type, search_term, tree_view):
        model = tree_view.model()
        found = False
    
        if search_type == "Identificación":
            for row in range(model.rowCount()):
                identificacion = model.item(row, model.columnCount() - 1).text().strip()  # Asumiendo que 'Identificacion' es la última columna
                if search_term.lower() == identificacion.lower():
                    index = model.indexFromItem(model.item(row, 0))
                    tree_view.setCurrentIndex(index)
                    tree_view.scrollTo(index, QTreeView.PositionAtCenter)
                    found = True
                    break

        elif search_type == "Nombre Completo":
            for row in range(model.rowCount()):
                nombre1 = model.item(row, 0).text().strip()
                nombre2 = model.item(row, 1).text().strip()
                apellido1 = model.item(row, 2).text().strip()
                apellido2 = model.item(row, 3).text().strip()
                nombre_completo = ' '.join(filter(None, [nombre1, nombre2, apellido1, apellido2]))
                if search_term.lower() in nombre_completo.lower():
                    index = model.indexFromItem(model.item(row, 0))
                    tree_view.setCurrentIndex(index)
                    tree_view.scrollTo(index, QTreeView.PositionAtCenter)
                    found = True
                    break
    
        if not found:
            QMessageBox.information(self, "Buscar Cliente", f"No se encontró ningún cliente que coincida con: {search_term} en {search_type}.")


    def seleccionar_cliente_en_dialog(self, tree_view, cotizacion, cliente_input, dialog):
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            model = tree_view.model()
            
            # Asumiendo que las columnas son: Nombre1, Nombre2, Apellido1, Apellido2, Identificacion
            nombre1 = model.item(row, 0).text().strip()
            nombre2 = model.item(row, 1).text().strip()
            apellido1 = model.item(row, 2).text().strip()
            apellido2 = model.item(row, 3).text().strip()
            
            # Concatenar el nombre completo sin comas ni otros separadores
            nombre_completo = ' '.join(filter(None, [nombre1, nombre2, apellido1, apellido2]))
            
            # Actualizar la cotización
            cotizacion['cliente'] = nombre_completo
            # Si deseas almacenar 'id_cliente', puedes agregarlo así:
            # cotizacion['id_cliente'] = model.item(row, 4).text()
            
            # Actualizar el QLineEdit
            cliente_input.setText(nombre_completo)
            
            # Guardar los cambios en la cotización
            self.guardar_cotizaciones()
            dialog.accept()
        else:
            QMessageBox.warning(self, "Seleccionar Cliente", "Seleccione un cliente de la lista.")

         
    def guardar_cotizaciones(self):
        try:
            with open('cotizaciones_data.json', 'w') as f:
                json.dump(cotizaciones, f, indent=4)
            print("Cotizaciones guardadas correctamente en 'cotizaciones_data.json'.")
        except Exception as e:
            QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar las cotizaciones.\nError: {e}")
            print(f"Error al guardar las cotizaciones: {e}")



    def agregar_cliente_desde_dialog(self, dialog):
        self.agregar_cliente()
        # Recargar los clientes en el diálogo
        self.cargar_clientes_guardados()
        QMessageBox.information(self, "Agregar Cliente", "Nuevo cliente agregado exitosamente. Por favor, realice la búsqueda nuevamente si es necesario.")


    def actualizar_cotizacion(self, item):
        row = item.row()
        column = item.column()
        value = item.text()
        model = item.model()
    
        # Depuración: Imprimir información del item cambiado
        print(f"Item cambiado - Fila: {row}, Columna: {column}, Nuevo Valor: {value}")
    
        # Obtener el ID de la cotización desde la primera columna (visible ahora)
        try:
            cotizacion_id = int(model.item(row, 0).text())
            print(f"ID de Cotización: {cotizacion_id}")
        except ValueError:
            QMessageBox.critical(self, "Error de Datos", "El ID de la cotización no es válido.")
            print("Error: El ID de la cotización no es válido.")
            return
    
        # Buscar la cotización en la lista global por ID
        cotizacion = next((c for c in cotizaciones if c['id'] == cotizacion_id), None)
        if cotizacion:
            print(f"Cotización encontrada: {cotizacion['nombre']}")
            if column == 3:  # Estado
                cotizacion['estado'] = value
                cotizacion['fecha'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                print(f"Estado actualizado a: {value}")
                print(f"Fecha actualizada a: {cotizacion['fecha']}")
                # Actualizar la fecha en la tabla
                item_fecha = model.item(row, 4)
                item_fecha.setText(cotizacion['fecha'])
                print("Fecha en la tabla actualizada.")
    
                # Guardar cambios en JSON
                self.guardar_cotizaciones()
        else:
            QMessageBox.warning(self, "Error", f"No se encontró la cotización con ID {cotizacion_id}.")
            print(f"Error: No se encontró la cotización con ID {cotizacion_id}.")


    
    def enviar_cotizacion(self, cotizacion_id):
        # Implementación para enviar la cotización
        print(f'Cotización {cotizacion_id} enviada (Funcionalidad pendiente).')

    def eliminar_cotizacion_cotizaciones_dialog(self, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            model = tree_view.model()
            # Obtener el ID de la cotización desde la primera columna (oculta)
            cotizacion_id = int(model.item(row, 0).text())
            # Buscar la cotización en la lista global por ID
            cotizacion = next((c for c in cotizaciones if c['id'] == cotizacion_id), None)
            if not cotizacion:
                QMessageBox.warning(self, "Error", f"No se encontró la cotización con ID {cotizacion_id}.")
                return
            # Confirmar eliminación
            reply = QMessageBox.question(self, 'Eliminar Cotización',
                                         f"¿Está seguro de que desea eliminar la cotización '{cotizacion['nombre']}'?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                # Eliminar de la lista de cotizaciones
                cotizaciones.remove(cotizacion)
                # Eliminar del modelo
                model.removeRow(row)
                # Eliminar el archivo Excel asociado
                file_path = os.path.join('cotizaciones', f"{cotizacion['nombre']}.xlsx")
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f'Archivo {file_path} eliminado.')
                else:
                    print(f'Archivo {file_path} no encontrado.')
                print("Cotización eliminada.")
                # Guardar cambios en JSON
                self.guardar_cotizaciones()
        else:
            QMessageBox.warning(self, "Eliminar Cotización", "Seleccione una cotización para eliminar.")

    
    def eliminar_cotizacion(self, tree_view, model):
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            cotizacion = cotizaciones[row]
            reply = QMessageBox.question(self, 'Eliminar Cotización',
                                         '¿Está seguro de que desea eliminar esta cotización?',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                # Eliminar de la lista de cotizaciones
                del cotizaciones[row]
                # Eliminar del modelo
                model.removeRow(row)
                # Eliminar el archivo Excel asociado
                file_path = os.path.join('cotizaciones', f"{cotizacion['nombre']}.xlsx")
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f'Archivo {file_path} eliminado.')
                else:
                    print(f'Archivo {file_path} no encontrado.')
                print("Cotización eliminada.")
                # Guardar cambios en JSON
                with open('cotizaciones_data.json', 'w') as f:
                    json.dump(cotizaciones, f, indent=4)
        else:
            QMessageBox.warning(self, "Eliminar Cotización", "Seleccione una cotización para eliminar.")
    
    



    




    def mostrar_menu_datos_cotizacion(self, position, tree_view):
        indexes = tree_view.selectedIndexes()
        if indexes:
            menu = QMenu()
            copiar_action = QAction("Copiar", self)
            copiar_action.triggered.connect(lambda: self.copiar_datos(tree_view))
            menu.addAction(copiar_action)
    
            detalles_action = QAction("Detalles", self)
            detalles_action.triggered.connect(lambda: self.mostrar_detalles_cotizacion(tree_view))
            menu.addAction(detalles_action)
    
            enviar_info_action = QAction("Enviar información de la cotización", self)
            enviar_info_action.triggered.connect(lambda: self.enviar_informacion_cotizacion(tree_view))
            menu.addAction(enviar_info_action)
    
            eliminar_action = QAction("Eliminar Cotización", self)
            eliminar_action.triggered.connect(lambda: self.eliminar_cotizacion_cotizaciones_dialog(tree_view))
            menu.addAction(eliminar_action)
    
            menu.exec_(tree_view.viewport().mapToGlobal(position))
        else:
            print("Clic fuera de una fila válida.")



    def modificar_nombre_cotizacion(self, tree_view, model):
        indexes = tree_view.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            cotizacion = cotizaciones[row]
            antiguo_nombre = cotizacion['nombre']
            nuevo_nombre, ok = QInputDialog.getText(self, "Modificar Nombre de Cotización", "Ingrese el nuevo nombre de la cotización:", text=antiguo_nombre)
            if ok and nuevo_nombre:
                if nuevo_nombre == antiguo_nombre:
                    # El nombre no cambió
                    return
                # Verificar si ya existe una cotización con el nuevo nombre
                nombres_existentes = [c['nombre'] for c in cotizaciones]
                if nuevo_nombre in nombres_existentes:
                    QMessageBox.warning(self, "Error", f"Ya existe una cotización con el nombre '{nuevo_nombre}'. Por favor, elija otro nombre.")
                    return
                # Renombrar el archivo Excel
                antiguo_file_path = os.path.join('cotizaciones', f"{antiguo_nombre}.xlsx")
                nuevo_file_path = os.path.join('cotizaciones', f"{nuevo_nombre}.xlsx")
                try:
                    if os.path.exists(antiguo_file_path):
                        shutil.move(antiguo_file_path, nuevo_file_path)
                        print(f"Archivo renombrado de {antiguo_file_path} a {nuevo_file_path}")
                    else:
                        print(f"Archivo {antiguo_file_path} no encontrado. Se creará {nuevo_file_path} al guardar la cotización.")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo renombrar el archivo: {e}")
                    return
                # Actualizar el nombre en la cotización
                cotizacion['nombre'] = nuevo_nombre
                cotizacion['fecha'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                # Actualizar el modelo
                item_nombre = model.item(row, 1)
                item_nombre.setText(nuevo_nombre)
                item_fecha = model.item(row, 3)
                item_fecha.setText(cotizacion['fecha'])
                # Guardar cambios en JSON
                with open('cotizaciones_data.json', 'w') as f:
                    json.dump(cotizaciones, f, indent=4)
            else:
                QMessageBox.warning(self, "Advertencia", "Debe ingresar un nombre válido para la cotización.")
        else:
            QMessageBox.warning(self, "Modificar Nombre", "Seleccione una cotización para modificar su nombre.")



    

    

    def modificar_precio_venta(self, tree_productos, cotizacion_id):
        indexes = tree_productos.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            item = tree_productos.model().item(row, 4)
            nuevo_precio, ok = QInputDialog.getDouble(self, "Modificar Precio de Venta", "Ingrese el nuevo precio de venta:")
            if ok:
                item.setData(nuevo_precio, Qt.DisplayRole)
                item.setData(nuevo_precio, Qt.EditRole)
                cotizacion = cotizaciones[cotizacion_id - 1]
                cotizacion['productos'][row][4] = nuevo_precio
                cotizacion['fecha'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.guardar_cotizacion_excel(cotizacion_id)
                print("Precio de venta modificado")

    def eliminar_producto_cotizacion(self, tree_productos, cotizacion, cotizacion_id):
        indexes = tree_productos.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            tree_productos.model().removeRow(row)
            del cotizacion['productos'][row]
            cotizacion['fecha'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.guardar_cotizacion_excel(cotizacion)
            print("Producto eliminado de la cotización")
        else:
            QMessageBox.warning(self, "Eliminar Producto", "Seleccione un producto para eliminar.")


    def guardar_cotizacion_excel(self, cotizacion):
        file_name = f"{cotizacion['nombre']}.xlsx"
        file_path = os.path.join('cotizaciones', file_name)
        if not os.path.exists('cotizaciones'):
            os.makedirs('cotizaciones')
    
        df = pd.DataFrame(cotizacion['productos'], columns=[
            "No.Producto", "Descripción", "Cantidad", "Precio Costo", "Precio Venta",
            "% Desc.", "Impuesto Cargo", "Impuesto Retención", "Valor Total"
        ])
    
        # Asegurarse de que 'No.Producto' sea int
        df['No.Producto'] = df['No.Producto'].astype(int)
    
        df.to_excel(file_path, index=False)
        print(f'Cotización guardada como Excel en: {file_path}')
    
        # Actualizar la información en 'cotizaciones_data.json'
        with open('cotizaciones_data.json', 'w') as f:
            json.dump(cotizaciones, f, indent=4)










    def guardar_cotizacion_pdf(self, cotizacion_id):
        # Implementación pendiente para generar PDF
        print(f'Cotización {cotizacion_id} guardada como PDF (Funcionalidad pendiente).')

    def cargar_cotizaciones_guardadas(self):
        global cotizaciones
        cotizaciones = []
        cotizaciones_folder = 'cotizaciones'
        if not os.path.exists(cotizaciones_folder):
            os.makedirs(cotizaciones_folder)
            print(f"Carpeta '{cotizaciones_folder}' creada.")
    
        cotizaciones_data_file = 'cotizaciones_data.json'
        if os.path.exists(cotizaciones_data_file):
            with open(cotizaciones_data_file, 'r') as f:
                cotizaciones = json.load(f)
                print(f"{len(cotizaciones)} cotizaciones cargadas desde '{cotizaciones_data_file}'.")
        else:
            print(f"No se encontró '{cotizaciones_data_file}'. Se creará uno nuevo al guardar cotizaciones.")
    
        # Asegurar que cada cotización tiene un ID único y los campos adicionales
        max_id = 0
        for cotizacion in cotizaciones:
            if 'id' not in cotizacion:
                max_id += 1
                cotizacion['id'] = max_id
            else:
                cotizacion['id'] = int(cotizacion['id'])
                if cotizacion['id'] > max_id:
                    max_id = cotizacion['id']
    
            # Añadir campos adicionales si no existen
            cotizacion.setdefault('cliente', '')
            cotizacion.setdefault('contacto', '')
            cotizacion.setdefault('responsable', '')
            cotizacion.setdefault('centro_costo', '')
            cotizacion.setdefault('tipo', 'C-1 Cotización')
            cotizacion.setdefault('fecha_elaboracion', datetime.now().strftime("%d/%m/%Y"))
            cotizacion.setdefault('moneda', 'COP - Peso colombiano')
            cotizacion.setdefault('encabezado', '')
    
        # Cargar los archivos Excel y actualizar las cotizaciones
        archivos_cotizaciones = [f for f in os.listdir(cotizaciones_folder) if f.endswith('.xlsx')]
        nombres_cotizaciones_excel = [os.path.splitext(f)[0] for f in archivos_cotizaciones]
    
        cotizaciones_dict = {c['id']: c for c in cotizaciones}
    
        for nombre_excel in nombres_cotizaciones_excel:
            file_path = os.path.join(cotizaciones_folder, f"{nombre_excel}.xlsx")
            df = pd.read_excel(file_path)
            productos = df.values.tolist()
            for producto in productos:
                while len(producto) < 9:
                    producto.append(0)
            # Buscar cotización por nombre
            cotizacion = next((c for c in cotizaciones if c['nombre'] == nombre_excel), None)
            if cotizacion:
                cotizacion['productos'] = productos
                print(f"Cotización '{nombre_excel}' actualizada con productos desde {file_path}")
            else:
                # Si no existe, crear una nueva cotización con ID único
                max_id += 1
                nueva_cotizacion = {
                    'id': max_id,
                    'nombre': nombre_excel,
                    'estado': 'Pendiente',
                    'fecha': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'productos': productos,
                    'cliente': '',
                    'contacto': '',
                    'responsable': '',
                    'centro_costo': '',
                    'tipo': 'C-1 Cotización',
                    'fecha_elaboracion': datetime.now().strftime("%d/%m/%Y"),
                    'moneda': 'COP - Peso colombiano',
                    'encabezado': ''
                }
                cotizaciones.append(nueva_cotizacion)
                cotizaciones_dict[max_id] = nueva_cotizacion
                print(f"Cotización '{nombre_excel}' cargada desde {file_path}")
    
        # Actualizar 'cotizaciones_data.json' con la información más reciente
        with open(cotizaciones_data_file, 'w') as f:
            json.dump(cotizaciones, f, indent=4)
            print(f"Archivo '{cotizaciones_data_file}' actualizado.")



class CustomTreeView(QTreeView):
    def __init__(self, parent=None):
        super(CustomTreeView, self).__init__(parent)
    
    def edit(self, index, trigger, event):
        # Verificar si el índice es válido
        if not index.isValid():
            return False
        
        # Obtener el modelo y el ítem correspondiente
        model = self.model()
        item = model.itemFromIndex(index)
        
        # Solo permitir editar si el ítem es editable
        if item.isEditable():
            return super(CustomTreeView, self).edit(index, trigger, event)
        else:
            return False



class ComboBoxDelegate(QStyledItemDelegate):
    def __init__(self, items, parent=None):
        super(ComboBoxDelegate, self).__init__(parent)
        self.items = items

    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        combo.addItems(self.items)
        return combo

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.DisplayRole)
        idx = self.items.index(value) if value in self.items else 0
        editor.setCurrentIndex(idx)

    def setModelData(self, editor, model, index):
        value = editor.currentText()
        model.setData(index, value, Qt.EditRole)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())






