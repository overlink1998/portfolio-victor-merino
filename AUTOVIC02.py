from pyshortcuts import make_shortcut
import tkinter as tk
import webbrowser
from tkinter import ttk, scrolledtext, filedialog, messagebox
from tkinter import filedialog, scrolledtext, messagebox, simpledialog
import threading
import keyboard
import time
import tkinter as tk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
import os
import pandas as pd
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import Select
import validators
from html import unescape
from googleapiclient.discovery import build
import base64
import re
from bs4 import BeautifulSoup
import os.path
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import glob
import json
from pandas import Timestamp
from datetime import datetime, timedelta
from PIL import Image, ImageTk
from tkinter import PhotoImage
registros_procesados = []
estado_registro = "no registrado"
# Flags para controlar la ejecución
pause = threading.Event()
stop = threading.Event()
cont = 0
cont2 = 0
numero_contrato = 0
numero = 0
import os
import sys
import inspect
import winshell
ultimo_informe = None

base_dir = "C:\\Users\\RIFIFI\\Documents"


token_path = os.path.join(base_dir, 'token.pickle')
credentials_path = os.path.join(base_dir, 'credentials.json')
url_file_path = os.path.join(base_dir, 'url_guardado.txt')
excel_path = os.path.join(base_dir, 'ultimo_excel.json')

 # Asegúrate de que el ícono exista en esta ruta

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'] 
def convert_float_columns_to_int(df):
    for col in df.select_dtypes(include=['float']).columns:
        # Convertir a entero solo si todos los valores flotantes tienen parte decimal cero
        if all(df[col].dropna().apply(lambda x: x.is_integer())):
            df[col] = df[col].dropna().astype(int)
            # Volver a introducir los valores NA
            df[col] = df[col].astype(pd.Int64Dtype())
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].str.strip()
    return df


# Inicia el WebDriver de Chrome
driver = None

def obtener_url_guardado():
    try:
        with open("url_guardado.txt", "r") as archivo:
            return archivo.read().strip()
    except FileNotFoundError:
        return None

def guardar_url_nuevo(url):
    with open("url_guardado.txt", "w") as archivo:
        archivo.write(url)
def service_gmail():
    creds = None
    # Intenta cargar las credenciales existentes desde el archivo token.pickle
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # Verifica si las credenciales están vencidas o no son válidas
    if not creds or not creds.valid:
        try:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                raise Exception("Necesario realizar nuevo login.")
        except Exception as e:
            # Elimina el archivo token.pickle si el refresco del token falla
            os.remove('token.pickle')
            # Inicia el proceso de autorización nuevamente
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            # Guarda las nuevas credenciales para la próxima ejecución
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    return service
def get_latest_email_link(service):
    results = service.users().messages().list(userId='me', maxResults=1).execute()
    messages = results.get('messages', [])
    
    if not messages:
        print('No messages found.')
        return None
    
    message = messages[0]
    msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
    
    email_content = ""
    # Decodificar el cuerpo del mensaje desde base64 y luego decodificar las entidades HTML
    if 'parts' in msg['payload']:
        for part in msg['payload']['parts']:
            body = part.get('body', {})
            data = body.get('data', '')
            text = base64.urlsafe_b64decode(data.encode('ASCII')).decode('utf-8')
            email_content += unescape(text)  # Decodifica las entidades HTML
    else:
        body = msg['payload'].get('body', {})
        data = body.get('data', '')
        text = base64.urlsafe_b64decode(data.encode('ASCII')).decode('utf-8')
        email_content = unescape(text)  # Decodifica las entidades HTML

    # Convertir el contenido del correo en una lista de líneas
    email_lines = email_content.split('\n')
    link = None
    # Ajusta la expresión regular para que coincida con ambos formatos de URL
    for line in email_lines:
        match = re.search(r'(https://catalogo-vpfe(?:-hab)?\.dian\.gov\.co/User/AuthToken\?pk=\d+\|\d+&rk=\d+&token=[\w-]+)', line)
        if match:
            link = match.group(0)
            break

    if link:
        return link
    else:
        print("No se encontró el enlace en el correo.")
        return None


def obtener_y_navegar_nuevo_link(driver, max_retries=2):
    """
    Esta función intenta obtener un nuevo enlace desde el correo y navegar a él.
    """
    gmail_service_instance = service_gmail()
    for attempt in range(max_retries):
        new_link = get_latest_email_link(gmail_service_instance)
        if new_link and validators.url(new_link):
            print(f"Navegando al nuevo enlace obtenido del correo: {new_link}")
            guardar_url_nuevo(new_link)
            driver.get(new_link)
            return True
        else:
            print("No se pudo obtener un nuevo enlace del correo o el enlace no es válido.")
    return False

def enter_dian_site_with_selenium(driver,user_code, company_code, max_retries=6):
    attempt = 0
    if stop.is_set():
        print("Detención solicitada, saliendo del ciclo de ejecución.")
        return
    while pause.is_set() and not stop.is_set():
        print("Ejecución pausada. Esperando para reanudar...") 
    while attempt < max_retries:
        try:
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
                print("Ejecución pausada. Esperando para reanudar...")
                time.sleep(1)  
            
            # Navega a la página de inicio de sesión o la página esperada aquí
            # Ejemplo: driver.get("URL de la página de DIAN")
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Representante legal')]"))).click()
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "UserCode"))).send_keys(user_code)
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "CompanyCode"))).send_keys(company_code)
            time.sleep(6)  
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Entrar')]"))).click()

            # Espera explícita para una condición de éxito o para detectar el captcha inválido
            try:
                WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "dian-alert-danger"))
                )
                print("Captcha inválido detectado. Reiniciando el proceso de ingreso.")
                # Si se detecta el captcha, refresca la página y continua el bucle de reintentos
                driver.refresh()
                continue
            except TimeoutException:
                # Si no se encuentra el captcha, procede a verificar el éxito del ingreso
                # Agrega aquí una espera explícita para una condición que confirme que has ingresado exitosamente
                print("Ingreso exitoso al sitio de la DIAN.")
                return True
        except Exception as e:
            print(f"Error durante el intento de ingreso: {e}")
        finally:
            # Incrementa el intento y, si no es el último, refresca la página para intentar de nuevo
            attempt += 1
            if attempt < max_retries:
                print(f"Reintentando... Intento {attempt}")
                time.sleep(2)
                
                
                
    print("No se pudo completar el proceso después de varios intentos.")
    return

def iniciar_reiniciar_navegador():
    global driver
    if driver:
        driver.quit() 
    try:
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...") 
            time.sleep(1)  # Espera activa. Ajusta según sea necesario
        driver = webdriver.Chrome()  # Asegúrate de que el PATH de tu chromedriver esté configurado correctamente
        driver.maximize_window()
        url_guardado = obtener_url_guardado()
        if url_guardado and validators.url(url_guardado):  # Se verifica que el URL guardado sea válido
            print(f"Usando URL guardado: {url_guardado}")
            driver.get(url_guardado)
        else:
            driver.get("https://catalogo-vpfe.dian.gov.co/User/CompanyLogin")
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  
    
        try:
            # Buscar por el ID del botón que indica la página inesperada
            driver.find_element(By.ID, "legalRepresentative")
            print("Se detectó la página inesperada.")
            while True:
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mainnav-toggle")))
                    print("La página actual es la esperada. Continuando con el proceso...")
                    break
                except TimeoutException:
                    enter_dian_site_with_selenium(driver, "21809134", "901249887", max_retries=6)
                    obtener_y_navegar_nuevo_link(driver, max_retries=2)
                    print("No se encuentra en la página esperada después de obtener el nuevo enlace.")
            time.sleep(3)
            cargar_navegar_pagina()
        except NoSuchElementException:
            # Si no se encuentra el botón, asumir que no se está en la página inesperada
            print("La página actual es la esperada.")
            time.sleep(3)
            cargar_navegar_pagina()
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  
    except Exception as e:
        print(f"Error durante la carga o navegación: {e}")
        iniciar_reiniciar_navegador()# Reinicia el navegador si ocurre un error
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  
    pass

def iniciar_reiniciar_navegador1():
    global driver
    if driver:
        driver.quit() 
    try:
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...") 
            time.sleep(1)  # Espera activa. Ajusta según sea necesario
        driver = webdriver.Chrome()  # Asegúrate de que el PATH de tu chromedriver esté configurado correctamente
        driver.maximize_window()
        url_guardado = obtener_url_guardado()
        if url_guardado and validators.url(url_guardado):  # Se verifica que el URL guardado sea válido
            print(f"Usando URL guardado: {url_guardado}")
            driver.get(url_guardado)
        else:
            driver.get("https://catalogo-vpfe.dian.gov.co/User/CompanyLogin")
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  
    
        try:
            # Buscar por el ID del botón que indica la página inesperada
            driver.find_element(By.ID, "legalRepresentative")
            print("Se detectó la página inesperada.")
            while True:
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mainnav-toggle")))
                    print("La página actual es la esperada. Continuando con el proceso...")
                    break
                except TimeoutException:
                    enter_dian_site_with_selenium(driver, "21809134", "901249887", max_retries=6)
                    obtener_y_navegar_nuevo_link(driver, max_retries=2)
                    print("No se encuentra en la página esperada después de obtener el nuevo enlace.")
            time.sleep(3)
            cargar_navegar_pagina1()
        except NoSuchElementException:
            # Si no se encuentra el botón, asumir que no se está en la página inesperada
            print("La página actual es la esperada.")
            time.sleep(3)
            cargar_navegar_pagina1()
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  
    except Exception as e:
        print(f"Error durante la carga o navegación: {e}")
        iniciar_reiniciar_navegador1()# Reinicia el navegador si ocurre un error
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  
    pass


def iniciar_reiniciar_navegador2():
    global driver
    if driver:
        driver.quit() 
    try:
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...") 
            time.sleep(1)  # Espera activa. Ajusta según sea necesario
        driver = webdriver.Chrome()  # Asegúrate de que el PATH de tu chromedriver esté configurado correctamente
        driver.maximize_window()
        url_guardado = obtener_url_guardado()
        if url_guardado and validators.url(url_guardado):  # Se verifica que el URL guardado sea válido
            print(f"Usando URL guardado: {url_guardado}")
            driver.get(url_guardado)
            
        else:
            driver.get("https://catalogo-vpfe.dian.gov.co/User/CompanyLogin")
        
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  
    
        try:
            # Buscar por el ID del botón que indica la página inesperada
            driver.find_element(By.ID, "legalRepresentative")
            print("Se detectó la página inesperada.")
            while True:
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mainnav-toggle")))
                    print("La página actual es la esperada. Continuando con el proceso...")
                    break
                except TimeoutException:
                    enter_dian_site_with_selenium(driver, "21809134", "901249887", max_retries=6)
                    obtener_y_navegar_nuevo_link(driver, max_retries=2)
                    print("No se encuentra en la página esperada después de obtener el nuevo enlace.")
            time.sleep(3)
            cargar_navegar_pagina2()
        except NoSuchElementException:
            # Si no se encuentra el botón, asumir que no se está en la página inesperada
            print("La página actual es la esperada.")
            time.sleep(3)
            cargar_navegar_pagina2()
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  
    except Exception as e:
        print(f"Error durante la carga o navegación: {e}")
        iniciar_reiniciar_navegador2()# Reinicia el navegador si ocurre un error
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  
    pass





def cargar_navegar_pagina():
    print("La página actual es la esperada. Continuando con el proceso...")
    try:
        time.sleep(2)
        try:
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
                print("Ejecución pausada. Esperando para reanudar...")
                time.sleep(1)  
            # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
            # y que está único en la página.
            print("Error al intentar hacer clic")
            link = WebDriverWait(driver, 70).until(
                #EC.element_to_be_clickable((By.LINK_TEXT, "Facturador Gratuito"))
                EC.visibility_of_element_located((By.ID, "Invoice"))
            )
            link.click()
            #EC.visibility_of_element_located((By.ID, "ResponsabilidadesTributarias"))
        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
            print(f"No se pudo hacer clic en 'Facturador Gratuito': {e}")
        try:
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
                print("Ejecución pausada. Esperando para reanudar...")
                time.sleep(1) 
            time.sleep(2)
            # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
            # y que está único en la página.
            link = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Ingreso"))
            )
            link.click()
        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
            print(f"Error al intentar hacer clic en 'Ingreso': {e}")
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
                print("Ejecución pausada. Esperando para reanudar...")
                time.sleep(1) 
    
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        original_window = driver.current_window_handle
        new_window = None
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
                print("Ejecución pausada. Esperando para reanudar...")
                time.sleep(1) 
    
        # Identificas la nueva ventana
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                new_window = window_handle
                break
    
        # Cierras la ventana original y cambias a la nueva ventana
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1) 
        if new_window:
            driver.switch_to.window(original_window)
            driver.close()
            driver.switch_to.window(new_window)
            # Realizas las operaciones necesarias en la nueva ventana
            # ...
        else:
            print("No se encontró una nueva ventana.")
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
               print("Ejecución pausada. Esperando para reanudar...")
               time.sleep(1) 

        # Espera a que la página se cargue y el botón 'Factura Electrónica' sea clickeable y haz clic
        try:
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
               print("Ejecución pausada. Esperando para reanudar...")
               time.sleep(1) 
            factura_electronica_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Factura Electrónica')]"))  # Asegúrate de que el XPATH coincida con el elemento real
            )
            factura_electronica_button.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Factura Electrónica': {e}")
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
               print("Ejecución pausada. Esperando para reanudar...")
               time.sleep(1) 
    
        try:
            # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
            # y que está único en la página.
            link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Configuración"))
            )
            link.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Configuración': {e}")
        try:
        # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
        # y que está único en la página.
            link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Adquiriente/comprador"))
            )
            link.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Adquiriente/comprador': {e}")
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
               print("Ejecución pausada. Esperando para reanudar...")
               time.sleep(1) 
    
        cont = 0
        print("Página cargada y navegada con éxito.")
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1) 
        time.sleep(2)
        
    except Exception as e:
        print(f"Error durante la carga o navegación: {e}")
        iniciar_reiniciar_navegador()
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1) 
    pass



def cargar_navegar_pagina1():
    global driver
    print("La página actual es la esperada. Continuando con el proceso...")
    if stop.is_set():
        print("Detención solicitada, saliendo del ciclo de ejecución.")
        return
    while pause.is_set() and not stop.is_set():
        print("Ejecución pausada. Esperando para reanudar...") 
    try:
        time.sleep(2)
        try:
            
            # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
            # y que está único en la página.
            print("Error al intentar hacer clic")
            link = WebDriverWait(driver, 70).until(
                #EC.element_to_be_clickable((By.LINK_TEXT, "Facturador Gratuito"))
                EC.visibility_of_element_located((By.ID, "Invoice"))
            )
            link.click()
            #EC.visibility_of_element_located((By.ID, "ResponsabilidadesTributarias"))
        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
            print(f"No se pudo hacer clic en 'Facturador Gratuito': {e}")
        try:
            time.sleep(2)
            # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
            # y que está único en la página.
            link = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Ingreso"))
            )
            link.click()
        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
            print(f"Error al intentar hacer clic en 'Ingreso': {e}")
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        original_window = driver.current_window_handle
        new_window = None
    
        # Identificas la nueva ventana
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                new_window = window_handle
                break
    
        # Cierras la ventana original y cambias a la nueva ventana
        if new_window:
            driver.switch_to.window(original_window)
            driver.close()
            driver.switch_to.window(new_window)
            # Realizas las operaciones necesarias en la nueva ventana
            # ...
        else:
            print("No se encontró una nueva ventana.")

        # Espera a que la página se cargue y el botón 'Factura Electrónica' sea clickeable y haz clic
        try:
            factura_electronica_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Factura Electrónica')]"))  # Asegúrate de que el XPATH coincida con el elemento real
            )
            factura_electronica_button.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Factura Electrónica': {e}")
    
        try:
            # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
            # y que está único en la página.
            link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Configuración"))
            )
            link.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Configuración': {e}")
        try:
        # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
        # y que está único en la página.
            link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Producto / Servicio"))
            )
            link.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Producto / Servicio': {e}")
    
        try:
            factura_electronica_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Crear nuevo ')]"))  # Asegúrate de que el XPATH coincida con el elemento real
            )
            factura_electronica_button.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Crear nuevo ': {e}")
        cont = 0
        print("Página cargada y navegada con éxito.")
        time.sleep(2)
        
    except Exception as e:
        print(f"Error durante la carga o navegación: {e}")
        iniciar_reiniciar_navegador1()# Reinicia el navegador si ocurre un error
        if stop.is_set():
            print("Detención solicitada. Finalizando ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  # Espera activa. Ajusta según sea necesario
    pass


def cargar_navegar_pagina2():
    global driver
    if stop.is_set():
        print("Detención solicitada, saliendo del ciclo de ejecución.")
        return
    while pause.is_set() and not stop.is_set():
        print("Ejecución pausada. Esperando para reanudar...") 
    try:
        time.sleep(2)
        try:
            
            # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
            # y que está único en la página.
            link = WebDriverWait(driver, 70).until(
                #EC.element_to_be_clickable((By.LINK_TEXT, "Facturador Gratuito"))
                EC.visibility_of_element_located((By.ID, "Invoice"))
            )
            link.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Facturador_Gratuito': {e}")
        try:
            time.sleep(2)
            # Suponiendo que 'Ingreso' es un texto en un elemento clickeable como un botón o un enlace
            # y que está único en la página.
            link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Ingreso"))
            )
            link.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Ingreso': {e}")
    
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        original_window = driver.current_window_handle
        new_window = None
    
        # Identificas la nueva ventana
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                new_window = window_handle
                break
    
        # Cierras la ventana original y cambias a la nueva ventana
        if new_window:
            driver.switch_to.window(original_window)
            driver.close()
            driver.switch_to.window(new_window)
            # Realizas las operaciones necesarias en la nueva ventana
            # ...
        else:
            print("No se encontró una nueva ventana.")

        # Espera a que la página se cargue y el botón 'Factura Electrónica' sea clickeable y haz clic
        try:
            factura_electronica_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Factura Electrónica')]"))  # Asegúrate de que el XPATH coincida con el elemento real
            )
            factura_electronica_button.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Factura Electrónica': {e}")
    
        print("Página cargada y navegada con éxito.")
        time.sleep(2)
        
    except Exception as e:
        print(f"Error durante la carga o navegación: {e}")
        iniciar_reiniciar_navegador2()# Reinicia el navegador si ocurre un error
        if stop.is_set():
            print("Detención solicitada. Finalizando ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  # Espera activa. Ajusta según sea necesario
    pass




def rellenar_campos(nombre_cliente, cedula_cliente, correo_cliente, direccion_cliente, telefono_cliente):
    if stop.is_set():
        print("Detención solicitada, saliendo del ciclo de ejecución.")
        return
    # Asumiendo que ya tienes los localizadores y el objeto driver configurado
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "NmbRecep"))).clear()
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "NmbRecep"))).send_keys(nombre_cliente)
    
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "DocRecep_NroDocRecep"))).clear()
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "DocRecep_NroDocRecep"))).send_keys(cedula_cliente)
    
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "ContactoReceptor_0__eMail"))).clear()
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "ContactoReceptor_0__eMail"))).send_keys(correo_cliente)
    
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "DomFiscalRcp_Calle"))).clear()
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "DomFiscalRcp_Calle"))).send_keys(direccion_cliente)
    
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "ContactoReceptor_0__Telefono"))).clear()
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "ContactoReceptor_0__Telefono"))).send_keys(telefono_cliente)

def verificar_datos_antes_de_guardar(nombre_cliente, cedula_cliente, correo_cliente, direccion_cliente, telefono_cliente):
    intentos = 3
    while intentos > 0:
        try:
            # Limpia y rellena los campos antes de verificar
            rellenar_campos(nombre_cliente, cedula_cliente, correo_cliente, direccion_cliente, telefono_cliente)
            
            nombre_ingresado = driver.find_element(By.ID, "NmbRecep").get_attribute('value').strip()
            cedula_ingresada = driver.find_element(By.ID, "DocRecep_NroDocRecep").get_attribute('value').strip()
            correo_ingresado = driver.find_element(By.ID, "ContactoReceptor_0__eMail").get_attribute('value').strip()
            direccion_ingresada = driver.find_element(By.ID, "DomFiscalRcp_Calle").get_attribute('value').strip()
            telefono_ingresado = driver.find_element(By.ID, "ContactoReceptor_0__Telefono").get_attribute('value').strip()
            
            # Convertir telefono_cliente a string y normalizar formato
            telefono_cliente_str = str(telefono_cliente).replace(" ", "").replace("-", "")
            telefono_ingresado_norm = telefono_ingresado.replace(" ", "").replace("-", "")

            if (nombre_ingresado == nombre_cliente and
                cedula_ingresada == cedula_cliente and
                correo_ingresado == correo_cliente and
                direccion_ingresada == direccion_cliente and
                telefono_ingresado_norm == telefono_cliente_str):
                print("Los datos ingresados son correctos.")
                return True
            
                
            else:
                print("Los datos ingresados no coinciden con los esperados.")
                # Los campos ya están limpios y rellenados en este punto
        except Exception as e:
            print(f"Error durante la verificación o el rellenado de los campos: {e}")
        intentos -= 1
        time.sleep(2)  # Da tiempo para que los campos se limpien y rellenen antes de verificar nuevamente
    
    print("No se pudieron verificar los datos después de varios intentos.")
    return False

# Nota: Asegúrate de ajustar las condiciones_para_verificar_datos con tu lógica de verificación real.

def ensure_accordion_content_visibility(driver, accordion_content_id, accordion_header_xpath):
    try:
        # Espera hasta que el contenido del acordeón sea visible.
        content_visible = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, accordion_content_id))
        )

        if content_visible:
            print(f"El contenido del acordeón con ID '{accordion_content_id}' ya está visible. No se requiere acción.")
        else:
            # Si el contenido no es visible, se asume que el acordeón está cerrado.
            # Encontrar el encabezado del acordeón y hacer clic para abrir.
            accordion_header = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, accordion_header_xpath))
            )
            accordion_header.click()
            print(f"El acordeón con ID '{accordion_content_id}' estaba cerrado y se abrió.")

    except TimeoutException:
        # Si el tiempo de espera expira, entonces intentar abrir el acordeón.
        accordion_header = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, accordion_header_xpath))
        )
        accordion_header.click()
        print(f"Timeout al buscar contenido del acordeón con ID '{accordion_content_id}'. Se intentó abrir el acordeón.")

    except Exception as e:
        print(f"Error al verificar la visibilidad del contenido del acordeón con ID '{accordion_content_id}': {e}")


def select_dropdown_option(driver, dropdown_selector, option_text):
        # Esperar a que la opción deseada sea clickeable y hacer clic en ella
    try:
        dropdown = WebDriverWait(driver, 10).until(
              EC.element_to_be_clickable((By.CSS_SELECTOR, dropdown_selector))
        )
        dropdown.click()

    # Esperar a que la opción deseada sea clickeable y hacer clic en ella
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//span[text()='{option_text}']"))
        )
    
        # Desplazarse hasta la opción y hacer clic en ella
        actions = ActionChains(driver)
        actions.move_to_element(option).click().perform()

    except Exception as e:
        print(f"No se pudo seleccionar la opción '{option_text}'. Error: {e}")



def fill_form(df, row):
    global driver, estado_registro, stop
    # Tu código para llenar el formulario aquí
    time.sleep(2)
    estado_registro = "No se ha registrado"
    if stop.is_set():
        print("Detención solicitada, saliendo del ciclo de ejecución.")
        return
    while pause.is_set() and not stop.is_set():
        print("Ejecución pausada. Esperando para reanudar...")
        time.sleep(1) 

    try:
        apart = str("MAGDALENA")
        ciudad = str("SANTA MARTA")
        adquiriente = str("2 - Persona Natural y asimiladas")
        time.sleep(2)
        #elemento_del_modal = WebDriverWait(driver, 60).until(
        #    EC.visibility_of_element_located((By.XPATH, "//button[contains(.,'Crear nuevo ')]"))
        #)
        #elemento_del_modal.click()
        button = driver.find_element(By.CSS_SELECTOR, "button[title='Nuevo Cliente']")
        driver.execute_script("arguments[0].click();", button)
        time.sleep(2)
        WebDriverWait(driver, 80).until(
          EC.visibility_of_element_located((By.ID, "NmbRecep"))
        ).send_keys(row['Nombre Cliente'])
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "DocRecep_NroDocRecep"))
        ).send_keys(row['Cedula'])
        WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "ContactoReceptor_0__eMail"))
        ).send_keys(row['Correo'])
        datos_adicionales_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//h3[normalize-space()='Datos adicionales si el Adquiriente lo requiere']"))
        )
        datos_adicionales_button.click()
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1) 
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "TipoContribuyenteR"))
        ).send_keys(adquiriente)
        campo_dirección =WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "DomFiscalRcp_Calle"))
        )
        campo_dirección.clear()

        campo_dirección =WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "ContactoReceptor_0__Telefono"))
        )
        campo_dirección.clear()
        modal = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "modal-body"))
        )
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", modal)
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "DomFiscalRcp_Calle"))
        ).send_keys(row['dirección'])
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "ContactoReceptor_0__Telefono"))
        ).send_keys(row['Telefono'])
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1) 
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "ClientDepartment"))
        ).send_keys(apart)
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "ClientMunicipality"))
        ).send_keys(ciudad)
        guardar_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "btn-save-client"))
        )

        if verificar_datos_antes_de_guardar(row['Nombre Cliente'], str(row['Cedula']), row['Correo'], row['dirección'], row['Telefono']):
            # Si los datos son correctos, haz clic en guardar
            guardar_button.click()
            if stop.is_set():
               print("Detención solicitada, saliendo del ciclo de ejecución.")
               return
            while pause.is_set() and not stop.is_set():
               print("Ejecución pausada. Esperando para reanudar...")
               time.sleep(1) 
        else:
            # Si los datos no son correctos, decides qué hacer (puedes lanzar una excepción, puedes intentar corregir los datos, etc.)
            # Por ejemplo, podrías querer detener la ejecución o simplemente no hacer clic en guardar.
            print("No se guardará el registro debido a datos incorrectos.")
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
                print("Ejecución pausada. Esperando para reanudar...")
                time.sleep(1) 
        #esta párte del codigo es cuando el cliente ya esta en el sistema y se necesita pasar al siguiente cliente 
        from selenium.webdriver.common.keys import Keys
        try:
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
                print("Ejecución pausada. Esperando para reanudar...")
                time.sleep(1) 
            # Asegúrate de que la ventana emergente esté enfocada
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "errorModal-title")))
    
            # Presiona la tecla ESC
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            driver.refresh()
            print("ya esta registrado.")
            estado_registro = "Ya está registrado"
            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
                print("Ejecución pausada. Esperando para reanudar...")
                time.sleep(1) 
            
        except Exception as e:
            print(f"se ha registrado el usuario: {e}")
            estado_registro = "Se ha registrado el usuario"    
        time.sleep(2)
    
    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
        print(f"Error al llenar el formulario: {e}")
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        

        iniciar_reiniciar_navegador()
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1) 
   
        
        fill_form(df, row)
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1) 
        
    if stop.is_set():
        print("Detención solicitada, saliendo del ciclo de ejecución.")
        return
    while pause.is_set() and not stop.is_set():
        print("Ejecución pausada. Esperando para reanudar...")
        time.sleep(1)  


    pass



def fill_form1(df, row):
    global driver, estado_registro, stop
    if stop.is_set():
        print("Detención solicitada, saliendo del ciclo de ejecución.")
        return
    while pause.is_set() and not stop.is_set():
        print("Ejecución pausada. Esperando para reanudar...") 
    try:
        estado_Registro = "No registrado"
        time.sleep(2)
        excel_data = df
        Nocontr = str(row[contrato])
        val = str((row[valor]) *num)
        adquiriente = str("2 - Persona Natural y asimiladas")
        tipo2 = " "
        resultado = tipo + tipo2 + Nocontr
        time.sleep(2)
        

        elemento_del_modal = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, "//button[contains(.,'Crear nuevo ')]"))
        )
        time.sleep(2)
        WebDriverWait(driver, 50).until(
           EC.visibility_of_element_located((By.ID, "tipoCodigoProducto"))
        ).send_keys(['999 - Estándar de adopción del contribuyente'])
        
        WebDriverWait(driver, 20).until(
           EC.visibility_of_element_located((By.ID, "CdgItem_0__VlrCodigo"))
        ).send_keys(tipo)
    
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "DscItem"))
        ).send_keys(resultado)
        
        campo_dirección =WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "PrcBrutoItem"))
        )
        campo_dirección.clear()
        WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "PrcBrutoItem"))
        ).send_keys(val)
        modal = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "modal-body"))
        )
        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1) 
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", modal)
        guardar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "btn-save-product"))
        )
        guardar_button.click()
        
        try:
            
            factura_electronica_button = WebDriverWait(driver, 30).until( 
               EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Crear nuevo ')]"))  # Asegúrate de que el XPATH coincida con el elemento real
            )
            factura_electronica_button.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Crear nuevo ': {e}")
        
        print("Formulario llenado con éxito.")
        estado_registro = "Se ha registrado el usuario"
        
        time.sleep(2)
     
    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
        print(f"Error al llenar el formulario: {e}")
        estado_registro = "No se ha podido registrar el usuario"
        
        if stop.is_set():
            print("Detención solicitada. Finalizando ejecución.")
            return  # O break, dependiendo del contexto
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  # Espera activa. Ajusta según sea necesario

        iniciar_reiniciar_navegador1()
            
        if stop.is_set():
            print("Detención solicitada. Finalizando ejecución.")
            return  # O break
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)
        fill_form1(df, row)
        if stop.is_set():
            print("Detención solicitada. Finalizando ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  # Espera activa. Ajusta según sea necesario

    pass


def ajustar_fecha(fecha_contrato):
    # Convertir el string a un objeto datetime
    if isinstance(fecha_contrato, Timestamp):
        fecha = fecha_contrato.to_pydatetime()
    else:
        # Asumimos que es una cadena y usamos strptime
        fecha = datetime.strptime(fecha_contrato, "%d/%m/%Y")
    # Fecha actual
    fecha_actual = datetime.now()
    
    # Fecha límite (9 días antes de la fecha actual)
    fecha_limite = fecha_actual
    
    # Si la fecha del contrato es mayor que la fecha actual, se ajusta a la fecha límite
    if fecha < fecha_actual:
        return fecha_limite.strftime("%d/%m/%Y")
    else:
        # Si la fecha del contrato es mayor que la fecha límite pero menor o igual a la fecha actual,
        # se devuelve como está.
        return fecha_actual.strftime("%d/%m/%Y")


def subir_fecha(driver, fecha_contrato):
    fecha_ajustada = ajustar_fecha(fecha_contrato)

    # Localizar el campo de fecha y limpiarlo
    campo_fecha = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "Documento_Encabezado_IdDoc_FechaEmis"))
    )
    campo_fecha.clear()

    # Enviar la fecha ajustada al campo
    campo_fecha.send_keys(fecha_ajustada)


def select_dropdown_option(driver, dropdown_selector, option_text):
        # Esperar a que la opción deseada sea clickeable y hacer clic en ella

    try:
        ensure_accordion_content_visibility(driver, "accordionReceptorData", "//h3[normalize-space()='3. Datos del adquirente/comprador']")
        dropdown = WebDriverWait(driver, 10).until(
              EC.element_to_be_clickable((By.CSS_SELECTOR, dropdown_selector))
        )
        dropdown.click()

    # Esperar a que la opción deseada sea clickeable y hacer clic en ella
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//span[text()='{option_text}']"))
        )
    
        # Desplazarse hasta la opción y hacer clic en ella
        actions1 = ActionChains(driver)
        actions1.move_to_element(option).click().perform()
        actions1.move_to_element(option).click().perform()
        option.click()

        
        ensure_accordion_content_visibility(driver, "accordionDetalleData", "//h3[normalize-space()='4. Detalles del producto / servicio']")
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                campo_cantidad = WebDriverWait(driver, 30).until(
                       EC.element_to_be_clickable((By.ID, "Documento_Detalle_0__QtyItem"))
                )
                campo_cantidad.click()  # Haz clic en el campo para asegurarte de que está activo
                campo_cantidad.clear()  # Borra el texto seleccionado
                val = 1
                campo_cantidad.send_keys(val)
                break  # Si todo salió bien, sal del ciclo
            except TimeoutException:
                
                print(f"Intento {attempt + 1} de {max_attempts}: no se pudo hacer clic o enviar texto al campo de cantidad. Reintentando...")
                if attempt < max_attempts - 1:
                    time.sleep(2)  # Espera un poco antes del próximo intento
                else:
                    print("No se pudo interactuar con el campo de cantidad después de varios intentos.")
            except Exception as e:
                print(f"Ocurrió un error inesperado: {e}")
                break 
        
        
    

    except TimeoutException as e:
        print(f"No se pudo seleccionar la opción")

def fill_form2(df, row):
    global driver, estado_registro, ensure_accordion_content_visibility, select_dropdown_option
    # Tu código para llenar el formulario aquí
    if stop.is_set():
        print("Detención solicitada, saliendo del ciclo de ejecución.")
        return
    while pause.is_set() and not stop.is_set():
        print("Ejecución pausada. Esperando para reanudar...") 
    try:
        estado_registro = "no registrado" 
        excel_data = df
        #apart = "14/04/2024"
        fecha_contrato = row[fecha]
        Nocontr = str(row[contrato])
        tipo2 = " "
        resultado = tipo + tipo2 + Nocontr
        subir_fecha(driver, fecha_contrato)
        time.sleep(2)
    
        ensure_accordion_content_visibility(driver, "accordionDetalleData", "//h3[normalize-space()='4. Detalles del producto / servicio']")
        try:
            # Hacer clic en el botón de búsqueda
            search_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div.input-group-btn a.btn.btn-info[data-toggle='modal']"))
            )
            search_button.click()

        except TimeoutException:
            # Si hay un TimeoutException, esperar para el próximo intento
            print(f"El campo de descripción no fue encontrado en el intento")
        max_attempts = 2

        for attempt in range(max_attempts):
            try:
                
                campo_descripcion = WebDriverWait(driver, 60).until(
                    EC.visibility_of_element_located((By.ID, "DscItem"))
                )
                campo_descripcion.click()
                campo_descripcion.clear()  # Limpia el campo antes de enviar texto
                campo_descripcion.send_keys(resultado)
                break  # Si todo es exitoso, rompe el ciclo
            except TimeoutException as e:
                print(f"Intento {attempt+1}/{max_attempts} falló. Razón: {str(e)}")
                if attempt < max_attempts - 1:
                    time.sleep(2)  # Espera antes del siguiente intento si no es el último
                else:
                    print("Máximo de intentos alcanzado. No se pudo realizar la acción en el campo.")
        
        time.sleep(10)
        actions = ActionChains(driver)
        actions.send_keys(Keys.TAB).perform()
        enlace_documento = WebDriverWait(driver, 100).until(
            EC.element_to_be_clickable((By.XPATH, f"//td/a[text()='{resultado}']"))
        )
        enlace_documento.click()
        time.sleep(2)
        
        try:
            # Esperar a que el elemento que intercepta el clic desaparezca (ejemplo: pantalla de carga)
            WebDriverWait(driver, 30).until(
                EC.invisibility_of_element_located((By.ID, "dian-loading"))
            )
            campo_cantidad = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.ID, "Documento_Detalle_0__QtyItem"))
            )
            campo_cantidad.click()  # Haz clic en el campo para asegurarte de que está activo
            campo_cantidad.send_keys(Keys.HOME)  # Mueve el cursor al inicio del campo
            campo_cantidad.send_keys(Keys.SHIFT + Keys.END)  # Selecciona todo el texto
            campo_cantidad.send_keys(Keys.BACK_SPACE)
            val = 1
            campo_cantidad.send_keys(val)
        except TimeoutException as e:
            print(f"No se pudo interactuar con el campo de cantidad debido a un error de tiempo de espera: {e}")
        except ElementClickInterceptedException as e:
            print(f"El clic en el campo de cantidad fue interceptado por otro elemento: {e}")
        except Exception as e:
            print(f"Ocurrió un error inesperado al intentar interactuar con el campo de cantidad: {e}")

        if stop.is_set():
            print("Detención solicitada, saliendo del ciclo de ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...") 
        
        ensure_accordion_content_visibility(driver, "collapseIdDocData", "//h3[normalize-space()='1. Datos del documento']")

        
        
        select_element = WebDriverWait(driver, 50).until(
            EC.visibility_of_element_located((By.ID, "Documento_Encabezado_IdDoc_RankKey"))  # Reemplaza con el ID real del elemento select
        )

        # Crea una instancia de Select con el elemento encontrado
        select = Select(select_element)

        # Selecciona la última opción por índice
        select.select_by_index(len(select.options) - 1)
    
        WebDriverWait(driver, 20).until(
           EC.visibility_of_element_located((By.ID, "Documento_Encabezado_IdDoc_MedioPago"))
        ).send_keys(['10 - Efectivo'])
        WebDriverWait(driver, 20).until(
           EC.visibility_of_element_located((By.ID, "Documento_Encabezado_IdDoc_TipoNegociacion"))
        ).send_keys(['1 - Contado'])
    
        #datos_adicionales_button = WebDriverWait(driver, 20).until(
        #    EC.element_to_be_clickable((By.XPATH, "//h3[normalize-space()='2. Datos del emisor/vendedor']"))
        #)
        #datos_adicionales_button.click()
        ensure_accordion_content_visibility(driver, "accordionEmisorData", "//h3[normalize-space()='2. Datos del emisor/vendedor']")
        WebDriverWait(driver, 20).until(
           EC.visibility_of_element_located((By.ID, "ResponsabilidadesTributarias"))
        ).send_keys(['ZZ - No aplica'])
        
        time.sleep(2)
        
        try: 
            menu_desplegable1 = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@data-id="TipoActividadEconomicaEmisor"]'))
            )
            menu_desplegable1.click()
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, "dropdown-menu"))
            )
            
            # Hace scroll hacia la opción deseada usando JavaScript
            driver.execute_script("""
                var button = arguments[0];
                var id = button.getAttribute('data-id');
                var select = document.getElementById(id);
                var options = select.options;
                for (var i = 0; i < options.length; i++) {
                options[i].selected = false;
                }
                $(select).selectpicker('refresh');
                """, menu_desplegable1)    
            menu_desplegable1 = WebDriverWait(driver, 20).until(
               EC.element_to_be_clickable((By.XPATH, '//button[@data-id="TipoActividadEconomicaEmisor"]'))
            )
            menu_desplegable1.click()
        except Exception as e:
            print(f"Error al verificar'{menu_desplegable1}': {e}")
        

        

    

        ensure_accordion_content_visibility(driver, "accordionReceptorData", "//h3[normalize-space()='3. Datos del adquirente/comprador']")
        
        busq_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "btnSearchReceiver"))
        )
        busq_button.click()

        
        campo_busqueda = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "input[type='search']"))
        )
        campo_busqueda.send_keys(row['Cedula'])
        cedula = str(row['Cedula'])
        enlace_documento = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, f"//td/a[text()='{cedula}']"))
        )
        enlace_documento.click()


        
        buttonf = WebDriverWait(driver, 40).until(
           EC.visibility_of_element_located((By.ID, "btnPrevisualizar"))
        )
        buttonf.click()
        
        #try:
        #    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "errorModal-title")))
        #    print("Mensaje de error detectado. Se intentará seleccionar la opción requerida.")
        #    select_dropdown_option(driver, "button[data-id='TipoResponsabilidadReceptor']", "R-99-PN - NO RESPONSABLE")
        #except Exception as e:
        #    print(f"se ha registrado el usuario: {e}")
        for attempt in range(max_attempts):
            try:
                # Intentar cerrar el mensaje de error si está presente
                error_message = WebDriverWait(driver, 30).until(
                    EC.visibility_of_element_located((By.ID, "errorModal-title"))
                )
                if error_message:
                    print(f"Mensaje de error detectado. Intento: {attempt + 1}")
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    select_dropdown_option(driver, "button[data-id='TipoResponsabilidadReceptor']", "R-99-PN - NO RESPONSABLE")
                    
                    time.sleep(1)
                    buttonf.click()
                else:
                    print("No se encontró mensaje de error.")
                    break
            except Exception as e:
                # Si no se encuentra el mensaje de error (Timeout), continuar con el siguiente intento
                print(f"No se detectó el mensaje de error en el intento.")
                break
            except Exception as e:
                # Capturar cualquier otra excepción y salir del ciclo
                print(f"Ocurrió un error en el intento {attempt}: {e}")
                break
               
            
            time.sleep(3)
        

        guardar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.ID, "btn-save"))
        )
        guardar_button.click()
        time.sleep(2)
        
        try:
            
            link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Factura Electrónica"))
            )
            link.click()
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Factura Electrónica': {e}")
        
        
        try:
            link = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Factura de Venta"))
            )
            link.click()
            estado_registro = "la factura ha sido completada"
            print("Se ha registrado el usuario.")
        except Exception as e:
            print(f"Error al intentar hacer clic en 'Factura de Venta': {e}")
            estado_registro = "No se ha podido registrar el usuario" 
        
    
    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
        print(f"Error al llenar el formulario: {e}")
        if stop.is_set():
            print("Detención solicitada. Finalizando ejecución.")
            return  # O break, dependiendo del contexto
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  # Espera activa. Ajusta según sea necesario

        iniciar_reiniciar_navegador2()

        
        if stop.is_set():
            print("Detención solicitada. Finalizando ejecución.")
            return  # O break

        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)

            
        if stop.is_set():
            print("Detención solicitada. Finalizando ejecución.")
            return  # O break
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)
        fill_form2(df, row)
        if stop.is_set():
            print("Detención solicitada. Finalizando ejecución.")
            return
        while pause.is_set() and not stop.is_set():
            print("Ejecución pausada. Esperando para reanudar...")
            time.sleep(1)  # Espera activa. Ajusta según sea necesario
        return "exito"
        
        
        

    if stop.is_set():
        print("Detención solicitada. Finalizando ejecución.")
        return  # O break

    while pause.is_set() and not stop.is_set():
        print("Ejecución pausada. Esperando para reanudar...")
        time.sleep(1)


    pass


def guardar_ruta_ultimo_informe():
    global ruta_del_ultimo_informe_generado
    if ruta_del_ultimo_informe_generado:
        with open('ruta_ultimo_informe.json', 'w') as f:
            json.dump({'ruta': ruta_del_ultimo_informe_generado}, f)
        print(f"Ruta del último informe guardada: {ruta_del_ultimo_informe_generado}")
    else:
        print("No hay ruta del último informe para guardar.")





#proceso normal
ruta_del_ultimo_informe_generado = None    
def guardar_registros_procesados(registros_a_guardar, tipo_automatizacion, tipo_proceso, nombre_excel):
    global ruta_del_ultimo_informe_generado
    df = pd.DataFrame(registros_a_guardar)
    ruta_guardado = os.path.join(os.getcwd(), 'Datos DIAN', nombre_excel)
    os.makedirs(os.path.dirname(ruta_guardado), exist_ok=True)
    df.to_excel(ruta_guardado, index=False)
    print(f"Registros procesados guardados en: {ruta_guardado}")
    ruta_del_ultimo_informe_generado = ruta_guardado
    guardar_ruta_ultimo_informe()


def cargar_ultimo_registro():
    try:
        ruta_archivo = os.path.join(os.getcwd(), 'Datos DIAN', 'ultimo_registro.xlsx')
        if os.path.exists(ruta_archivo):
            df_registro = pd.read_excel(ruta_archivo)
            registro = df_registro.iloc[-1]  # Tomamos el último registro
            actualizar_ultimo_registro(registro['contrato'], registro['numero'], registro['fecha'], registro['tipo_automatizacion'], registro['tipo_proceso'])
    except FileNotFoundError:
        print("No hay registro anterior guardado.")

        


    
def ejecutar_proceso(df, tipo_automatizacion, tipo_proceso):
    global driver, pause, stop, cont, window, is_automation_running, root, contrato, valor, num, tipo, fecha, fecha_contrato
    cargar_ultimo_registro()
    nombre_excel_base = os.path.splitext(os.path.basename(archivo_excel))[0]
    registros_procesados = []
    tipo_proceso_actual = tipo_proceso
    tipo_automatizacion_actual = tipo_automatizacion
    try:
        is_automation_running = True
        if tipo_automatizacion == "Adquirientes": 
            if tipo_proceso == "Contrato":
                iniciar_reiniciar_navegador()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["No Contrato"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                    # Llama a actualizar GUI en el hilo principal
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
            if tipo_proceso == "Retiro":
                iniciar_reiniciar_navegador()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["No Contrato"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    print(f"Registros a guardar: {len(registros_procesados)}")
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
            if tipo_proceso == "Prorroga":
                iniciar_reiniciar_navegador()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["No Contrato"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
            if tipo_proceso == "Venta":
                iniciar_reiniciar_navegador()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["CODIGO"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
        if tipo_automatizacion == "Producto": 
            if tipo_proceso == "Contrato":
                tipo = (str("CON"))
                contrato = "No Contrato"
                valor = "Valor"
                num = 1000
                iniciar_reiniciar_navegador1()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["No Contrato"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form1(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador1()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form1(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                    # Llama a actualizar GUI en el hilo principal
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
            if tipo_proceso == "Retiro":
                tipo = (str("RET"))
                contrato = "No Contrato"
                valor = "Valor"
                num = 1000
                iniciar_reiniciar_navegador1()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["No Contrato"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form1(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador1()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form1(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                    # Llama a actualizar GUI en el hilo principal
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
            if tipo_proceso == "Prorroga":
                tipo = (str("PRO"))
                contrato = "No Contrato"
                valor = "Valor Pagado"
                num = 1
                iniciar_reiniciar_navegador1()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["No Contrato"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form1(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador1()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form1(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                    # Llama a actualizar GUI en el hilo principal
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
    
            if tipo_proceso == "Venta":
                tipo = (str("VEN"))
                contrato = "CODIGO"
                valor = "total"
                num = 1000
                iniciar_reiniciar_navegador1()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["CODIGO"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form1(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador1()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form1(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                    # Llama a actualizar GUI en el hilo principal
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
            

        if tipo_automatizacion == "Factura Venta": 
            if tipo_proceso == "Contrato":
                tipo = (str("CON"))
                contrato = "No Contrato"
                valor = "Valor"
                fecha = "Fecha  Contrato"
                
                iniciar_reiniciar_navegador2()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row[contrato]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form2(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador2()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form2(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                    # Llama a actualizar GUI en el hilo principal
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
            if tipo_proceso == "Retiro":
                tipo = (str("RET"))
                contrato = "No Contrato"
                valor = "Valor"
                fecha = "Fecha Retirado"
                iniciar_reiniciar_navegador2()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["No Contrato"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form2(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador2()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form2(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                    # Llama a actualizar GUI en el hilo principal
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
            if tipo_proceso == "Prorroga":
                tipo = (str("PRO"))
                contrato = "No Contrato"
                valor = "Valor Pagado"
                fecha = "Fecha"
                iniciar_reiniciar_navegador2()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["No Contrato"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form2(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador2()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form2(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                    # Llama a actualizar GUI en el hilo principal
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
    
            if tipo_proceso == "Venta":
                tipo = (str("VEN"))
                contrato = "CODIGO"
                valor = "total"
                fecha = "FECHA VENTA"
                iniciar_reiniciar_navegador2()
                if stop.is_set():
                    print("Detención solicitada, saliendo del ciclo de ejecución.")
                    return
                while pause.is_set() and not stop.is_set():
                    print("Ejecución pausada. Esperando para reanudar...")
                    time.sleep(1)  
                total_rows = len(df)
                for index, row in df.iloc[cont:].iterrows():
                    cont += 1
                    numero_contrato = row["CODIGO"]
                    print(f"Contador: {cont}, No Contrato: {numero_contrato}")
                    try:
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form2(df, row) 
                    except (TimeoutException, WebDriverException, NoSuchElementException) as e:
                        print(f"Ocurrió un error: {e}")
                        iniciar_reiniciar_navegador2()
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        fill_form2(df, row) 
                        if stop.is_set():
                            print("Detención solicitada, saliendo del ciclo de ejecución.")
                            return
                        while pause.is_set() and not stop.is_set():
                            print("Ejecución pausada. Esperando para reanudar...")
                            time.sleep(1)  
                    if cont == total_rows:
                        print("la automatizacion ha terminado.")
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                        # Llama a actualizar GUI en el hilo principal
                        root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
                        guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                        registros_procesados.append({
                        
                        "Número de registro procesado": cont,
                        "No Contrato": numero_contrato,
                        "Estado_Registro": estado_registro  # Usa el valor actualizado
                        })
                        break
                    if stop.is_set():
                        print("Detención solicitada. Finalizando ejecución.")
                        return  # O break

                    while pause.is_set() and not stop.is_set():
                        print("Ejecución pausada. Esperando para reanudar...")
                        time.sleep(1)
                    registros_procesados.append({
                    
                    "Número de registro procesado": cont,
                    "No Contrato": numero_contrato,
                    "Estado_Registro": estado_registro  # Usa el valor actualizado
                    })
                    # Suponiendo que tipo_proceso_actual es una variable global o pasada como parámetro que contiene el tipo de proceso actual
                    #guardar_ultimo_registro(numero_contrato, cont, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo_proceso_actual)
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    #guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_proceso_actual)
                    # Llama a actualizar GUI en el hilo principal
                    root.after(0, actualizar_ultimo_registro, numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)
                    guardar_ultimo_registro(numero_contrato, cont, fecha_actual, tipo_automatizacion_actual, tipo_proceso_actual)

                    print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")


            if stop.is_set():
                print("Detención solicitada, saliendo del ciclo de ejecución.")
                return
            while pause.is_set() and not stop.is_set():
                print("Ejecución pausada. Esperando para reanudar...")
                time.sleep(1)
        

    finally:
        # Asegúrate de que la limpieza se realiza independientemente de cómo se salga de la función
        finalizar_automatizacion()
        stop.clear()
        print(f"el ultimo contrato fue la columna: {cont}, No Contrato: {numero_contrato}, Estado_Registro: {estado_registro}")
        if registros_procesados:
            fecha_actual = datetime.now().strftime("%Y-%m-%d")

            nombre_archivo = f"{nombre_excel_base}_{tipo_automatizacion_actual}_{tipo_proceso_actual}_{fecha_actual}.xlsx"
            guardar_registros_procesados(registros_procesados, tipo_automatizacion, tipo_proceso, nombre_archivo)
        else:
            print("No se procesaron registros.")
        
        
        
    numero_entry['state'] = 'normal'
    if stop.is_set():
        print("Detención solicitada, saliendo del ciclo de ejecución.")
        return
    while pause.is_set() and not stop.is_set():
        print("Ejecución pausada. Esperando para reanudar...")
        time.sleep(1)
    pass
    
# Variable global para almacenar la ruta del archivo Excel
archivo_excel = None
driver = None
is_automation_running = False
ultimo_contrato_var = None
ultimo_numero_var = None
ultima_fecha_var = None
tipo_proceso_actual = None
tipo_automatización_actual = None
tipo_proceso_var = None
tipo_automatizacion_var = None
panel_mensajes = None
ruta_informe_excel_var = None
automatizaciones = {
    "Adquirientes": {
        "Contrato": {"sheet_name": "Hoja1"},
        "Retiro": {"sheet_name": "Hoja2"},
        "Prorroga": {"sheet_name": "Hoja3"},
        "Venta": {"sheet_name": "Hoja4"}
    },
    "Producto": {
        "Contrato": {"sheet_name": "Hoja1"},
        "Retiro": {"sheet_name": "Hoja2"},
        "Prorroga": {"sheet_name": "Hoja3"},
        "Venta": {"sheet_name": "Hoja4"}
    },
    "Factura Venta": {
        "Contrato": {"sheet_name": "Hoja1"},
        "Retiro": {"sheet_name": "Hoja2"},
        "Prorroga": {"sheet_name": "Hoja3"},
        "Venta": {"sheet_name": "Hoja4"}
    }
}
def iniciar_automatizacion(tipo_automatizacion, tipo_proceso):
    global archivo_excel, driver, is_automation_running, save, cont, numero_entry
    root.protocol("WM_DELETE_WINDOW", lambda: None)
    #messagebox.showinfo("inhabilitacion para salir del programa", "detenga la automatización o  espere que se complete, se prohibe el cierre para evitar problemas")
    cont = 0
    numero_str = numero_var.get()  # Obtiene el valor de la variable de la entrada de texto
    if numero_str.isdigit():
        numero = int(numero_str)
        if numero > 0:
            save = numero
            cont = numero - 1
            numero_entry.configure(state='disabled')  # Deshabilitar la entrada de número
        else:
            messagebox.showerror("Error", "El número debe ser mayor que cero.")
            return
    else:
        messagebox.showerror("Error", "Ingrese un número válido.")
        return 
    if archivo_excel and not is_automation_running:
        # Asegúrate de que no estás tratando de iniciar más de un proceso simultáneamente
        if driver and driver.service.process:  # Comprueba si el driver ya está activo
            messagebox.showerror("Error", "Una automatización ya está en curso.")
            
            return
        if not is_automation_running:
            # Comprobaciones adicionales aquí...
            df = pd.read_excel(archivo_excel, sheet_name=automatizaciones[tipo_automatizacion][tipo_proceso]["sheet_name"])
            df = convert_float_columns_to_int(df)
            proceso_thread = threading.Thread(target=ejecutar_proceso, args=(df, tipo_automatizacion, tipo_proceso), daemon=True)
            proceso_thread.start()
        else:
            print("Una automatización ya está en curso.")
    else:
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, carga un archivo Excel antes de iniciar la automatización.")
            numero_entry['state'] = 'normal'
            return
        if is_automation_running:
            messagebox.showerror("Error", "Una automatización ya está en curso.")
            return
        print(f"Iniciando automatización para el proceso: {tipo_automatizacion, tipo_proceso}")

def finalizar_automatizacion():
    global is_automation_running, driver, numero_entry

    if driver:
        try:
            driver.quit()
              # Intenta cerrar el navegador
        except Exception as e:
            print(f"No se pudo cerrar el navegador: {e}")
        driver = None
    is_automation_running = False
    numero_entry.configure(state='normal')  # Habilitar la entrada de número
    print("Automatización finalizada o detenida.")   
    root.protocol("WM_DELETE_WINDOW", on_close)
      
    

def cargar_excel():
    global archivo_excel
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filepath:
        archivo_excel = filepath
        print(f"Archivo seleccionado: {archivo_excel}")
        # Usa lambda para pasar argumentos a la función
        root.after(0, lambda: actualizar_ruta_excel_en_gui(archivo_excel))
    else:
        print("Carga de archivo cancelada")

    
# Función para mostrar mensajes en el panel de manera segura
def print(mensaje):
    panel_mensajes['state'] = 'normal'
    panel_mensajes.insert(tk.END, mensaje + "\n")
    panel_mensajes.see(tk.END)
    panel_mensajes['state'] = 'disabled'

    
def toggle_pause():
    if pause.is_set():
        pause.clear()
        print("Reanudando ejecución.")
    else:
        pause.set()
        print("Pausando ejecución.")
        

        
def stop_execution1(event=None):
    global is_automation_running, stop
    if is_automation_running:
        root.protocol("WM_DELETE_WINDOW", lambda: None)  # Desactiva el evento de cierre
        print("Deteniendo ejecución, espere unos segundos.")
        messagebox.showinfo("Detener Automatización", "Espere a que se detenga la automatización")
        print("Deteniendo ejecución, espere unos segundos.")
        stop.set()
        driver.quit()
        # Asegúrate de unirte a los hilos si es necesario, y de que la limpieza del driver ocurra aquí si no está manejada por el hilo de automatización
    else:
        print("No hay una automatización en curso para detener.")

def stop_execution(root):
    print("Deteniendo ejecución, espere unos segundos.")
    # Aquí tu código para detener la ejecución, asegurándote de que todos los threads y recursos estén cerrados correctamente
    root.destroy()
    
 # Define una StringVar para almacenar el número

def actualizar_ultimo_registro(contrato, numero, fecha, tipo_automatizacion, tipo_proceso):
    ultimo_contrato_var.set(contrato)
    ultimo_numero_var.set(numero)
    ultima_fecha_var.set(fecha)
    tipo_automatizacion_var.set(tipo_automatizacion)
    tipo_proceso_var.set(tipo_proceso)


def cargar_ultimo_registro():
    try:
        with open('ultimo_registro.json', 'r') as file:
            registro = json.load(file)
            # Usar el método get para proporcionar un valor predeterminado si la clave no existe
            contrato = registro.get('contrato', 'Desconocido')
            numero = registro.get('numero', 'Desconocido')
            fecha = registro.get('fecha', 'Desconocido')
            tipo_automatizacion = registro.get('tipo_automatizacion', 'Desconocido')
            tipo_proceso = registro.get('tipo_proceso', 'Desconocido')
            
            actualizar_ultimo_registro(contrato, numero, fecha, tipo_automatizacion, tipo_proceso)
    except FileNotFoundError:
        print("No hay registro anterior guardado.")


def guardar_ultimo_registro(contrato, numero, fecha, tipo_automatizacion_actual, tipo_proceso_actual):
    
    registro = {
        'contrato': contrato,
        'numero': numero,
        'fecha': fecha,
        'tipo_automatizacion': tipo_automatizacion_actual,
        'tipo_proceso': tipo_proceso_actual
    }
    
    with open('ultimo_registro.json', 'w') as file:
        json.dump(registro, file)
    print("Registro guardado correctamente.")  
def guardar_ruta_excel():
    with open('config.txt', 'w') as file:
        file.write(ruta_excel_var.get())        
def actualizar_ruta_excel_en_gui(ruta):
    if ruta:
        ruta_excel_var.set(ruta)  # Actualiza el widget de Tkinter con la nueva ruta
        print(f"Ruta del archivo Excel actualizada: {ruta}")
    else:
        ruta_excel_var.set("No se ha cargado ningún archivo Excel")
        print("No se ha cargado ningún archivo Excel")


def cargar_ruta_excel():
    try:
        with open('config.txt', 'r') as file:
            ruta_excel_var.set(file.read())
    except FileNotFoundError:
        print("Archivo de configuración no encontrado.")

def cargar_ultima_ruta_excel():
    global archivo_excel
    try:
        with open('ultimo_excel.json', 'r') as file:
            datos = json.load(file)
            archivo_excel = datos.get('ultimo_excel', None)
            cargar_ruta_excel()  # Update the GUI with the loaded path
            print(f"Última ruta de Excel cargada: {archivo_excel}")
    except FileNotFoundError:
        print("No previous Excel file path found.")

def guardar_ultima_ruta_excel():
    global archivo_excel  # Asegúrate de que esta variable contiene la ruta actual del archivo Excel
    datos = {'ultimo_excel': archivo_excel}
    with open('ultimo_excel.json', 'w') as file:
        json.dump(datos, file)
    print(f"Ruta del último Excel guardada: {archivo_excel}")
    
def abrir_ultimo_informe():
    global ruta_del_ultimo_informe_generado
    if ruta_del_ultimo_informe_generado and os.path.exists(ruta_del_ultimo_informe_generado):
        webbrowser.open(ruta_del_ultimo_informe_generado)
    else:
        messagebox.showerror("Error", "No se encontró el último informe o aún no se ha generado.")

def on_close():
    if messagebox.askokcancel("Salir", "¿Quieres salir de la aplicación?"):
        guardar_ruta_excel()
        guardar_ruta_ultimo_informe() 
        guardar_ultima_ruta_excel()
        
        if is_automation_running:
            stop.set()
            finalizar_automatizacion()
        root.destroy()
def cargar_ultima_ruta_informe():
    global ruta_del_ultimo_informe_generado
    try:
        with open('ruta_ultimo_informe.json', 'r') as file:
            datos = json.load(file)
            ruta_del_ultimo_informe_generado = datos['ruta']
            if os.path.exists(ruta_del_ultimo_informe_generado):
                print(f"Último informe cargado: {ruta_del_ultimo_informe_generado}")
            else:
                print("El archivo del último informe no existe o la ruta es incorrecta.")
    except (FileNotFoundError, KeyError):
        print("No se encontró el archivo de ruta del último informe o está malformado.")




def create_gui():
    global logo_image, ruta_excel_var, root, panel_mensajes, numero_var, numero_entry, ultimo_contrato_var, ultimo_numero_var, ultima_fecha_var, tipo_automatizacion_var, tipo_proceso_var
    
    # Configurar el estilo de la GUI
    root = tk.Tk()
    root.title("Automatización DIAN COMPRAVENTA LA ESMERALDA")
    root.configure(bg='black')

    icon_path = 'icono.ico'  # Asegúrate de que la ruta sea correcta
    
    # Intenta establecer la imagen como ícono de la ventana
    try:
        root.iconbitmap(icon_path)
        
    except tk.TclError as e:
        print(f"No se pudo cargar el ícono de la aplicación: {e}")
        
        
        
        
    window_width = 1300
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (window_width / 2))
    y_coordinate = int((screen_height / 2) - (window_height / 2))
    root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
            
    # Define un estilo para los bordes verdes
    border_green = {'highlightbackground': 'green', 'highlightcolor': 'green', 'highlightthickness': 1, 'bd': 0}

    main_frame = tk.Frame(root, bg='black')
    main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    frame_numero = tk.LabelFrame(main_frame, text="Número", fg='green', **border_green, bg='black')
    frame_numero.pack(fill=tk.X, padx=5, pady=5)
    numero_var = tk.StringVar(value="1")
    numero_entry = tk.Entry(frame_numero, textvariable=numero_var, insertbackground='white', fg='white', bg='black')
    numero_entry.pack(side=tk.LEFT, padx=10)
    
    frame_excel_path = tk.LabelFrame(main_frame, text="Ruta del Archivo Excel", fg='green', **border_green, bg='black')
    frame_excel_path.pack(fill=tk.X, padx=5, pady=5)
    ruta_excel_var = tk.StringVar(value="N/A")

    tk.Label(frame_excel_path, text="Ruta Actual:", fg='green', bg='black').pack(side=tk.LEFT)
    tk.Label(frame_excel_path, textvariable=ruta_excel_var, fg='white', bg='black').pack(side=tk.LEFT, padx=10, expand=True)
    
    frame_registro = tk.LabelFrame(main_frame, text="Último Registro", fg='green', **border_green, bg='black')
    frame_registro.pack(fill=tk.X, padx=5, pady=5)
    
    ultimo_contrato_var = tk.StringVar(value="N/A")
    ultimo_numero_var = tk.StringVar(value="N/A")
    ultima_fecha_var = tk.StringVar(value="N/A")
    tipo_automatizacion_var = tk.StringVar(value="N/A")
    tipo_proceso_var = tk.StringVar(value="N/A")

    tk.Label(frame_registro, text="Último Contrato:", fg='white', bg='black').pack(side=tk.LEFT)
    tk.Label(frame_registro, textvariable=ultimo_contrato_var, fg='white', bg='black').pack(side=tk.LEFT, padx=10)
    tk.Label(frame_registro, text="Número:", fg='white', bg='black').pack(side=tk.LEFT)
    tk.Label(frame_registro, textvariable=ultimo_numero_var, fg='white', bg='black').pack(side=tk.LEFT, padx=10)
    tk.Label(frame_registro, text="Fecha:", fg='white', bg='black').pack(side=tk.LEFT)
    tk.Label(frame_registro, textvariable=ultima_fecha_var, fg='white', bg='black').pack(side=tk.LEFT, padx=10)
    tk.Label(frame_registro, text="Tipo de automatización:", fg='white', bg='black').pack(side=tk.LEFT)
    tk.Label(frame_registro, textvariable=tipo_automatizacion_var, fg='white', bg='black').pack(side=tk.LEFT, padx=10)
    tk.Label(frame_registro, text="Tipo de proceso:", fg='white', bg='black').pack(side=tk.LEFT)
    tk.Label(frame_registro, textvariable=tipo_proceso_var, fg='white', bg='black').pack(side=tk.LEFT, padx=10)

    panel_frame = tk.Frame(root, bg='black')
    panel_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
    
    panel_mensajes = scrolledtext.ScrolledText(panel_frame, state='disabled', bg='black', fg='white')
    panel_mensajes.pack(fill=tk.BOTH, expand=True)
    
    for seccion in ["Adquirientes", "Producto", "Factura Venta"]:
        frame_seccion = tk.LabelFrame(main_frame, text=seccion, fg='white', **border_green, bg='black')
        frame_seccion.pack(fill=tk.X, padx=5, pady=5)
        for proceso in ["Contrato", "Retiro", "Prorroga", "Venta"]:
            boton = tk.Button(frame_seccion, text=proceso, bg='green', fg='white', 
                              command=lambda s=seccion, p=proceso: iniciar_automatizacion(s, p))
            boton.pack(side=tk.LEFT, padx=10)
    
    frame_acciones = tk.Frame(main_frame, bg='black')
    frame_acciones.pack(fill=tk.X, padx=5, pady=5)

    boton_abrir_informe = tk.Button(frame_acciones, text="Abrir Último Informe", bg='green', fg='white', command=abrir_ultimo_informe)
    boton_abrir_informe.pack(side=tk.LEFT, padx=10)

    boton_cargar = tk.Button(frame_acciones, text="Cargar Archivo Excel", bg='green', fg='white', command=cargar_excel)
    boton_cargar.pack(side=tk.LEFT, padx=10)

    boton_pausar = tk.Button(frame_acciones, text="Pausar/Reanudar", bg='green', fg='white', command=toggle_pause)
    boton_pausar.pack(side=tk.LEFT, padx=10)

    boton_detener = tk.Button(frame_acciones, text="Detener", bg='green', fg='white', command=stop_execution1)
    boton_detener.pack(side=tk.LEFT, padx=10)
    cargar_ruta_excel()
    cargar_ultimo_registro()
    cargar_ultima_ruta_excel()
    guardar_ruta_ultimo_informe()
    cargar_ultima_ruta_informe()
    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    
    create_gui()


    




             
