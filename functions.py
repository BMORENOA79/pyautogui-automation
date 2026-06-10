import pyautogui
import time
import os
import xlwings as xw
import glob
import take_screenshots

# Definir la base una sola vez
BASE_DIR = os.path.dirname(os.path.abspath(__file__))



def limpiar_carpeta_reportes(ruta_carpeta):
    print(f"🧹 Limpiando archivos de Excel en: {ruta_carpeta}")
    
    # Definimos las extensiones que queremos borrar
    extensiones = ['*.xlsx', '*.xls', '*.csv', '*.xlsm']
    archivos_borrados = 0

    try:
        for ext in extensiones:
            # Buscamos todos los archivos que coincidan con la extensión
            archivos = glob.glob(os.path.join(ruta_carpeta, ext))
            
            for archivo in archivos:
                # Opcional: Evitar borrar tu archivo maestro si está en la misma carpeta
                if "tagetik.xlsm" in archivo.lower():
                    continue
                
                try:
                    os.remove(archivo)
                    print(f"Eliminado: {os.path.basename(archivo)}")
                    archivos_borrados += 1
                except Exception as e:
                    print(f"⚠️ No se pudo eliminar {archivo}: {e}")
        
        print(f"✅ Limpieza terminada. Se eliminaron {archivos_borrados} archivos.")
        time.sleep(2)
        
    except Exception as e:
        print(f"❌ Error crítico al acceder a la carpeta: {e}")

def esperar_imagen(nombre_archivo, timeout=20, confidence=0.7):
        # Construimos la ruta dinámica
        ruta = os.path.join(BASE_DIR, 'img', nombre_archivo)
        
        inicio = time.time()
        print(f"Buscando {nombre_archivo}...")

        while True:
            try:
                # Intentamos localizar la imagen
                posicion = pyautogui.locateCenterOnScreen(ruta, confidence=confidence)
                
                if posicion is not None:
                    print(f"✅ Imagen encontrada en {posicion}")
                    time.sleep(1)
                    return posicion
                    
            except (pyautogui.ImageNotFoundException, Exception):
                # Si no la encuentra, lanza la excepción. 
                # Simplemente la ignoramos para que el bucle continúe.
                pass
            
            # Verificar si ya pasamos el tiempo límite
            if time.time() - inicio > timeout:
                print(f"❌ Tiempo agotado: No se encontró {nombre_archivo}")
                return None
            
            # Esperar un poco antes de volver a intentar
            time.sleep(0.5)

def esperar_caja_imagen(nombre_img, timeout=20, confidence=0.8):
    """Espera una imagen y devuelve su recuadro (Box) para calcular offsets."""
    ruta = os.path.join(BASE_DIR, 'img', nombre_img)
    inicio = time.time()
    
    while time.time() - inicio < timeout:
        try:
            # Buscamos la caja (left, top, width, height)
            caja = pyautogui.locateOnScreen(ruta, confidence=confidence)
            if caja:
                return caja
        except:
            pass # Ignora errores temporales de lectura
        
        time.sleep(0.5) # Pausa para no saturar el CPU
    
    return None

def take_uic(report_name):
    caja_parametro = esperar_caja_imagen('close_bar.png', confidence=0.8)
    x_clic = caja_parametro.left + (caja_parametro.width - 15)
    y_clic = caja_parametro.top + (caja_parametro.height / 2)
    time.sleep(2)
    pyautogui.click(x_clic, y_clic)
    take_screenshots.capturar_reporte_completo(report_name)

def consolidar_reportes():

    # 1. Configuración
    carpeta_reportes = r'C:\Users\bmorenoa\Documents'
    archivo_maestro_ruta = os.path.join(os.getcwd(), 'USCEM Financial Statements.xlsx')

    app = xw.App(visible=False)

    try:
        if os.path.exists(archivo_maestro_ruta):
            try: os.remove(archivo_maestro_ruta)
            except: pass
        
        maestro_wb = app.books.add()
        archivos = [f for f in os.listdir(carpeta_reportes) if f.endswith('.xlsx')]

        for nombre_archivo in archivos:
            ruta_origen = os.path.join(carpeta_reportes, nombre_archivo)
            print(f"Procesando: {nombre_archivo}...")
            
            origen_wb = app.books.open(ruta_origen)
            
            for hoja in origen_wb.sheets:
                prefijo = nombre_archivo.split('_')[0]
                nuevo_nombre = f"{prefijo}_{hoja.name}"[:31]
                
                # Copiar la hoja completa
                hoja.copy(after=maestro_wb.sheets[-1])
                maestro_wb.sheets[-1].name = nuevo_nombre
                
            origen_wb.close()

        # --- PASO CRÍTICO: Romper vínculos en todo el libro maestro ---
        # Obtenemos la lista de vínculos externos (archivos de origen)
        vinc_externos = maestro_wb.api.LinkSources(Type=1) # 1 = xlExcelLinks
        
        if vinc_externos:
            print(f"Encontrados {len(vinc_externos)} vínculos externos. Rompiendo...")
            for vinculo in vinc_externos:
                # Esta instrucción convierte todas las fórmulas ligadas a ese vínculo en valores
                maestro_wb.api.BreakLink(Name=vinculo, Type=1)
            print("✅ Vínculos rotos exitosamente.")
        else:
            print("No se detectaron vínculos externos pendientes.")

        # Limpieza de la hoja inicial
        if 'Sheet1' in [h.name for h in maestro_wb.sheets]:
            maestro_wb.sheets['Sheet1'].delete()

        maestro_wb.save(archivo_maestro_ruta)
        print(f"\n✅ Proceso completado. Archivo guardado como: {archivo_maestro_ruta}")


    finally:
        app.quit()

def cerrar_libros_abiertos():
    
    print("Iniciando limpieza profunda de Excel...")
    os.system("taskkill /f /im excel.exe")