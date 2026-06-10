import xlwings as xw
import os
import tagetik_bot
import time
from functions import consolidar_reportes, cerrar_libros_abiertos

def ejecutar_proceso_completo():
    """Esta es la función principal"""
    try:
        # 1. Llamar a la descarga (pyautogui)
        print("Iniciando descarga con PyAutoGUI...")
        tagetik_bot.generar_reportes() 

        # 2. Cerrar todos los libros abiertos
        print('Cerrando libros abiertos')
        cerrar_libros_abiertos()
        time.sleep(2)

        # 3. Llamar a la consolidación (xlwings)
        print("Iniciando consolidación...")
        consolidar_reportes()
        
        print("✅ Proceso completado con éxito.")
        
    except Exception as e:
        print(f"❌ Error durante el proceso: {str(e)}")

if __name__ == "__main__":
    ejecutar_proceso_completo()