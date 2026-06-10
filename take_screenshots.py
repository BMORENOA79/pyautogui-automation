import pyautogui
import xlwings as xw
import os
import time

# Configuración de ruta
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCREENSHOT_DIR = os.path.join(BASE_DIR, 'screenshots')

# Crear la carpeta si no existe
if not os.path.exists(SCREENSHOT_DIR):
    os.makedirs(SCREENSHOT_DIR)

def capturar_reporte_completo(nombre_base):
    try:
        wb = xw.books.active
        sheet = wb.sheets.active
        time.sleep(2) # Pausa para que la ventana se acomode
        # --- CAPTURA INICIAL ---
        # Ir a la celda A1 para asegurar que estamos al inicio
        wb.sheets.active.range("A1").select()
        time.sleep(0.5)
        
        ruta_inicial = os.path.join(SCREENSHOT_DIR, f"{nombre_base}_inicial.png")
        pyautogui.screenshot(ruta_inicial)
        print(f"📸 Captura inicial guardada: {ruta_inicial}")

        # --- DESPLAZAMIENTO AL FINAL ---
        # Usamos xlwings para encontrar la última celda con datos
        # Esto simula presionar Ctrl + Fin
        ultima_fila_e = sheet.range('E' + str(sheet.cells.last_cell.row)).end('up').row
        u_col = sheet.range(ultima_fila_e, sheet.cells.last_cell.column).end('left').column
        celda_final = sheet.cells(ultima_fila_e, u_col)
        celda_final.select()
        # Un pequeño scroll adicional para asegurar que vemos el final real
        wb.app.api.ActiveWindow.SmallScroll(Up=5) 
        time.sleep(1) # Esperar a que el scroll termine

        # --- CAPTURA FINAL ---
        ruta_final = os.path.join(SCREENSHOT_DIR, f"{nombre_base}_final.png")
        pyautogui.screenshot(ruta_final)
        print(f"📸 Captura final guardada: {ruta_final}")

        pyautogui.hotkey('alt', 'tab')

    except Exception as e:
        print(f"❌ Error al tomar screenshots de {nombre_base}: {e}")

def captura_inputs(report):
    try:
        ruta = os.path.join(SCREENSHOT_DIR, f"{report}_inputs.png")
        time.sleep(1)
        pyautogui.screenshot(ruta)
        print(f"📸 Captura de los inputs guardada: {ruta}")

    except Exception as e:
        print(f"❌ Error al tomar screenshots de {report}: {e}")
    

if __name__ == "__main__":
    capturar_reporte_completo('report_317')
