import pyautogui
import time
import os
import xlwings as xw
import glob
import take_screenshots
from functions import esperar_caja_imagen, esperar_imagen, limpiar_carpeta_reportes, take_uic

carpeta_reportes = r'C:\Users\bmorenoa\Documents'

def generar_reportes():

    def report_317():

        boton = esperar_imagen('tagetik_addin.png', timeout=20, confidence=0.9)
        pyautogui.click(boton)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/navigation.png', confidence=0.8)
        pyautogui.click(boton_coords)

        caja_parametro = esperar_caja_imagen('300 financial statements.png', confidence=0.9)
        #Dar click en el nodo de la izquierda
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/317 IS.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)

        caja_parametro = esperar_imagen('parameters.png', confidence=0.9)
        iconos = list(pyautogui.locateAllOnScreen('img/downArrow.png', confidence=0.8))
        opcion = iconos[1] 
        centro = pyautogui.center(opcion)
        pyautogui.click(centro)

        time.sleep(2)
        caja_parametro = pyautogui.locateOnScreen('img/EN05.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(2)
        caja_parametro = pyautogui.locateOnScreen('img/BM.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(2)
        caja_parametro = pyautogui.locateOnScreen('img/USBM.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/USA.png', confidence=0.7)
        pyautogui.doubleClick(boton_coords)

        time.sleep(2)
        opcion = iconos[2] 
        centro = pyautogui.center(opcion)
        pyautogui.click(centro)

        time.sleep(2)
        caja_parametro = pyautogui.locateOnScreen('img/conso scenario.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        #click en el consolidation scenario 2026-USD
        time.sleep(2)
        green_iconos = list(pyautogui.locateAllOnScreen('img/green-icon.png', confidence=0.7))
        ultima_opcion = green_iconos[-1] 
        centro = pyautogui.center(ultima_opcion)
        pyautogui.doubleClick(centro)

        time.sleep(2)
        opcion = iconos[3] 
        centro = pyautogui.center(opcion)
        pyautogui.click(centro)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/month.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)

        time.sleep(2)
        opcion = iconos[4] 
        centro = pyautogui.center(opcion)
        pyautogui.click(centro)

        time.sleep(2)
        caja_parametro = pyautogui.locateOnScreen('img/100.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/total.png', confidence=0.7)
        pyautogui.doubleClick(boton_coords)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/ok.png', confidence=0.7)
        take_screenshots.captura_inputs('report_315')
        pyautogui.click(boton_coords)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        #screenshots
        take_uic('report_317')

    def report_316():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/316 BS.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)
       
        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_316')
        pyautogui.click(boton)
        time.sleep(1)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_316')

    def report_315():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/315 IS.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)

        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_317')
        pyautogui.click(boton)
        time.sleep(1)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_315')
  
    def report_201():

        time.sleep(3)
        caja_parametro = pyautogui.locateOnScreen('img/200 reports.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/201.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)


        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_201')
        pyautogui.click(boton)
        time.sleep(1)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_201')

    def report_202():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/202.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)


        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_202')
        pyautogui.click(boton)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_202')
  
    def report_203():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/203.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)


        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_203')
        pyautogui.click(boton)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(1)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_203')

    def report_204():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/204.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)


        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_204')
        pyautogui.click(boton)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(1)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_204')
    
    def report_205():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/205.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)


        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_205')
        pyautogui.click(boton)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_205')
    
    def report_206():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/206.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)


        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_206')
        pyautogui.click(boton)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_206')
    
    def report_212B():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/200 reports.png', confidence=0.9)
        pyautogui.click(boton_coords)

        time.sleep(1)
        pyautogui.press('pagedown')
        time.sleep(1)
        pyautogui.press('pagedown')

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/212B.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)

        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        pyautogui.click(boton)
        time.sleep(1)

        boton = esperar_imagen('export.png', timeout=30, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

    def report_210B():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/300 financial statements.png', confidence=0.9)
        pyautogui.click(boton_coords)

        time.sleep(1)
        pyautogui.press('pageup')
        time.sleep(1)
        pyautogui.press('up')

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/210B.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)

        #Cambiar a original amounts
        time.sleep(6)
        caja_parametro = esperar_imagen('parameters.png', confidence=0.9)
        iconos = list(pyautogui.locateAllOnScreen('img/downArrow.png', confidence=0.8))
        opcion = iconos[2] 
        centro = pyautogui.center(opcion)
        pyautogui.click(centro)

        caja_parametro = esperar_caja_imagen('original_scenario.png', confidence=0.9)
        #Dar click en el nodo de la izquierda
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(2)
        green_iconos = list(pyautogui.locateAllOnScreen('img/green-icon.png', confidence=0.7))
        ultima_opcion = green_iconos[-1] 
        centro = pyautogui.center(ultima_opcion)
        pyautogui.doubleClick(centro)

        time.sleep(2)
        opcion = iconos[4] 
        centro = pyautogui.center(opcion)
        pyautogui.click(centro)

        time.sleep(2)
        caja_parametro = pyautogui.locateOnScreen('img/100.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        caja_parametro = esperar_caja_imagen('total.png', confidence=0.9)
        #Dar click en el nodo de la izquierda
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/precon_amounts.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)

        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_210B')
        pyautogui.click(boton)
        time.sleep(1)   

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_210B')
    
    def report_211B():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/211B.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)

        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_211B')
        pyautogui.click(boton)
        time.sleep(1)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_211B')

    def report_404():

        time.sleep(5)
        boton_coords = pyautogui.locateCenterOnScreen('img/210B.png', confidence=0.9)
        pyautogui.click(boton_coords)

        time.sleep(2)
        pyautogui.press('pagedown')
        time.sleep(1)
        pyautogui.press('pagedown')
        time.sleep(1)
        pyautogui.press('pagedown')
        time.sleep(1)
        
        caja_parametro = pyautogui.locateOnScreen('img/400 reports.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/404.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)

        caja_parametro = esperar_imagen('parameters.png', confidence=0.9)
        iconos = list(pyautogui.locateAllOnScreen('img/downArrow.png', confidence=0.8))
        opcion = iconos[1] 
        centro = pyautogui.center(opcion)
        pyautogui.click(centro)

        time.sleep(2)
        caja_parametro = pyautogui.locateOnScreen('img/conso scenario.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        #click en el consolidation scenario 2026-USD
        time.sleep(2)
        green_iconos = list(pyautogui.locateAllOnScreen('img/green-icon.png', confidence=0.7))
        ultima_opcion = green_iconos[-1] 
        centro = pyautogui.center(ultima_opcion)
        pyautogui.doubleClick(centro)

        take_screenshots.captura_inputs('report_404')

        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        pyautogui.click(boton)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)

        take_uic('report_404')
    
    def report_402():

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/402.png', confidence=0.9)
        pyautogui.doubleClick(boton_coords)

        #Select consolidated amounts
        time.sleep(6)
        caja_parametro = esperar_imagen('parameters.png', confidence=0.9)
        iconos = list(pyautogui.locateAllOnScreen('img/downArrow.png', confidence=0.8))
        opcion = iconos[1] 
        centro = pyautogui.center(opcion)
        pyautogui.click(centro)

        time.sleep(2)
        caja_parametro = pyautogui.locateOnScreen('img/conso scenario.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        #click en el consolidation scenario 2026-USD
        time.sleep(2)
        green_iconos = list(pyautogui.locateAllOnScreen('img/green-icon.png', confidence=0.7))
        ultima_opcion = green_iconos[-1] 
        centro = pyautogui.center(ultima_opcion)
        pyautogui.doubleClick(centro)

        time.sleep(2)
        caja_parametro = esperar_imagen('parameters.png', confidence=0.9)
        iconos = list(pyautogui.locateAllOnScreen('img/downArrow.png', confidence=0.8))
        opcion = iconos[4] 
        centro = pyautogui.center(opcion)
        pyautogui.click(centro)

        time.sleep(2)
        caja_parametro = pyautogui.locateOnScreen('img/100.png', confidence=0.7)
        x_clic = caja_parametro.left + 15
        y_clic = caja_parametro.top + (caja_parametro.height / 2)
        pyautogui.click(x_clic, y_clic)

        time.sleep(2)
        boton_coords = pyautogui.locateCenterOnScreen('img/total.png', confidence=0.7)
        pyautogui.doubleClick(boton_coords)


        boton = esperar_imagen('ok.png', timeout=20, confidence=0.7)
        take_screenshots.captura_inputs('report_402')
        pyautogui.click(boton)

        time.sleep(6)

        boton = esperar_imagen('export.png', timeout=20, confidence=0.7)
        time.sleep(2)
        pyautogui.click(boton)

        time.sleep(3)
        boton_coords = pyautogui.locateCenterOnScreen('img/save.png', confidence=0.7)
        pyautogui.click(boton_coords)
        time.sleep(3)

        #screenshots
        take_uic('report_402')

    #Execute
    limpiar_carpeta_reportes(carpeta_reportes)
    report_317()
    report_316()
    report_315()
    report_201()
    report_202()
    report_203()
    report_204()
    report_205()
    report_206()
    report_212B()
    report_210B()
    report_211B()
    report_404()
    report_402()


if __name__ == "__main__":
    generar_reportes()
