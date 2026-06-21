import pyautogui
import time
import os
import xlwings as xw
import glob
import take_screenshots
from functions import esperar_caja_imagen, esperar_imagen, limpiar_carpeta_reportes, take_uic, genenate_report

carpeta_reportes = r'C:\Users\bmorenoa\Documents'

def generar_reportes():
    
    limpiar_carpeta_reportes(carpeta_reportes)

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

        #Parameters are selected by the user

        caja_parametro = esperar_imagen('parameters.png', confidence=0.9)

        response = pyautogui.confirm(text='Click ok in this dialog when paramenters are done!', title='Paramenters', buttons=['Ok', 'Cancel'])

        if response == 'Cancel':
            return
        # iconos = list(pyautogui.locateAllOnScreen('img/downArrow.png', confidence=0.8))
        # opcion = iconos[1] 
        # centro = pyautogui.center(opcion)
        # pyautogui.click(centro)

        # time.sleep(2)
        # caja_parametro = pyautogui.locateOnScreen('img/EN05.png', confidence=0.7)
        # x_clic = caja_parametro.left + 15
        # y_clic = caja_parametro.top + (caja_parametro.height / 2)
        # pyautogui.click(x_clic, y_clic)

        # time.sleep(2)
        # caja_parametro = pyautogui.locateOnScreen('img/BM.png', confidence=0.7)
        # x_clic = caja_parametro.left + 15
        # y_clic = caja_parametro.top + (caja_parametro.height / 2)
        # pyautogui.click(x_clic, y_clic)

        # time.sleep(2)
        # caja_parametro = pyautogui.locateOnScreen('img/USBM.png', confidence=0.7)
        # x_clic = caja_parametro.left + 15
        # y_clic = caja_parametro.top + (caja_parametro.height / 2)
        # pyautogui.click(x_clic, y_clic)

        # time.sleep(2)
        # boton_coords = pyautogui.locateCenterOnScreen('img/USA.png', confidence=0.7)
        # pyautogui.doubleClick(boton_coords)

        # time.sleep(2)
        # opcion = iconos[2] 
        # centro = pyautogui.center(opcion)
        # pyautogui.click(centro)

        # time.sleep(2)
        # caja_parametro = pyautogui.locateOnScreen('img/conso scenario.png', confidence=0.7)
        # x_clic = caja_parametro.left + 15
        # y_clic = caja_parametro.top + (caja_parametro.height / 2)
        # pyautogui.click(x_clic, y_clic)

        # #click en el consolidation scenario 2026-USD
        # time.sleep(2)
        # green_iconos = list(pyautogui.locateAllOnScreen('img/green-icon.png', confidence=0.7))
        # ultima_opcion = green_iconos[-1] 
        # centro = pyautogui.center(ultima_opcion)
        # pyautogui.doubleClick(centro)

        # time.sleep(2)
        # opcion = iconos[3] 
        # centro = pyautogui.center(opcion)
        # pyautogui.click(centro)

        # time.sleep(2)
        # boton_coords = pyautogui.locateCenterOnScreen('img/month.png', confidence=0.9)
        # pyautogui.doubleClick(boton_coords)

        # time.sleep(2)
        # opcion = iconos[4] 
        # centro = pyautogui.center(opcion)
        # pyautogui.click(centro)

        # time.sleep(2)
        # caja_parametro = pyautogui.locateOnScreen('img/100.png', confidence=0.7)
        # x_clic = caja_parametro.left + 15
        # y_clic = caja_parametro.top + (caja_parametro.height / 2)
        # pyautogui.click(x_clic, y_clic)

        # time.sleep(2)
        # boton_coords = pyautogui.locateCenterOnScreen('img/total.png', confidence=0.7)
        # pyautogui.doubleClick(boton_coords)

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

    report_317()
    genenate_report(img='316 BS', report='report_316')
    genenate_report(img='315 IS', report='report_315')

    #Open 200 Reports
    time.sleep(3)
    caja_parametro = pyautogui.locateOnScreen('img/200 reports.png', confidence=0.7)
    x_clic = caja_parametro.left + 15
    y_clic = caja_parametro.top + (caja_parametro.height / 2)
    pyautogui.click(x_clic, y_clic)

    genenate_report(img='201', report='report_201')
    genenate_report(img='202', report='report_202')
    genenate_report(img='203', report='report_203')
    genenate_report(img='204', report='report_204')
    genenate_report(img='205', report='report_205')
    genenate_report(img='206', report='report_206')

    #Display 21* Reports

    def report_210B():

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
        boton_coords = esperar_imagen('save.png', timeout=20, confidence=0.7)
        time.sleep(1)
        pyautogui.click(boton_coords)

        take_uic('report_210B')
    
    report_210B()

    time.sleep(2)
    boton_coords = pyautogui.locateCenterOnScreen('img/200 reports 2.png', confidence=0.8)
    pyautogui.click(boton_coords)
    time.sleep(1)
    pyautogui.press('pagedown')
    time.sleep(1)
    pyautogui.press('pagedown')

    genenate_report(img='211B', report='report_211B')

    def report_404():

        time.sleep(5)
        boton_coords = pyautogui.locateCenterOnScreen('img/300 financial statements.png', confidence=0.9)
        pyautogui.click(boton_coords)
        time.sleep(1)
        pyautogui.press('pagedown')
        
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

        #click en el consolidation scenario 2026-USD-CAD
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
        boton_coords = esperar_imagen('save.png', timeout=20, confidence=0.7)
        time.sleep(1)
        pyautogui.click(boton_coords)

        take_uic('report_404')
    report_404()

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

        #click en el consolidation scenario 2026-USD-CAD
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
        boton_coords = esperar_imagen('save.png', timeout=20, confidence=0.7)
        time.sleep(1)
        pyautogui.click(boton_coords)
        time.sleep(3)

        #screenshots
        take_uic('report_402')
    report_402()
    

if __name__ == "__main__":
    generar_reportes()
