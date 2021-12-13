# -*- coding: utf-8 -*-
"""
Created on Tue May 25 16:45:21 2021

@author: leoesleo1111@gmail.com
"""

import random
from selenium import webdriver
from time import sleep
import shutil, os
from os import remove

#PARAMETROS DE ENTRADA
#****************************
anioDesde = 1992
anioHasta = 2020
mesParam  = 'Destino Diciembre'
path_descargas = 'C:/Users/CAP04/Downloads/'
#****************************

def trasladoFichero(fichero):
    """
    La funcion pide como parametro el archivo que se va a buscar en descargas para pasarlo a la carpeta ficheros.  
    Si existe un archivo con el mismo nombre, lo sobrescribirá.
    """    
    path_ficheros = 'ficheros'
    res = 'error'
    #fichero = 'Destino Diciembre 2019'
    #print(fichero,"##############################")
    with os.scandir(path_descargas) as ficheros:
        ficheros = [fichero.name for fichero in ficheros if fichero.is_file()]
    for i in range(len(ficheros)):
        if ficheros[i].rfind(fichero) != -1: #validar si existe
            try:
                shutil.move(path_descargas+ficheros[i], path_ficheros) #corta y pega en ficheros
                remove(path_ficheros+'/'+fichero+'.xlsx')
                os.rename(path_ficheros+'/'+ficheros[i], path_ficheros+'/'+fichero+'.xlsx')
                print("---------Archivo renombrado")
                res = 'ok'
            except:
                res = 'error'
    return res            

#rango de años
if anioDesde > anioHasta:
    print("El anioDesde debe ser menor a anioHasta")
else:    
    driver = webdriver.Chrome("chromedriver.exe")
    driver.get("https://www.aerocivil.gov.co/atencion/estadisticas-de-las-actividades-aeronauticas/bases-de-datos")
    
    try:
        boton = driver.find_element_by_xpath('//div[@class="chordion chordion1"]')
        boton.click()
        sleep(random.uniform(1.0, 2.0))    
        #crear vector con los meses a buscar            
        vector_meses = []
        for i in range(anioHasta-anioDesde):
            vector_meses.append(mesParam+' '+str(anioDesde+i))
        vector_meses.append(mesParam+' '+str(anioDesde+i+1))
        
        #crear lista con todos los meses de ​Origen - Destino
        lista_mes = boton.find_elements_by_xpath('.//li[@class="dfwp-item"]')
        for i in lista_mes:
            mes = i.find_element_by_xpath('.//h2[@class="title-article"]').text
            mesbuscado = mes[mes.find('-')+1:].strip()
            print("Buscando:",mesParam,"en:",mesbuscado,"[",mesbuscado in vector_meses,"]")
            if mesbuscado in vector_meses:
                try:
                    i.find_element_by_xpath('.//a[@class="tool-doc download"]').click()    
                except Exception as e:
                    print("---------ERROR; No se encuentra el botón de descarga del periodo: ",mes[mes.find('-')+1:].strip() + str(e))    
                try: 
                    sleep(random.uniform(10.0, 15.0)) 
                    res = trasladoFichero(mesbuscado)
                    res = 'ok'
                    print("---------Descargado: ",mesbuscado,"estado del traslado: ",res) 
                except Exception as e:                               
                    print("---------ERROR; Problemas para trasladar fichero desde DESCARGAS" + str(e))                  
    except Exception as e: 
        print("No se encuentra el botón ​Origen - Destino" + str(e))
driver.quit()

