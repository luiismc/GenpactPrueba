#Codigo hecho por Luis Marin prueba Genpact
from pathlib import Path 
import sys
import os
import time
import logging
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import glob
import shutil
import asyncio
from concurrent.futures import ThreadPoolExecutor
import xlwings as xw  # pip install xlwings
_executor = ThreadPoolExecutor(1)






#Clases
class Archivo:
  def __init__(self,nombre, extension,file,path):
    self.nombre = nombre
    self.extension = extension
    self.file = file
    self.path = path

class AsyncWrite(threading.Thread):

    def __init__(self,path):
        self.path = path
        # calling superclass init
        threading.Thread.__init__(self)
    
 
    def run(self):
       SOURCE_DIR = self.path+'/Processed'
       excel_files = list(Path(SOURCE_DIR).glob('*.xlsx'))
       combined_wb = xw.Book()

       for excel_file in excel_files:
           wb = xw.Book(excel_file)
           for sheet in wb.sheets:
              sheet.copy(after=combined_wb.sheets[0])
           wb.close()

       combined_wb.sheets[0].delete()
       combined_wb.save(f'masterbook.xlsx')
       if len(combined_wb.app.books) == 1:
            combined_wb.app.quit()
       else:
            combined_wb.close()





#Funciones
def procesadorArchivos(archivos):
    for archivo in archivos:
         clasificar(archivo)

def clasificar(archivo):
      if archivo.extension == 'xlsx' or archivo.extension == 'xlsm' or archivo.extension == 'xls' :
          
          if os.path.exists(path+'/Processed'):
            print("procesando")
            shutil.move(path+'/'+archivo.file, path+'/Processed/'+archivo.file)
            print("Se movio")
    
          else:
              print("Se Creo directorio y se movio")
              os.makedirs(path+'/Processed')
              shutil.move(path+'/'+archivo.file, path+'/Processed/'+archivo.file)
      else:
          print("nada que procesar")
          if os.path.exists(path+'/Not applicable'): 
              shutil.move(path+'/'+archivo.file, path+'/Not applicable/'+archivo.file)
              print("Moving away")
              
          else:
            os.makedirs(path+'/Not applicable')
            shutil.move(path+'/'+archivo.file, path+'/Not applicable/'+archivo.file)
            print("Moving away")     

    


#Eventos: Archivo nuevo en carpeta 
def on_created(event):
    print("Archivo nuevo detectado, consiguiendo informacion..")
    files = os.listdir(path)
    mis_archivos=[]
    for file in files:
        filename,extension = os.path.splitext(file)
        extension = extension[1:]
        if extension :
            mis_archivos.append(Archivo(filename,extension,file,path))
            print(filename)
    print("Moviendo archivos..")
    background = AsyncWrite(path)
    procesadorArchivos(mis_archivos)
    background.start()
    background.join()
    print("Listo")
   
    

    
def on_moved(event):
    print("Se clasifico un archivo")
    
def pathreboot():
    event_handler = FileSystemEventHandler()
    # calling functions
    event_handler.on_created = on_created
    event_handler.on_moved = on_moved

    path = input("Introduzca la ruta de la carpeta: ")
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()

    try:
        print("Monitoreando")
        files = os.listdir(path)
        mis_archivos=[]
        for file in files:
            filename,extension = os.path.splitext(file)
            extension = extension[1:]
            if extension :
                mis_archivos.append(Archivo(filename,extension,file,path))
                print(filename)

  
        procesadorArchivos(mis_archivos)
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("terminado")
    observer.join()


if __name__ == "__main__":
    event_handler = FileSystemEventHandler()
    # calling functions
    event_handler.on_created = on_created
    event_handler.on_moved = on_moved

    path = input("Introduzca la ruta de la carpeta: ")
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()

    try:
        print("Monitoreando")
        files = os.listdir(path)
        mis_archivos=[]
        for file in files:
            filename,extension = os.path.splitext(file)
            extension = extension[1:]
            if extension :
                mis_archivos.append(Archivo(filename,extension,file,path))
                print(filename)

    
        print("Moviendo archivos..")
        background = AsyncWrite(path)
        procesadorArchivos(mis_archivos)
        background.start()       
        background.join()
        print("Listo")

        print("puede cambiar el directorio con ctrl+c")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("terminado")
    observer.join()
    pathreboot()

