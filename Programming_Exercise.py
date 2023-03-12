#requiere:
#pip install watchdog
#pip install xlwings

from pathlib import Path 
import sys
import os
import time
import logging
import threading
import glob
import shutil
import asyncio
import xlwings as xw  
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from concurrent.futures import ThreadPoolExecutor


_executor = ThreadPoolExecutor(1)


#Classes
class File:
  def __init__(self, file_name, file_ext,file,path):
    self.file_name= file_name
    self.file_ext = file_ext
    self.file = file
    self.path = path

class AsyncWrite(threading.Thread):

    def __init__(self,path):
        self.path = path
      
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



def process_files(files):
    for archivo in files:
         clasificar(archivo)

#Clasifica archivos
def clasificar(archivo):
      if archivo.file_ext == 'xlsx' or archivo.file_ext == 'xlsm' or archivo.file_ext == 'xls' :
          
          if os.path.exists(path+'/Processed'):
            print("Processing")
            shutil.move(path+'/'+archivo.file, path+'/Processed/'+archivo.file)
            print("Moved")
    
          else:
              print("Folder created and file moved")
              os.makedirs(path+'/Processed')
              shutil.move(path+'/'+archivo.file, path+'/Processed/'+archivo.file)
      else:
          print("Nothing to process")
          if os.path.exists(path+'/Not applicable'): 
              shutil.move(path+'/'+archivo.file, path+'/Not applicable/'+archivo.file)
              print("Moving")
              
          else:
            os.makedirs(path+'/Not applicable')
            shutil.move(path+'/'+archivo.file, path+'/Not applicable/'+archivo.file)
            print("Moving")     


#Mueve archivos a carpeta
def on_created(event):
    print("New File detected, getting the data...")
    files = os.listdir(path)
    my_files=[]
    for file in files:
        filename,file_ext = os.path.splitext(file)
        file_ext = file_ext[1:]
        if file_ext :
            my_files.append(File(filename,file_ext,file,path))
            print(filename)
    print("Moving Files...")
    background = AsyncWrite(path)
    process_files(my_files)
    background.start()
    background.join()
    print("Ready")
   
        
def on_moved(event):
    print("The File has Been Clasified")
    

if __name__ == "__main__":
    event_handler = FileSystemEventHandler()
    # Llamada de Funciones
    event_handler.on_created = on_created
    event_handler.on_moved = on_moved

    path = input("Write the folder path: ")
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()

    try:
        print("Monitoring")
        files = os.listdir(path)
        my_files=[]
        for file in files:
            filename,file_ext = os.path.splitext(file)
            file_ext = file_ext[1:]
            if file_ext :
                my_files.append(File(filename,file_ext,file,path))
                print(filename)

    
        print("Moving Files..")
        background = AsyncWrite(path)
        process_files(my_files)
        background.start()       
        background.join()
        print("Ready")

        
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("Finish")
    observer.join()
