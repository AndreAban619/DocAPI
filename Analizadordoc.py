#   Se instala el DocxTemplate con pip install docxtpl
from docxtpl import DocxTemplate #importamos la libreria
import os
import comtypes.client
import time
import urllib.parse
import requests
url = 'http://localhost:3000/api/tokens'
caracteres = "[]"
data= requests.get(url)
if data.status_code == 200:
    data = data.json()
    print(data)
format_code = 17

time_start = time.time()

# Creamos el MS word app
word_app = comtypes.client.CreateObject('Word.Application')
word_app.Visible = False

tpl = DocxTemplate("Documentos\FORMATOS ISAI HACKATHON.docx") #plantilla a utilizar
context = {'ISAI_Articulo' : 1 }#partes a reemplazar de la plantilla la izquierda es lo tokens y la derecha es la información por la cual se reemplaza por {{}}
copias = 1

for i in range(0, copias, 1): # numero de copias del documento introducido
    print(i)   
    tpl.render(context) # El render introduce el valor
    tpl.save("generated_doc"+str(i+1)+".docx") # Nombre y formato a guardar
   
# conversion
for i in range(0, copias, 1): # numero de copias del documento introducido
    print(i)
    file_input = os.path.abspath('generated_doc'+str(i+1)+'.docx')
    file_output = os.path.abspath('generated_doc'+str(i+1)+'.pdf')
    word_file = word_app.Documents.Open(file_input)
    word_file.SaveAs(file_output,FileFormat=format_code)
    word_file.Close()

# cerrar la direccion de la aplicación
word_app.Quit()

time_end = time.time()

print(f'Time used for conversion {time_end - time_start}')


































    

