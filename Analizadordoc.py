#   Se instala el DocxTemplate con pip install docxtpl
#André Vicente Canul Abán
from docxtpl import DocxTemplate #importamos la libreria
import os
import comtypes.client
import time
import urllib.parse
import requests
import json
url = 'http://localhost:3000/api/tokens'
format_code = 17
time_start = time.time()
data= requests.get(url)
if data.status_code == 200:
    data = data.json()
    for e in data:
  #nUESTRO dICCCIONARIO      
        context = {
        'id_usuario':e['id_usuario'] ,
        "F2_Municipio" : e[ "F2_Municipio"],
        "Cata_Nombre": e["Cata_Nombre"],
        "Cata_NotoEsc": e[ "Cata_NotoEsc"],
        "Cata_No.": e["Cata_No."],
        "Cata_Domicilio": e["Cata_Domicilio"],
        "RPP_Escritura":e["RPP_Escritura"],
        "RPP_Numero": e["RPP_Numero"],
        "RPP_Fecha": e["RPP_Fecha"],
        "RPP_No.Inscripcion": e["RPP_No.Inscripcion"],
        "RPP_FolioElectronico":e["RPP_FolioElectronico"],
        "Ubi_Clase_Predio": e[ "Ubi_Clase_Predio"],
        "Ubi_Calle": e["Ubi_Calle"],
        "Ubi_Numero": e["Ubi_Numero"],
        "Ubi_Colonia": e["Ubi_Colonia"],
        "Ubi_Localidad": e["Ubi_Localidad"],
        "Ubi_Municipio": e["Ubi_Municipio"],
        "Objeto_Man_Descripcion": e[ "Objeto_Man_Descripcion"],
        "Pro_Ant_Nombre": e[ "Pro_Ant_Nombre"],
        "Pro_Ant_Curp": e["Pro_Ant_Curp"],
        "Pro_Ant_RFC": e["Pro_Ant_RFC"],
        "Pro_Ant_Domicilio":e["Pro_Ant_Domicilio"],
        "Pro_Act_Nombre": e["Pro_Act_Nombre"],
        "Pro_Act_Curp": e["Pro_Act_Curp"],
        "Pro_Act_RFC": e["Pro_Act_RFC"],
        "Pro_Act_Domicilio": e["Pro_Act_Domicilio"],
        "Dat_Cata_Areas": e["Dat_Cata_Areas"],
        "Dat_Cata_Aerreno": e["Dat_Cata_Aerreno"],
        "Dat_Cata_Aonstruccion":e["Dat_Cata_Aonstruccion"],
        "Dat_Cata_Apredio": e["Dat_Cata_Apredio"],
        "Dat_Cata_Acatastral": e["Dat_Cata_Acatastral"],
        "Dat_Cata_Defecha": e["Dat_Cata_Defecha"],
        "Dat_Cata_fecha": e["Dat_Cata_fecha"],
        "id_plantilla": e["id_plantilla"],
        "Credito_Folio": e["Credito_Folio"],
        "Credito_CANDESOL": e["Credito_CANDESOL"],
        "Credito_Plazo": e["Credito_Plazo"],
        "Credito_TasaComision": e["Credito_TasaComision"],
        "FechaSubscripcion": e["FechaSubscripcion"],
        "Credito_TipoPlazo": e["Credito_TipoPlazo"],
        "Credito_MontoCredito": e["Credito_MontoCredito"],
        "Pagare_FechaVencimientoPrimero": e["Pagare_FechaVencimientoPrimero"],
        "Cliente_ApellidoPaterno": e["Cliente_ApellidoPaterno"],
        "Cliente_ApellidoMaterno": e["Cliente_ApellidoMaterno"],
        "Cliente_FechaNacimiento": e["Cliente_FechaNacimiento"],
        "Cliente_PaisNacimiento": e["Cliente_PaisNacimiento"],
        "Cliente_LugarNacimiento": e["Cliente_LugarNacimiento"],
        "Cliente_DirPais": e["Cliente_DirPais"],
        "Cliente_Nacionalidad":e["Cliente_Nacionalidad"],
        "Cliente_Sexo":e["Cliente_Sexo"],
        "Cliente_RFC": e["Cliente_RFC"],
        "Empleo_Ocupacion": e["Empleo_Ocupacion"],
        "Cliente_ActividadEconomica":e["Cliente_ActividadEconomica"],
        "Cliente_TelCelular": e["Cliente_TelCelular"],
        "Cliente_TelParticular": e["Cliente_TelParticular"],
        "Cliente_eCorreo": e["Cliente_eCorreo"],
        "Cliente_CURP": e["Cliente_CURP"],
        "Cliente_NumeroSerieFiel":e["Cliente_NumeroSerieFiel"],
        "Cliente_EdoCivil": e["Cliente_EdoCivil"],
        "Cliente_NumDependientesEco": e["Cliente_NumDependientesEco"],
        "Bancario_CLABE": e["Bancario_CLABE"],
        "Bancario_RazonSocial":e["Bancario_RazonSocial"],
        "Cliente_DirCalles": e["Cliente_DirCalles"],
        "Cliente_DirNumero": e["Cliente_DirNumero"],
        "Cliente_DirNumeroInt":e["Cliente_DirNumeroInt"],
        "Cliente_DirColonia": e["Cliente_DirColonia"],
        "Cliente_DirMunicipio": e["Cliente_DirMunicipio"],
        "Cliente_DirCiudad": e["Cliente_DirCiudad"],
        "Cliente_DirEstado": e["Cliente_DirEstado"],
        "Cliente_DirCP": e["Cliente_DirCP"],
        "Empleo_EmpresaNombre": e["Empleo_EmpresaNombre"],
        "Empleo_Puesto": e["Empleo_Puesto"],
        "Cliente_TotalIngresos": e[ "Cliente_TotalIngresos"],
        "Informacion_TotalGastos":e[ "Informacion_TotalGastos"],
        "ISAI_Enc_Mun": e[ "ISAI_Enc_Mun"],
        "ISAI_Articulo": e["ISAI_Articulo"],
        "ISAI_Fundamentos": e["ISAI_Fundamentos"],
        "ISAI_Acta": e["ISAI_Acta"],
        "ISAI_Fecha_Escritura": e["ISAI_Fecha_Escritura"],
        "ISAI_Concepto_Adg": e["ISAI_Concepto_Adg"],
        "Enaj_Nombre_Completo": e["Enaj_Nombre_Completo"],
        "Enaj_Domicilio_Notificaciones": e["Enaj_Domicilio_Notificaciones"],
        "Enaj_Ciudad_Completo": e["Enaj_Ciudad_Completo"],
        "Enaj_C.P_Completo": e["Enaj_C.P_Completo"],
        "Enaj_Curp_Completo": e["Enaj_Curp_Completo"],
        "Enaj_Rfc_Completo": e["Enaj_Rfc_Completo"],
        "Adqu_Ciudad_Completo": e["Adqu_Ciudad_Completo"],
        "Adqu_C.P_Completo": e["Adqu_C.P_Completo"],
        "Adqu_Curp_Completo": e["Adqu_Curp_Completo"],
        "Adqu_Rfc_Completo": e["Adqu_Rfc_Completo"],
        "Inmu_Valor_Operacion":e["Inmu_Valor_Operacion"],
        "Inmu_Avaluo_Aperacion": e["Inmu_Avaluo_Aperacion"],
        "Inmu_Valor_Cperacion": e[ "Inmu_Valor_Cperacion"],
        "Inmu_Impuesto_Pperacion": e["Inmu_Impuesto_Pperacion"]  
            }

# Creamos el MS word app
word_app = comtypes.client.CreateObject('Word.Application')
word_app.Visible = False
#tpl = DocxTemplate("Ejemplo solicitud de crédito HACKATHON.docx")
#tpl = DocxTemplate("F2 FORMATO CUALQUIER MUNICIPIO CON CATASTRO HACKATHON.docx")
tpl = DocxTemplate("DocAPI\FORMATOS ISAI HACKATHON 3.docx") #plantilla a utilizar
 #partes a reemplazar de la plantilla la izquierda es lo tokens y la derecha es la información por la cual se reemplaza por {{}}
copias = 1

for i in range(0, copias, 1): # numero de copias del documento introducido
    #print(context)   
    tpl.render(context) # El render introduce el valor
    tpl.save("Documentogen"+str(i+1)+".docx") # Nombre y formato a guardar
   
# conversion
for i in range(0, copias, 1): # numero de copias del documento introducido
    print(i)
    file_input = os.path.abspath('Documentogen'+str(i+1)+'.docx')
    file_output = os.path.abspath('Documentogen'+str(i+1)+'.pdf')
    word_file = word_app.Documents.Open(file_input)
    word_file.SaveAs(file_output,FileFormat=format_code)
    word_file.Close()

# cerrar la direccion de la aplicación
word_app.Quit()

time_end = time.time()

print(f'Time used for conversion {time_end - time_start}')


































    

