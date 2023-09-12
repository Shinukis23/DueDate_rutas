###Programa Funcional OK Septiembre 11/2023
# Programa para calcular due-date y dias de atraso para reporte DeliveryPickup 

import pandas as pd
import numpy as np
import xlwt
import openpyxl
import time
import xlsxwriter
import gspread
import pygsheets
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime, timedelta
from Fun_PromesaCliente import *
import io
import os
from urllib.parse import urlparse
from google.oauth2 import service_account
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

import gspread
import sys
import warnings
import gspread_dataframe as gd
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)



if len(sys.argv) != 3  :
    print("Uso: ProCliente.py (Archivo.csv) ((-) Horas)")
    print("Archivo.csv        Es el archivo que se quiere modificar, puede incluir ruta")
    print("(-) Horas          Cantidad de horas que se quieren restar/sumar")
    print(len(sys.argv))
    quit()
else:
    print("Verificando contenido del archivo .....")


file_id0 = "1F0L_aHVNNhGuV-KNnuT6nCr_X1Af3l3E" #cortes2023.xlxs
file_id1 = "15vHlzGFgi9MjxyclqmNArvheijJhLSK5" #tiempos.xls

scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]

credentials = ServiceAccountCredentials.from_json_keyfile_name("monitor-eficiencia-3a13458926a2.json", scopes) #access the json key you downloaded earlier 
file = gspread.authorize(credentials)# authenticate the JSON key with gspread

nombre = sys.argv[1]

try:
    hora = int(sys.argv[2])
except:
   print("Uso: Incluya las horas que desea sumar o restar")
   quit()

ds = pd.read_csv(nombre)
####columnas = sys.argv[2]
####columnas = columnas.split(',')
columnas=[11,12,18,19,20]
nombre1 = nombre.split('.')
nombre2 = nombre.split('_')
#print(nombre2)
nombre1 = nombre1[0] + ".xlsx"

# Define the Drive API client
service = build("drive", "v3", credentials=credentials)

# Define the URL to download the file from
file_url = service.files().get(fileId=file_id0, fields="webContentLink").execute()["webContentLink"]
parsed_url = urlparse(file_url)

# Define the filename to save the downloaded file as
filename = f"cortes2023.xlsx"

# Download the file
try:
    request = service.files().get_media(fileId=file_id0)
    file = io.BytesIO()
    downloader = io.BytesIO()
    downloader.write(request.execute())
    downloader.seek(0)
    with open(filename, "wb") as f:
        f.write(downloader.getbuffer())
    print(f"File downloaded as {filename}")
except HttpError as error:
    print(f"An error occurred: {error}")

# Define the URL to download the file from
file_url = service.files().get(fileId=file_id1, fields="webContentLink").execute()["webContentLink"]
parsed_url = urlparse(file_url)

# Define the filename to save the downloaded file as
filename = f"tiempos.xls"

# Download the file
try:
    request = service.files().get_media(fileId=file_id1)
    file = io.BytesIO()
    downloader = io.BytesIO()
    downloader.write(request.execute())
    downloader.seek(0)
    with open(filename, "wb") as f:
        f.write(downloader.getbuffer())
    print(f"File downloaded as {filename}")
except HttpError as error:
    print(f"An error occurred: {error}")    


df = pd.read_excel(r'Tiempos.xls')
dc = pd.read_excel(r'Cortes2023.xlsx')
#dp = pd.read_excel(r'periodos.xlsx')

df1 = pd.DataFrame({
    "Unnamed: 0": "FALSE",
    "Job #": '',
    "Order #": '',
    "Type": '',
    "Customer":'',
    "Interchange":'',
    "Store #":'',
    "Stock #":'',
    "Year":'',
    "Model":'',
    "Price":'',
    "Created":'',
    "Due":'',
    "Route":'',
    "Salesperson":'',
    "Driver":'',
    "Event":'',
    "Reason":'',
    "Date":'',
    "Delivery Time":'',
    "Pickup Time":'',
    "Due Date":'',
    "Dias de atraso":'',
    "Conciliacion":'',
}, index=["Dummy"])
																						

#date = pd.to_datetime(sys.argv[4])
# Funcion Main para buscar todos los trabajos

#dc= dc.set_index('DIA')
df= df.set_index('Store')
#dc = dc.to_dict()

dc = dict(dc.set_index('DIA').groupby(level = 0).\
    apply(lambda x : x.to_dict(orient= 'list')))
#print(dc)
ds['Pickup Time'] = pd.to_datetime(ds['Pickup Time'], errors='coerce') 
#print('Pickup time',ds['Pickup Time'])
ds2 = timeFix(columnas,hora,ds)
#dia = date.weekday()
#fechaa =date + timedelta(hours = 16)
#print(fechaa)


ds2['Fecha Compromiso']=" "
ds2['Due Date']=" "
ds2['Dias de atraso']=" "
#print(ds2['Created'])
#print(type(ds2['Created']))
ds2['Created'] = pd.to_datetime(ds2['Created'], errors='coerce')
#ds2['Created'] = pd.to_datetime(ds2['Created'])
ds2['Dia']= ds2['Created'].dt.dayofweek
#ds2['Dia'] = ds2['Created'].dt.day_name()
#print('Dia: =====>',ds2['Dia'])
ds2['Dia'].mask(ds2['Dia'] == 6, 0, inplace=True)

ds2['tiempo'] = pd.to_datetime(ds2['Created']).dt.time
ds2['Fecha'] = pd.to_datetime(ds2['Created']).dt.date
ds2['Conciliacion']=" "

#Asigno el valor de ruta
for i in range(len(ds2)) :
    St =ds2['Store #'][i]
    Rt =ds2['Route'][i]
    #ds2['Fecha Compromiso'][i]=df.at[St,Rt]
    if ( pd.isna(Rt)== False):
        ds2.loc[i,'Fecha Compromiso']=df.at[St,Rt]

    
ds2['Conciliacion'].mask(ds2['Fecha Compromiso'] == 99, 'Revisar', inplace=True)    

#print('Fecha conciliacion', ds2['Conciliacion'])
#print('Created', ds2['Created'])
def tabla(i,tiempo,c,b):
 if ds2['tiempo'][i] < tiempo.time() :
        a=dc.get(ds2['Dia'][i])
        delt = a.get(c)        
        #ds2['Due Date'][i]= pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = delt[0])
        ds2.loc[i, 'Due Date'] = pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = delt[0])
        #print('Dias de atraso: =====>',ds2['Due Date'])
 else :
        a=dc.get(ds2['Dia'][i])
        delt = a.get(b)      
        #ds2['Due Date'][i]= pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = delt[0])
        ds2.loc[i, 'Due Date'] = pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = delt[0])
        #print('Dias de atraso: =====>',ds2['Due Date'])
def tabla1(i,tiempo,c,b):
 if ds2['tiempo'][i] < tiempo.time() :
        #ds2['Due Date'][i]= pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = 0) 
        ds2.loc[i, 'Due Date'] = pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = 0) 
        #print('Dias de atraso: =====>',ds2['Due Date'])    
 else :      
        #ds2['Due Date'][i]= pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = 0)
        ds2.loc[i, 'Due Date'] = pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = 0)
        #print('Dias de atraso: =====>',ds2['Due Date'])
        
tiempo1 = datetime(2022,1,1,12,31,00) # asigno tiempos iniciales para comparar 12:31
tiempo2 = datetime(2022,1,1,13,1,00) # asigno tiempos iniciales para comparar 13:01
tiempo3 = datetime(2022,1,1,14,1,00) # asigno tiempos iniciales para comparar 14:01
tiempo4 = datetime(2022,1,1,16,1,00) # asigno tiempos iniciales para comparar 16:01
tiempo5 = datetime(2022,1,1,17,1,00) # asigno tiempos iniciales para comparar 17:01  Todas las tiendas cierre 
tiempo6 = datetime(2022,1,1,15,1,00) # asigno tiempos iniciales para comparar 15:01 Economy Sabado

for i in range(len(ds2)):
 if ds2['Fecha Compromiso'][i] != 99:
     if ds2['Dia'][i] in range(0,6) :   
      if ds2['Fecha Compromiso'][i] == 1 and ds2['Dia'][i]==5: 
        tabla(i,tiempo2,'1.2','1.3')
      elif ds2['Fecha Compromiso'][i] == 1:
        tabla(i,tiempo4,1,'1.1')     
      if ds2['Fecha Compromiso'][i] == 2 and ds2['Dia'][i]==5: 
        tabla(i,tiempo2,'2.2','2.3')
      elif ds2['Fecha Compromiso'][i] == 2:  
        tabla(i,tiempo4,2,'2.1')
      if ds2['Fecha Compromiso'][i] == 3 and ds2['Dia'][i]==5:
        tabla(i,tiempo2,'3.2','3.3')
      elif ds2['Fecha Compromiso'][i] == 3:
        tabla(i,tiempo4,3,'3.1')      
      if ds2['Fecha Compromiso'][i] == 4: 
        tabla(i,tiempo3,4,'4.1')     
      if ds2['Fecha Compromiso'][i] == 5:    
        tabla(i,tiempo1,5,'5.1')  
      if ds2['Fecha Compromiso'][i] == 6:    
        tabla(i,tiempo4,6,'6.1')         
      if ds2['Fecha Compromiso'][i] == 7:    
        tabla(i,tiempo3,7,'7.1')  
      if ds2['Fecha Compromiso'][i] == 8:    
        tabla(i,tiempo1,8,'8.1')
      if ds2['Fecha Compromiso'][i] == 9:    
        tabla(i,tiempo4,9,'9.1')
      if ds2['Fecha Compromiso'][i] == 10:  
        tabla(i,tiempo4,10,'10.1')
      if ds2['Fecha Compromiso'][i] == 11: 
        tabla(i,tiempo4,11,'11.1') 
      if ds2['Fecha Compromiso'][i] == 12:    
        tabla(i,tiempo1,12,'12.1')
      if ds2['Fecha Compromiso'][i] == 13:
        tabla(i,tiempo4,13,'13.1')
      if ds2['Fecha Compromiso'][i] == 14:    
        tabla(i,tiempo3,14,'14.1')
      if ds2['Fecha Compromiso'][i] == 15:    
        tabla(i,tiempo1,15,'15.1')
      if ds2['Fecha Compromiso'][i] == 16:    
        tabla(i,tiempo4,16,'16.1')  

      if ds2['Fecha Compromiso'][i] == 17 and ds2['Dia'][i]==5: 
        tabla(i,tiempo3,'17.2','17.3')     
      elif ds2['Fecha Compromiso'][i] == 17: 
        tabla(i,tiempo5,17,'17.1') 
      if ds2['Fecha Compromiso'][i] == 18 and ds2['Dia'][i]==5:
        tabla(i,tiempo6,'18.2','18.3')           
      elif ds2['Fecha Compromiso'][i] == 18:     
        tabla(i,tiempo5,18,'18.1')
      if ds2['Fecha Compromiso'][i] == 19 and ds2['Dia'][i]==5: 
        tabla(i,tiempo2,'19.2','19.3')   
      elif ds2['Fecha Compromiso'][i] == 19:   
        tabla(i,tiempo5,19,'19.1') 
      if ds2['Fecha Compromiso'][i] == 20 and ds2['Dia'][i]==5: 
        tabla(i,tiempo3,'20.2','20.3')    
      elif ds2['Fecha Compromiso'][i] == 20:   
        tabla(i,tiempo5,20,'20.1')        
            
 elif ds2['Fecha Compromiso'][i] == 99: 
     tabla1(i,tiempo4,1,'1.1')
   
#print('Dias de atraso: =====>',ds2['Due Date'])     
ds2['Due Date'] = pd.to_datetime(ds2['Due Date'], errors='coerce')     
ds2['Delivery Time'] = pd.to_datetime(ds2['Delivery Time'], errors='coerce')     
ds2['Pickup Time'] = pd.to_datetime(ds2['Pickup Time'], errors='coerce') 
#print('Pickup time',ds2['Pickup Time'])
ds2['Due Date'] = pd.to_datetime(ds2['Due Date']).dt.date
ds2['temp1'] = pd.to_datetime(ds2['Delivery Time']).dt.date
ds2['temp'] = pd.to_datetime(ds2['Pickup Time']).dt.date

ds2['Due Date'] = pd.to_datetime(ds2['Due Date'], errors='coerce')
ds2['Due'] = pd.to_datetime(ds2['Due'], errors='coerce')
#print('Due Date 295:',ds2['Due Date'])    
#print('Due 295:',ds2['Due'])    

for j in range(len(ds2)):
 if pd.isna(ds2.loc[j,'Pickup Time']) is False:
  #ds2['temp1'][j] = ds2['temp'][j]
  ds2.loc[j, 'temp1'] = ds2['temp'][j]
 if ds2['Due'][j] > ds2['Due Date'][j]:
  #ds2['Due Date'][j]= ds2['Due'][j]
  ds2.loc[j, 'Due Date'] = ds2['Due'][j]






#print(ds2['temp1'])
ds2['temp1'] = pd.to_datetime(ds2['temp1'], errors='coerce') 
#print(ds2['temp1'])
#print('Dias de atraso: =====>',ds2['Due Date'])
ds2['Dias de atraso']= ds2['temp1'] - ds2['Due Date']
#print(ds2.head())
ds2.to_excel("Verifica.xlsx")
#print(ds2['Due'])

#print(ds2['Due'])
#print(ds2['Due Date'])
#ds2['Diferencia DueDates']= ds2['Due'] - ds2['Due Date']
ds2['Diferencia DueDates']=  (ds2['Due']- ds2['Due Date']).dt.days 
#print(ds2['Dias de atraso'])
#print(ds2['Diferencia DueDates'])
del ds2["Dia"]
del ds2['tiempo']
del ds2['Fecha']
del ds2['temp']
del ds2['temp1']
del ds2['Fecha Compromiso']
#print(len(ds2['Unnamed: 0'])+1)
#ds2['Unnamed: 0'][len(ds2['Unnamed: 0'])+1] = "Fin"
########################ds2 = pd.concat([ds2, df1], axis=1)
#ds2 = ds2.append(df1)
#ds2.reindex(ds2.columns[ds2.columns != 'Conciliacion'].union(['Conciliacion']), axis=1)

writer = pd.ExcelWriter(nombre1, engine='xlsxwriter')
# Convert the dataframe to an XlsxWriter Excel object.
print("Creando archivo", nombre1)
ds2.to_excel(writer, sheet_name='BD-2022',header=True, index = False)

while True:
    try:
        writer.close()
    except xlsxwriter.exceptions.FileCreateError as e:
        decision = input("Exception caught in workbook.close(): %s\n"
                         "Please close the file if it is open in Excel.\n"
                         "Try to write file again? [Y/n]: " % e)
        if decision != 'n':
            continue

    break

insertRow = ["","","","","","","","","","","","","","","","","","","","","","","","","","","",]
###print("Conectando a google sheets .....")
###ss = file.open("EficienciaReporte")
###print("Conexion exitosa")


###_bdatos_ = ss.worksheet("ABCD123")
#_.values.append(5000)
###print(_bdatos_.row_count)
###print(_bdatos_.col_count)

###data= _bdatos_.get_all_values()
###print("total de datos :",len(data))
###print(ds2.shape[0])
###_bdatos_.add_rows(ds2.shape[0])
#_bdatos_.append_row(insertRow)
###print("Actualizando google sheets .....")
#time.sleep(5)
###gd.set_with_dataframe(worksheet=_bdatos_,dataframe=ds2,include_index=False,include_column_header=False,row=len(data),resize=False)
#gd.set_with_dataframe(worksheet=_bdatos_,dataframe=ds2,include_index=False,include_column_header=False,row=len(data),resize=False)

