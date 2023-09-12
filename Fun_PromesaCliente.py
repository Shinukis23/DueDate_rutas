from datetime import datetime, timedelta
import pandas as pd

USAstores = [1,2,3,5,6,7,8]
DesarmeTJ = [4,14]
Economy = 10

def fechasCorte1(date):
 # Convierte argumento de entrada a fecha, y calculo fechas de cortes
 datey = date - timedelta(days=1)
 datet = date + timedelta(days=1)
 dateS = date - timedelta(days=2)
 dateF = date - timedelta(days=3)
 
 # Arreglo de fechas de cortes de tiendas Tapatio
 cortes = [ datey + timedelta(hours = 16)+ timedelta(minutes = 1),
            date + timedelta(hours = 16),
            datey + timedelta(hours = 13)+ timedelta(minutes = 1),
            date + timedelta(hours = 13),
            datey + timedelta(hours = 12)+ timedelta(minutes = 31),
            date + timedelta(hours = 12) + timedelta(minutes= 30),
            datey + timedelta(hours = 14)+ timedelta(minutes = 1),
            date + timedelta(hours = 14),
            dateS + timedelta(hours = 13)+ timedelta(minutes = 1),
            dateF + timedelta(hours = 12)+ timedelta(minutes = 31),
            dateS + timedelta(hours = 14)+ timedelta(minutes = 1),
            date + timedelta(hours = 23)+ timedelta(minutes = 59)] # Sabado todo el dia 23:59

 cortes = pd.to_datetime(cortes)
 return cortes,datet
    
def timeFix(col,h,ds):
 # Sumo o resto horas para arreglar hora de reporte de Produccion
 for i in range(len(col)):
     n = int(col[i])
     ds.iloc[:, n] = pd.to_datetime(ds.iloc[:,n])+ timedelta(hours=h)
 ds3 = ds.copy()
 return ds3




def borraColumnas(col,value,ds2):
    # Borro los registros que no se ocupan del reporte de produccion
    for i in range(len(value)):
     indexDeleted = ds2[ds2[col] == value[i]].index
     ds2.drop(indexDeleted,inplace=True)
    indexDeleted = ds2[ds2['Part Price'] <  0].index #borro regresos de dinero (retornos)
    ds2.drop(indexDeleted,inplace=True)


def rangoFechas(df,date):
    #print(date)
    #print(df['Created'].min())
    #print(df['Created'].max())
    if date <= df['Created'].min() or date >= df['Created'].max():
        print("La fecha no se encuentra en el rango de fechas del archivo")
        print("[",df['Created'].min()," - ", df['Created'].max(),"]")
        exit()
    
# Selecciona los horarios de corte de acuerdo a la tienda    
"""def horaCortes(store,cortes,date):
    if date.weekday() == 0:
     if store in range(1,3):
       cuts = [cortes[8],cortes[1]]
     elif store in range(6,9):
       cuts = [cortes[8],cortes[1]]
     else :
         if (store == 14 or  store == 4):
             cuts = [cortes[9],cortes[5]]
         elif store == 10:
             cuts = [cortes[10],cortes[7]]
    elif date.weekday() == 5:
     if store in range(1,3):
       cuts = [cortes[0],cortes[3]]
     elif store in range(6,9):
       cuts = [cortes[0],cortes[3]]
     else :
         if (store == 14 or  store == 4):
             cuts = [cortes[6],cortes[7]]
         elif store == 10:
             cuts = [cortes[6],cortes[7]]
    else :
     if store in range(1,3):
        cuts = [cortes[0],cortes[1]]
     elif store in range(6,9):
        cuts = [cortes[0],cortes[1]]
     else :
          if (store == 14 or  store == 4):
              cuts = [cortes[4],cortes[5]]
          elif store == 10:
              cuts = [cortes[6],cortes[7]]
    return cuts"""
# Selecciona los horarios de corte de acuerdo a la tienda    
def horaCortes(store,cortes,date):
    if date.weekday() == 0:  # (0 lunes) (1 martes) (2 miercoles) (3 jueves) (4 vienres) (5 sabado) 
        if store in USAstores: #range(1,3) or store in range(6,9):
            cuts = [cortes[8],cortes[1]]
        elif store in DesarmeTJ:#(store == 14 or  store == 4):
            cuts = [cortes[9],cortes[5]]
        elif store == Economy:
            cuts = [cortes[10],cortes[7]]
    elif date.weekday() == 5:      # 5 SABADO
        if store in USAstores: # range(1,3) or store in range(6,9):
            cuts = [cortes[0],cortes[3]]
        elif store in DesarmeTJ:#  (store == 14 or  store == 4):  
            cuts = [cortes[4],cortes[11]]  
        elif store == Economy:
            cuts = [cortes[6],cortes[7]]
    else :
        if store in USAstores:#range(1,3) or store in range(6,9):
            cuts = [cortes[0],cortes[1]]
        elif store in DesarmeTJ:#s(store == 14 or  store == 4):
            cuts = [cortes[4],cortes[5]]
        elif store == Economy:
            cuts = [cortes[6],cortes[7]]
    return cuts


#busca los trabajos de cada tienda de acuerdo a los horarios de corte
def trabajos1(ds2,store,i,cortes,date,Jobs,JobStore,hojas,datet):
    cut = horaCortes(store,cortes,date)
    Jobs.append(str(datetime.date(date)))
    Jobs.append(hojas[i])
    Jobs.append(len(ds2[(ds2['Part Store #'] == store) & (ds2['Created'] >= cut[0]) & (ds2['Created'] <= cut[1])] ))
    Jobs.append(len(ds2[(ds2['Part Store #'] == store) & (ds2['Created'] >= cut[0]) & (ds2['Created'] <= cut[1]) & (ds2['Pulled Finished'] < datet)] ))
    Jobs.append(len(ds2[(ds2['Part Store #'] == store) & (ds2['Created'] >= cut[0]) & (ds2['Created'] <= cut[1]) & (ds2['Job Status'] == 'Pulling Part')] ))
    Jobs.append(len(ds2[(ds2['Part Store #'] == store) & (ds2['Created'] >= cut[0]) & (ds2['Created'] <= cut[1]) & (ds2['Job Status'] == 'Unassigned')] )) 
    Jobs.append(len(ds2[(ds2['Part Store #'] == store) & (ds2['Created'] >= cut[1]) & (ds2['Created'] < datet)] ))
    Jobs.append(len(ds2[(ds2['Part Store #'] == store) & (ds2['Created'] >= cut[1]) & (ds2['Created'] < datet) & (ds2['Pulled Finished'] < datet)] ))
    JobStore[i].extend(Jobs)




