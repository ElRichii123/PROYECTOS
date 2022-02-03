#from nis import cat
import os
from numpy import nan
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import tkinter as tk
from datetime import datetime
from functools import partial
import sys

now = datetime.now()
months = ("Enero", "Febrero", "Marzo", "Abri", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
mes = months[now.month - 1]
mes_pasado = months[now.month - 2]

def corregir_dni(cuils):
    dnis = []
    for i in cuils:
        cadena = list(str(i))
        del cadena[0]
        del cadena[0]
        del cadena[len(cadena)-1]
        cadena = ''.join(cadena)
        dnis.append(cadena)
    return dnis

def generar_primertipo(base, valor,nombre):
    if valor == 1951:
        df = base[base["empresa"] == 1951]
        df.to_excel(ruta_Destino["text"]+"/"+nombre+".xlsx", index=False,sheet_name = "base")
    else:
        df = base[base["empresa"] == 703]
        df = df[df["convenio"] == valor]
        df.to_excel(ruta_Destino["text"]+"/"+nombre+".xlsx", index=False, sheet_name = "base")
    return

def obtenerArchivoPasado(area):
    if str(mes_pasado) == "Diciembre":
        #Estamos en el siguiente año
        return dir_actual + "/"+area+str(mes_pasado)+str(now.year-1)+".xlsx"
    else:
        return dir_actual + "/"+area+str(mes_pasado)+str(now.year)+".xlsx"

def generar_rrhh(base):
    path = pd.ExcelWriter(ruta_Destino["text"]+"/"+"RRHH-"+str(mes)+str(now.year)+".xlsx")
    df = base[base["empresa"] == 703]
    valores = [4,12]
    df = df[df.convenio.isin(valores)]
    df_pasado = pd.read_excel(obtenerArchivoPasado("RRHH-"),sheet_name="resumen")
    df_aux = df[df["CUIL"].isin(df_pasado["CUIL"])]
    df_resumen = pd.DataFrame(columns=df_pasado.columns)
    df_resumen["CUIL"], df_resumen["DNI"], df_resumen["NOMBRE"], df_resumen["MONTO"] = df_aux["CUIL"], corregir_dni(df_aux["CUIL"]), df_aux["NOMBRE"], df_aux["TOTAL"]
    total = sum(df_resumen["MONTO"])
    df_resumen = df_resumen.append({"CUIL": " ", "DNI": " ", "NOMBRE": "TOTAL","MONTO":total},ignore_index = True)
    df_resumen["MONTO"] = df_resumen["MONTO"].round(2)
    df.to_excel(path, index=False,sheet_name ="base")
    df_resumen.to_excel(path, index=False,sheet_name ="resumen")
    path.save()
    path.close()
    return

def generar_jerq(base):
    path = pd.ExcelWriter(ruta_Destino["text"]+"/"+"Jerarquicos-"+str(mes)+str(now.year)+".xlsx")
    df = base[base["empresa"] == 703]
    df = df[df["convenio"] == 5]
    
    df_pasado = pd.read_excel(obtenerArchivoPasado("Jerarquicos-"),sheet_name="resumen")
    df_resumen = pd.DataFrame(columns=df_pasado.columns)
    df_resumen["legajo o DNI a descontar"], df_resumen["Afiliado"] = df_pasado["legajo o DNI a descontar"], df_pasado["Afiliado"]
    df_resumen["monto individual"] = obtener_mtos_new(df,df_pasado)
    df_resumen["monto total"] = calcular_mtos_new(df_resumen)
    df_resumen["monto individual"] = df_resumen["monto individual"].round(2)
    df_resumen["monto total"] = df_resumen["monto total"].round(2)
    df.to_excel(path, index=False,sheet_name ="base")
    df_resumen.to_excel(path, index=False,sheet_name ="resumen")
    path.save()
    path.close()
    return

def calcular_mtos_new(df):
    legajos = []
    mtos_new = []
    for i, fila in df.iterrows():
        if fila["legajo o DNI a descontar"] not in legajos:
            mtos_new.append(fila["monto individual"])
        else:
            suma = fila["monto individual"] + mtos_new[len(mtos_new)-1]
            mtos_new.append(suma)
            mtos_new[len(mtos_new) - 2] = 0
        legajos.append(fila["legajo o DNI a descontar"])
    total = sum(mtos_new) - df.loc[df.index[-1], "monto individual"]
    mtos_new[len(mtos_new)-1] = total
    
    return mtos_new
 
def obtener_mtos_new(df,resumen):
    mtos_new = []
    for i in resumen["Afiliado"]:
        aux = df[df["NOMBRE"] == i]
        a = aux["TOTAL"].tolist()
        if len(a) == 0:
            mtos_new.append(0)
            continue
        mtos_new.append(a[0])
    total = sum(mtos_new)
    mtos_new[len(mtos_new)-1] = total
    return mtos_new

def generar_segundotipo(base,var):
    if var == 1:
        path = pd.ExcelWriter(ruta_Destino["text"]+"/"+"Medicos-"+str(mes)+str(now.year)+".xlsx")
        df = base[base["empresa"] == 703]
        df = df[df["convenio"] == 1]
        df_pasado = pd.read_excel(obtenerArchivoPasado("Medicos-"),sheet_name="resumen")
    else:
        path = pd.ExcelWriter(ruta_Destino["text"]+"/"+"Empleados-"+str(mes)+str(now.year)+".xlsx")
        df = base[base["empresa"] == 703]
        df = df[(df["convenio"] == 7) | (df["convenio"] == 11)]
        df_pasado = pd.read_excel(obtenerArchivoPasado("Empleados-"),sheet_name="resumen")
    #Obtenemos los valores para la hoja resumen
    ctas_new, mtos_new = calcular_mtos_new_2(df)
    ctas_new.append("Total general")
    mtos_new.append(sum(mtos_new))
    df_resumen = pd.DataFrame(data={"Cuenta Actual": ctas_new,"Monto Actual":mtos_new})
    #Corregimos las cuentas considerando bajas y altas y lo guardamos en la hoja comparativa
    ctas_old, mtos_old, ctas_new, mtos_new, com = corregir_ctas(df_pasado["Cuenta Actual"], df_pasado["Monto Actual"], ctas_new, mtos_new)
    df_cmp = pd.DataFrame(data={"Cuenta Anterior":ctas_old,"Monto Anterior":mtos_old,"Diferencias": " " ,"Cuenta Actual": ctas_new,"Monto Actual":mtos_new,"Comentarios":com})
    df_cmp["Diferencias"] = df_cmp["Monto Actual"] - (df_cmp["Monto Anterior"] * ((globals()["aumento"] / 100) + 1))
    df_cmp["Diferencias"] = df_cmp["Diferencias"].round(2)
    #Buscar locale currency para tener string en moneda
    df_cmp["Monto Actual"] = df_cmp["Monto Actual"].round(2)
    df_cmp["Monto Anterior"] = df_cmp["Monto Anterior"].round(2)
    df_resumen["Monto Actual"] = df_resumen["Monto Actual"].round(2)
    df.to_excel(path, index=False,sheet_name ="base")
    df_cmp.to_excel(path, index=False,sheet_name ="comparativa")
    df_resumen.to_excel(path, index=False,sheet_name ="resumen")
    path.save()
    path.close()
    return   

def corregir_ctas(a,b,c,d):
    Cta_Old = list(a)
    Mto_Old = list(b)
    Cta_New = list(c)
    Mto_New = list(d)
    Com = [" "] * len(Mto_New)
    for i in range(len(Cta_New)):
        if  Cta_Old[i] not in Cta_New :
            #Baja de cuenta
            Cta_New.insert(i, Cta_Old[i])
            Mto_New.insert(i, 0)
            Com.insert(i,"BAJA CUENTA")
        elif Cta_New[i] not in Cta_Old:
            #Alta de cuenta
            Cta_Old.insert(i, Cta_New[i])
            Mto_Old.insert(i, 0)
            Com[i] = ("ALTA CUENTA")
    #df = pd.DataFrame(data={"Cuenta Anterior":Cta_Old,"Monto Anterior":Mto_Old,"Diferencias":" ","Cuenta Actual":Cta_New,"Monto Actual": Mto_New})
    return Cta_Old,Mto_Old,Cta_New,Mto_New,Com

def convertirMonFloat(num):
    num = str(num)
    num = list(num)
    list(num).pop(0)
    return "".join(num)

def calcular_mtos_new_2(df):
    legajos = []
    mtos_new = []
    for i, fila in df.iterrows():
        if fila["Legajo"] not in legajos:
            legajos.append(int(fila["Legajo"]))
            mtos_new.append(fila["TOTAL"])
        else:
            suma = fila["TOTAL"] + mtos_new[len(mtos_new)-1]
            mtos_new[len(mtos_new) - 1] = suma
    return legajos,mtos_new



def actualizar_dir_actual():
    global dir_actual
    aux = ruta_Base["text"].split("/")
    aux.pop(-1)
    dir_actual = "/".join(aux)


def abrirArchivo(var):
    ruta = filedialog.askopenfilename(initialdir = dir_actual)
    var['text'] = ruta
    aux = ruta_Base["text"].split("/")
    aux.pop(-1)
    rutas["text"]  = "/".join(aux)

def abrirCarpeta(var):
    ruta = filedialog.askdirectory(initialdir = dir_actual)
    var['text'] = ruta

def main():
    actualizar_dir_actual()
    try:
        df_original = pd.read_excel(ruta_Base["text"],sheet_name="base_HPC___mutual_2")
        df_original = df_original.sort_values("Legajo",na_position="first")
        generar_primertipo(df_original,8, "canjes-"+str(mes)+str(now.year)) #Generar archivo canjes
        generar_primertipo(df_original, 1951, "MutualMedicos-"+str(mes)+str(now.year)) #Generar archivo Mutual
        generar_primertipo(df_original, 9, "Residentes-"+str(mes)+str(now.year))#Generar archivo Residentes
        generar_rrhh(df_original)#Generar archivo RRHH
        generar_jerq(df_original)#Generar archivo Jerarquicos
        generar_segundotipo(df_original,7)#Generar archivo Empleados
        generar_segundotipo(df_original,1)#Generar archivo Medicos
        messagebox.showinfo("Descuentos", "Archivos generados")
        sys.exit()
    except Exception as e:
        messagebox.showerror("Descuentos", e)
    

def abrirAumento():
    newWindow = tk.Toplevel(raiz)
    newWindow.geometry("500x100")
    tk.Label(newWindow, text="Ingrese el aumento en numero entero(Ej: 10%-> 10)").grid(row=0)
    e1 = tk.Entry(newWindow,font=("Calibri",12))
    e1.grid(row=0, column=1)
    ttk.Button(newWindow,text="Confimar", command=partial(confirmarEntry,e1,newWindow)).place(x=200,y=50).grid(row=3,column=1,sticky=tk.W,pady=4)


def confirmarEntry(e,ventana):
    globals()["aumento"] = int(e.get())
    ventana.state(newstate="withdraw")

if __name__== "__main__":
    raiz = Tk()
    raiz.title("Descuentos OSPE")
    raiz.geometry("350x300")
    dir_actual = os.path.abspath(os.path.dirname(__file__))
    Label(raiz, text= "BIENVENIDO").pack(side = TOP)
    ruta_Destino = Label(raiz,text = "Ubicacion de la carpeta donde desea guardar los archivos")
    ruta_Destino.place(x=10,y=50)
    ruta_Base = Label(raiz,text = "Ubicacion del archivo base y los archivos del mes pasado ")
    ruta_Base.place(x=10,y=100)
    Label(raiz, text= "Los archivos del mes pasado estan en: ").place(x=10,y=125)
    rutas = Label(raiz,text = " ")
    rutas.place(x=10,y=150)
    aumento = 0
    ttk.Button(raiz, text = "Seleccionar Carpeta de destino", command = partial(abrirCarpeta,ruta_Destino)).place(x=10,y=20)
    ttk.Button(raiz, text = "Seleccionar Archivo Base", command = partial(abrirArchivo,ruta_Base)).place(x=10,y=70)
    ttk.Button(raiz,text="Hubo aumento?", command=abrirAumento).place(x=10,y = 275)
    ttk.Button(raiz, text = "Generar archivos", command = main).pack(side= BOTTOM)
    raiz.mainloop()
    
