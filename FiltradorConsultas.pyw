from functools import partial
import sys
import pandas as pd
from datetime import datetime
import os
from os import system
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import tkinter as tk
from pathlib import Path

#Determinamos que el archivo del cual vamos a tomar los datos, esta en la misma carpeta que el programa
dir = os.path.abspath(os.path.dirname(__file__))

def corregir_fecha(v):
    entry = list(v.split(" "))
    año = entry[-1]
    mes_nom = entry[-3]
    mes = str(es_mes(mes_nom))
    dia = entry[-5]
    fecha_str = dia + "/" + mes + "/"+ año
    fecha = datetime.strptime(fecha_str, '%d/%m/%Y').date()
    return fecha

def es_mes(pal):
    meses = [" ","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
    for i in range(largo(meses)):
        if meses[i] == pal:
            return i

def largo(v):
    i = 0
    for j in v:
        i += 1
    return i


def corregir_ortografia(v):
    v = list(v)
    final = []
    dicc = "Ã¡±³©º‰“"
    correc = ["á","ñ","ó","é","ú","É","Ó","Í"]
    for i in range(largo(v)):
        if v[i] == "Ã":
            continue
        elif  v[i] not in dicc and v[i-1] == "Ã":
            if v[i] == '':
                final.append("Á")
            else:
                final.append("í")
            continue
        elif v[i] in dicc and v[i-1] == "Ã":
            error = "Ã" + v[i]
            pos = buscar_pos(error)
            agregar = correc[pos]
            final.append(agregar)
        else:
            final.append(v[i])
    final = "".join(final)
    return final

def buscar_pos(e):
    error = ["Ã¡","Ã±","Ã³","Ã©","Ãº","Ã‰","Ã“","Ã"]
    for i in range(largo(error)):
        if error[i] == e:
            return i

def recorrer_fila(fila):
    for i in range(largo(fila)):
        fila[i] = corregir_ortografia(str(fila[i]))
    return fila

def corregir_cba(valor):
    if valor == "cordoba" or valor == "córdoba" or valor == "Cordoba" or valor == "CORDOBA":
        valor = "Córdoba"
    return valor

def corregir_lnd(valor):
    v = list(valor)
    final = v[8:len(v)+1]
    #final =[]
    #malos = ["0","1","2","3","4","5","6","7"]
    #for i in range(largo(v)):
    #    if str(i) not in malos:
    #        final.append(v[i])
    final = "".join(final)
    return final
        

def llenar_na(valor):
    if valor == "nan":
        valor = " "
    return valor

def remover_dupl(df):
    keep_names = set()
    keep_mails = set()
    keep_rows = list()
    for irow, name in enumerate(df["Nombre"]):
        mail = df.iloc[irow,8]
        if name not in keep_names and mail not in keep_mails:
            keep_names.add(name)
            keep_mails.add(mail)
            keep_rows.append(irow)
    return df.iloc[keep_rows, :]

def principal_pi():
    try:
        df = pd.read_excel(ruta['text'],sheet_name="Hoja1")
        #Corregimos los errores ortograficos
        df = df.apply(recorrer_fila, axis=1)
        #Le damos el formato deseado a la fecha
        df["Fecha"] = df["Fecha"].apply(corregir_fecha)
        #Ordenamos y redefinimos columnas
        df = df[["Fecha","Nombre Completo","Nacionalidad","Localidad",df.columns[-2],"Desde",df.columns[-1],"Teléfono","E-mail"]]
        df.columns = ["Fecha","Nombre","Pais","Localidad","Tipo de servicio","Desde","Mensaje","Teléfono","E-mail"]
        #Llenamos los mensajes "nan" por " "
        df["Mensaje"] = df["Mensaje"].apply(llenar_na)
        #Corregimos Valores de Córdoba
        df["Localidad"] = df["Localidad"].apply(corregir_cba)
        df = df.sort_values(by=["Fecha"])
        filtrar(df)
        messagebox.showinfo("Filtrado De Consultas", "FINALIZADO CON EXITO!! :)\n\n Los archivos se guardaron en: \n"+ruta['text'])
        sys.exit()
    except Exception as e:
        messagebox.showerror("Filtrado De Consultas", e)
        

def filtrar(df_orig):
    #Dividimos el archivo en los dataframe que correspondan
    args = df_orig[df_orig["Pais"] == "Argentina"]
    df_inter = df_orig[df_orig["Pais"] != "Argentina"]
    df_nacionales = args[args["Desde"] == "Landing TELEMEDICINA ARG"]
    df_nacionales_ac = args[args["Desde"] != "Landing TELEMEDICINA ARG"]
    #Reordenamos el archivo de nacionales alta complejidad
    df_nacionales_ac = df_nacionales_ac[["Fecha","Nombre","Pais","Localidad","Desde","Tipo de servicio","Mensaje","Teléfono","E-mail"]]
    df_nacionales_ac["Desde"] = df_nacionales_ac["Desde"].apply(corregir_lnd)
    #Eliminamos los duplicados de cada dataframe
    df_inter = remover_dupl(df_inter)
    df_nacionales = remover_dupl(df_nacionales)
    df_nacionales_ac = remover_dupl(df_nacionales_ac)
    #Guardamos
    with pd.ExcelWriter(ruta['text']) as writer:
        df_inter.to_excel(writer,index = False,sheet_name="Internacionales")
        df_nacionales.to_excel(writer,index = False, sheet_name="Nacionales")
        df_nacionales_ac.to_excel(writer,index = False,sheet_name="Nacionales AC")
        writer.sheets["Nacionales AC"].column_dimensions["A"].width = 20
        writer.sheets["Internacionales"].column_dimensions["A"].width = 20
        writer.sheets["Nacionales"].column_dimensions["A"].width = 20
        writer.save()
        writer.close()
    return
    
def abrirArchivo(var):
    ruta = filedialog.askopenfilename(initialdir = os.path.abspath(os.path.dirname(__file__)))
    var['text'] = ruta

if __name__ == "__main__":
    raiz = Tk()
    raiz.title("Filtrado De Consultas")
    raiz.geometry("350x300")
    Label(raiz, text= "BIENVENIDO").pack(side = TOP)
    ruta = Label(raiz,text = "Ubicacion del archivo Excel donde se encuentran los datos")
    ruta.place(x=10,y=50)
    rutas = Label(raiz,text = " ")
    rutas.place(x=10,y=150)
    ttk.Button(raiz, text = "Seleccionar Archivo Excel", command = partial(abrirArchivo,ruta)).place(x=10,y=70)
    ttk.Button(raiz, text = "INICIAR FILTRADO", command = principal_pi).pack(side= BOTTOM)
    raiz.mainloop()