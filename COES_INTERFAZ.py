# -*- coding: utf-8 -*-
"""
Created on Tue Dec  7 08:56:20 2021

@author: USER
"""

#INterfaz grafica Perseo 3.0
import tkinter as tk
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.pyplot as plt
from tkinter import ttk
import os
def EVO_MAX(HORA_GLOBAL,F1,F2):
    fig, axes = plt.subplots(ncols=1, nrows=1,figsize=(4.5,3.5))
    PD_TOTAL=pd.DataFrame()
    PD_TOTAL["HORA"]=HORA_GLOBAL
    for T in range(F1,F2+1):
        NEW=[]
        b=os.listdir(os.path.join(os.getcwd(),str(T)))
        #X_BAR=[x[18:-10] for x in b]
        mes=len(b)
        for i in range(mes):
            file_iter=os.path.join(os.getcwd(),str(T),os.listdir(os.path.join(os.getcwd(),str(T)))[i])
            df_iter=pd.read_excel(file_iter)
            df_iter.columns=["NADA","HORA","DEMANDA"]
            df_u=df_iter.loc[range(22,118),["HORA","DEMANDA"]]
            df_u.reset_index(drop=True,inplace=True)
            #df_u[df_u["DEMANDA"]==df_u.max()["DEMANDA"]]
            NEW.append(df_u.max()["DEMANDA"])
        for i in range(mes):
            if NEW[i]==max(NEW):
                file_iter=os.path.join(os.getcwd(),str(T),os.listdir(os.path.join(os.getcwd(),str(T)))[i])
                df_iter=pd.read_excel(file_iter)
                df_iter.columns=["NADA","HORA","DEMANDA"]
                df_u=df_iter.loc[range(22,118),["HORA","DEMANDA"]]
                df_u.reset_index(drop=True,inplace=True)
                PD_TOTAL[str(T)]=df_u["DEMANDA"]
                    #df_u.plot(ax=axes[1],kind="area",x="HORA",y="DEMANDA").set_ylim(0,7000)
            else:
                pass
    PD_TOTAL.plot(ax=axes,x="HORA",stacked=False)
    return fig

def HORA_MAYOR(HORA_GLOBAL,INCIDENCIA,F1,F2):
    fig, axes = plt.subplots(ncols=1, nrows=1,figsize=(4.5,3.5))
    for T in range(F1,F2+1):
        b=os.listdir(os.path.join(os.getcwd(),str(T)))
        #X_BAR=[x[18:-10] for x in b]
        mes=len(b)
        for i in range(mes):
            file_iter=os.path.join(os.getcwd(),str(T),os.listdir(os.path.join(os.getcwd(),str(T)))[i])
            df_iter=pd.read_excel(file_iter)
            df_iter.columns=["NADA","HORA","DEMANDA"]
            df_u=df_iter.loc[range(22,118),["HORA","DEMANDA"]]
            df_u.reset_index(drop=True,inplace=True)
            df_u[df_u["DEMANDA"]==df_u.max()["DEMANDA"]]
            for j in HORA_GLOBAL:
                if j==df_u[df_u["DEMANDA"]==df_u.max()["DEMANDA"]]["HORA"].values[0]:
                    INCIDENCIA[j]=INCIDENCIA[j]+1;
                else:
                    pass
    df_INCIDENCIA = pd.DataFrame([[key, INCIDENCIA[key]] for key in INCIDENCIA.keys()], columns=['HORA', 'CANTIDAD'])
    df_INCIDENCIA.plot(ax=axes,kind="area",x="HORA",y="CANTIDAD")
    return fig

def EVO_HOR():
    F1=int(fecha1.get())
    F2=int(fecha2.get())
    file_path=os.path.join(os.getcwd(),"2015",os.listdir(os.path.join(os.getcwd(),"2015"))[0])
    df=pd.read_excel(file_path)
    df.columns=["NADA","HORA","DEMANDA"]
    fina=df.loc[range(22,118),["HORA","DEMANDA"]]
    fina.reset_index(drop=True,inplace=True)
    HORA_GLOBAL=fina["HORA"].to_list()
    HORA_GLOBAL
    CANTIDAD=[]
    for i in HORA_GLOBAL:
        CANTIDAD.append(0)
        INCIDENCIA=dict(zip(HORA_GLOBAL,CANTIDAD))
    #MAYOR INCIDENCIA
    aux_1=HORA_MAYOR(HORA_GLOBAL,INCIDENCIA,F1,F2)
    aux_2=EVO_MAX(HORA_GLOBAL,F1,F2)
    canvas_2=FigureCanvasTkAgg(aux_2,master=GRA_3)
    canvas_2.draw()
    canvas_2.get_tk_widget().place(x=0,y=0)
    canvas=FigureCanvasTkAgg(aux_1,master=GRA_2)
    canvas.draw()
    canvas.get_tk_widget().place(x=0,y=0)
    plt.show()
    
def DEM_MEN():
    fig, axes = plt.subplots(ncols=2, nrows=1,figsize=(10.5,3.5))
    Y_BAR=[]
    b=os.listdir(os.path.join(os.getcwd(),lstDesplegable.get()))
    X_BAR=[x[18:-10] for x in b]
    mes=len(b)
    #X_BAR=["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SET","OCT","NOV","DIC"]
    #graficos de barras
    for i in range(mes):
        file_iter=os.path.join(os.getcwd(),lstDesplegable.get(),os.listdir(os.path.join(os.getcwd(),lstDesplegable.get()))[i])
        df_iter=pd.read_excel(file_iter)
        df_iter.columns=["NADA","HORA","DEMANDA"]
        df_u=df_iter.loc[range(22,118),["HORA","DEMANDA"]]
        df_u.reset_index(drop=True,inplace=True)
        Y_BAR.append(df_u.max()["DEMANDA"])
        #dia de mayor demanda
    for i in range(mes):
        if Y_BAR[i]==max(Y_BAR):
            file_iter=os.path.join(os.getcwd(),lstDesplegable.get(),os.listdir(os.path.join(os.getcwd(),lstDesplegable.get()))[i])
            df_iter=pd.read_excel(file_iter)
            df_iter.columns=["NADA","HORA","DEMANDA"]
            df_u=df_iter.loc[range(22,118),["HORA","DEMANDA"]]
            df_u.reset_index(drop=True,inplace=True)
            df_u.plot(ax=axes[1],kind="area",x="HORA",y="DEMANDA").set_ylim(0,7300)
        else:
            pass
    axes[0].bar(X_BAR,Y_BAR)
    axes[0].set_ylim(5000,7300)
    axes[0].tick_params(axis="x", labelrotation = 45)
    axes[0].set_title("Demanda mensual máxima del "+lstDesplegable.get())
    axes[1].set_title("Diagrama de carga del dia de maxima demanda de "+lstDesplegable.get())
    plt.subplots_adjust(wspace=0.4, hspace=0.35)
    canvas=FigureCanvasTkAgg(fig,master=GRA_1)
    canvas.draw()
    canvas.get_tk_widget().place(x=-50,y=0)
    plt.show()
    texto.insert(0.0, df_u[df_u["DEMANDA"]==df_u.max()["DEMANDA"]])
    texto.insert(0.0,"\nAño: "+lstDesplegable.get()+"\n")

def emergente():
    #crea lista desplegable
    #lstDesplegable=ttk.Combobox(COES, width=23)
    tipo_usuario=[x for x in range(int(fecha1.get()),int(fecha2.get())+1)]
    lstDesplegable['values']= tipo_usuario
    #lstDesplegable.place(x=55,y=310)
    #lstDesplegable.set(2017)
    #boton
#definiendo la ventana
COES=tk.Tk()
COES.title("Análisis de Demanada Máxima")
COES.geometry("1000x620")
COES.config(bg="midnightblue")

#variable de los usuarios y contraseñas
fecha1= tk.StringVar()
fecha2= tk.StringVar()

#creando los frames que contendran los elementos
Fechas=tk.Frame(COES)
Fechas.place(x=10,y=15)
Fechas.config(width=250,height=550,bg="lightgray")

GRA_1=tk.Frame(COES)
GRA_1.place(x=290,y=310)
GRA_1.config(width=700,height=255,bg="lightgray")

GRA_2=tk.Frame(COES)
GRA_2.place(x=290,y=15)
GRA_2.config(width=330,height=255,bg="lightgray")

GRA_3=tk.Frame(COES)
GRA_3.place(x=650,y=15)
GRA_3.config(width=330,height=255,bg="lightgray")
#Etiqueta de las fechas 
title=tk.Label(Fechas, text="Análisis de Maxima Demanda")
title.place(x=10,y=15)
title.config(bg="yellow",width=32,height=2,fg="black")

label_1=tk.Label(Fechas, text="al")
label_1.place(x=100,y=65)
label_1.config(bg="yellow",width=5,height=1,fg="black")

label_2=tk.Label(Fechas, text="Máxima demanda mensual")
label_2.place(x=10,y=255)
label_2.config(bg="yellow",width=32,height=1,fg="black")

label_3=tk.Label(Fechas, text="Horas de mayor incidencia \n Evolución de la demanda máxima")
label_3.place(x=10,y=140)
label_3.config(bg="yellow",width=32,height=3,fg="black")

TITULO_APP=tk.Label(COES, text="PERSEO 3.0 By: Grupo 5 - Yucra")
TITULO_APP.place(x=-25,y=570)
TITULO_APP.config(bg="midnightblue",width=32,height=2,fg="Yellow",font=("Curier",14))
#Entrada de Texto
F_ini=tk.Entry(Fechas)
F_ini.config(justify="center",bd=1,textvariable=fecha1,width=12)
F_ini.place(x=10,y=65)

F_fin=tk.Entry(Fechas)
F_fin.config(justify="center",bd=1,textvariable=fecha2,width=12)
F_fin.place(x=155,y=65)

#Boton de calculo 
boton=tk.Button(Fechas,text="ANALIZAR",command=emergente)
boton.config(width=8,height=1)
boton.place(x=90,y=100)

boton_2=tk.Button(Fechas,text="Calcular",command=DEM_MEN)
boton_2.config(width=8,height=1)
boton_2.place(x=100,y=400)

boton_3=tk.Button(Fechas,text="Mostrar",command=EVO_HOR)
boton_3.config(width=8,height=1)
boton_3.place(x=100,y=200)
#Lista desplegable
lstDesplegable=ttk.Combobox(COES, width=23)
lstDesplegable.place(x=55,y=310)
#CONSOLA DE TEXTO
texto=tk.Text(Fechas)
texto.config(width=26,height=5,padx=10,pady=10,state="normal")
texto.place(x=10,y=440)


tk.mainloop()