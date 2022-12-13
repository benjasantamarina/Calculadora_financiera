from tkinter import ttk
from cProfile import label
from cgitb import text
from logging import root
from tkinter import *
from tkinter import messagebox
from tkinter.tix import INTEGER
from turtle import st
from os import *


#---------------Funciones-------------------------

        
def plan_financiacion (capital, plazo, tasa, cliente):
    
    import pandas as pd
    
    df = pd.DataFrame()
    
    try:
        capital = float(capital.get())
        plazo = int(plazo.get())
    except:
        messagebox.showinfo("Error", "Error, el capital ingresado es incorrecto, por favor ingrese un valor numerico")
        
    else:
        try:
            tasa = float(tasa.get())
        except:
            messagebox.showinfo("Error", "Error, la TNA ingresada es incorrecta, por favor ingrese un valor numerico")
        else:
            try:
                cuotas = int(plazo/30)
                lista = list(range(1, (cuotas + 1)))
                df["Cuota"] = lista
                df["Monto"] = round(float(capital/cuotas),2)
                df["Plazo"] = list(range(30, (plazo+1), 30))
                df["Tasa"] = round((((1 + ((tasa/100)/12))**(df["Plazo"]/30))-1),4)
                df["Interes"] = round((df["Monto"] * df["Tasa"]),2)
                df["Total"] = round((df["Monto"] + df["Interes"]),2)
                df.loc["Total"] = 0
                df.iloc[-1,-1] = df["Total"].sum(axis = 0)
                df.iloc[-1,-2] = df["Interes"].sum(axis = 0)
                df.iloc[-1,-5] = capital
                df.iloc[-1,-3] = (df.iloc[-1, -2])/(df.iloc[-1, -5])
                df.iloc[-1,-4] = ""
                df.iloc[-1,-6] = "Total"
                df.set_index("Cuota", inplace=True)
                df["Tasa"] = df["Tasa"].apply(a_porcentaje)
                df.to_excel(f"C:/Users/Benja/Desktop/Plan de pago/Plan de pago - {cliente.get()}.xlsx")
                messagebox.showinfo("Plan de financiacion", "Programa ejecutado con exito")
            except:
                error = sys.exc_info()[1]
                messagebox.showinfo("ERROR", error)
    
    
def a_porcentaje (numero):
    numero = "{0:.2f}%".format(numero * 100)
    return numero

def salirAplicacion():

    valor=messagebox.askquestion("Salir", "¿Desea salir de la aplicación")

    if valor=="yes":
        root.destroy()
        
def limpiarCampos():
    capital.set("")
    plazo.set("")
    tasa.set("")
    cliente.set("")
    
#---------------Raiz------------------------------

raiz = Tk()
raiz.title("Plan de Financiación")
# raiz.iconbitmap("C:/Users/Usuario/Desktop/Plan financiacion/logo.ico")
raiz.config(bg = "#BCC4E1")


capital = StringVar()
plazo = StringVar()
tasa = StringVar()
cliente = StringVar()


#---------------Frame-----------------------------

mi_frame = Frame()
mi_frame.pack(fill = "both", expand = "True")
mi_frame.config(bg = "#BCC4E1")
mi_frame.config(width = "300", height = "350")


#---------------Label-----------------------------

capital_texto = Label(mi_frame, text="Capital a financiar: ", bg="#BCC4E1")
capital_texto.grid(row=0, column=0, padx=10, pady=10)
                           
plazo_texto = Label(mi_frame, text="Plazo financiacion: ", bg="#BCC4E1")
plazo_texto.grid(row=2, column=0, padx=10, pady=5)

tasa_texto = Label(mi_frame, text="TNA: ", bg="#BCC4E1")
tasa_texto.grid(row=4, column=0, padx=10, pady=5)

cliente_texto = Label(mi_frame, text="Cliente: ", bg="#BCC4E1")
cliente_texto.grid(row=6, column=0, padx=10, pady=5)

#---------------Entry-----------------------------

capital_entrada = Entry(mi_frame, textvariable = capital)
capital_entrada.grid(row=1, column=0, padx = 10, pady = 5)
capital_entrada.config(justify = "right")

plazo_combo = ttk.Combobox(mi_frame, state="readonly",values=[30, 60, 90, 120, 150, 180, 210], textvariable = plazo)
plazo_combo.grid(row=3, column=0, padx=10, pady=5)
plazo_combo.config(width = "17")

tasa_entrada = Entry(mi_frame, textvariable = tasa)
tasa_entrada.grid(row=5, column=0, padx=10, pady=5)
tasa_entrada.config(justify = "right")

cliente_entrada = Entry(mi_frame, textvariable = cliente)
cliente_entrada.grid(row=7, column=0, padx=10, pady=10)

#---------------Botones---------------------------

frameInferior=Frame(raiz)
frameInferior.pack()
frameInferior.config(bg = "#BCC4E1")

boton_calculo = Button(frameInferior, text = "Calcular", command = lambda:plan_financiacion(capital, plazo, tasa, cliente))
boton_calculo.grid(row=0, column=1, sticky="e", padx = 10, pady = 10)

boton_limpiar = Button(frameInferior, text = "Borrar", command = limpiarCampos)
boton_limpiar.grid(row=0, column=0, sticky="e", padx = 10, pady = 10)

raiz.mainloop()