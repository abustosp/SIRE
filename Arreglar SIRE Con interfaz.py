import pandas as pd
import numpy as np
import os
import customtkinter as ctk
from tkinter import filedialog
from PIL import Image
import tkinter

ctk.set_default_color_theme("green")
ctk.set_appearance_mode("dark")


# Crear la clase App
class App(ctk.CTk):

    # Inicializar la clase
    def __init__(self):
        
        # Inicializar la clase padre
        super().__init__()

        # Crea la ventana principal
        self.title("Arreglar SIRE 1.0 por Agustín Bustos")
        self.geometry("400x200")
        self.grid_columnconfigure(0 , weight=1)

        #obtener el directorio de trabajo
        #directorio = os.getcwd()

        #Agregar un cuadro con el logo
        self.logo = ctk.CTkImage(Image.open(("BIN/ABP blanco sin fondo.png")) , size=(40 , 40))
        self.logo_imagen = ctk.CTkLabel(self, text="", image=self.logo)
        self.logo_imagen.grid(row=0, column=0, pady=5)

        #Agregar un texto de bienvenida justificado a la izquierda
        self.bienvenida = ctk.CTkLabel(self, text="Arreglar archivos de SIRE 1.0 por Agustín Bustos")
        self.bienvenida.grid(row=1, column=0, pady=5)
    
        def Seleccionar_Archivo_y_procesar():

            #Seleccionar el archivo
            Archivo = filedialog.askopenfile(parent=self)

            #si no se selecciona ningun archivo, dejar Archivo vacio
            if Archivo == None:
                Archivo = ""

            ############### Leer archivos Origninal de SIRE y MC #############

            if Archivo != "":

                #leer el archivo original de SIRE
                SIRE_o = pd.read_fwf(Archivo,
                                    header=None,
                                    widths=[4, 36, 3, 3, 10, 2, 1, 30, 14, 14, 1, 6, 10, 2, 10, 5, 1, 8, 12, 12, 14, 16, 30, 11, 25, 10, 14, 1],
                                    names=['VERSIÓN', 'CÓDIGO DE TRAZABILIDAD', 'IMPUESTO', 'RÉGIMEN', 'FECHA RETENCIÓN', 'CONDICIÓN', 'IMPOSIBILIDAD DE RETENCIÓN', 'MOTIVO NO RETENCIÓN', 'IMPORTE RETENCIÓN', 'BASE DE CÁLCULO', 'RÉGIMEN DE EXCLUSIÓN', '% DE EXCLUSIÓN', 'FECHA PUBL O FINAL DE LA VIGENCIA', 'TIPO CBTE', 'FECHA CBTE', 'Pto de venta', '-', 'Nro de Cbte', 'COE', 'COE ORIGINAL', 'CAE', 'IMPORTE COMPROBANTE', 'MOTIVO EMISIÓN DE NOTA DE CRÉDITO/AJUSTE', 'RETENIDO CLAVE', 'CERTIFICADO ORIGINAL NRO', 'CERTIFICADO ORIGINAL FECHA RETEN', 'CERTIFICADO ORIGINAL IMPORTE', 'MOTIVO DE LA ANULACIÓN'],
                                    decimal="," ,
                                    thousands=""
                                    )

                #leer todos los archivos .xlsx de la carpeta "Base/MCR" y unirlos en un solo DataFrame
                path = "Base/MCR"
                all_files = os.listdir(path)
                all_files = [file for file in all_files if file.endswith('.xlsx')]
                all_files.sort(reverse=True)
                
                #Crear un DataFrame vacio
                MCr = pd.DataFrame()

                #Leer cada archivo y agregarlo al DataFrame vacio
                for file in all_files:
                    df = pd.read_excel(path + "/" + file , skiprows=1)
                    #ordenar dataframe en orden descendente de indice
                    df = df.sort_index( ascending=False)
                    MCr = pd.concat([MCr, df])

                del df , file , all_files , path

                print('''
Realizando arreglos...
                ''')

                ############ Crear columnas auxiliares en Mis Comprobantes Recibidos

                MCr = MCr[MCr["Fecha"] != "Fecha"] #Filtrar las filas que contienen la palabra "fecha" en la columna de Fecha (aparece al consolidar todos los archivos CSV de Mis Comprobantes) 
                #MCr = MCr.sort_index( ascending=False)
                MCr["AUX"] = MCr["Punto de Venta"].astype(str) + " - " + MCr["Número Desde"].astype(str) + " - " + MCr["Nro. Doc. Emisor"].astype(str)
                MCr[['Tipo Nro' , "Tipo Descripción"]] = MCr["Tipo"].str.split(' - ', expand=True)
                MCr = MCr[~MCr["Tipo Nro"].isin(("11","12","13","15"))]  #### Filtrar los datos que no se incluyen en una lista

                #Convertir las columas 'Imp. Neto Gravado' , 'IVA' , 'Imp. Total' y 'Tipo Cambio' a tipo float
                MCr["Imp. Neto Gravado"] = MCr["Imp. Neto Gravado"].astype(float)
                MCr["IVA"] = MCr["IVA"].astype(float)
                MCr["Imp. Total"] = MCr["Imp. Total"].astype(float)
                MCr["Tipo Cambio"] = MCr["Tipo Cambio"].astype(float)

                ###### Pasar a Pesos los comprobantes en Moneda Extranjera
                MCr["Imp. Neto Gravado MCr"] = (MCr["Imp. Neto Gravado"] * MCr["Tipo Cambio"]).round(2) # Redondear a 2 decimales
                MCr["IVA MCr"] = (MCr["IVA"] * MCr["Tipo Cambio"]).round(2) # Redondear a 2 decimales
                MCr["Imp. Total MCr"] = (MCr["Imp. Total"] * MCr["Tipo Cambio"]).round(2) # Redondear a 2 decimales

                #### Seleccionar columnas de MCr
                MCr = MCr[["AUX" , "Fecha", "Imp. Neto Gravado MCr", "IVA MCr", "Imp. Total MCr"]]
                MCr = MCr.rename(columns={"Fecha":"Fecha MCr"})

                #Crear un nuevo dataframe y modificar los datos
                Sire_Modificado = SIRE_o
                Sire_Modificado["Pto de venta"] = Sire_Modificado["Pto de venta"].astype(str)

                #Modificar la columna 'Pto de venta' para que no tenga el 11 o 12 al principio en caso que tenga una longitud de 5 caracteres
                Sire_Modificado.loc[((Sire_Modificado["Pto de venta"].astype(str).str.len() == 5) & (Sire_Modificado["Pto de venta"].str.startswith("11") | Sire_Modificado["Pto de venta"].str.startswith("12"))) , "Pto de venta"] = Sire_Modificado["Pto de venta"].astype(str).str[2:5]
                Sire_Modificado["Pto de venta"] = Sire_Modificado["Pto de venta"].astype(int)
                Sire_Modificado["AUX"] = Sire_Modificado["Pto de venta"].astype(str) + " - " + Sire_Modificado["Nro de Cbte"].astype(str) + " - " + Sire_Modificado["RETENIDO CLAVE"].astype(str)

                # Unir las tablas con su AUX #
                Sire_Modificado = pd.merge(
                    left= Sire_Modificado,
                    right= MCr,
                    left_on="AUX",
                    right_on="AUX",
                    how="left"
                    )

                #crear columnas auxiliares de control para verificar diferencias
                Sire_Modificado["Diferencia de IVA (IVA MCr - Base de Cálculo)"] =  Sire_Modificado["IVA MCr"] - Sire_Modificado["BASE DE CÁLCULO"]
                Sire_Modificado["Diferencia de Total CBTE (Total MCr - Total SIRe)"] =  Sire_Modificado["Imp. Total MCr"] - Sire_Modificado["IMPORTE COMPROBANTE"]
                Sire_Modificado["% RET efectivo"] = (Sire_Modificado["IMPORTE RETENCIÓN"] / Sire_Modificado["BASE DE CÁLCULO"]).round(2)

                ####### Reemplazar valores de IVA en Tabla original
                Sire_Modificado.loc[Sire_Modificado["IVA MCr"].notnull() , ["BASE DE CÁLCULO"]] = Sire_Modificado["IVA MCr"]
                Sire_Modificado.loc[Sire_Modificado["Fecha MCr"].notnull() , ["FECHA CBTE"]] = Sire_Modificado["Fecha MCr"]
                Sire_Modificado.loc[Sire_Modificado["Imp. Total MCr"].notnull() , ["IMPORTE COMPROBANTE"]] = Sire_Modificado["Imp. Total MCr"]

                ######### Calcular el % de retención teorico
                Sire_Modificado["% RET"] = np.NAN #Crear columna de % RET y rellenar con NaN
                Sire_Modificado.loc[Sire_Modificado["Imp. Total MCr"] < 24000 , ["% RET"]] = 1
                Sire_Modificado.loc[(Sire_Modificado["Imp. Total MCr"] > 24000) & (Sire_Modificado["RÉGIMEN"] == 212) , ["% RET"]] = 0.8
                Sire_Modificado.loc[(Sire_Modificado["Imp. Total MCr"] > 24000) & (Sire_Modificado["RÉGIMEN"] == 214) , ["% RET"]] = 0.5

                #Calcular el % de retención efectiva y la diferencia con lo teorico
                Sire_Modificado["Diferencia RET efectivo"] = Sire_Modificado["% RET efectivo"] - Sire_Modificado["% RET"]

                #Eliminar columna '-'
                Sire_Modificado = Sire_Modificado.drop(columns=["-"])

                #Crear la carpeta 'Generado' si no existe
                if not os.path.exists("Generado"):
                    os.makedirs("Generado")

                #Guardar el archivo en la carpeta 'Generado'
                Sire_Modificado.to_excel("Generado/Retenciones_Modificadas.xlsx", index=False)
                MCr.to_excel("Generado/MCr.xlsx", index=False)

                print('''Proceso finalizado: Archivos generados correctamente en la carpeta 'Generado'
                ''')
                
        #Crea un botón para seleccionar la carpeta desde una ventana del explorador de archivos
        boton = ctk.CTkButton(self, text="Seleccionar TXT de SIRE a corregir", command=Seleccionar_Archivo_y_procesar)
        boton.grid(row=2, column=0, pady=10)

        #Crear un botón para salir de la aplicación al lado del botón anterior
        boton2 = ctk.CTkButton(self, text="Salir", command=self.quit)
        boton2.grid(row=3, column=0, pady=10)

# Inicia el bucle principal de la ventana
if __name__ == "__main__":
    app = App()
    app.mainloop()