import xlwings as xw
"Importa la librería xlwings, que permite interactuar con archivos de Excel desde Python."

import numpy as np
"Importa la librería NumPy como np para trabajar con arrays numéricos y funciones matemáticas."

import pandas as pd
"Importa la librería pandas como pd para trabajar con estructuras de datos como DataFrames."

import matplotlib.pyplot as plt
"Importa la librería matplotlib.pyplot para crear gráficos."

from model.pvt import Rs,Pb, rho_oil,Co, mu,Bo
"Importa la librería matplotlib.pyplot para crear gráficos."

np.set_printoptions(precision=3, suppress=True)
"Configura NumPy para mostrar solo 3 decimales y suprimir la notación científica al imprimir arrays."

xw.Book("pvt.xlsm").set_mock_caller()
"Establece el archivo pvt.xlsm como archivo activo para pruebas (mock caller) si se ejecuta fuera de Excel."

wb = xw.Book.caller()
"Obtiene el libro activo cuando el código es ejecutado desde Excel."


sheet = wb.sheets["pvt"]
"Selecciona la hoja llamada pvt dentro del archivo de Excel."


API_cell='API'
"Define nombres de celdas para facilitar la lectura de datos desde Excel."

Gamma_Gas_cell='Gamma_Gas'
Observations_cell='observations'
Pressure_Start_cell='Pressure_Start'
Pressure_End_cell='Pressure_End'
Rsb_cell='Rsb'
T_cell='T'
Bubble_Point='Bubble_Point'

Bo_values_cell='Bo_Values'
Co_values_cell='Co_Values'
Mu_O_values_cell='Mu_O_Values'
Rho_O_values_cell='Rho_O_Values'
Rs_values_cell='Rs_Values'
Pressure_values_cell='Pressure_Values'

Plot_cell='Plot'

API = float(sheet[API_cell].value)
"Lee los valores desde las celdas especificadas y los convierte al tipo de dato adecuado (float o int)."

Gamma_Gas = float(sheet[Gamma_Gas_cell].value)
Observations = int(sheet[Observations_cell].value)
Pressure_Start = float(sheet[Pressure_Start_cell].value)
Pressure_End = float(sheet[Pressure_End_cell].value)
Rsb = float(sheet[Rsb_cell].value)
T = float(sheet[T_cell].value)
Plot = sheet[Plot_cell].value

Gamma_Oil = 141.5/(131.5+API)
"Calcula la gravedad específica del petróleo (relación entre la densidad del petróleo y la del agua)."

pressure=np.linspace(Pressure_Start,Pressure_End,Observations)
"Genera un arreglo con valores de presión distribuidos linealmente entre los límites definidos."

Pressure_Bubble_Point=Pb(Rsb,API, T, Gamma_Gas)
"Calcula la presión del punto de burbuja con la función Pb"

sheet[Bubble_Point].value = f'{Pressure_Bubble_Point:.3f}'
"Escribe el valor de la presión de burbuja en la hoja de Excel con tres decimales."


Gas_Solubility_Array = Rs(pressure,API,T,Gamma_Gas,Rsb)
"Calcula la solubilidad del gas (Rs) para cada valor de presión."

values_sheet=wb.sheets['pvt_values']
"Accede a la hoja "pvt_values" donde se guardarán los resultados."

values_sheet[Rs_values_cell].options(np.array, transpose=True).value=Gas_Solubility_Array
"Escribe los valores de Rs en la hoja de Excel en forma de columna."

Compressibility = Co(pressure,Rsb,Gamma_Gas,API,T)
"Calcula la compresibilidad del crudo para los valores de presión."

values_sheet[Co_values_cell].options(np.array, transpose=True).value=Compressibility
"Escribe los valores de compresibilidad en la hoja de Excel."

Oil_density = rho_oil(pressure,
              API,
              Rsb,
              Pressure_Bubble_Point,
              Gamma_Oil,
              Gamma_Gas,
              T,
              Compressibility)
"Calcula la densidad del petróleo usando los parámetros definidos."

values_sheet[Rho_O_values_cell].options(np.array, transpose=True).value=Oil_density
"Escribe los valores de densidad del petróleo en la hoja de Excel."

Mu_O_values=mu(pressure,
              Pressure_Bubble_Point,
              Gas_Solubility_Array,
              API,
              T)
values_sheet[Mu_O_values_cell].options(np.array, transpose=True).value=Mu_O_values
"Escribe los valores de viscosidad en la hoja de Excel."

Volumetric_factor=Bo(pressure,
                     API,
                     T,
                     Rsb,
                     Pressure_Bubble_Point,
                     Gamma_Gas,
                     Gamma_Oil,
                     Compressibility)
"Calcula el factor volumétrico del petróleo."

values_sheet[Bo_values_cell].options(np.array, transpose=True).value=Volumetric_factor
"Escribe los valores del factor volumétrico en la hoja de Excel."

values_sheet[Pressure_values_cell].options(np.array, transpose=True).value=pressure
"Escribe los valores de presión en la hoja de Excel."

Results_df=values_sheet.range('A1').options(pd.DataFrame, expand='table').value
"Lee un rango de datos desde Excel y lo convierte en un DataFrame de pandas."


fig, ax = plt.subplots()
"Crea una figura y un eje (ax) para graficar."

if Plot=='Rs':
    ax.plot(pressure, Results_df['RS'])
elif Plot=='Bo':
    ax.plot(pressure, Results_df['Bo'])
elif Plot=='Rho_O':
    ax.plot(pressure, Oil_density)
elif Plot=='Mu_O':
    ax.plot(pressure, Results_df['Mu_O'])
elif Plot=='Co':
    ax.plot(pressure, Results_df['Co'])



ax.set_title(f"{Plot} vs pressure")
"Establece el título del gráfico."

ax.set_xlabel('Pressure')
"Establece la etiqueta del eje X."

ax.set_ylabel(f'{Plot}')
"Establece la etiqueta del eje Y con la propiedad elegida."

ax.grid()
"Activa la cuadrícula en el gráfico."

print(Pressure_Bubble_Point)
"Imprime en consola la presión del punto de burbuja."

sheet.pictures.add(fig, name='MyPlot', update=True,
                     left=sheet.range('B18').left, top=sheet.range('B18 ').top)
"Inserta el gráfico generado en la hoja de Excel, posicionándolo cerca de la celda B18."