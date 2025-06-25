import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from model.pvt import Rs,Pb, rho_oil,Co, mu,Bo
np.set_printoptions(precision=3, suppress=True)

xw.Book("pvt.xlsm").set_mock_caller()
wb = xw.Book.caller()
sheet = wb.sheets["pvt"]

API_cell='API'
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
Gamma_Gas = float(sheet[Gamma_Gas_cell].value)
Observations = int(sheet[Observations_cell].value)
Pressure_Start = float(sheet[Pressure_Start_cell].value)
Pressure_End = float(sheet[Pressure_End_cell].value)
Rsb = float(sheet[Rsb_cell].value)
T = float(sheet[T_cell].value)
Plot = sheet[Plot_cell].value

Gamma_Oil = 141.5/(131.5+API)

pressure=np.linspace(Pressure_Start,Pressure_End,Observations)
Pressure_Bubble_Point=Pb(Rsb,API, T, Gamma_Gas)
sheet[Bubble_Point].value = f'{Pressure_Bubble_Point:.3f}'


Gas_Solubility_Array = Rs(pressure,API,T,Gamma_Gas,Rsb)
values_sheet=wb.sheets['pvt_values']
values_sheet[Rs_values_cell].options(np.array, transpose=True).value=Gas_Solubility_Array

Compressibility = Co(pressure,Rsb,Gamma_Gas,API,T)
values_sheet[Co_values_cell].options(np.array, transpose=True).value=Compressibility

Oil_density = rho_oil(pressure,
              API,
              Rsb,
              Pressure_Bubble_Point,
              Gamma_Oil,
              Gamma_Gas,
              T,
              Compressibility)
values_sheet[Rho_O_values_cell].options(np.array, transpose=True).value=Oil_density

Mu_O_values=mu(pressure,
              Pressure_Bubble_Point,
              Gas_Solubility_Array,
              API,
              T)
values_sheet[Mu_O_values_cell].options(np.array, transpose=True).value=Mu_O_values

Volumetric_factor=Bo(pressure,
                     API,
                     T,
                     Rsb,
                     Pressure_Bubble_Point,
                     Gamma_Gas,
                     Gamma_Oil,
                     Compressibility)
values_sheet[Bo_values_cell].options(np.array, transpose=True).value=Volumetric_factor

values_sheet[Pressure_values_cell].options(np.array, transpose=True).value=pressure

Results_df=values_sheet.range('A1').options(pd.DataFrame, expand='table').value


fig, ax = plt.subplots()
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
ax.set_xlabel('Pressure')
ax.set_ylabel(f'{Plot}')
ax.grid()
print(Pressure_Bubble_Point)
sheet.pictures.add(fig, name='MyPlot', update=True,
                     left=sheet.range('B18').left, top=sheet.range('B18 ').top)