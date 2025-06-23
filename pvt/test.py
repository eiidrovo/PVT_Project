from pvt.model.pvt import Rs,Pb, rho_oil,Co, mu,Bo
import numpy as np
import matplotlib.pyplot as plt
observations=1000
pressure_start = 14.7#14.7
pressure_end= 4409

#IMPUT DATA
API = 38.982
Gamma_Gas = 0.65
Gas_Solubility_Bubble_Point = 1124
Temperature = 140


Gamma_Oil = 141.5/(131.5+API)



#Variable Values
pressure=np.linspace(pressure_start,pressure_end,observations)


Pressure_Bubble_Point=(Pb(Gas_Solubility_Bubble_Point,API, Temperature, Gamma_Gas))
Gas_Solubility_Array = Rs(pressure,API,Temperature,Gamma_Gas,Gas_Solubility_Bubble_Point)




Compressibility = Co(pressure,Gas_Solubility_Bubble_Point,Gamma_Gas,API,Temperature)

Oil_density = rho_oil(pressure,
              API,
              Gas_Solubility_Bubble_Point,
              Pressure_Bubble_Point,
              Gamma_Oil,
              Gamma_Gas,
              Temperature,
              Compressibility)

viscosidad=mu(pressure,
              Pressure_Bubble_Point,
              Gas_Solubility_Array,
              API,
              Temperature)

Volumetric_factor=Bo(pressure,
                     API,
                     Temperature,
                     Gas_Solubility_Bubble_Point,
                     Pressure_Bubble_Point,
                     Gamma_Gas,
                     Gamma_Oil,
                     Compressibility)

plt.plot(pressure, Volumetric_factor)
plt.show()
