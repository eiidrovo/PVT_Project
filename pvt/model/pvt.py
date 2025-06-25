import numpy as np
"Carga la librería NumPy para trabajar con operaciones vectorizadas y arrays numéricos."

"Calcula el factor de volumen del petróleo (Bo) en función de presión y otras propiedades. Ajusta su valor si la presión supera la de burbuja."
def Bo(p, api, t, rsb, pb, gamma_gas, gamma_oil, co):

    '''

    Parameters
    ----------
    p
    api
    t
    rsb
    pb
    gamma_gas
    gamma_oil
    co

    Returns
    -------

    '''
    rs = Rs(p, api, t, gamma_gas, rsb)
    bo = 0.9759 + 0.000120 * (rs * (gamma_gas / gamma_oil) + 1.25 * t) ** 1.2
    index = np.where(p >= pb)
    bo[index] = bo[index[0] - 1] * np.exp(co[index] * (pb - p[index]))
    return bo

"Determina la cantidad de gas disuelto (Rs) dependiendo de la presión, API, temperatura y gravedad específica del gas. Se limita a Rsb cuando la presión supera el punto de burbuja."
def Rs(p, api, temperature, gamma_gas, rsb, pb=None):
    '''

    Parameters
    ----------
    p
    api
    temperature
    gamma_gas
    rsb
    pb

    Returns
    -------

    '''
    x = 0.0125 * api - 0.00091 * temperature
    rs = gamma_gas * ((p / 18.2 + 1.4) * 10**x) ** 1.2048
    if isinstance(p, np.ndarray):
        rs[np.where(rs >= rsb)] = rsb
    elif pb and p >= pb:
        rs = rsb
    return rs

"Estima la presión en la que el gas comienza a separarse del petróleo (Pb) usando Rsb, API, temperatura y gravedad del gas."
def Pb(rsb, api, temperature, gamma_gas):
    '''

    Parameters
    ----------
    rsb
    api
    temperature
    gamma_gas

    Returns
    -------

    '''
    a = 0.00091 * temperature - 0.0125 * api
    pb = 18.2 * (rsb / gamma_gas) ** 0.83 * 10**a - 1.4
    return pb

"Calcula la compresibilidad isotérmica del crudo (Co) como función empírica de Rsb, presión, gravedad del gas, API y temperatura."
def Co(p, rsb, gamma_g, api, t):
    '''

    Parameters
    ----------
    p
    rsb
    gamma_g
    api
    t

    Returns
    -------

    '''
    co = (
        1.705e-7 * rsb**0.69357 * gamma_g**0.1885 * api**0.3272 * t**0.6729 * p**-0.5906
    )
    return co

"Estima la densidad del petróleo (ρo) en función de presión, Rs, gravedades específicas y temperatura. Ajusta la densidad si la presión supera el punto de burbuja usando la compresibilidad."
def rho_oil(
    pressure,
    api,
    rsb,
    bubble_point,
    gamma_oil,
    gamma_gas,
    temperature,
    comprensibility=None,
):
    '''

    Parameters
    ----------
    pressure
    api
    rsb
    bubble_point
    gamma_oil
    gamma_gas
    temperature
    comprensibility

    Returns
    -------

    '''
    rs = Rs(pressure, api, temperature, gamma_gas, rsb)
    rho_o = (62.4 * gamma_oil + 0.0316 * rs * gamma_gas) / (
        0.972
        + 0.000147
        * (rs * (gamma_gas / gamma_oil) ** 0.25 + 1.25 * temperature) ** 1.175
    )

    if comprensibility is not None:
        for pressure_index in np.where(pressure > bubble_point):
            rho_o[pressure_index] = rho_o[
                np.where(pressure > bubble_point)[0] - 1
            ] ** np.exp(
                comprensibility[pressure_index]
                * (pressure[pressure_index] - bubble_point)
            )
    return rho_o

"Calcula la viscosidad del petróleo (μo) considerando Rs, API, temperatura y presión. Usa una fórmula empírica para ajustar el valor cuando la presión es mayor a la de burbuja."
def mu(p, pb, rs, api, t):
    '''

    Parameters
    ----------
    p
    pb
    rs
    api
    t

    Returns
    -------

    '''
    x = 10 ** (3.0324 - 0.02023 * api) * t ** (-1.163)
    mu_od = 10**x - 1

    a = 10.715 * (rs + 100) ** (-0.515)
    b = 5.44 * (rs + 150) ** (-0.338)
    mu_ob = a * mu_od**b

    index = np.where(p > pb)
    m = 2.6 * p[index] ** 1.187 * np.exp(-11.513 - 8.98e-5 * p[index])
    mu_ob[index] = mu_ob[index] * (p[index] / pb) ** m

    return mu_ob
