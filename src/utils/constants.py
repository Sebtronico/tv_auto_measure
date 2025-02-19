"""
##########################################################################
##########################################################################

Diccionarios para televisión

##########################################################################
##########################################################################
"""
# Definición del diccionario con las frecuencias centrales de los canales
TV_TABLE = {
    2:  57,   3: 63,   4: 69,   5: 79,   6: 85,   7: 177,  8: 183,
    9:  189, 10: 195, 11: 201, 12: 207, 13: 213, 14: 473, 15: 479,
    16: 485, 17: 491, 18: 497, 19: 503, 20: 509, 21: 515, 22: 521,
    23: 527, 24: 533, 25: 539, 26: 545, 27: 551, 28: 557, 29: 563,
    30: 569, 31: 575, 32: 581, 33: 587, 34: 593, 35: 599, 36: 605,
    38: 617, 39: 623, 40: 629, 41: 635, 42: 641, 43: 647, 44: 653,
    45: 659, 46: 665, 47: 671, 48: 677, 49: 683, 50: 689, 51: 695
}

# Diccionario de las variables que se leen en cada modo de medida en tv digital
MODE_PARAMETERS = {
    'OVER': ['LEVel', 'CFOFfset', 'BROFfset', 'PERatio', 'BERLdpc', 'BBCH', 'FERatio', 'ESRatio'],
    'MERR': ['IMBalance', 'QERRor', 'CSUPpression', 'MRLO', 'MPLO', 'MRPLp', 'MPPLp', 'ERPLp', 'EPPLp'],
    'DSP': ['SALower', 'SAUPper'],
    'APHase': ['AMPLitude', 'PHASe'],
    'AGRoup': ['GDELay'],
}

# Tabla de conversión de los valores FEC.
FEC_TABLE = {
    'R1_2':'1/2',
    'R3_5':'3/5',
    'R2_3':'2/3',
    'R3_4':'3/4',
    'R4_5':'4/4',
    'R5_6':'5/6',
    '---' :'NA'
}

# Tabla de conversión de los valores de intervalo de guardas.
GINTERVAL_TABLE = {
    'G1_4'   :'1/4',
    'G19_128':'19/28',
    'G1_8'   :'1/8',
    'G19_256':'19/256',
    'G1_16'  :'1/16',
    'G1_32'  :'1/32',
    'G1_128' :'1/128',
    '---'    :'NA'
}

# Tabla de conversión de los valores de modulación.
MODULATION_TABLE = {
    'QAM16' :'16QAM',
    'QAM64' :'64QAM',
    'QAM256':'256QAM',
    'QPSK'  :'QPSK',
    'BPSK'  :'BPSK',
    '---'   :'NA'
}

# Tabla de conversión de los valores de FFT.
FFT_MODE_TABLE = {
    'F2K' :'2k',
    'F4K' :'4k',
    'F8K' :'8k',
    'F1K' :'1k',
    'F8KE':'8k ext',
    'F16K':'16k',
    'F16E':'16k ext',
    'F32K':'32k',
    'F32E':'32k ext',
    '---' :'NA'
}

#Definición del diccionario para el procesamiento de la medición de txCheck
TXCHECK_PARAMETERS = {
    'LEVel':        [35.0,      83.75,      70, 'Level'],
    'CFOFfset':     [30000,     0,          20, 'Carrier Freq Offset'],
    'BROFfset':     [20.0,      0.0,        20, 'Bit Rate Offset'],
    'MRLO':         [24.0,      43.0,       20, 'MER(L1,rms)'],
    'MPLO':         [10.0,      28.0,       20, 'MER(L1,peak)'],
    'MRPLp':        [24.0,      43.0,       20, 'MER(PLP,rms)'],
    'MPPLp':        [10.0,      28.0,       20, 'MER(PLP,peak)'],
    'ERPLp':        [4.40,      0.50,       20, 'EVM(PLP, rms)'],
    'EPPLp':        [22.00,     2.50,       20, 'EVM(PLP, peak)'],
    'AMPLitude':    [3.00,      0.10,       45, 'Amplitude'],
    'PHASe':        [45.00,     0.50,       45, 'Phase'],
    'GDELay':       [300e-9,    6e-9,       45, 'Group Delay'],
    'IMBalance':    [2.00,      0.00,       60, 'Amplitude Imbalance'],
    'QERRor':       [2.00,      0.00,       60, 'Quadrature Error'],
    'CSUPpression': [15.0,      50.0,       60, 'Carrier Suppression'],
    'BERLdpc':      [1e-2,      1e-11,      20, 'BER before LDPC'],
    'BBCH':         [1e-5,      1e-11,      20, 'BER before BCH'],
    'FERatio':      [1e-10,     1e-11,      20, 'BBFRAME error ratio'],
    'ESRatio':      [10.0,      0.0,        20, 'Errored second ratio'],
    'PERatio':      [1e-7,      1e-11,      80, 'Packet Error Ratio'],
    'SALower':      [35.0,      53.0,       65, 'Shoulder Att Lower'],
    'SAUPper':      [35.0,      53.0,       65, 'Shoulder Att Upper'],
}

"""
##########################################################################
##########################################################################

Diccionarios para banco de mediciones

##########################################################################
##########################################################################
"""


# Definición del diccionario con los parámetros de cada banda para el ETL
BANDS_ETL = {
    #Band      Fi      Ff      VBW   RBW  ref    tr. mode   unit 
    '700':    [703,    803,    100,  30, -25,   'AVERage', 'DBM'],
    '850':    [824,    894,    100,  30, -25,   'AVERage', 'DBM'],
    '815':    [806,    824,    100,  30, -25,   'AVERage', 'DBM'],
    '1900':   [1850,   1990,   100,  30, -25,   'AVERage', 'DBM'],
    'AWS1':   [1755,   1780,   30,   10, -25,   'MAXHold', 'DBM'],
    'AWS2':   [2155,   2180,   30,   10, -25,   'MAXHold', 'DBM'],
    'AWS3':   [2170,   2200,   30,   10, -25,   'MAXHold', 'DBM'],
    '2.5':    [2500,   2690,   100,  30, -25,   'AVERage', 'DBM'],
    '900_1':  [894,    915,    30,   10, -60,   'MAXHold', 'DBM'],
    '900_2':  [900,    928,    30,   10, -60,   'AVERage', 'DBM'],
    '3500':   [3300,   3700,   100,  30, -40,   'MAXHold', 'DBM'],
    '2.4GHz': [2400,   2483.5, 30,   10, -60,   'MAXHold', 'DBM'],
    '5GHz':   [5180,   5825,   100,  30, -60,   'MAXHold', 'DBM'],
    '2300MHz':[2300,   2400,   100,  30, -60,   'AVERage', 'DBM'],
    'Enlace': [300,    330,    30,   10,  82,   'AVERage', 'DBUV'],
    'tv_b1':  [54,     88,     30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_b3':  [174,    216,    30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_d':   [470,    512,    30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_b4':  [512,    608,    30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_b4_5':[614,    656,    30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_b5':  [656,    698,    30,   10,  82,   'AVERage', 'DBUVm'],
    'FM':     [88,     108,    30,   10,  82,   'AVERage', 'DBUVm'],
}

# Definición del diccionario con los parámetros de cada banda
BANDS_FXH = {
    #Band      Fi      Ff      VBW   RBW  ref    tr. mode   unit 
    '700':    [703,    803,    100,  30, -25,   'AVERage', 'DBM'],
    '850':    [824,    894,    100,  30, -25,   'AVERage', 'DBM'],
    '815':    [806,    824,    100,  30, -25,   'AVERage', 'DBM'],
    '1900':   [1850,   1990,   100,  30, -25,   'AVERage', 'DBM'],
    'AWS1':   [1755,   1780,   30,   10, -25,   'MAXHold', 'DBM'],
    'AWS2':   [2155,   2180,   30,   10, -25,   'MAXHold', 'DBM'],
    'AWS3':   [2170,   2200,   30,   10, -25,   'MAXHold', 'DBM'],
    '2.5':    [2500,   2690,   100,  30, -25,   'AVERage', 'DBM'],
    '900_1':  [894,    915,    30,   10, -60,   'MAXHold', 'DBM'],
    '900_2':  [900,    928,    30,   10, -60,   'AVERage', 'DBM'],
    '3500':   [3300,   3700,   100,  30, -40,   'MAXHold', 'DBM'],
    '2.4GHz': [2400,   2483.5, 30,   10, -60,   'MAXHold', 'DBM'],
    '5GHz':   [5180,   5825,   100,  30, -60,   'MAXHold', 'DBM'],
    '2300MHz':[2300,   2400,   100,  30, -60,   'AVERage', 'DBM'],
    'Enlace': [300,    330,    30,   10,  82,   'AVERage', 'DBUV'],
    'tv_b1':  [54,     88,     30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_b3':  [174,    216,    30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_d':   [470,    512,    30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_b4':  [512,    608,    30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_b4_5':[614,    656,    30,   10,  82,   'AVERage', 'DBUVm'],
    'tv_b5':  [656,    698,    30,   10,  82,   'AVERage', 'DBUVm'],
    'FM':     [88,     108,    30,   10,  82,   'AVERage', 'DBUVm'],
}

UNITS = {
    'DBM': 'dBm',
    'DBUV': 'dBμV',
    'DBUVm': 'dBμV/m'
}


"""
##########################################################################
##########################################################################

Diccionarios para lectura de preingeniería

##########################################################################
##########################################################################
"""

# Definición del diccionario con las columnas de búsqueda de estación y acimuth
# en la hoja 'Coordenadas y Acimuths Ppal'
PRINCIPAL = {
    'CCNP': 'Az\n(°)',
    'RTVC': 'Az\n(°).1'
}

# Definición del diccionario con las columnas de búsqueda de estación y acimuth
# en la hoja 'Acimuths Adicionales'
ADDITIONAL = {
    'Tx_DA1_CCNP':          'Az 1\n(°)', 
    'Tx_DA1_RTVC':          'Az 2\n(°)',
    'Tx_DA2_CCNP':          'Az 3\n(°)',
    'Tx_DA2_RTVC':          'Az 4\n(°)',
    'Tx_A1_RTVC':           'Az 5\n(°)',
    'Tx_A1_Regional':       'Az 6\n(°)',
    'Tx_A1_CCNP':           'Az 7\n(°)',
    'Tx_A2_RTVC':           'Az 8\n(°)',
    'Tx_A2_Regional':       'Az 9\n(°)',
    'Tx_A2_CCNP':           'Az 10\n(°)',
    'Tx_A1_2do_Regional':   'Az 11\n(°)',
    'Tx_AD1_Local':         'Az 12\n(°)',
    'Tx_DA3_CCNP':          'Az 13\n(°)',
    'Tx_DA3_RTVC':          'Az 14\n(°)',
}

# Diccionario de iteración para obtención de datos en la hoja 'Canalización Ppal'
SEARCH_PRINCIPALS = {
    'Analógico': {
        'Estación Públicos TV Analógica': {
            'C1': 'Canal 1',
            'CI': 'Canal Institucional',
            'SC': 'Señal Colombia',
        },

        'Estación Regional TV Analógica': {
            'CR': 'Canal Regional',
        },

        'Estación Privados TV Analógica': {
            'RCN': 'RCN',
            'CRC': 'Caracol',
        },
    },

    'Digital': {
        'Estación TDT CCNP': {
        'RCND': 'RCN',
        'CRCD': 'Caracol',
        },

        'Estación TDT RTVC': {
        'RTVC': 'RTVC',
        'REG1': 'Canal Regional'
        }
    }
}

# Diccionario de iteración para obtención de datos en la hoja 'Canalización Ppal'
SEARCH_ADDITIONALS = {
    'Analógico': {
        'Estación Públicos TV Analógica': {
            'C1': 'Canal 1',
            'CI': 'Canal Institucional',
            'SC': 'Señal Colombia',
        },

        'Estación Regional TV Analógica': {
            'CR': 'Canal Regional',
        },

        'Estación Privados TV Analógica': {
            'RCN': 'RCN',
            'CRC': 'Caracol',
        },

        'Estación 2do_CR': {
            '2do_CR': 'Regional 2'
        },

        'Estación Analógica Canal Local': {
            'CH_Local1': 'Local_1_Analógica'
        }
    },

    'Digital': {
        'Estación TDT CCNP': {
            'RCND': 'RCN',
            'CRCD': 'Caracol',
        },

        'Estación TDT RTVC': {
            'RTVC': 'RTVC',
            'REG1': 'Canal Regional'
        },

        'Estación TDT 2do_CR': {
            '2doCR TDT': 'Regional 2 TDT'
        },

        'Estación Local TDT': {
            'CH_Local TDT': 'Local TDT'
        }
    }
}

# Definición del diccionario de servicios de televisión por operador.
TV_SERVICES = {
    'RTVC':	        ['Señal Colombia',  'Canal Institucional',  'Canal 1'],
    'Caracol':	    ['Caracol HD',      'Caracol HD 2',         'La Kalle', 'Caracol Móvil'],
    'RCN':	        ['RCN HD',          'RCN HD 2',             'RCN Móvil'],
    'City TV':	    ['City TV',         'El Tiempo TV'],
    'Telecaribe':	['Telecaribe',      'Telecaribe +'],
    'TRO':	        ['TRO',             'TRO HD2'],
    'Teveandina':	['Canal Trece',     'Canal Trece +'],
    'Teleantioquia':['Teleantioquia',   'Teleantioquia HD2'],
    'Telesantiago':	['Telesantiago'],
    'Telecafé':	    ['Telecafé',        'Telecafé 2'],
    'Teleislas':	['Teleislas',       'Raizal TV'],
    'Canal Capital':['Canal Capital',   'Eureka Capital'],
    'Telepasto':	['Telepasto'],
    'Telepacífico':	['Telepacífico HD', 'Origen Channel']
}