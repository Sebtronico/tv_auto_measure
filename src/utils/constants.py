# Diccionarios y constantes

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
            'Local_1': 'Local_1_Analógica'
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
            'Local_TDT': 'Local TDT'
        }
    }
}

# Definición del diccionario de servicios de televisión por operador.
TV_SERVICES = {
    'RTVC':	        ['Señal Colombia',  'Canal Institucional',  'Canal 1'],
    'Caracol':	    ['Caracol HD',      'Caracol HD 2',         'La Kalle', 'Caracol Móvil'],
    'RCN':	        ['RCN HD',          'RCN HD 2',             'RCN Móvil'],
    'CityTV':	    ['City TV',         'El Tiempo TV'],
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