import pandas as pd
import configparser

class Report():
    def __init__(self, orig, dest):
        self._orig = orig
        self._dest = dest

    def read_data(self):
        self._data = pd.read_excel(self._orig)

def read_config(path):
    config = configparser.ConfigParser()
    config.read(path)
    return config

def save_config(params, filename):
    config = configparser.ConfigParser()
    config.add_section('Principal')
    config['Principal']['Origen'] = params['-ORIG-']
    config['Principal']['Destino'] = params['-DEST-']
    with open(filename, 'w') as configfile:
        config.write(configfile)