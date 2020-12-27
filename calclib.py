import configparser

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