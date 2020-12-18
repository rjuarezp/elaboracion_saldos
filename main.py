import PySimpleGUI as sg
import os
import calclib

ININAME = 'config.ini'



sg.theme('DarkAmber')   # Add a touch of color

# All the stuff inside your window.
layout = [
            [sg.Text('Directorio de origen', size=(20,1)), sg.InputText(key='-ORIG-'), sg.FolderBrowse(button_text='Seleccionar')],
            [sg.Text('Directorio de destino', size=(20,1)), sg.InputText(key='-DEST-'), sg.FolderBrowse(button_text='Seleccionar')],
            [sg.Button('Procesar'), sg.Button('Salir'), sg.Text(size=(45,1)), sg.Button('Guardar config')]
        ]

# Create the Window
window = sg.Window('Generador de reportes', layout, finalize=True)
# Event Loop to process "events" and get the "values" of the inputs

# Read contents of INI file
if os.path.exists(ININAME):
    init_data = calclib.read_config(ININAME)
    window['-ORIG-'].Update(init_data['Principal']['Origen'])
    window['-DEST-'].Update(init_data['Principal']['Destino'])

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Salir': # if user closes window or clicks cancel
        break
    elif event == 'Procesar':
        print('You entered ', values['-ORIG-'], values['-DEST-'])
        print(values, event)
    elif event == 'Guardar config':
        calclib.save_config(values, ININAME)
        print(values)


window.close()