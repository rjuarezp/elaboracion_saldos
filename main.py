import PySimpleGUI as sg

sg.theme('DarkAmber')   # Add a touch of color
# All the stuff inside your window.
layout = [  [sg.Text('Analizar datos')],
            [sg.Text('Directorio de origen', size=(20,1)), sg.InputText(key='-ORIG-'), sg.FolderBrowse()],
            [sg.Text('Directorio de destino', size=(20,1)), sg.InputText(key='-DEST-'), sg.FolderBrowse()],
            [sg.Button('Procesar'), sg.Button('Cancel')] ]

# Create the Window
window = sg.Window('Título de la ventana', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    elif event == 'Procesar':
        print('You entered ', values['-ORIG-'], values['-DEST-'])

window.close()