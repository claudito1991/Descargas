import pandas as pd
import PySimpleGUI as sg
import os

working_directory = os.getcwd()
layout = [
 [sg.Text('Filename')],
 [sg.Input(key="-FILE_PATH-"), sg.FileBrowse(initial_folder=working_directory)], 
 [sg.Button("Procesar")],
 [sg.Exit()], 
 [sg.Text(key='-OUT-LOG-')]]

window = sg.Window("Procesamiento de descargas", layout)

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break

    elif event == "Procesar":
            file_path = values["-FILE_PATH-"]
            excel_file = pd.read_excel(file_path)
            print(excel_file)
            window['-OUT-LOG-'].update("Se cargó el archivo necesario")
            

window.close()