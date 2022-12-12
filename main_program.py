import pandas as pd
import PySimpleGUI as sg
import os
import xlsxwriter

import functions 

working_directory = os.getcwd()
left_col = [
 [sg.Text('Filename')],
 [sg.Input(key="-FILE_PATH-"), sg.FileBrowse(initial_folder=working_directory)], 
 [sg.Button("Procesar")],
 [sg.Text(key='-OUT-LOG-')]]


def Proceso(fecha, archivo):
    df_filtrado_por_fecha = functions.mask_dataframe_single_date(archivo, fecha)
    repartos = functions.cantidad_de_repartos(df_filtrado_por_fecha)
    df_menos_columnas, articulos_unicos = functions.dataframe_cleaning(df_filtrado_por_fecha)
    df_final = functions.suma_y_merge(df_menos_columnas, articulos_unicos)
    return df_final,repartos

def Exportar_excel(dataframe):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('salida/descargas.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    dataframe.to_excel(writer, sheet_name='Sheet1')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

right_col = [[ sg.Text('Fecha DES', size=(12,1))],
            [ sg.Input(key='-FECHA-', size=(12,1))], [sg.Exit()]]


layout = [[sg.Column(left_col, element_justification='left'),sg.Column(right_col, element_justification='left') ] ]

window = sg.Window("Procesamiento de descargas", layout)

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break

    elif event == "Procesar":
            file_path = values["-FILE_PATH-"]
            excel_file = pd.read_excel(file_path)
            window['-OUT-LOG-'].update("Se cargó el archivo necesario")
            fecha = values['-FECHA-']
            df_final,repartos = Proceso(fecha,excel_file)
            Exportar_excel(df_final)
            print(repartos)
window.close()