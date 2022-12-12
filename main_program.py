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


def Make_directory():
    #python program to check if a directory exists
    working_directory = os.getcwd()
    # Check whether the specified path exists or not
    isExist = os.path.exists(working_directory+"/salida")
    if not isExist:
        # Create a new directory because it does not exist
        os.makedirs(working_directory+"/salida")
        print("The new directory is created!")

def Reverse_string_from_user(string):
    return string[-4:] + "-" + string[:2]+"-"+string[3:5]
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
    dataframe.to_excel(writer, sheet_name='descargas consolidado')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

def CreateTable(dataframe):
    
    dataframe = dataframe.reset_index()
    print("longitud de tabla: ", len(dataframe) )
    longitud_tabla = len(dataframe)
    headers = list(dataframe.columns.astype(str))
    #headers=[StringFunction(x) for x in headers]
    table_values = dataframe.values.tolist()
    table = [[sg.Table(values = table_values, headings=headers,
                max_col_width=35,
                auto_size_columns=True,
                justification='center',
                num_rows=longitud_tabla,
                key='-TABLE-',
                row_height=35)]]
    layout = [[sg.Column(table, element_justification='left') ] ]
  
    window = sg.Window('Descargas consolidado', layout,resizable=True)
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
    window.close()




right_col = [[ sg.Text('Fecha DES', size=(12,1))],
            [ sg.Input(key='-FECHA-', size=(12,1))], [sg.Exit()]]


layout = [[sg.Column(left_col, element_justification='left'),sg.Column(right_col, element_justification='left') ] ]

window = sg.Window("Procesamiento de descargas", layout)

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break

    elif event == "Procesar":
            Make_directory()
            file_path = values["-FILE_PATH-"]
            excel_file = pd.read_excel(file_path)
            window['-OUT-LOG-'].update("Se cargó el archivo necesario")
            fecha = Reverse_string_from_user(values['-FECHA-'])
            df_final,repartos = Proceso(fecha,excel_file)
            Exportar_excel(df_final)
            print(repartos)
            CreateTable(df_final)
window.close()