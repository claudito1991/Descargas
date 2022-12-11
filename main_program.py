import pandas as pd
import PySimpleGUI as sg

event, values = sg.Window('Get filename example', [
 [sg.Text('Filename')],
 [sg.Input(), sg.FileBrowse()], 
 [sg.OK(), sg.Cancel()]
 [sg.Button("Procesar")] ]).read(close=True)