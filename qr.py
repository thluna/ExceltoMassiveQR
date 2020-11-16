import PySimpleGUI as sg
import tkinter, re, PIL
import pandas, xlrd, qrcode, os

def verificar_carpeta():
    if not(os.path.isdir('/Users/contr/Desktop/CodigosQR/')):
        os.mkdir('/Users/contr/Desktop/CodigosQR/')
    
#declarar objetos dentro de la ventana
layout = [  [sg.Text('Selecciona un archivo de excel para generar los códigos QR')],
            [sg.Text('Ingresar la ubicación del archivo:'), sg.InputText(key='input_ruta'),sg.Button('Examinar')],
            [sg.Button('Cerrar Ventana'), sg.Button('Generar')] ]

# Crear la ventana
window = sg.Window('Massive QR', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == 'Examinar':
        #abrir explorador de archivos
        root = tkinter.Tk() #esto se hace solo para eliminar la ventanita de Tkinter 
        root.withdraw() #ahora se cierra 
        file_path = tkinter.filedialog.askopenfilename() #abre el explorador de archivos y guarda la seleccion en la variable!
        window['input_ruta'].update(file_path)#escribe la ubicacion seleccionada

    if event == sg.WIN_CLOSED or event == 'Cerrar Ventana': # if user closes window or clicks cancel
        break

    if event == 'Generar':
        file_path = values['input_ruta']
        df = pandas.read_excel(file_path)  #extrae la hoja de excel en una variable
        #print(df.iloc[0,0]) #acceder a una casilla en especifico
        columnas = df.shape[1] 
        renglones = df.shape[0]
        codigoqr = [0 for x in range(renglones)]
        verificar_carpeta()
        for qr_actual in range(renglones):
            texto = ''
            for info in range(columnas):
                texto = texto + str(df.iloc[qr_actual,info]) + ' '
            codigoqr[qr_actual] = qrcode.make(texto)
            codigoqr[qr_actual].save('/Users/contr/Desktop/CodigosQR/'+ str(qr_actual) + '.png')
        
window.close()
