import PySimpleGUI as sg
import tkinter, re
import PIL, PIL.ImageFont
import pandas, xlrd, qrcode, os

qr = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_L,
    box_size=10,
    border=10,
)

#declarar objetos dentro de la ventana
layout = [  [sg.Text('Selecciona un archivo de excel para generar los códigos QR')],
            [sg.Text('Ingresar la ubicación del Excel:              '), sg.InputText(key='input_ruta'),sg.Button('Examinar')],
            [sg.Text('Ingresa la ubicación para guardar tus QR'),sg.InputText(key='carpeta_salida'),sg.Button('Examinar',key="ex_salida")],
            [sg.Button('Cerrar Ventana'), sg.Button('Generar')] ]

# Crear la ventana
window = sg.Window('Excel to Massive QR by @contreraspablo9', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == 'Examinar':
        #abrir explorador de archivos
        root = tkinter.Tk() #esto se hace solo para eliminar la ventanita de Tkinter 
        root.withdraw() #ahora se cierra 
        file_path = tkinter.filedialog.askopenfilename() #abre el explorador de archivos y guarda la seleccion en la variable!
        window['input_ruta'].update(file_path)#escribe la ubicacion seleccionada

    if event == 'ex_salida':
        root = tkinter.Tk() #esto se hace solo para eliminar la ventanita de Tkinter 
        root.withdraw() #ahora se cierra 
        salida_path = tkinter.filedialog.askdirectory() #abre el explorador de archivos y guarda la carpeta en la variable!
        window['carpeta_salida'].update(salida_path)#escribe la ruta seleccionada


    if event == sg.WIN_CLOSED or event == 'Cerrar Ventana': # if user closes window or clicks cancel
        break

    if event == 'Generar':
        file_path = values['input_ruta']
        df = pandas.read_excel(file_path)  #extrae la hoja de excel en una variable

        #verificar que exista la carpeta de destino:
        salida_path = salida_path + str('/QRCodes/')
        if not(os.path.isdir(salida_path)):
            os.mkdir(salida_path)


        #print(df.iloc[0,0]) #acceder a una casilla en especifico
        columnas = df.shape[1] 
        renglones = df.shape[0]
        codigoqr = [0 for x in range(renglones)]

        #generar codigos qr:
        for qr_actual in range(renglones):
            texto = ''
            for info in range(columnas):
                texto = texto + str(df.iloc[qr_actual,info]) + ' '
            codigoqr[qr_actual] = qrcode.make(texto)#generar qr
            draw = PIL.ImageDraw.Draw(codigoqr[qr_actual])
            font = PIL.ImageFont.truetype("arial", 16)
            draw.text((10, 10),str(qr_actual./),font=font)
            codigoqr[qr_actual].save(salida_path+ str(qr_actual) + '.png') #guardar qr
        
window.close()

