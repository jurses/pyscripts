#!/usr/bin/env python3
'''
Copyright 2018 Raúl Ulises Martín Hernández

Se concede permiso por la presente, libre de cargos, a cualquier persona que obtenga una copia de este software y de los archivos de documentación asociados (el "Software"), a utilizar el Software sin restricción, incluyendo sin limitación los derechos a usar, copiar, modificar, fusionar, publicar, distribuir, sublicenciar, y/o vender copias del Software, y a permitir a las personas a las que se les proporcione el Software a hacer lo mismo, sujeto a las siguientes condiciones:

El aviso de copyright anterior y este aviso de permiso se incluirán en todas las copias o partes sustanciales del Software.

EL SOFTWARE SE PROPORCIONA "COMO ESTÁ", SIN GARANTÍA DE NINGÚN TIPO, EXPRESA O IMPLÍCITA, INCLUYENDO PERO NO LIMITADO A GARANTÍAS DE COMERCIALIZACIÓN, IDONEIDAD PARA UN PROPÓSITO PARTICULAR E INCUMPLIMIENTO. EN NINGÚN CASO LOS AUTORES O PROPIETARIOS DE LOS DERECHOS DE AUTOR SERÁN RESPONSABLES DE NINGUNA RECLAMACIÓN, DAÑOS U OTRAS RESPONSABILIDADES, YA SEA EN UNA ACCIÓN DE CONTRATO, AGRAVIO O CUALQUIER OTRO MOTIVO, DERIVADAS DE, FUERA DE O EN CONEXIÓN CON EL SOFTWARE O SU USO U OTRO TIPO DE ACCIONES EN EL SOFTWARE. 
'''

from tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.utils import *

def openFile():
    global fileName
    fileName = filedialog.askopenfilename(initialdir = ".", title = "Select file", filetypes = (("Archivos Excel","*.xlsx"), ("Cualquier archivo","*.*")))

def removeWholeRow(cell, ws):
    for cellRows in ws.iter_rows(min_col = column_index_from_string("A"), max_col = 70, min_row = cell.row, max_row = cell.row):
        for x in cellRows:
            x.value = None

def removeUntilNextChange():
    wb = load_workbook(filename=fileName, data_only=True)
    ws = wb[e1.get()]
    col = column_index_from_string(e2.get())
    row_init = int(e3.get())
    row_end = int(e4.get())

    newValue = None
    for cellRows in ws.iter_rows(min_col = col, max_col = col, min_row = row_init, max_row = row_end):
        if newValue != cellRows[0].value:
            newValue = cellRows[0].value
        else:
            if allRow.get():
                removeWholeRow(cellRows[0], ws)

            cellRows[0].value = None

    wb.save("{}_modBorraRepes.xlsx".format(fileName))
        
root = Tk()
root.title("Herramientas Excel")
global allRow
allRow = BooleanVar()
fileName = None

root.geometry("500x500")
root.resizable(0, 0)

Label(root, text="Hoja del libro").grid(row = 0)
Label(root, text="Columna").grid(row = 1)
Label(root, text="Fila de inicio").grid(row = 2)
Label(root, text="Fila final").grid(row = 3)
#Label(root, text="Nombre del archivo: {}".format(fileName)).grid(row = 5)
Button(root, text="Aceptar", command=removeUntilNextChange).grid(row = 4, column = 1, sticky=W, pady=4)
Button(root, text="Salir", command=root.quit).grid(row = 4, column = 2, sticky=W, pady=4)
Button(root, text = "Abrir archivo", command = openFile).grid(row = 4, column = 0, sticky = W, pady = 4)
Checkbutton(root, text = "Aplicar eliminación a toda la fila", variable = allRow, onvalue = TRUE, offvalue = FALSE).grid(row = 3, column = 2, sticky = W, pady =4)

e1 = Entry(root)
e2 = Entry(root)
e3 = Entry(root)
e4 = Entry(root)

e1.grid(row = 0, column = 1)
e2.grid(row = 1, column = 1)
e3.grid(row = 2, column = 1)
e4.grid(row = 3, column = 1)

mainloop()