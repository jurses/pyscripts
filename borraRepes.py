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
import re

global allRow

def openFile():
    global fileName
    fileName = filedialog.askopenfilename(initialdir = ".", title = "Select file", filetypes = (("Archivos Excel","*.xlsx"), ("Cualquier archivo","*.*")))
    print(fileName)

def cellIsInVCells(cell, vCells):
    for v in vCells:
        if cell.value == v.value:
            return True
    
    return False

def getVCells(row, refCols, min_col = 0):
    vCells = []
    for col in refCols:
        vCells.append(row[col - min_col])
    
    return vCells

#   Obtiene las columnas de referencias pasadas por una cadena
#   Para ser usadas con openpyxl
def getVRefCols(columns_string):
    columns_string = columns_string.replace(" ", "")
    vCols = []
    if len(columns_string) > 1:
        for v in columns_string.split(','):
            vCols.append(column_index_from_string(v))
    else:
        vCols.append(column_index_from_string(columns_string))
    
    return vCols

def getCornerMatrix(matrix):
    matrixSelected = []

    matrix = matrix.replace(" ", "")
    matrix = matrix.split(':')
    
    matrixSelected.append("".join(re.split("[^a-zA-Z]*", matrix[0])))
    matrixSelected.append("".join(re.split("[^0-9]*", matrix[0])))
    matrixSelected.append("".join(re.split("[^a-zA-Z]*", matrix[1])))
    matrixSelected.append("".join(re.split("[^0-9]*", matrix[1])))

    return matrixSelected

def removeUntilNextChange():
    assert(fileName)
    wb = load_workbook(filename=fileName, data_only=True)
    ws = wb[e1.get()]
    print(e1.get())
    matrixSelected = getCornerMatrix(e2.get())
    vRefCols = getVRefCols(e5.get())
    print(e5.get())

    currentValue = None

    for col in vRefCols:
        vCell = getVCells(cellRows, vRefCols, column_index_from_string(col))
        if cellIsInVCells(cellRows[col], vCell):
            if allRow.get():


    for cellRows in ws.iter_rows(   min_col = column_index_from_string(matrixSelected[0]),
                                    max_col = column_index_from_string(matrixSelected[2]), 
                                    min_row = int(matrixSelected[1]),
                                    max_row = int(matrixSelected[3])
                                ):
        for col in vRefCols:
            vCell = getVCells(cellRows, vRefCols, column_index_from_string(col))
            if cellIsInVCells(cellRows[col], vCell):
                if allRow.get():
                else:
                    cellRows[col] = None

    wb.save("{}_modBorraRepes.xlsx".format(fileName))
        
root = Tk()
root.title("Herramientas Excel")

root.geometry("500x500")
root.resizable(0, 0)

allRow = BooleanVar()


l1 = Label(root, text="Hoja del libro").grid(row = 0)
l2 = Label(root, text = "Seleccione la matriz\nEjemplo: D3:F5").grid(row = 1)
l5 = Label(root, text = "Columnas referentes\nSeparadas por ','").grid(row = 2)
Button(root, text="Aceptar", command=removeUntilNextChange).grid(row = 3, column = 1, sticky=W, pady=4)
Button(root, text="Salir", command=root.quit).grid(row = 3, column = 2, sticky=W, pady=4)
Button(root, text = "Abrir archivo", command = openFile).grid(row = 3, column = 0, sticky = W, pady = 4)
Checkbutton(root, text = "Aplicar eliminación a toda la fila", variable = allRow, onvalue = TRUE, offvalue = FALSE).grid(row = 1, column = 2, sticky = W, pady = 4)

e1 = Entry(root)
e2 = Entry(root)
e5 = Entry(root)

e1.grid(row = 0, column = 1)
e2.grid(row = 1, column = 1)
e5.grid(row = 2, column = 1)


mainloop()