import openpyxl
import os

relative_path_DB = "DB/"
absolute_path_DB = os.path.abspath(relative_path_DB)

def Save_tab1(cell, data):
    book = openpyxl.open(absolute_path_DB + "/raschet_turboagregata_predvaritelny.xlsx")
    sheet = book.active

    # row = 2
    # for datas in data:
    #     sheet[3][row].value = datas
    #     row+=1
    sheet[cell] = data

    book.save(absolute_path_DB + "/raschet_turboagregata_predvaritelny.xlsx")
    book.close()