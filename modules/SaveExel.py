import openpyxl

def Save_tab1(cell, data):
    book = openpyxl.open("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    sheet = book.active

    # row = 2
    # for datas in data:
    #     sheet[3][row].value = datas
    #     row+=1
    sheet[cell] = data


    book.save("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    book.close()