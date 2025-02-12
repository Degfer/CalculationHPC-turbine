# module CalPPT
import openpyxl
from iapws import IAPWS97

p0 = None
t0 = None
pk = None

def DIP(p0, t0, pk):
    steam0 = IAPWS97(P=p0, T=t0+273)
    h0 = steam0.h
    s0 = steam0.s
    v0 = steam0.v

    steamk = IAPWS97(P=pk, s=s0)
    hk = steamk.h
    tk = steamk.T
    vk = steamk.v
    return h0, s0, v0, hk, tk, vk

def start():
    book = openpyxl.open("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    sheet = book.active

    p0 = sheet['E3'].value
    t0 = sheet['F3'].value
    pk = sheet['G3'].value

    h0, s0, v0, hk, tk, vk = DIP(p0, t0, pk)
    sheet['D9'] = h0
    sheet['E9'] = s0
    sheet['F9'] = v0

    sheet['C14'] = tk-273
    sheet['D14'] = hk
    sheet['F14'] = vk

    book.save("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    book.close()


if __name__ == "__main__":
    print("This is the module - Сalculation of process parameters in a turbine")