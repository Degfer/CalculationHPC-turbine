# module CalPPT2
import math
import random
import openpyxl
# import pandas as pd
from pycel import ExcelCompiler

from iapws import IAPWS97

from tkinter import *
from tkinter import ttk

# Finding pairs of steam meters at point 1t - (Индекс 1t)
def FinPairSM_i1t(h1t, s1t):
    steam_conval = IAPWS97(h=h1t, s=s1t)
    p1t = steam_conval.P
    t1t = steam_conval.T
    v1t = steam_conval.v
    return h1t, s1t, v1t, p1t, t1t 

# Calculation of the specific volume for the auxiliary table number 1
def CS_VL_AT_N1(h2_rs, s11t, H0st1,  H0st2, H0st3, H0st4, H0st5, H0st10):
    z1_steam_conval = IAPWS97(h=(h2_rs-H0st1), s=s11t)
    z2_steam_conval = IAPWS97(h=(h2_rs-H0st2), s=s11t)
    z3_steam_conval = IAPWS97(h=(h2_rs-H0st3), s=s11t)
    z4_steam_conval = IAPWS97(h=(h2_rs-H0st4), s=s11t)
    z5_steam_conval = IAPWS97(h=(h2_rs-H0st5), s=s11t)
    z10_steam_conval = IAPWS97(h=(h2_rs-H0st10), s=s11t)
    
    z1_v1t = z1_steam_conval.v
    z2_v1t = z2_steam_conval.v
    z3_v1t = z3_steam_conval.v
    z4_v1t = z4_steam_conval.v
    z5_v1t = z5_steam_conval.v
    z10_v1t = z10_steam_conval.v

    return z1_v1t, z2_v1t, z3_v1t, z4_v1t, z5_v1t, z10_v1t 

# Selecting the height value of the last step
def SL_VLS_Random(dk, F2_divide_PI):
    discrepancy = 1
    while True:
        l2_recal = random.uniform(0, 1)
        dll = (dk/1000+l2_recal)*l2_recal
        discrepancy = (dll-F2_divide_PI)*1000
        if (discrepancy < 0.001) and (discrepancy > 0): 
            break
    return l2_recal, discrepancy

# Calculation of the specific volume for the auxiliary table number 2
def CS_VL_AT_N2(h_k, s_k, H0z1, H0z2, H0z3, H0z4, H0z5):
    Hz1_steam_conval = IAPWS97(h=(h_k+H0z1), s=s_k)
    Hz2_steam_conval = IAPWS97(h=(h_k+H0z2), s=s_k)
    Hz3_steam_conval = IAPWS97(h=(h_k+H0z3), s=s_k)
    Hz4_steam_conval = IAPWS97(h=(h_k+H0z4), s=s_k)
    Hz5_steam_conval = IAPWS97(h=(h_k+H0z5), s=s_k)

    Hz1_v1t = Hz1_steam_conval.v
    Hz2_v1t = Hz2_steam_conval.v
    Hz3_v1t = Hz3_steam_conval.v
    Hz4_v1t = Hz4_steam_conval.v
    Hz5_v1t = Hz5_steam_conval.v

    return Hz1_v1t, Hz2_v1t, Hz3_v1t, Hz4_v1t, Hz5_v1t

# Selection of the value of the root diameter of the last stage
def SL_VRD_LS_Random(dz1, excel):
    while True:
        dz = random.uniform(dz1, 2)

        book = openpyxl.open("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
        sheet = book.active
        
        sheet['T131'] = dz

        # Save exel file (DB)
        book.save("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
        book.close()

        discrepancy_dz = excel.evaluate('Лист1!T142')

        if (discrepancy_dz < 0.01) and (discrepancy_dz > 0):
            break
    return dz, discrepancy_dz 

def start(content, root):
    book = openpyxl.open("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    sheet = book.active

    # Configuration parameters
    ro = sheet['E97'].value
    alpha1_eff = sheet['E98'].value
    u_div_cfi = sheet['E99'].value
    μ1 = sheet['C106'].value

    # Uploading data from DB
    G0 = sheet['C91'].value
    H0_rs = sheet['E38'].value
    
    # Load file
    excel = ExcelCompiler(filename="C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")

    # Load data for FinPairSM_i1t
    h1t = excel.evaluate('Лист1!L17')
    s1t = excel.evaluate('Лист1!M17')

    # FinPairSM_i1t - Finding pairs of steam meters at point 1t 
    h1t, s1t, v1t, p1t, t1t = FinPairSM_i1t(h1t, s1t)

    print('Нахождепние парметров в точке 1t', 'p1t=', p1t, 't1t=', t1t-273, 'h1t=', h1t, 's1t=', s1t, 'v1t=', v1t)
    sheet['J17'] = p1t
    sheet['K17'] = t1t-273
    sheet['N17'] = v1t

    # To display the data in the point 2.1
    C1t = excel.evaluate('Лист1!G103')
    el1 = excel.evaluate('Лист1!C105')
    e_opt = excel.evaluate('Лист1!C107')
    e_recal = excel.evaluate('Лист1!D108')
    l1 = excel.evaluate('Лист1!C110')
    print('2.1 Определение диаметра и высоты лопаток регулирующей ступени', 'C1t=', C1t, 'el1=', el1, 'e_opt=', e_opt, 'e_recal=', e_recal, 'l1=', l1)

    # To display the data in the point 3.1
    d1 = excel.evaluate('Лист1!C118')
    print('3.1 Диаметр первой нерегулируемой ступени', 'd1=', d1)

    # To display the data in the point 3.2
    H01 = excel.evaluate('Лист1!C120')
    print('3.2 Располагаемый теплоперепад на ступень', 'H01=', H01)

    # To display the data in the point 3.3
    С1t1 = excel.evaluate('Лист1!C124')
    print('3.3 Скорость на выходе из сопловой решетки', 'С1t1=', С1t1)

    # Load data for CS_VL_AT_N1
    h2_rs = excel.evaluate('Лист1!L12')
    s11t =  excel.evaluate('Лист1!M18')

    H0st1 = excel.evaluate('Лист1!M115')
    H0st2 = excel.evaluate('Лист1!N115')
    H0st3 = excel.evaluate('Лист1!O115')
    H0st4 = excel.evaluate('Лист1!P115')
    H0st5 = excel.evaluate('Лист1!Q115')
    H0st10 = excel.evaluate('Лист1!R115')

    # CS_VL_AT_N1 - Calculation of the specific volume for the auxiliary table number 1
    z1_v1t, z2_v1t, z3_v1t, z4_v1t, z5_v1t, z10_v1t = CS_VL_AT_N1(h2_rs, s11t, H0st1,  H0st2, H0st3, H0st4, H0st5, H0st10)

    print('Расчет удельного объема для вспомогательной таблицы №1', 'z1_v1t=', z1_v1t, 'z2_v1t=', z2_v1t, 'z3_v1t=', z3_v1t, 'z4_v1t=', z4_v1t, 'z5_v1t=', z5_v1t, 'z10_v1t=', z10_v1t)
    sheet['M124'] = z1_v1t
    sheet['N124'] = z2_v1t
    sheet['O124'] = z3_v1t
    sheet['P124'] = z4_v1t
    sheet['Q124'] = z5_v1t
    sheet['R124'] = z10_v1t

    # Save exel file (DB)
    book.save("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    book.close()

    # To display the data in the point 3.4
    l1_recal = excel.evaluate('Лист1!C127')
    print('3.4 Высота сопловой решетки', 'l1_recal=', l1_recal)

    # To display the data in the point 3.5
    l2 = excel.evaluate('Лист1!C130')
    print('3.5 Высота рабочей решетки', 'l2=', l2)

    # To display the data in the point 3.6
    dk = excel.evaluate('Лист1!C133')
    print('3.6 Корневой диаметр', 'dk=', dk)

    # To display the data in the point 4.1
    F2 = excel.evaluate('Лист1!C139')
    c2 = excel.evaluate('Лист1!F137')
    print('4.1 Выходная площадь рабочей решетки', 'F2=', F2, 'c2=', c2)

    book = openpyxl.open("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    sheet = book.active

    # Load data for SL_VLS_Random
    F2_divide_PI =  excel.evaluate('Лист1!J140')

    # SL_VLS_Random - Selecting the height value of the last step
    l2_recal, discrepancy = SL_VLS_Random(dk, F2_divide_PI)

    print('Решим квадратное уравнение', 'Отсюда l2=', l2_recal, 'невязка=', discrepancy)
    sheet['F140'] = l2_recal

    # Load data for CS_VL_AT_N2
    h_k = excel.evaluate('Лист1!L14')
    s_k =  excel.evaluate('Лист1!M14')

    H0z1 = excel.evaluate('Лист1!P136')
    H0z2 = excel.evaluate('Лист1!Q136')
    H0z3 = excel.evaluate('Лист1!R136')
    H0z4 = excel.evaluate('Лист1!S136')
    H0z5 = excel.evaluate('Лист1!T136')

    # CS_VL_AT_N2 - Calculation of the specific volume for the auxiliary table number 2
    Hz1_v1t, Hz2_v1t, Hz3_v1t, Hz4_v1t, Hz5_v1t = CS_VL_AT_N2(h_k, s_k, H0z1,  H0z2, H0z3, H0z4, H0z5)

    print('Расчет удельного объема для вспомогательной таблицы №2', 'Hz1_v1t=', Hz1_v1t, 'Hz2_v1t=', Hz2_v1t, 'Hz3_v1t=', Hz3_v1t, 'Hz4_v1t=', Hz4_v1t, 'Hz5_v1t=', Hz5_v1t)
    sheet['P139'] = Hz1_v1t
    sheet['Q139'] = Hz2_v1t
    sheet['R139'] = Hz3_v1t
    sheet['S139'] = Hz4_v1t
    sheet['T139'] = Hz5_v1t

    # Save exel file (DB)
    book.save("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    book.close()

    # Load data for SL_VRD_LS_Random
    dz1 = excel.evaluate('Лист1!P131')

    # # SL_VRD_LS_Random - Selection of the value of the root diameter of the last stage
    # dz, discrepancy_dz = SL_VRD_LS_Random(dz1, excel)

    # print('Подбор корневого диаметра последней ступени', 'dz=', dz, 'невязка=', discrepancy_dz)

if __name__ == "__main__":
    print("This is the module - Сalculation of process parameters in a turbine 2")