# module CalPPT
import math
import openpyxl
from iapws import IAPWS97

from tkinter import *
from tkinter import ttk

# Determining the initial parameters - (Индекс 0 и Индекс kt'')
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
 
# Disposable heat transfer
def DHT(ho, hk):
    H0 = ho - hk
    return H0

# Nominal flow rate
def NFR(H0, N, ηoi_ηm_ηg):
    G0 = N*1000/(H0*ηoi_ηm_ηg)
    return G0

# Electrical power calculation
def EPC(H0, G0, ηoi_ηm_ηg):
    N = G0*H0*ηoi_ηm_ηg
    return N

# Assessment of pressure losses in the steam inlet organs
def APL_SIO(p0, del_p_div_p0):
    p_streak = p0*(1-del_p_div_p0)
    return p_streak

# Steam parameters after control valves - (Индекс 0 (штрих))
def SP_CV(p_streak, h0):
    steam_conval = IAPWS97(P=p_streak, h=h0)
    t_conval = steam_conval.T
    h_conval = h0
    s_conval = steam_conval.s
    v_conval = steam_conval.v
    return t_conval, h_conval, s_conval, v_conval

# Heat transfer of the regulating stage
def HTofRS(d_rs, n, u_div_cfi):
    H0_rs = (d_rs*float(math.pi)*n/float(u_div_cfi))**2/2000
    return H0_rs

# Steam parameters after the control stage (excluding losses) - (Индекс 2рсt)
def SM_CS_ExLos(H0_rs, h_conval, s_conval):
    h2rst = h_conval - H0_rs
    steam_conval = IAPWS97(h=h2rst, s=s_conval)
    s2rst = s_conval
    p2rst = steam_conval.P
    t2rst = steam_conval.T
    v2rst = steam_conval.v
    return h2rst, s2rst, v2rst, p2rst, t2rst

# Cost-effectiveness of the control stage
def COS_OF_CS(k_x_i, G0, v0):
    η0i_rs = k_x_i*(0.8-0.15/(G0*v0))                                               
    return η0i_rs

# Available thermal differential of the control stage
def ATD_CS(H0_rs, η0i_rs):
    Hi_rs = H0_rs*η0i_rs
    return Hi_rs

# Parameters after the control stage (taking into account losses) - (Индекс 2рс)
def PA_CS_TIAL(Hi_rs, h0, p2rst):
    p2rs = p2rst
    h2rs = h0 - Hi_rs
    steam_conval = IAPWS97(P=p2rs, h=h2rs)
    t2rs = steam_conval.T
    s2rs = steam_conval.s
    v2rs = steam_conval.v
    return h2rs, s2rs, v2rs, p2rs, t2rs

# Steam parameters at the output of an unregulated group of stages (excluding losses) - (Индекс kt)
def SP_UNG_ExL(pk, s2rs):
    p_kt = pk
    s_kt = s2rs
    steam_conval = IAPWS97(P=p_kt, s=s_kt)
    t_kt = steam_conval.T
    h_kt = steam_conval.h
    v_kt = steam_conval.v
    return h_kt, s_kt, v_kt, p_kt, t_kt

# Disposable heat transfer of a group of non-regulating steps
def DHT_GNoNREG(h2rs, h_kt):
    H0x = h2rs - h_kt
    return H0x

# Cost-effectiveness of a group of unregulated steps
def CosEff_GR_UnnReg(G0, v2rs, v_kt, H0x, k_vl):
    Gsr = G0
    v_sr = math.sqrt(v2rs*v_kt)
    η0i_gr = (0.92-0.2/(Gsr*v_sr))*(1+(H0x-700)/20000)*k_vl
    return Gsr, v_sr, η0i_gr

# Disposable thermal difference of a non-controllable group of steps
def DT_Diff_NonContrl_GS(H0x, η0i_gr):
    Hix = H0x*η0i_gr
    return Hix

# Steam parameters at the output of an unregulated group of stages (taking into account losses) - (Индекс k)
def SP_Out_UnReg_TIAL(pk, Hix, h2rs, h0):
    p_k = pk
    h_k = h2rs - Hix
    steam_conval = IAPWS97(P=p_k, h=h_k)
    t_k = steam_conval.T
    s_k = steam_conval.s
    v_k = steam_conval.v
    h_indk_diff = h0 - h_k
    return h_k, s_k, v_k, p_k, t_k, h_indk_diff

# EFFICIENCY of the flow part of the turbine
def EFFICIE_FlPaT(h_k, h_conval, H0):
    η0iCVD = (h_conval-h_k)/H0
    return η0iCVD

# Rec_NomFR -  Recalculating the nominal flow rate
def Rec_NomFR(h_indk_diff, η0iCVD, ηm, ηg, N, data_N):
    if data_N == True:
        G0_RetCal = N*1000/(η0iCVD*h_indk_diff*ηm*ηg)
    elif data_N == False:
        G0_RetCal = N/(η0iCVD*h_indk_diff*ηm*ηg)
    else:
        return print("Error for N")
    return G0_RetCal

def start(content, root):
    book = openpyxl.open("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    sheet = book.active

    # Configuration parameters
    ηoi_ηm_ηg = sheet['H21'].value
    del_p_div_p0 = sheet['C27'].value
    u_div_cfi = sheet['E37'].value
    k_x_i = sheet['D45'].value
    k_vl = sheet['F68'].value
    ηm = sheet['C86'].value
    ηg = sheet['C87'].value

    # Input parameters
    p0 = sheet['E3'].value
    t0 = sheet['F3'].value
    pk = sheet['G3'].value

    N = sheet['D3'].value
    G0 = sheet['K3'].value

    n = sheet['I3'].value
    H0_rs = sheet['J3'].value
    d_rs = sheet['H3'].value

    # DIP - Determining the initial parameters
    h0, s0, v0, hk, tk, vk = DIP(p0, t0, pk)

    line1 = ttk.Separator(content, orient=HORIZONTAL).grid(column=0, row=3, rowspan=2, columnspan=15, sticky='EW')

    print('1.1 Нахождение начальных парметров:', 'h0=', h0, 's0=', s0, 'v0=', v0)
    sheet['D9'] = h0
    sheet['E9'] = s0
    sheet['F9'] = v0
    DIPtext0 = ttk.Label(content, 
    text='1.1 Нахождение начальных парметров '+ '\n' 
    + 'h0=' + ' ' + str(h0) + '\n' 
    + ' s0=' + ' ' + str(s0) + '\n' 
    + ' v0=' + ' ' + str(v0))
    DIPtext0.grid(column=0, row=5, padx=5, rowspan=2, columnspan=3, sticky='EW')

    print('1.2 Нахождение конечных парметров:', 'hk=', hk, 'tk=', tk-273, 'vk=', vk)
    sheet['C14'] = tk-273
    sheet['D14'] = hk
    sheet['F14'] = vk
    DIPtextk = ttk.Label(content, 
    text='1.2 Нахождение конечных парметров: ' + '\n' 
    + 'hk=' + ' ' + str(hk) + '\n' 
    + ' tk=' + ' ' + str(tk-273) + '\n'
    + ' vk=' + ' ' + str(vk))
    DIPtextk.grid(column=0, row=8, padx=5, rowspan=2, columnspan=3, sticky='EW')

    line2 = ttk.Separator(content, orient=HORIZONTAL).grid(column=0, row=11, columnspan=15, sticky='EW')

    # DHT - Disposable heat transfer
    H0 = DHT(h0, hk)

    print('1.3 Располагаемый теплоперепад:', 'H0=', H0)
    sheet['C18'] = H0
    DHTtext = ttk.Label(content, 
    text=
    '1.3 Располагаемый теплоперепад: '+ '\n' + 
    'H0=' + ' ' + str(H0))
    DHTtext.grid(column=0, row=14, padx=5, rowspan=5, columnspan=3, sticky='EW')

    line3 = ttk.Separator(content, orient=HORIZONTAL).grid(column=0, row=19, columnspan=15, sticky='EW')

    # Checking whether the capacity is initially set for recalculation
    data_N = True

    # If calculation N or G0
    if N != 0 and G0 == 0:
        # NFR - Nominal flow rate
        G0 = NFR(H0, N, ηoi_ηm_ηg)

        print('1.4 Номинальная расход:', 'G0=', G0)
        sheet['C22'] = G0  
        sheet['H24'] = N
        FPaCTexts = ttk.Label(content, text=
        '1.4 Номинальная расход: '+ '\n' + 
        'G0=' + ' ' + str(G0))
        FPaCTexts.grid(column=0, row=21, padx=5, rowspan=5, columnspan=3, sticky='EW')

    elif G0 != 0 and N == 0: 
        # EPC - Electrical power calculation
        N = EPC(H0, G0, ηoi_ηm_ηg)

        data_N = False

        print('1.4 Находим электрическую мощносчть:', 'Nэ=', N)
        sheet['H24'] = N
        sheet['C22'] = G0  
        FPaCTexts = ttk.Label(content, text=
        '1.4 Находим электрическую мощносчть: '+ '\n' + 
        'Nэ=' + ' ' + str(N))
        FPaCTexts.grid(column=0, row=21, padx=5, rowspan=5, columnspan=3, sticky='EW')

    else: 
        return print("Error for N or G0")

    line4 = ttk.Separator(content, orient=HORIZONTAL).grid(column=0, row=26, columnspan=15, sticky='EW')

    # APL_SIO - Assessment of pressure losses in the steam inlet organs
    p_streak = APL_SIO(p0, del_p_div_p0)

    print('2. Оценка потерь давления в паровпускных органах:', 'p0`=', p_streak)
    sheet['F27'] = p_streak
    APLSIOText = ttk.Label(content, text=
    '2. Оценка потерь давления в паровпускных органах: '+ '\n' + 
    'p0`= ' + ' ' + str(p_streak))
    APLSIOText.grid(column=0, row=27, padx=5, rowspan=5, columnspan=3, sticky='EW')

    line5 = ttk.Separator(content, orient=HORIZONTAL).grid(column=0, row=32, columnspan=15, sticky='EW')

    # SP_CV - Steam parameters after control valves
    t_conval, h_conval, s_conval, v_conval = SP_CV(p_streak, h0)

    print('3. Параметры пара после регулирующих клапанов:', 'p0`=', p_streak, 't0`=', t_conval-273, 'h0`=', h_conval, 's0`=', s_conval, 'v0`=', v_conval)
    sheet['C32'] = t_conval-273
    sheet['D32'] = h_conval
    sheet['E32'] = s_conval
    sheet['F32'] = v_conval
    SPCVtext = ttk.Label(content, 
    text=
    '3. Параметры пара после регулирующих клапанов: ' + '\n' + 
    'p0`=' + ' ' + str(p_streak) + '\n' + 
    't0`=' + ' ' + str(t_conval-273) + '\n' + 
    'h0`=' + ' ' + str(h_conval) + '\n' +
    's0`=' + ' ' + str(s_conval) + '\n' +
    'v0`=' + ' ' + str(v_conval))
    SPCVtext.grid(column=0, row=35, padx=5, rowspan=2, columnspan=3, sticky='EW')

    # If calculation H0_rs
    if H0_rs == 0 and d_rs != 0:
        #We find H0_rs by d_rs and n and u_div_cfi

        #HTofRS - Heat transfer of the regulating stage
        H0_rs = HTofRS(d_rs, n, u_div_cfi)

        print('4.1.1 Находим теплоперепад в рег. ступени', 'H0рс =', H0_rs)
        sheet['E38'] = H0_rs

    elif H0_rs != 0 and d_rs == 0: 
        #The heat transfer H0_rs is initially set
        print('4.1.1 Находим теплоперепад в рег. ступени', 'H0рс =', H0_rs)
        sheet['E38'] = H0_rs

    else: 
        return print("Error for H0_rs or d_rs")

    
    # SM_CS_ExLos - Steam parameters after the control stage (excluding losses)
    h2rst, s2rst, v2rst, p2rst, t2rst = SM_CS_ExLos(H0_rs, h_conval, s_conval)

    print('4.1 Параметры пара после регулирующей ступени (без учета потерь)', 'p2рсt=', p2rst, 't2рсt=', t2rst-273, 'h2рсt=', h2rst, 's2рсt=', s2rst, 'v2рсt=', v2rst)
    sheet['B42'] = p2rst
    sheet['C42'] = t2rst-273
    sheet['D42'] = h2rst
    sheet['E42'] = s2rst
    sheet['F42'] = v2rst

    # COS_OF_CS - Cost-effectiveness of the control stage
    η0i_rs = COS_OF_CS(k_x_i, G0, v0)

    print('4.2 Экономичность регулирующей ступени', 'η0iр.c. =', η0i_rs)
    sheet['D46'] = η0i_rs

    # ATD_CS - Available thermal differential of the control stage
    Hi_rs = ATD_CS(H0_rs, η0i_rs)

    print('4.3 Располагаемый тепловой перепад регулирующей ступени', 'Hipc =', Hi_rs)
    sheet['C50'] = Hi_rs

    # PA_CS_TIAL - Parameters after the control stage (taking into account losses)
    h2rs, s2rs, v2rs, p2rs, t2rs = PA_CS_TIAL(Hi_rs, h0, p2rst)

    print('4.4 Параметры после регулирующей ступени (с учетом потерь)', 'p2рс=', p2rs, 't2рс=', t2rs-273, 'h2рс=', h2rs, 's2рс=', s2rs, 'v2рс=', v2rs)
    sheet['B55'] = p2rs
    sheet['C55'] = t2rs-273
    sheet['D55'] = h2rs
    sheet['E55'] = s2rs
    sheet['F55'] = v2rs

    # SP_UNG_ExL - Steam parameters at the output of an unregulated group of stages (excluding losses)
    h_kt, s_kt, v_kt, p_kt, t_kt = SP_UNG_ExL(pk, s2rs)

    print('5. Параметры пара на выходе из нерегулируещей группы ступеней (без учета потерь)', 'p_kt=', p_kt, 't_kt=', t_kt-273, 'h_kt=', h_kt, 's_kt=', s_kt, 'v_kt=', v_kt)
    sheet['B60'] = p_kt
    sheet['C60'] = t_kt-273
    sheet['D60'] = h_kt
    sheet['E60'] = s_kt
    sheet['F60'] = v_kt

    # DHT_GNoNREG - Disposable heat transfer of a group of non-regulating steps
    H0x = DHT_GNoNREG(h2rs, h_kt)

    print('5.1 Располагаемый теплоперепад группы нерегулирующих ступеней', 'H0х =', H0x)
    sheet['C64'] = H0x

    # CosEff_GR_UnnReg - Cost-effectiveness of a group of unregulated steps
    Gsr, v_sr, η0i_gr = CosEff_GR_UnnReg(G0, v2rs, v_kt, H0x, k_vl)

    print('5.2 Экономичность группы нерегулируещих ступеней', 'Gср=G0=', Gsr, 'vср=', v_sr, 'η0iгр=', η0i_gr)
    sheet['C68'] = Gsr
    sheet['C69'] = v_sr
    sheet['C70'] = η0i_gr

    # DT_Diff_NonContrl_GS - Disposable thermal difference of a non-controllable group of steps
    Hix = DT_Diff_NonContrl_GS(H0x, η0i_gr)

    print('5.3 Располагаемый тепловой перепад неругулируещей группы ступеней', 'Hix=', Hix)
    sheet['C74'] = Hix

    # SP_Out_UnReg_TIAL - Steam parameters at the output of an unregulated group of stages (taking into account losses)
    h_k, s_k, v_k, p_k, t_k, h_indk_diff = SP_Out_UnReg_TIAL(pk, Hix, h2rs, h0)

    print('5.4 Параметры пара на выходе из нерегулируещей группы ступеней (с учетом потерь)', 'p_k=', p_k, 't_k=', t_k-273, 'h_k=', h_k, 's_k=', s_k, 'v_k=', v_k, 'h_indk_diff=', h_indk_diff)
    sheet['B79'] = p_k
    sheet['C79'] = t_k-273
    sheet['D79'] = h_k
    sheet['E79'] = s_k
    sheet['F79'] = v_k
    sheet['D80'] = h_indk_diff

    # EFFICIE_FlPaT - EFFICIENCY of the flow part of the turbine
    η0iCVD = EFFICIE_FlPaT(h_k, h_conval, H0)

    print('6. КПД проточной части турбины', 'η0iЦВД=', η0iCVD)
    sheet['C83'] = η0iCVD

    # Rec_NomFR -  Recalculating the nominal flow rate
    G0_RetCal = Rec_NomFR(h_indk_diff, η0iCVD, ηm, ηg, N, data_N)

    print('8. Пересчитываем номинальный расход', 'G0=', G0_RetCal)
    sheet['C91'] = G0_RetCal

    # Save exel file (DB)
    book.save("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    book.close()

if __name__ == "__main__":
    print("This is the module - Сalculation of process parameters in a turbine")