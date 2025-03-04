# module CalPPT
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

def start(content):
    book = openpyxl.open("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    sheet = book.active

    # Configuration parameters
    ηoi_ηm_ηg = sheet['H21'].value
    del_p_div_p0 = sheet['C27'].value

    # Input parameters
    p0 = sheet['E3'].value
    t0 = sheet['F3'].value
    pk = sheet['G3'].value

    N = sheet['D3'].value
    G0 = sheet['K3'].value

    n = sheet['I3'].value
    Ho_rs = sheet['J3'].value

    # DIP - Determining the initial parameters
    h0, s0, v0, hk, tk, vk = DIP(p0, t0, pk)

    print('1.1 Нахождение начальных парметров:', 'h0=', h0, 's0=', s0, 'v0=', v0)
    sheet['D9'] = h0
    sheet['E9'] = s0
    sheet['F9'] = v0
    DIPtext = ttk.Label(content, text='1.1 Нахождение начальных парметров ' + '\n' + 'h0=' + ' ' + str(h0) + '\n' + ' s0=' + ' ' + str(s0) + '\n' + ' v0=' + ' ' + str(v0))
    DIPtext.grid(column=0, row=5, padx=5, rowspan=2, columnspan=3)

    print('1.2 Нахождение конечных парметров:', 'hk=', hk, 'tk=', tk-273, 'vk=', vk)
    sheet['C14'] = tk-273
    sheet['D14'] = hk
    sheet['F14'] = vk

    # DHT - Disposable heat transfer
    H0 = DHT(h0, hk)

    print('1.3 Располагаемый теплоперепад:', 'H0=', H0)
    sheet['C18'] = H0

    # If calculation N or G0
    if N != 0:
        # NFR - Nominal flow rate
        G0 = NFR(H0, N, ηoi_ηm_ηg)

        print('1.4 Номинальная расход:', 'G0=', G0)
        sheet['C22'] = G0

    elif G0 != 0:
        # EPC - Electrical power calculation
        N = EPC(H0, G0, ηoi_ηm_ηg)

        print('1.4 Находим электрическую мощносчть:', 'Nэ=', N)
        sheet['H24'] = N

    else: print("Error for N or G0")

    # APL_SIO - Assessment of pressure losses in the steam inlet organs
    p_streak = APL_SIO(p0, del_p_div_p0)

    print('2. Оценка потерь давления в паровпускных органах:', 'p0`=', p_streak)
    sheet['F27'] = p_streak

    # SP_CV - Steam parameters after control valves
    t_conval, h_conval, s_conval, v_conval = SP_CV(p_streak, h0)

    print('3. Параметры пара после регулирующих клапанов:', 'p0`=', p_streak, 't0`=', t_conval-273, 'h0`=', h_conval, 's0`=', s_conval, 'v0`=', v_conval)
    sheet['C32'] = t_conval-273
    sheet['D32'] = h_conval
    sheet['E32'] = s_conval
    sheet['F32'] = v_conval

    book.save("C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\DB\\raschet_turboagregata_predvaritelny.xlsx")
    book.close()


if __name__ == "__main__":
    print("This is the module - Сalculation of process parameters in a turbine")