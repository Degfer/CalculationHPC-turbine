import sys

from tkinter import *
from tkinter import ttk
#import tkinter as tk
# from tkinter.scrolledtext import ScrolledText

sys.path.insert(0, "C:\\Users\\Дэн\\Desktop\\Дипломная работа\\CalculationHPC-turbine\\modules")

import CalPPT
import SaveExel

# Сохранение данных в БД - exel
def Save_tab1_DB(content, root):
    SaveExel.Save_tab1('B3', str(Marker1_count.get()))
    SaveExel.Save_tab1('C3', str(Marker2_count.get()))
    SaveExel.Save_tab1('D3', float(N_count.get()))
    SaveExel.Save_tab1('E3', float(p0_count.get()))
    SaveExel.Save_tab1('F3', float(t0_count.get()))
    SaveExel.Save_tab1('G3', float(pk_count.get()))
    SaveExel.Save_tab1('I3', float(n_count.get()))
    SaveExel.Save_tab1('J3', float(H0rs_count.get()))
    SaveExel.Save_tab1('K3', float(G0_count.get()))
    CalPPT.start(content, root)
    print("End")

# Иницилизация приложения
root = Tk()
root.title("Расчет ЦВД паравой турбины")
root.geometry("820x480")

# Значение парметров таблицы
content = ttk.Frame(root)
frame = ttk.Frame(content, width=100, height=30)

# Маркировка турбины
Marker = ttk.Label(content, text="Маркировка турбины")
Marker1_count = ttk.Combobox(content,
    state="readonly",
    values=["К-300-23,5", "Т-100-130"], width=8
)
Marker2_count = ttk.Combobox(content,
    state="readonly",
    values=["ЛМЗ, ЦВД", "УТЗ, ЦВД"], width=8
)

# Мощность
N = ttk.Label(content, text="N")
N_count = ttk.Entry(content, width=10)

# Расход
G0 = ttk.Label(content, text="G0, кг/c")
G0_count = ttk.Entry(content, width=10)

# Изначальное давление
p0 = ttk.Label(content, text="p0, мПа")
p0_count = ttk.Entry(content, width=10)

# Изначальная температура
t0 = ttk.Label(content, text="t0, °C")
t0_count = ttk.Entry(content, width=10)

# Конечное давление
pk = ttk.Label(content, text="pk, мПа")
pk_count = ttk.Entry(content, width=10)

# Теплоперепад на РС
H0rs = ttk.Label(content, text="H0рс, кДж/кг")
H0rs_count = ttk.Entry(content, width=10)

# Частота
n = ttk.Label(content, text="n, с^-1")
n_count = ttk.Entry(content, width=10)

# Кнопка
btn = ttk.Button(content, text="Расчет", command= lambda: Save_tab1_DB(content, root))

#Параметры окна
content.grid(column=0, row=0)
frame.grid(column=0, row=0)

# Парметры столбцов
Marker.grid(column=0, row=0, columnspan=2, padx=10)
Marker1_count.grid(column=0, row=1, columnspan=1, padx=10)
Marker2_count.grid(column=1, row=1, columnspan=1,  padx=10)

N.grid(column=2, row=0, padx=10)
N_count.grid(column=2, row=1, padx=10)

G0.grid(column=3, row=0, padx=10)
G0_count.grid(column=3, row=1, padx=10)

p0.grid(column=4, row=0, padx=10)
p0_count.grid(column=4, row=1, padx=10)

t0.grid(column=5, row=0, padx=10)
t0_count.grid(column=5, row=1, padx=10)

pk.grid(column=6, row=0, padx=10)
pk_count.grid(column=6, row=1, padx=10)

H0rs.grid(column=7, row=0, padx=10)
H0rs_count.grid(column=7, row=1, padx=10)

n.grid(column=8, row=0, padx=10)
n_count.grid(column=8, row=1, padx=10)

btn.grid(column=8, row=2, pady=20)

root.mainloop()