from tkinter import *
from tkinter import ttk

# Иницилизация приложения
root = Tk()
root.title("Расчет ЦВД паравой турбины")
root.geometry("1080x480")

# Значение парметров таблицы
content = ttk.Frame(root)
frame = ttk.Frame(content, width=100, height=30)

# Мощность
N = ttk.Label(content, text="N")
N_count = ttk.Entry(content)

# Расход
G0 = ttk.Label(content, text="G0")
G0_count = ttk.Entry(content)

# Изначальное давление
p0 = ttk.Label(content, text="p0")
p0_count = ttk.Entry(content)

# Изначальная температура
t0 = ttk.Label(content, text="t0")
t0_count = ttk.Entry(content)

# Конечное давление
pk = ttk.Label(content, text="pk")
pk_count = ttk.Entry(content)

# Теплоперепад на РС
H0rs = ttk.Label(content, text="H0рс")
H0rs_count = ttk.Entry(content)

# Частота
n = ttk.Label(content, text="n")
n_count = ttk.Entry(content)

onevar = BooleanVar(value=True)
twovar = BooleanVar(value=False)
threevar = BooleanVar(value=True)

one = ttk.Checkbutton(content, text="One", variable=onevar, onvalue=True)

content.grid(column=0, row=0)
frame.grid(column=0, row=0)

# Парметры столбцов
N.grid(column=1, row=0, padx=10)
N_count.grid(column=1, row=1, padx=10)

G0.grid(column=2, row=0, padx=10)
G0_count.grid(column=2, row=1, padx=10)

p0.grid(column=3, row=0, padx=10)
p0_count.grid(column=3, row=1, padx=10)

t0.grid(column=4, row=0, padx=10)
t0_count.grid(column=4, row=1, padx=10)

pk.grid(column=5, row=0, padx=10)
pk_count.grid(column=5, row=1, padx=10)

H0rs.grid(column=6, row=0, padx=10)
H0rs_count.grid(column=6, row=1, padx=10)

one.grid(column=0, row=3)
###

root.mainloop()