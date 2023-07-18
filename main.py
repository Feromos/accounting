import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import xlrd
import openpyxl

window = Tk()
window.state('zoomed')

c = {'Дата': [], 'Документ': [], 'Дебет': [], 'Кредит': [], 'Дебет_УПП': [], 'Кредит_УПП': [], 'Дебет_Отклонение': [],
     'Кредит_Отклонение': []}
a1 = {}
b1 = {}
r = {'Продажа': 'Приход', 'Возврат': 'Корректировка', 'Платежное': 'Оплата', 'Закупка': 'Продажа'}


def open1():
    global a
    file = tk.filedialog.askopenfilename()
    a = pd.read_excel(file)
    inf1.config(text=file)


def open2():
    global b
    file = tk.filedialog.askopenfilename()
    b = pd.read_excel(file)
    inf2.config(text=file)


def main_():
    global a
    global b
    while True:
        for i in a:
            for j in range(len(a[i])):
                if ('Дата' in str(a[i][j]) or 'дата' in str(a[i][j])) and 'Дата' not in a1:
                    r1 = 0
                    r2 = 0
                    for k in range(j + 1, len(a[i])):
                        if str(a[i][k]).count('.') == 2 and not r1:
                            r1 = k
                        if str(a[i][k]).count('.') != 2 and r1:
                            r2 = k
                            break
                    a1['Дата'] = list(a[i][r1 + 1:r2])
                else:
                    if 'Дата' in a1:
                        if ('Документ' in str(a[i][j]) or 'документ' in str(a[i][j])) and 'Документ' not in a1:
                            a1['Документ'] = list(a[i][r1:r2])
                            if str(a1['Документ'][0]) == 'nan':
                                del a1['Документ']
                        if ('Дебет' in str(a[i][j]) or 'дебет' in str(a[i][j])) and 'Дебет' not in a1:
                            a1['Дебет'] = list(a[i][r1:r2])
                            for y in range(len(a1['Дебет'])):
                                if type(a1['Дебет'][y]) != float and type(a1['Дебет'][y]) != int:
                                    a1['Дебет'][y] = 0
                        if ('Кредит' in str(a[i][j]) or 'кредит' in str(a[i][j])) and 'Кредит' not in a1:
                            a1['Кредит'] = list(a[i][r1:r2])
                            for y in range(len(a1['Кредит'])):
                                if type(a1['Кредит'][y]) != float and type(a1['Кредит'][y]) != int:
                                    a1['Кредит'][y] = 0
                                if a1['Кредит'][y] < 0:
                                    a1['Дебет'][y] = -a1['Кредит'][y]
                                    a1['Кредит'][y] = 0
                                if a1['Дебет'][y] < 0:
                                    a1['Кредит'][y] = -a1['Дебет'][y]
                                    a1['Дебет'][y] = 0
        if 'Дата' in a1 and 'Документ' in a1 and 'Дебет' in a1 and 'Кредит' in a1:
            break
    for i in b:
        for j in range(len(b[i])):
            if ('Дата' in str(b[i][j]) or 'дата' in str(b[i][j])) and 'Дата' not in b1:
                r1 = 0
                r2 = 0
                for k in range(j + 1, len(b[i])):
                    if str(b[i][k]).count('.') == 2 and not r1:
                        r1 = k
                    if str(b[i][k]).count('.') != 2 and r1:
                        r2 = k
                        break
                b1['Дата'] = list(b[i][r1:r2])
            if 'Дата' in b1:
                if ('Документ' in str(b[i][j]) or 'документ' in str(b[i][j])) and 'Документ' not in b1:
                    b1['Документ'] = list(b[i][r1:r2])
                if ('Дебет' in str(b[i][j]) or 'дебет' in str(b[i][j])) and 'Дебет' not in b1:
                    b1['Дебет'] = list(b[i][r1:r2])
                    for y in range(len(b1['Дебет'])):
                        if type(b1['Дебет'][y]) != float and type(b1['Дебет'][y]) != int:
                            b1['Дебет'][y] = 0
                if ('Кредит' in str(b[i][j]) or 'кредит' in str(b[i][j])) and 'Кредит' not in b1:
                    b1['Кредит'] = list(b[i][r1:r2])
                    for y in range(len(b1['Кредит'])):
                        if type(b1['Кредит'][y]) != float and type(b1['Кредит'][y]) != int:
                            b1['Кредит'][y] = 0
                        if b1['Кредит'][y] < 0:
                            b1['Дебет'][y] = -b1['Кредит'][y]
                            b1['Кредит'][y] = 0
                        if b1['Дебет'][y] < 0:
                            b1['Кредит'][y] = -b1['Дебет'][y]
                            b1['Дебет'][y] = 0
    for i in range(len(a1['Дата'])):
        if a1['Документ'][i] not in c['Документ']:
            for y in a1:
                c[y].append(a1[y][i])
            k1 = str(a1['Документ'][i]).split()
            k2 = []
            for y in k1:
                s1 = ''
                for o in y:
                    if o.isdigit():
                        s1 += o
                    else:
                        if o == '.':
                            s1 = ''
                            break
                        if s1:
                            break
                if s1:
                    k2.append(s1)
            t = True
            f = True
            for k in k2:
                for j in range(len(b1['Дата'])):
                    if b1['Документ'][j] and k in b1['Документ'][j]:
                        t1 = False
                        for y in r:
                            if k1[0] == y and r[y] in b1['Документ'][j]:
                                t1 = True
                                break
                        if t1:
                            b1['Документ'][j] = None
                            if f:
                                c['Дебет_УПП'].append(b1['Дебет'][j])
                                c['Кредит_УПП'].append(b1['Кредит'][j])
                            else:
                                c['Дебет_УПП'][-1] += b1['Дебет'][j]
                                c['Кредит_УПП'][-1] += b1['Кредит'][j]
                            t = False
                            f = False
            if not t:
                c['Дебет_Отклонение'].append(round(c['Дебет'][-1], 2) - round(c['Кредит_УПП'][-1], 2))
                c['Кредит_Отклонение'].append(round(c['Кредит'][-1], 2) - round(c['Дебет_УПП'][-1], 2))
                if c['Дебет_Отклонение'][-1] == 0:
                    c['Дебет_Отклонение'][-1] = None
                if c['Кредит_Отклонение'][-1] == 0:
                    c['Кредит_Отклонение'][-1] = None
            if t:
                c['Дебет_УПП'].append(-1)
                c['Кредит_УПП'].append(-1)
                c['Дебет_Отклонение'].append(-1)
                c['Кредит_Отклонение'].append(-1)
        else:
            for j in range(len(c['Документ']) - 1, -1, -1):
                if a1['Документ'][i] == c['Документ'][j]:
                    c['Дебет'][j] += a1['Дебет'][i]
                    c['Кредит'][j] += a1['Кредит'][i]
    c['Дата'].append('Ошибки:')
    c['Документ'].append(None)
    c['Дебет'].append(None)
    c['Кредит'].append(None)
    c['Дебет_УПП'].append(None)
    c['Кредит_УПП'].append(None)
    c['Дебет_Отклонение'].append(None)
    c['Кредит_Отклонение'].append(None)
    for i in range(len(b1['Документ'])):
        if b1['Документ'][i]:
            c['Дата'].append(b1['Дата'][i])
            c['Документ'].append(b1['Документ'][i])
            c['Дебет'].append(-1)
            c['Кредит'].append(-1)
            c['Дебет_УПП'].append(b1['Дебет'][i])
            c['Кредит_УПП'].append(b1['Кредит'][i])
            c['Дебет_Отклонение'].append(-1)
            c['Кредит_Отклонение'].append(-1)
    c_1 = pd.DataFrame(c)
    c_1.to_excel('Акт сверки.xlsx', index=False)
    messagebox.showinfo('bmi-pythonguides', 'Файл сохранен')


inf1 = Label(text='Пусто', font='Times 18')
inf1.grid(row=0, column=1)
button1 = Button(text="Акт контрагента", font='Times 18', command=open1)
button1.grid(column=0, row=0)

inf2 = Label(text='Пусто', font='Times 18')
inf2.grid(row=1, column=1)
button2 = Button(text="Акт из УПП", font='Times 18', command=open2)
button2.grid(column=0, row=1)

button1 = Button(text="Сформировать акт", font='Times 18', command=main_)
button1.grid(column=0, row=2)

window.mainloop()
