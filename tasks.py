import re
from openpyxl import load_workbook
import matplotlib.pyplot as plt  # Библиотека для работы с графиками
import numpy as np


class Tasks:

    def __init__(self):
        # Load in the workbook
        self.wb = load_workbook('./data/Book1.xlsx')

    def createGraph(self, x, y, z, z2):
        plt.plot(x, y, color="red", marker="o", ms=5, label="label")  # Создаем график с переменными в качестве x, y
        plt.plot(z, z2, color="blue", marker="o", ms=5, label="measurements")  # Создаем точку вычислений на графике
        plt.legend()
        plt.xlabel("$deep(m)$", fontsize=16)
        plt.ylabel("$Vm^3$", fontsize=16)
        plt.show()

    def writeValuesToExcel(self, sheet, z, z2, t):
        sheet.cell(row=3, column=7, value=z)
        sheet.cell(row=4, column=7, value=z2)
        sheet.cell(row=5, column=7, value=t)
        self.wb.save(filename='./data/Book1.xlsx')

    def task1(self, sounding, tank, answer_t, answer_m3):

        if sounding == '' and tank == '':
            answer_t['text'] = ''
            answer_m3['text'] = 'Please enter values into the fields'
            return answer_t, answer_m3
        elif tank == '':
            answer_m3['text'] = 'Please enter the valid value into the Tank name field'
            return answer_m3
        elif sounding == '':
            answer_m3['text'] = 'Please enter the valid value into the Sounding field'
            return answer_m3
        elif bool(re.search(r'[^\W\d]', sounding)) == True:   # Check the string on the letters
            answer_m3['text'] = 'Please enter the valid value into the Sounding field'
            return answer_m3

        sounding_number = float(sounding)
        tank_name = str(tank)

        l1 = str(('10', '9', '12', '13', '14', '21', '22', '23', 'Overflow', '7', '6', '5', 'Sludge', 'Bilge', '17',
                  'Waste', 'MEoil'))  # Преобразуем в строку чтобы подхватила библиотека re

        call1 = bool(re.search(tank_name, l1, re.IGNORECASE))  # Используем библиотеку re чтобы найти параметр кол в
        # строке л1, применяя re.IGNORECASE чтобы игнорировать капс. Используем это чтобы пользователь вводил
        # значение игнорируя капс

        if call1 is True:
            a = re.search(tank_name, l1, re.IGNORECASE).group()  # Передаем в переменную искаемое слово с помощью group()
            sheet = self.wb[a]  # Get a sheet by name
        else:
            answer_m3['text'] = 'Uncorrected name of tank!'
            return answer_m3

        # Создаем списки чтобы в них поместить координаты
        x = []
        y = []
        # Получаем координаты Х из ексель файла с измерения глубины
        for i in range(3, 50):
            x1 = sheet.cell(row=i, column=2).value
            if x1 is not None:  # Если значение ячейки не равно None то добавляем значение в список
                x.append(x1)
            else:
                continue
        # Получаем обьем Y топливного танка из Ексель файла
        for i in range(3, 50):
            y1 = sheet.cell(row=i, column=4).value
            if y1 is not None:
                y.append(y1)
            else:
                continue
        k = sheet.cell(row=2, column=7).value  # Получаем значение плотности из таблицы

        z_max = x[-1]  # Находим последнее значение списка глубины, чтобы оптимизировать работу на всех страницах Екселя

        z = 0  # Значение на графике

        if 0.00 < sounding_number < z_max:
            z = float(np.interp(sounding_number, x, y))  # Считаем с помощью интерпритаций значение точки на графике, в нашем случае z2
            answer_m3['text'] = f'You have {"{0:.3f}".format(z)} m3'  # {"{0:.3f}".format(z)} use only 2 numbers after the dot in the z2
            t = k * z  # Вычисленние тонн
            answer_t['text'] = f'You have {"{0:.3f}".format(t)} t' # {"{0:.3f}".format(z)} use only 2 numbers after the dot in the t
            self.writeValuesToExcel(sheet, sounding_number, z, t)
            self.createGraph(x, y, sounding_number, z)
            return answer_m3, answer_t
        else:
            answer_m3['text'] = 'Uncorrected sounding, use format 0.00'
            answer_t['text'] = ''
            return answer_m3, answer_t
