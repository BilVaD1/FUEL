from openpyxl import load_workbook
import datetime

class Summary:

    def __init__(self):
        # Load in the workbook
        self.wb = load_workbook('./data/Summary_report.xlsx')

    def writeValuesToExcel(self, tank_name, total_capacity, current_capacity):
        sheet = self.wb['Summary']  # Get a sheet by name

        current_datetime = datetime.datetime.now()

        remaining_capacity = float(total_capacity) - current_capacity
        percentage = round(remaining_capacity/total_capacity * 100, 2)

        for i in range(2, 19):
            name = sheet.cell(row=2, column=i).value
            if str(name) == str(tank_name):  # Если значение ячейки равно названию танка то записываем в него все данные
                print(tank_name, name)
                sheet.cell(row=1, column=i, value=current_datetime)
                sheet.cell(row=3, column=i, value=total_capacity)
                sheet.cell(row=4, column=i, value=current_capacity)
                sheet.cell(row=5, column=i, value=remaining_capacity)
                sheet.cell(row=6, column=i, value=percentage)
                break
            else:
                print(tank_name, name)
                continue

        self.wb.save(filename='./data/Summary_report.xlsx')