from openpyxl import load_workbook
import datetime

class Summary:

    def __init__(self, tanks_capacity=None, current_capacity=None, remaining_capacity=None, tank_names=None, Notification_Block=None):
        self.remaining_capacity = remaining_capacity
        self.tanks_capacity = tanks_capacity
        self.current_capacity = current_capacity
        self.tank_names = tank_names
        self.Notification_Block = Notification_Block
        # Load in the workbook
        self.wb = load_workbook('./data/Summary_report.xlsx')
        self.sheet = self.wb['Summary']

    def writeValuesToExcel(self, tank_name, total_capacity, current_capacity):

        current_datetime = datetime.datetime.now()
        #today = datetime.date.today()

        remaining_capacity = float(total_capacity) - current_capacity
        percentage = round(remaining_capacity/total_capacity * 100, 2)

        for i in range(2, 19):
            name = self.sheet.cell(row=2, column=i).value
            if str(name) == str(tank_name):  # Если значение ячейки равно названию танка то записываем в него все данные
                print(tank_name, name)
                self.sheet.cell(row=1, column=i, value=current_datetime)
                self.sheet.cell(row=3, column=i, value=total_capacity)
                self.sheet.cell(row=4, column=i, value=current_capacity)
                self.sheet.cell(row=5, column=i, value=remaining_capacity)
                self.sheet.cell(row=6, column=i, value=percentage)
                break
            else:
                print(tank_name, name)
                continue

        self.wb.save(filename=f'./data/Summary_report.xlsx')
        #self.wb.save(filename=f'./reports/Summary_report_{today}.xlsx')

    def generate_report(self):
        today = datetime.date.today()
        self.wb.save(filename=f'./reports/Summary_report_{today}.xlsx')


    def writeValuesToUI(self, total, current, remaining, tanks):
        self.tanks_capacity['text'] = f'Tank(s) capacity: {str(total)}'
        self.current_capacity['text'] = f'Current capacity: {str(current)}'
        self.remaining_capacity['text'] = f'Remaining capacity: {str(remaining)}'
        self.tank_names['text'] = f'Tank(s) name(s): {tanks}'


    def foundSelectedValues(self, tanks):
        total = []
        current = []
        remaining = []

        for tank in tanks:
            for i in range(2, 19):
                name = self.sheet.cell(row=2, column=i).value
                if str(name) == str(tank):  # Если значение ячейки равно названию танка то записываем в него все данные
                    total_value = self.sheet.cell(row=3, column=i).value
                    current_value = self.sheet.cell(row=4, column=i).value
                    remaining_value = self.sheet.cell(row=5, column=i).value
                    if total_value is not None and current_value is not None and remaining_value is not None:
                        total.append(self.sheet.cell(row=3, column=i).value)
                        current.append(self.sheet.cell(row=4, column=i).value)
                        remaining.append(self.sheet.cell(row=5, column=i).value)
                        break
                    elif total_value is None: 
                        self.Notification_Block['text'] = f'Error in the total volume of tank {name}'
                        self.Notification_Block['bg'] = f'red'
                        return self.tanks_capacity
                    elif current_value is not None:
                        self.Notification_Block['text'] = f'Error in the current volume of tank {name}'
                        self.Notification_Block['bg'] = f'red'
                        return self.tanks_capacity
                    elif remaining_value is  None:
                        self.Notification_Block['text'] = f'Error in the remaining volume of tank {name}'
                        self.Notification_Block['bg'] = f'red'
                        return self.tanks_capacity
        
        #print(sum(total), sum(current), sum(remaining))

        self.writeValuesToUI(sum(total), sum(current), sum(remaining), tanks)