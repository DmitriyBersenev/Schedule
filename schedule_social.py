from xlutils.copy import copy
from xlrd import open_workbook
import calendar
import random
import math
from datetime import datetime

workbook = open_workbook('Соц_сети.xlsx')
Sheet = workbook.sheet_by_name('Исходные данные')
Settings = workbook.sheet_by_name('Распределение смен')

read_book = open_workbook('Соц_сети.xlsx')
write_b_copy = copy(read_book)
save = write_b_copy.get_sheet(0)

month = calendar.Calendar().monthdayscalendar(int(Settings.cell_value(rowx=1, colx=34)),
                                              int(Settings.cell_value(rowx=0, colx=34)))
days_in_month = max(month[len(month) - 1])

celebrations = [int(day) - 1 for day in Settings.cell_value(rowx=2, colx=34).split(',') if day != '']
holidays = [day - 1 for week in month for day in week[5:] if day != 0] + celebrations


class Shift:
    """ Смена. time - время смены, count_shifts - список c количеством определенных смен в день,
         """

    def __init__(self, time):
        self.time = time
        self.count_shifts = [int(count) - sum([1 for shift in Sheet.col_values(index + 1) if shift == time])
                             for index, count in
                             enumerate(Settings.row_values(Settings.col_values(0).index(time), start_colx=1,
                                                           end_colx=days_in_month + 1))]
        self.sum_shifts = sum(self.count_shifts)


class WeekdayShift(Shift):

    def arrange_shift(self, workers):
        for day in range(days_in_month):
            while self.count_shifts[day] > 0:
                random.shuffle(workers)
                for worker in workers:
                    if worker.worker_days[day] == 1:
                        if self.time not in workers[0].worker_days[day-(len(workers)-1) if day-(len(workers)-1) > 0 else 0:day]:
                            worker.worker_days[day] = self.time
                            self.count_shifts[day] -= 1
                            break
                continue
                for worker in workers:
                    if worker.worker_days[day] == 1:
                        if self.time not in workers[0].worker_days[day-(len(workers)-2) if day-(len(workers)-1) > 0 else 0:day]:
                            worker.worker_days[day] = self.time
                            self.count_shifts[day] -= 1
                            break
                continue
                for worker in workers:
                    if worker.worker_days[day] == 1:
                        worker.worker_days[day] = self.time
                        self.count_shifts[day] -= 1
                        break

class Worker:
    def __init__(self, name):
        self.name = name
        self.row = Sheet.col_values(0).index(name)
        self.worker_days = [1 if day == '' else day
                            for day in Sheet.row_values(self.row, start_colx=1, end_colx=days_in_month + 1)]
        self.count_shifts = {time: len(
            [shift for shift in Sheet.row_values(self.row, start_colx=1, end_colx=days_in_month + 1) if shift == time])
            for time in Settings.col_values(0) if time != ''}


workers = [Worker('Ивашко Юлия Юрьевна'), Worker('Бежан Диана Васильевна'), Worker('Поскребышева Мария Сергеевна')]
a = WeekdayShift('15:00 00:00')
a.arrange_shift(workers)



for worker in workers:
    for day in range(days_in_month):
        if worker.worker_days[day] == '21:00 08:00':
            save.write(worker.row + 1, day + 1, 11)
        elif len(str(worker.worker_days[day])) > 8 or (worker.worker_days[day] == 1 and day not in holidays):
            save.write(worker.row + 1, day + 1, 8)

        if worker.worker_days[day] in [0, 1, 'Вых']:
            save.write(worker.row, day + 1, '')
        else:
            save.write(worker.row, day + 1, worker.worker_days[day])

write_b_copy.save('Готовый шаблон соцсетей.xls')
print('Наслаждайтесь! График составлен ;)')