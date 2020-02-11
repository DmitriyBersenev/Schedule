from xlutils.copy import copy
from xlrd import open_workbook
from time import time
import calendar
import random
import math

WORKBOOK = open_workbook('Шаблон.xlsx')
WRITE_BY_COPY = copy(WORKBOOK)
SAVE = WRITE_BY_COPY.get_sheet(0)


class Config:
    def __init__(self):
        self.sheet = WORKBOOK.sheet_by_name('Исходные данные')
        self.count_shifts = WORKBOOK.sheet_by_name('Распределение смен')
        self.wishes = WORKBOOK.sheet_by_name('Пожелания по сменам')
        self.exceptions = WORKBOOK.sheet_by_name('Исключения')
        self.experience = WORKBOOK.sheet_by_name('Опытность')
        self.all_shifts = self.all_shifts()
        self.days_in_month = max(self.month()[-1])
        self.celebrations = [int(day) - 1 for day in self.count_shifts.cell_value(rowx=2, colx=34).split(',') if day]
        self.holidays = [day - 1 for week in self.month() for day in week[5:] if day] + self.celebrations
        self.saturdays = [week[5] for week in self.month()]
        self.sundays = [week[6] for week in self.month()]
        self.norma_hours = (self.days_in_month - len(self.holidays)) * 8
        self.teams = [team for team in self.count_shifts.col_values(34, start_rowx=4, end_rowx=11) if team]
        self.all_workers_name = [worker for worker in self.sheet.col_values(0) if len(str(worker)) > 15]
        self.all_workers_wishes = self.workers_wishes(self.all_workers_name)
        self.hours_between_shifts = self.hours_between_shifts()
        self.day_between_night = self.cell_founder('Дней между ночными сменами')
        self.max_weekend = self.cell_founder('Смен в выходные дни в месяц')
        self.day_between_weekend = self.cell_founder('Дней между сменами в выходные дни')
        self.excel_last_shifts = self.sheet.col_values(colx=35)
        self.sum_shifts_in_excel = self.sum_shifts_in_excel()
        self.excel_shift_row = self.excel_shifts_row()
        self.evening_in_day = self.evening_in_day()
        self.weekday_shifts_on_team = self.weekday_shifts_on_team()

    def all_shifts(self):
        return {shift: self.count_shifts.col_values(0).index(shift) for shift in self.count_shifts.col_values(0) if
                shift}

    def month(self):
        return calendar.Calendar().monthdayscalendar(int(self.count_shifts.cell_value(rowx=1, colx=34)),
                                                     int(self.count_shifts.cell_value(rowx=0, colx=34)))

    def workers_wishes(self, workers_name):
        """ Возвращает {Спец: лист с пожеланиями}"""

        def worker_wishes(name):
            """ Возвращает 1 или 0 - ходит/не ходит в определенную смену """
            name_search = 'По умолчанию' if name not in self.wishes.col_values(0) else name
            index_name = self.wishes.col_values(0).index(name_search)
            return [wish for wish in self.wishes.row_values(index_name) if type(wish) is float]

        return {worker: worker_wishes(worker) for worker in workers_name}

    def hours_between_shifts(self):
        cell_founder = self.count_shifts.col_values(33).index('Минимум часов между сменами')
        return int(self.count_shifts.cell_value(rowx=cell_founder, colx=34))

    def sum_shifts_in_excel(self):
        """ {День: {Смена: количество пожеланий по этой смене}} """

        def sum_shift(day, time):
            return sum([1 for shift in self.sheet.col_values(day + 1) if shift == time])

        def count_shift(day):
            return {time: sum_shift(day, time) for time in self.all_shifts}

        return {day: count_shift(day) for day in range(self.days_in_month)}

    def excel_shifts_row(self):
        def shift_row(time):
            return self.count_shifts.row_values(self.all_shifts[time], start_colx=1, end_colx=self.days_in_month + 1)

        return {time: shift_row(time) for time in self.all_shifts}

    def evening_in_day(self):
        """ {День: сумма вечерних смен} """

        def evening_shift(time, day):
            return self.count_shifts.cell_value(rowx=self.all_shifts[time], colx=day + 1)

        def sum_evening_shifts(day):
            return sum([evening_shift(time, day) for time in self.all_shifts if int(time[:2]) >= 14])

        return {day: sum_evening_shifts(day) for day in range(self.days_in_month)}

    def cell_founder(self, text):
        found_cell = self.count_shifts.col_values(33).index(text)
        return int(self.count_shifts.cell_value(rowx=found_cell, colx=34))

    def weekday_shifts_on_team(self):
        """ Среднее количество смен на команду в выходные """

        def average_shifts(day):
            count_shifts = sum(self.count_shifts.col_values(day + 1, start_rowx=2, end_rowx=len(self.all_shifts) + 2))
            return math.ceil(count_shifts / len(self.teams))

        return {day: average_shifts(day) for day in self.holidays}


class Team:
    """ name - имя команды """

    def __init__(self, name, config):
        self.name = name
        self.conf = config
        self.start_row = self.conf.sheet.col_values(0).index(name) + 2
        self.end_row = self.team_end_row()
        self.team_in_sheet = self.team_in_sheet()
        self.count_workers_in_day = self.count_workers_in_day()
        self.evening_shifts = self.count_evening_shifts()
        self.days_with_night = self.days_with_night()
        self.team_day_shifts = self.team_day_shifts()
        self.shifts_team = self.shifts_team()
        self.work_in_shift = self.work_in_shift()
        self.workers = self.init_workers()

    def team_in_sheet(self):
        """ {Колонка в эксель: список в диапазоне команды} """

        def col_in_sheet(col):
            return self.conf.sheet.col_values(col, start_rowx=self.start_row, end_rowx=self.end_row)

        return {col: col_in_sheet(col) for col in range(conf.days_in_month + 1)}

    def team_end_row(self):
        """ Возвращает последний ряд команды в таблице эксель """
        excel_list = self.conf.sheet.col_values(0, start_rowx=self.start_row)
        for row in range(len(excel_list))[::2]:
            if row + 2 > len(excel_list) or not excel_list[row]:
                return row + self.start_row + 1

    def count_workers_in_day(self):
        """ {День: количество работающих в день} """

        def workers_in_day(day):
            return len([cell for cell in self.team_in_sheet[day + 1] if cell in self.conf.all_shifts])

        return {day: workers_in_day(day) for day in range(self.conf.days_in_month)}

    def count_evening_shifts(self):
        """ {День: количество поздних смен у команды} """

        def evening_in_day(day):
            return len([cell for cell in self.team_in_sheet[day + 1] if len(str(cell)) > 10 and int(cell[:2]) > 13])

        return {day: evening_in_day(day) for day in range(self.conf.days_in_month)}

    def days_with_night(self):
        """ {День: True или False, если есть ночная в этот день} """

        def with_night(day):
            return '21:00 08:00' in [shift for shift in self.team_in_sheet[day + 1]]

        return {day: with_night(day) for day in range(self.conf.days_in_month)}

    def team_day_shifts(self):
        """ {День: {Смена: Среднее количество смен на команду}} """

        def count_shift_in_day(time, day):
            """ Количество определенной смены в день """
            return self.conf.count_shifts.cell_value(rowx=self.conf.all_shifts[time], colx=day + 1)

        def shifts_in_schedule(time, day):
            """ Смены команды, что уже стоят в графике """
            return len([shift for shift in self.team_in_sheet[day + 1] if shift == time])

        def average_shifts_in_day(day):
            """ Среднее количество всех смен в конкретный день """
            return {
                time: math.ceil(count_shift_in_day(time, day) / len(self.conf.teams)) - shifts_in_schedule(time,
                                                                                                           day) + 1
                for time in self.conf.all_shifts}  # Прибавил 1, чтобы увеличить возможность составления

        return {day: average_shifts_in_day(day) for day in range(self.conf.days_in_month)}

    def shifts_team(self):
        """ {Смена: Среднее количество смен на команду в месяц} """
        team_workers = len([worker for worker in self.team_in_sheet[0] if worker])

        def sum_shifts(time):
            """ Количество определенных смен в месяц """
            return sum(self.conf.count_shifts.row_values(self.conf.all_shifts[time], start_colx=1,
                                                         end_colx=self.conf.days_in_month + 1))

        return {shift: math.ceil(team_workers * sum_shifts(shift) / len(self.conf.all_workers_name)) for shift in
                self.conf.all_shifts}

    def work_in_shift(self):
        """ {Смена: Количество людей от команды в смену} """
        workers_name = [worker for worker in self.team_in_sheet[0] if worker]

        def workers_in_shift(time):
            """ Количество людей от команды в смену """
            return sum(
                [self.conf.all_workers_wishes[worker][self.conf.all_shifts[time] - 2] for worker in workers_name])

        return {time: workers_in_shift(time) for time in self.conf.all_shifts}

    def init_workers(self):
        """ Инициализирует класс специалистов """
        excel_workers = [worker for worker in self.team_in_sheet[0] if worker]
        return [Worker(worker, self.name, self.shifts_team, self.work_in_shift, self.conf) for worker in excel_workers]


class Worker:
    """ name - ФИО, team - команда. row - ряд в экселе, exp - опыт специалиста(False - меньше полугода),
        shifts_team - количество определенных смен за месяц у команды,
        work_in_shift - количество работающих людей в эту смену,
        all_except - исключения(False - без выходных и ночных смен),
        night_except - исключение(False - без ночных смен), worker_days - список с календарем специалиста,
        last_night - дней до предыдущей ночи, count_last_month_shifts - смен в конце прошлого месяца """

    def __init__(self, name, team, shifts_team, work_in_shift, config):
        self.name = name
        self.team = team
        self.shifts_team = shifts_team
        self.work_in_shift = work_in_shift
        self.conf = config
        self.row = self.conf.sheet.col_values(0).index(name)
        self.excel_row = self.conf.sheet.row_values(self.row)
        self.exp = self.conf.experience.cell_value(rowx=self.conf.experience.col_values(0).index(name), colx=2)  # -!
        self.all_except = name not in self.conf.exceptions.col_values(0)  # -!
        self.night_except = name not in self.conf.exceptions.col_values(2)  # -!
        self.start_hours = self.start_hours()
        self.worker_hours = self.worker_hours()
        self.worker_days = [1 if not day else day for day in self.worker_row()]
        self.last_night = int(self.excel_row[34]) if self.excel_row[34] else 0  # -!
        self.count_last_month_shifts = int(self.excel_row[33])  # -!
        self.count_shifts = self.count_worker_shifts()
        self.weekend_days = self.worker_weekend_days()
        self.shift_in_month = self.init_shift_in_month()

    def worker_row(self):
        """ Список с данными из экселья в месячном диапазоне """
        return self.excel_row[1:self.conf.days_in_month + 1]

    def start_hours(self):
        """ Начальная норма часов специалиста """
        days_off = [day for num, day in enumerate(self.worker_row()) if
                    (day == 'о' or day == 'б') and num not in self.conf.holidays]
        return self.conf.norma_hours - len(days_off)

    def worker_hours(self):
        """ Реальная норма часов специалиста """
        count_excel_shifts = len([day for day in self.worker_row() if day in self.conf.all_shifts])
        return self.start_hours - count_excel_shifts * 8

    def count_worker_shifts(self):
        """ {Смена: количество этих смен у специалиста} """
        shifts_counter = {}
        for time in self.conf.all_shifts:
            shifts_counter[time] = len([shift for shift in self.worker_row() if shift == time])
        return shifts_counter

    def worker_weekend_days(self):
        """ Количество смен у специалиста в выходные дни """
        days = [self.excel_row[day + 1] for day in self.conf.holidays if
                self.excel_row[day + 1] not in ['о', 'б', 'Вых', '']]
        return len(days)

    def init_shift_in_month(self):
        """ {Смена: количество смен пропорционально рабочим часам} """

        def average_shift(time):
            """ Среднее количество смены в месяц """
            if self.conf.all_workers_wishes[self.name][self.conf.all_shifts[time] - 2] and self.work_in_shift[time]:
                shifts = math.ceil((self.shifts_team[time] / self.work_in_shift[time]) * k_hours * k_shifts)
                return shifts - len([shift for shift in self.worker_days if time == shift])
            return 0

        k_hours = self.start_hours / self.conf.norma_hours
        k_shifts = (len(self.conf.all_shifts)) / sum(self.conf.all_workers_wishes[self.name])
        return {time: average_shift(time) for time in self.conf.all_shifts}


class Shift:
    """ time - время смены """

    def __init__(self, time, config):
        self.time = time
        self.conf = config
        self.count_shifts = self.count_shifts()
        self.count_evening_shifts = self.count_evening_shifts()

    def count_shifts(self):
        """ Список с количеством определенных смен в день """
        return [int(count) - conf.sum_shifts_in_excel[day][self.time] if count else 0 for day, count in
                enumerate(conf.excel_shift_row[self.time])]

    def count_evening_shifts(self):
        """ Считает максимум вечерних смен в день на команду """
        return {day: math.ceil(self.conf.evening_in_day[day] / len(self.conf.teams)) for day in
                range(self.conf.days_in_month)}

    def next_last_shift(self, day, worker):
        """ Проверяет часы между следующей/предыдущей сменой """

        def start_conv(time):
            """ Конвертирует текст смены в число = часу ее начала """
            return int(time[: time.find(':')])

        def end_conv(time):
            """ Конвертирует текст смены в число = часу ее конца """
            return int(time[5: time.find(':', 5)])

        if worker.worker_days[day] != 1:
            return False
        if start_conv(self.time) < 13:
            if not day:
                if len(str(conf.excel_last_shifts[worker.row])) > 9:
                    clock = end_conv(conf.excel_last_shifts[worker.row])
                    return (clock - 24 if clock > 6 else clock) + self.conf.hours_between_shifts < start_conv(self.time)
            else:
                if len(str(worker.worker_days[day - 1])) > 9:
                    clock = end_conv(worker.worker_days[day - 1])
                    return (clock - 24 if clock > 6 else clock) + self.conf.hours_between_shifts < start_conv(self.time)
        else:
            if day != self.conf.days_in_month - 1:
                if len(str(worker.worker_days[day + 1])) > 9:
                    clock = end_conv(self.time)
                    return (clock - 24 if clock > 6 else clock) + self.conf.hours_between_shifts < start_conv(
                        worker.worker_days[day + 1])
        return True

    def check_vacation(self, check_day, worker_days):
        """  Не ставит смены в выходые, если отпуск в начале недели """
        if check_day + 1 in self.conf.saturdays and check_day + 2 <= self.conf.days_in_month:
            return worker_days[check_day + 2] != 'о'
        if check_day + 1 in self.conf.sundays and check_day + 1 <= self.conf.days_in_month:
            return worker_days[check_day + 1] != 'о'
        return True

    def check_more_five(self, check_day, worker_days, last_month_days):
        """ Проверяет, больше 5ти смен подряд у спеца. Возвращает True / False """

        def look_forward():
            """ Считает количество смен подряд после конкретной даты """
            for i, shift in enumerate(worker_days[check_day + 1: check_day + 6]):
                if len(str(shift)) < 5:
                    return i
            return 5 if check_day + 1 != len(worker_days) else 0

        def look_back():
            """ Считает количество смен подряд перед конкретной датой """
            back_list = worker_days[0 if check_day - 5 < 0 else check_day - 5: check_day]
            back_list.reverse()
            for i, shift in enumerate(back_list):
                if len(str(shift)) < 5:
                    return i
            return len(back_list)

        if check_day < 7:
            def start_look_back():
                """ Считает количество смен подряд перед конкретной датой, либо из предыдущего месяца """
                shifts_before_day = len([day for day in worker_days[: check_day] if len(str(day)) > 5])
                return look_back() if shifts_before_day < check_day else last_month_days + look_back()

            return start_look_back() + look_forward() < 5
        return look_back() + look_forward() < 5

    def install_shift(self, day, worker, excepting_teams=None):
        """ Инсталлирует показатели по всем сменам """
        worker.worker_days[day] = self.time
        worker.count_shifts[self.time] += 1
        worker.shift_in_month[self.time] -= 1
        self.count_shifts[day] -= 1
        teams[worker.team].count_workers_in_day[day] += 1
        teams[worker.team].team_day_shifts[day][self.time] -= 1
        if int(self.time[:2]) >= 14:
            teams[worker.team].evening_shifts[day] += 1
        if day in self.conf.holidays:
            worker.weekend_days += 1
        if self.time == '21:00 08:00':
            worker.worker_hours -= 11
            if day + 2 <= self.conf.days_in_month:
                worker.worker_days[day + 1] = 0
            if worker.team not in excepting_teams:
                excepting_teams.append(worker.team)
        else:
            worker.worker_hours -= 8


class NightShift(Shift):
    """ Ночные смены. except_to - количество дней между ночными сменами """

    def __init__(self, time, config):
        super().__init__(time, config)

    def check_except_to(self, day, worker):
        """ True, если ночная была не раньше установленного периода """
        except_to = worker.worker_days[day - self.conf.day_between_night: self.conf.day_between_night + day]
        return self.time not in except_to and worker.last_night + day > self.conf.day_between_night

    def arrange_shifts(self, workers, teams):
        """ Расставляет смены в ночь """
        for day in range(self.conf.days_in_month):
            excepting_teams = [team.name for team in teams.values() if team.days_with_night[day]]
            experience = True
            for count in range(self.count_shifts[day]):
                if self.time == '13:00 22:00':
                    print(self.time, self.count_shifts[day], day)
                random.shuffle(workers)
                for index, worker in enumerate(workers):
                    if all([worker.worker_days[day] == 1,
                            worker.worker_days[day + 1 if day + 2 < self.conf.days_in_month else 0] == 1,
                            worker.worker_days[0 if day - 1 < 0 else day - 1] != 'о',
                            worker.team not in excepting_teams,
                            experience or worker.exp,
                            worker.count_shifts[self.time] < round(sum(self.count_shifts) / len(workers)) + 1,
                            self.check_except_to(day, worker),
                            self.check_vacation(day, worker.worker_days)]):
                        self.install_shift(day, worker, excepting_teams)
                        experience = worker.exp
                        break
                    if index + 1 == len(workers):
                        return True
        return False


class WeekendShift(Shift):
    """ max_weekend - максимум смен в выходные за месяц """

    def __init__(self, time, config):
        super().__init__(time, config)

    def check_between(self, day, worker):
        """ Определяет период между сменами в выходные дни """
        holidays = self.conf.holidays
        before_day = holidays[holidays.index(day) - self.conf.day_between_weekend: holidays.index(day)]
        after_day = holidays[holidays.index(day) + 1: holidays.index(day) + self.conf.day_between_weekend + 1]
        return not sum([1 for day in before_day + after_day if worker.worker_days[day] in self.conf.all_shifts])

    def arrange_shifts(self, workers):
        """ Расставляет смены в выходные дни """

        for day in self.conf.holidays:
            for count in range(self.count_shifts[day]):
                random.shuffle(workers)

                for index, worker in enumerate(workers):
                    if all([self.check_between(day, worker),
                            worker.shift_in_month[self.time] > 0,
                            self.conf.weekday_shifts_on_team[day] > teams[worker.team].count_workers_in_day[day],
                            self.conf.max_weekend > worker.weekend_days,
                            self.next_last_shift(day, worker),
                            self.check_more_five(day, worker.worker_days,
                                                 worker.count_last_month_shifts) if day < 29 else True,
                            self.check_vacation(day, worker.worker_days)]):
                        self.install_shift(day, worker)

                        break
                    if index + 1 == len(work_at_weekend):
                        #print(day + 1, self.time)
                        return True

        return False


class WeekdayShift(Shift):
    """ Будние дни """

    def arrange_shifts(self, workers):
        """ Расставляет смены в будние дни """
        for day in range(self.conf.days_in_month):
            if day not in self.conf.holidays:
                random.shuffle(workers)
                for count in range(self.count_shifts[day]):
                    for index, worker in enumerate(workers):
                        if all([self.next_last_shift(day, worker),
                                worker.shift_in_month[self.time] > 0,
                                worker.worker_hours + 4 > 0,
                                teams[worker.team].team_day_shifts[day][self.time] > 0,
                                self.check_more_five(day, worker.worker_days,
                                                     worker.count_last_month_shifts) if day < 29 else True]):
                            if int(self.time[:2]) >= 14:
                                if self.count_evening_shifts[day] > teams[worker.team].evening_shifts[day]:
                                    self.install_shift(day, worker)
                                    break
                                continue
                            self.install_shift(day, worker)
                            break
                        if index + 1 == len(workers):
                            return True, day, self.time, self.count_shifts[day]
        return False


conf = Config()

tic = time()
k_day = 0
k_shift = '18:00 03:00'
k_count_shift = 20
back = True
while back:

    teams = {team: Team(team, conf) for team in conf.teams}

    all_workers = []
    for team in teams.values():
        for worker in team.workers:
            all_workers.append(worker)

    work_at_night = [worker for worker in all_workers if worker.all_except and worker.night_except]
    night_shift = NightShift('21:00 08:00', conf)

    w_s = ['18:00 03:00', '16:00 01:00', '14:00 23:00', '07:00 16:00', '08:00 17:00', '09:00 18:00', '10:00 19:00',
           '11:00 20:00', '12:00 21:00', '13:00 22:00']

    work_at_weekend = [worker for worker in all_workers if worker.all_except]
    weekend_shifts = [WeekendShift(time, conf) for time in w_s if
                      time not in ['', '21:00 08:00']]

    weekday_shifts = [WeekdayShift(time, conf) for time in w_s]

    if night_shift.arrange_shifts(work_at_night, teams):
        continue
    # print('Night good!')
    counter_shift = [0]
    last_shift_check = True
    for i, shift in enumerate(weekend_shifts):
        if not shift.arrange_shifts(work_at_weekend) and i in counter_shift:
            counter_shift.append(i + 1)
            continue
        elif i + 1 == len(w_s):
            last_shift_check = False
    if len(counter_shift) < len(weekend_shifts) or not last_shift_check:
        continue

    counter_shift = [0]
    for i, shift in enumerate(weekday_shifts):
        back = True
        func_result = shift.arrange_shifts(all_workers)
        if not func_result and i in counter_shift:
            back = False
            counter_shift.append(i + 1)
            continue

        if k_day <= func_result[1]:
            if k_day < func_result[1]:
                k_shift = '18:00 03:00'
            if w_s.index(k_shift) <= w_s.index(func_result[2]):
                k_shift = func_result[2]
                print(func_result[1] + 1, func_result[2], func_result[3])
                print('Продолжаю ставить смены...')
            k_day = func_result[1]

        sum_shifts = [shift.count_shifts for shift in weekday_shifts]
        break

for worker in all_workers:
    for day in range(conf.days_in_month):
        if worker.worker_days[day] == '21:00 08:00':
            SAVE.write(worker.row + 1, day + 1, 11)
        elif len(str(worker.worker_days[day])) > 8 or (worker.worker_days[day] == 1 and day not in conf.holidays):
            SAVE.write(worker.row + 1, day + 1, 8)
        if worker.worker_days[day] in [0, 1, 'Вых']:
            SAVE.write(worker.row, day + 1, '')
        else:
            SAVE.write(worker.row, day + 1, worker.worker_days[day])

WRITE_BY_COPY.save('Готовый шаблон.xls')
print('Наслаждайтесь! График составлен ;)')
toc = time()
print((toc - tic), 'sec')
