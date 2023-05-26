import tehkas_connect
from datetime import datetime, timedelta, time
import pandas as pd
from pandas.io.excel import ExcelWriter
from stats_to_excel import remove_excel_sheet, create_chart_all_tickets
import stats_to_excel

def tickets_in_work(reference):
    '''
    Считаем количество обращений в работе
    :param reference:справочник в котром смотрим обращения
    :return:словарь в формате "Месяц ГОД": количесвтво обращений
    '''
    in_work = ['Р']
    in_work = tehkas_connect.return_query(in_work, "СостОбращения", reference)
    in_work = reference.AddWhere(in_work)
    reference.Open()
    reference.First()
    created_mounth = dict()
    months = {"01": "Январь", "02": "Февраль", "03": "Март", "04": "Апрель", "05": "Май", "06": "Июнь", "07": "Июль",
              "08": "Август", "09": "Сентябрь", "10": "Октябрь", "11": "Ноябрь", "12": "Декабрь"}

    while not reference.EOF:
        created_date = reference.Requisites("ДатОткр").AsString
        mounth = f"{months[reference.Requisites('ДатОткр').AsString[3:5]]} {reference.Requisites('ДатОткр').AsString[6:]}"
        if mounth not in created_mounth:
            created_mounth[mounth] = 1
        else:
            created_mounth[mounth] += 1
        tehkas_connect.next_ticket(reference)
    reference.DelWhere(in_work)
    return created_mounth

def old_tickets(reference, config, days, active=True):
    '''
    Считаем старые обращения в команде
    :param reference: справочник для перебора
    :param config: список команды из файла config.json
    :param days: количество дней с даты регистрации старого обращения
    :param active: указатель находится ли в работе обращение
    :return: словарь в формате "Фамилия Имя": количесвтво обращений
    '''

    if active:
        in_work = ['Р']
    else:
        in_work = ['Р', 'К', 'И', 'П']
    in_work = tehkas_connect.return_query(in_work, "СостОбращения", reference)
    in_work = reference.AddWhere(in_work)
    reference.Open()
    reference.First()

    employee_list = dict()

    while not reference.EOF:
        created_date = datetime.strptime(reference.Requisites("ДатОткр").AsString, '%d.%m.%Y')
        difference_in_days = datetime.today() - created_date
        if difference_in_days > timedelta(days=days):
            if config["command_tab_num"][reference.Requisites("Работник").AsString] not in employee_list:
                employee_list[config["command_tab_num"][reference.Requisites("Работник").AsString]] = 1
            else:
                employee_list[config["command_tab_num"][reference.Requisites("Работник").AsString]] += 1
        tehkas_connect.next_ticket(reference)
    reference.DelWhere(in_work)
    return employee_list

def getTicketsCountByType(ticket_type, reference):
    '''
    Считает количество обращений в списке по типу
    :param ticket_type: тип обращений, ожидается список
    :param reference: справочник, где ищем
    :return:
    '''
    tickets_type = ticket_type
    tickets_type = command = tehkas_connect.return_query(tickets_type, "ТипОбращения", reference)
    tickets_type = reference.AddWhere(tickets_type)
    reference.Open()
    tickets_count = reference.RecordCount
    reference.Close()
    reference.DelWhere(tickets_type)
    return tickets_count

def registred_yesterday(reference):
    '''
    Обращения, зарегистрированные за предыдущий день

    !!!
    Потроганные обращения:
    1. Смотрим таблицу состояний. Если в ней есть переходы в состояние "На контроле" или "Переадресовано" то
    с обращением работали
    2. Если это запрос, то в выполненных работах смотрим работы с наименованием "ТМ" от имени сотрудника из команды
    !!!

    :param reference: справочник с обращениями
    :return: словарь в формате "зарегистрировано": количесвтво обращений, "в работе": количесвтво обращений
    '''

    yesterday = tehkas_connect.return_query(get_yesterday(), "ДатОткр", reference)
    yesterday = reference.AddWhere(yesterday)

    incoming = dict()
    incoming["incidents"] = getTicketsCountByType(["И"], reference)
    incoming["requests"] = getTicketsCountByType(["З"], reference)
    incoming["consultation"] = getTicketsCountByType(["К"], reference)
    incoming["problems"] = getTicketsCountByType(["П"], reference)

    reference.DelWhere(yesterday)
    return incoming

def not_closed(reference):
    ticket_status = ['Р', 'К', 'И', 'П']
    ticket_status = tehkas_connect.return_query(ticket_status, "СостОбращения", reference)
    ticket_status = reference.AddWhere(ticket_status)

    tickets_type = ['И', 'К', 'З']
    tickets_type = command = tehkas_connect.return_query(tickets_type, "ТипОбращения", reference)
    tickets_type = reference.AddWhere(tickets_type)

    reference.Open()
    reference.First()

    tickets_list = {"not_closed": 0, "older_four_weeks": 0}
    list_of_old_tickets = list()

    while not reference.EOF:
        tickets_list["not_closed"] += 1
        created_date = datetime.strptime(reference.Requisites("ДатОткр").AsString, '%d.%m.%Y')
        difference_in_days = datetime.today() - created_date

        if difference_in_days >= timedelta(days=28):
            tickets_list["older_four_weeks"] += 1
            list_of_old_tickets.append([reference.Requisites("ДатОткр").AsString,
                                        reference.Requisites("Код").AsString,
                                        reference.Requisites("Содержание").AsString,
                                        reference.Requisites("Работник").DisplayText,
                                        reference.Requisites("СостОбращения").DisplayText])
        tehkas_connect.next_ticket(reference)
    reference.DelWhere(ticket_status)
    reference.DelWhere(tickets_type)

    return tickets_list, list_of_old_tickets

def time_zones(reference, tickets_type):
    #tickets_type = ['И']
    tickets_type = tehkas_connect.return_query(tickets_type, "ТипОбращения", reference)
    tickets_type = reference.AddWhere(tickets_type)
    ticket_status = ['Р', 'К', 'И', 'П']
    ticket_status = tehkas_connect.return_query(ticket_status, "СостОбращения", reference)
    ticket_status = reference.AddWhere(ticket_status)
    ticket_time = pd.DataFrame(columns=["Номер обращения", "Описание", "Состояние обращения", "Ответсвенный",
                                        "Время в работе", "Приоритет", "Организация"])
    HOLYDAYS = ["23.2", "24.2", "8.3", "1.5", "8.5", "9.5", "12.6", "6.11"]  #TODO: вынести в конфиг

    reference.Open()
    reference.First()

    time_zone = {"red": 0, "yellow": 0, "sandy": 0, "green": 0}

    while not reference.EOF:
        reference.OpenRecord()
        detail = reference.DetailDataSet(4)
        detail.First()
        time_in_work_duration = 0
        start = datetime.now()
        end = datetime.now()
        has_start = False
        has_end = False

        while not detail.EOF:
            res = 0
            if detail.Requisites("СостОбращенияТ4").AsString == "В работе":
                start = datetime.strptime(str(detail.Requisites('ДатаВремяT4').AsString), '%d.%m.%Y %H:%M:%S')
                has_start = True

            elif (detail.Requisites("СостОбращенияТ4").AsString in ["На контроле", "Переадресовано"]) and has_start:
                end = datetime.strptime(str(detail.Requisites('ДатаВремяT4').AsString), '%d.%m.%Y %H:%M:%S')
                has_end = True


            detail.Next()
            if detail.EOF:
                if has_end == True:
                    pass
                else:
                    end = datetime.now()
                    has_end = True

            if has_start and has_end:
                time_start = time.fromisoformat('09:00')
                time_end = time.fromisoformat('17:00')
                t = start
                while (t := t + timedelta(minutes=1)) <= end:
                    if t.weekday() < 5 and time_start <= t.time() <= time_end and f"{t.day}.{t.month}" not in HOLYDAYS:
                        res += 1
                has_start = False
                has_end = False

            time_in_work_duration += res
        ticket_time.loc[len(ticket_time.index)] = [reference.Requisites("Код").AsString,
                                                   reference.Requisites("Содержание").AsString,
                                                   reference.Requisites("СостОбращения").AsString,
                                                   reference.Requisites("Работник").DisplayText,
                                                   round(time_in_work_duration/60, 2),
                                                   reference.Requisites("Строка3").AsString,
                                                   reference.Requisites("Организация").DisplayText,
                                                   ]
        ticket_time = ticket_time.sort_values(by=["Время в работе"], ascending=False)

        reference.Cancel()
        reference.CloseRecord()

        if time_in_work_duration < 480:
            time_zone['green'] += 1
        elif time_in_work_duration >= 480 and time_in_work_duration < 960:
            time_zone['sandy'] += 1
        elif time_in_work_duration >= 960 and time_in_work_duration < 1440:
            time_zone['yellow'] += 1
        elif time_in_work_duration >= 1440:
            time_zone['red'] += 1

        tehkas_connect.next_ticket(reference)
    reference.DelWhere(tickets_type)
    reference.DelWhere(tickets_type)
    return time_zone, ticket_time

def get_negative_grades(reference, COMMAND):
    '''
    Считает плохие оценки, полученные вчера
    :param reference: справочник обращения
    :param COMMAND: список команды
    :return:количество плохих оценок
    '''
    grades = tehkas_connect.Reference("REQUEST_SOLUTION_MARKS")  # Справочник оценки

    yesterday = list()  #Если сегодня понедельник посчитаем и субботу с воскресеньем

    #TODO добваил плохие оценки за сегодня. Если будет не надо - убрать
    yesterday.append(f"{datetime.today().day}.{datetime.today().month}.{datetime.today().year}")
    if datetime.today().isoweekday() == 1:
        for day in range(1,4):
            day_ago = datetime.today() - timedelta(days=day)
            day_ago = f"{day_ago.day}.{day_ago.month}.{day_ago.year}"
            yesterday.append(day_ago)
    else:
        day_ago = datetime.today() - timedelta(days=1)
        yesterday.append(f"{day_ago.day}.{day_ago.month}.{day_ago.year}")


    grade_get_yesterday = tehkas_connect.return_query(yesterday, "ДатЗакр", grades)
    grade_get_yesterday = grades.AddWhere(grade_get_yesterday)

    # Получим список обращений по которым надо смотреть оценки
    tickets_list = list()
    get_grade_time = datetime.today() - timedelta(days=10)
    get_grade_time = f"{get_grade_time.day}.{get_grade_time.month}.{get_grade_time.year}"
    closed_after = "({}.{} >= '{}')".format(reference.TableName, reference.Requisites("ДатЗакр").FieldName,
                                            get_grade_time)

    ticket_close_yesterday = reference.AddWhere(closed_after)

    command_only = tehkas_connect.return_query(COMMAND, "Работник", reference)
    command_only = reference.AddWhere(command_only)
    reference.Open()
    reference.First()
    while not reference.EOF:
        tickets_list.append(reference.Requisites("Код").AsInteger)
        tehkas_connect.next_ticket(reference)
    reference.Close()
    reference.DelWhere(ticket_close_yesterday)
    reference.DelWhere(command_only)

    good_marks = 0
    bad_marks = 0
    grades.Open()
    grades.First()
    while not grades.EOF:
        grades.OpenRecord()
        if grades.Requisites("Обращение").AsInteger in tickets_list:
            if grades.Requisites("ISBIntNumber").AsInteger == 4:
                good_marks += 1
            else:
                bad_marks += 1
        tehkas_connect.next_ticket(grades)
    grades.Close()
    return bad_marks

def get_time_for_request(reference):
    '''
    Считаем время, потраченное на запросы
    :param reference:
    :return:
    '''

    yesterday = tehkas_connect.return_query(get_yesterday(), "ДатОткр", reference)
    yesterday = reference.AddWhere(yesterday)

    request = tehkas_connect.return_query(['З'], "ТипОбращения", reference)
    request = reference.AddWhere(request)

    time = 0

    reference.Open()
    reference.First()

    while not reference.EOF:
        time += float(reference.Requisites("Сумма").AsString)
        tehkas_connect.next_ticket(reference)

    reference.DelWhere(yesterday)
    reference.DelWhere(request)
    return round(time, 2)

def get_yesterday():
    '''
    Считаем вчерашний день.
    Если сегодня понедельник посчитаем и субботу с воскресеньем
    :return: Список с датами
    '''
    yesterday = list()
    #TODO когда докажем правильность вернуть на вчера
    yesterday.append(f"{datetime.today().day}.{datetime.today().month}.{datetime.today().year}")

    # if datetime.today().isoweekday() == 1:
    #     for day in range(1, 4):
    #         day_ago = datetime.today() - timedelta(days=day)
    #         day_ago = f"{day_ago.day}.{day_ago.month}.{day_ago.year}"
    #         yesterday.append(day_ago)
    # else:
    #     day_ago = datetime.today() - timedelta(days=1)
    #     yesterday.append(f"{day_ago.day}.{day_ago.month}.{day_ago.year}")
    return yesterday