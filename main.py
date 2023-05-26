import json
import os
import tehkas_connect
from datetime import datetime, timedelta, time
import pandas as pd
from pandas.io.excel import ExcelWriter
from stats_to_excel import remove_excel_sheet, create_chart_all_tickets
import stats_to_excel
from calculation_of_statistics import *

def get_conf_and_start():
    directory = '.'
    for config in os.listdir(directory):
        if config.endswith('.json'):
            config_name = config.split(sep='.')[0]
            for data_file in os.listdir(directory):
                if data_file.endswith('.xlsx'):
                    data_file_name = data_file.split(sep='.')[0]
                    if config_name == data_file_name:
                        calculate_stats(config, data_file)
                        #calculate_stats(os.path.join(directory, config), os.path.join(directory, data_file))
                        #print(os.path.join(directory, config), os.path.join(directory, data_file))


def calculate_stats(config_path, data_file_path):
    with open(config_path, "r", encoding="utf-8") as config_file:
        config = json.loads(config_file.read())

    tickets = tehkas_connect.Reference("ПДД")
    COMMAND = [employee for employee in config["command_for_tehkas"].values()]
    command = tehkas_connect.return_query(COMMAND, "Работник", tickets)
    command = tickets.AddWhere(command)

    # все коды поступления обращений кроме чат-бота
    not_chat_bot = [1761155, 1761153, 1761152, 6753420, 1761156, 3010066, 4521240, 1768268, 4081742, 1761154]
    not_chat_bot = tehkas_connect.return_query(not_chat_bot, "ИсточникОбращения", tickets)
    not_chat_bot = tickets.AddWhere(not_chat_bot)

    # Убираем анонимки
    not_anonymous = '30262732'   # анонимки (Код-338)
    not_anonymous = f"{tickets.TableName}.{tickets.Requisites('ОбластьПоддержки').FieldName} <> {not_anonymous}"
    not_anonymous = tickets.AddWhere(not_anonymous)

    # текущая дата
    current_day = f'{datetime.today().day}.{datetime.today().month if len(str(datetime.today().month)) != 1 else "0" + str(datetime.today().month)}.{datetime.today().year}'
    # вставляю в таблицу завтрашний день с пустыми значениями. эксель без него почему то не отрисовывает последний день
    if (datetime.today() + timedelta(days=1)).isoweekday() == 6:
        tomorrow = datetime.today() + timedelta(days=3)
    else:
        tomorrow = datetime.today() + timedelta(days=1)
    tomorrow = f'{tomorrow.day}.{tomorrow.month if len(str(tomorrow.month)) != 1 else "0" + str(tomorrow.month)}.{tomorrow.year}'
    print(f"Почитал конфиг {config_path}, открыл ТехКАС")

    # данные для графиков
    statistics = pd.read_excel(data_file_path, sheet_name='tables')
    statistics.index = statistics['Обращения в работе (В работе). В разрезе месяца поступления']


    incoming = registred_yesterday(tickets)
    statistics.loc["Инциденты", current_day] = incoming['incidents']
    statistics.loc["Консультации", current_day] = incoming['consultation']
    statistics.loc["Запросы", current_day] = incoming['requests']
    statistics.loc["Проблемы", current_day] = incoming['problems']
    statistics.loc["Поступило всего", current_day] = incoming['incidents'] + incoming['requests'] +\
                                                     incoming['consultation'] + incoming['problems']
    statistics.loc["Поступило всего", tomorrow] = 0
    statistics.loc["Затрачено в часах", current_day] = get_time_for_request(tickets)

    print("Посчитал поток")

    # записываем обращения в работе по месяцам
    total_in_work = 0
    for mounth, count in tickets_in_work(tickets).items():
        statistics.loc[mounth, current_day] = count
        total_in_work += count
    statistics.loc["Всего в работе", current_day] = total_in_work

    # записываем данные для обращений старше 2 недель и 3 недель
    count_14_days = 0
    count_21_days = 0
    for day in [14, 21]:
        for employee, count in old_tickets(tickets, config, day).items():
            statistics.loc[(employee + f"_{day}"), current_day] = count
            if day == 14:
                count_14_days += count
            else:
                count_21_days += count

    statistics.loc["Старше 2 недель", current_day] = count_14_days
    statistics.loc["Старше 3 недель", current_day] = count_21_days
    old_tickets_count, old_tickets_list = not_closed(tickets)
    statistics.loc['Старше 4 недель', current_day] = old_tickets_count['older_four_weeks']
    statistics.loc['Хвост', current_day] = old_tickets_count['not_closed']
    print("Посчитал старье")


    #Записываем цветные зоны
    time_zone, ticket_time = time_zones(tickets, ['И'])
    statistics.loc['0-8', current_day] = time_zone['green']
    statistics.loc['8-16', current_day] = time_zone['sandy']
    statistics.loc['16-24', current_day] = time_zone['yellow']
    statistics.loc['>24', current_day] = time_zone['red']
    statistics.loc["Поступившие",current_day] = get_negative_grades(tickets, COMMAND)
    print("Посчитал цветные зоны")

    time_zone_cons, ticket_time_consultation = time_zones(tickets, ["К"])

    ticket_time.sort_values("Время в работе", ascending=False)
    ticket_time_consultation.sort_values("Время в работе", ascending=False)

    print("Посчитал все статы команды, пошел записывать в ексель")

    stats_to_excel.remove_excel_sheet("Графики", data_file_path)
    stats_to_excel.create_excel_sheet("Графики", data_file_path)

    stats_to_excel.remove_excel_sheet('tables', data_file_path)
    stats_to_excel.remove_excel_sheet('Инциденты', data_file_path)
    stats_to_excel.remove_excel_sheet('Консультации', data_file_path)

    with ExcelWriter(data_file_path, mode='a') as writer:
        statistics.to_excel(writer, sheet_name='tables')
        ticket_time.to_excel(writer, sheet_name='Инциденты', index=False)
        ticket_time_consultation.to_excel(writer, sheet_name='Консультации', index=False)
    stats_to_excel.style_text_sheet_list(data_file_path, 'Инциденты')
    stats_to_excel.style_text_sheet_list(data_file_path, 'Консультации')
    stats_to_excel.style_text_sheet_data_table(data_file_path, 'tables')

    stats_to_excel.remove_excel_sheet("Старые обращения", data_file_path)
    stats_to_excel.insert_old_tickets(data_file_path, old_tickets_list)

    stats_to_excel.write_chars_to_file(data_file_path)

    print(f'Файл {data_file_path} записан')
    print("Все готово, вы работаете великолепно")

if __name__ == "__main__":
    #calculate_stats()
    get_conf_and_start()
