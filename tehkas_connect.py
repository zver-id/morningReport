import datetime

import pythoncom
import win32com.client
"""
Необходимые пакеты для импорта в venv:
pypiwin32
pywin32 - в PyCharm подтягиевается автоматически

Справочник "Обращения в службу поддержки" - "ПДД"
"""
class Reference:
    def __new__(cls, reference_name):
        """
        Создание соединение с ТехКАС
        :param reference_name: название справочника в ТехКАС
        :return: справочник ТехКАС
        """
        pythoncom.CoInitialize()
        win32com.SetupEnvironment
        win32com.gen_py
        lp = win32com.client.DispatchEx("SBLogon.LoginPoint")
        ap = lp.GetApplication("systemcode=TEHKASNPO")

        # подключаем справочник "Обращения в службу поддержки"
        reference = ap.ReferencesFactory.ReferenceFactory(reference_name).GetComponent()
        return reference

def return_query(attribute_list, props_name, reference):
    """
    Cбор запроса для метода AddWhere из списка параметров
    :param attribute_list: список параметров
    :param props_name: имя реквизита, по перечню которого собираем запрос
    :param reference: справочник, в котором ведется поиск
    :return: запрос для применения в методе AddWhere
    """
    add_where = ""
    for attribute in attribute_list:
        if add_where:
            add_where = add_where + " or {}.{} = '{}'".format(
                reference.TableName, reference.Requisites(props_name).FieldName, attribute)
        else:
            add_where = add_where + "({}.{} = '{}'".format(
                reference.TableName, reference.Requisites(props_name).FieldName, attribute)
        if len(attribute_list) == attribute_list.index(attribute) + 1:
            add_where = add_where + ")"
    return add_where

def next_ticket(reference):
    """
    При переборе справочника отменяем изменения, закрываем обращение, выбираем следующее
    :param reference: Справочник, который перебирается
    """
    reference.Cancel()
    reference.CloseRecord()
    reference.Next()

def duration_in_work(reference):
    reference.OpenRecord()
    detail = reference.DetailDataSet(4)
    detail.First()
    time_in_work_duration = 0
    start = datetime.datetime.now()
    end = datetime.datetime.now()
    while not detail.EOF:
        res = 0
        if detail.Requisites("СостОбращенияТ4").AsString == "В работе":
            start = detail.Requisites('ДатаВремяT4').AsDate
        elif detail.Requisites("СостОбращенияТ4").AsString == "На контроле":
            end = detail.Requisites('ДатаВремяT4').AsDate
            time_start = time.fromisoformat('09:00')
            time_end = time.fromisoformat('17:00')
            res = 0
            t = start
            while (t := t + datetime.timedelta(minutes=1)) <= end:
                if t.weekday() < 5 and time_start <= t.time() <= time_end:
                    res += 1

        time_in_work_duration += res
        # print(detail.Requisites("РаботникТ4").AsString)
        # print(detail.Requisites("СостОбращенияТ4").AsString)
        # print(detail.Requisites('ДатаВремяT4').AsDate)
        detail.Next()
    reference.Cancel()
    reference.CloseRecord()
    return round(time_in_work_duration/60, 2)
